/**
 * Integra a planilha com uma agenda do Google Calendar.
 *
 * Intervalos nomeados esperados:
 *   ID      - uma unica celula com o ID da agenda do Google Calendar.
 *   TAREFAS - colunas: titulo, recorrencia, data inicial, detalhe, descricao/link.
 *   MÊS     - uma unica celula com o numero do mes (1 a 12).
 */

const AGENDA_CALENDAR = Object.freeze({
  MENU: 'Agenda',
  ID_RANGE: 'ID',
  TASKS_RANGE: 'TAREFAS',
  MONTH_RANGE: 'MÊS',
  YEARS_TO_CREATE: 2,
  UPDATE_STATE_VERSION: 6,
  UPDATE_HANDLER: 'continuarAtualizacaoAgenda',
  UPDATE_STATE_PROPERTY: 'AGENDA_CALENDAR_UPDATE_STATE',
  DELETE_BATCH_SIZE: 10,
  CREATE_BATCH_SIZE: 10,
  CALENDAR_LIST_PAGE_SIZE: 250,
  MAX_BATCH_MILLISECONDS: 60000,
  CALENDAR_RETRY_ATTEMPTS: 3,
  MAX_NO_PROGRESS_MILLISECONDS: 15 * 60 * 1000,
  SAFETY_TRIGGER_DELAY_MILLISECONDS: 7 * 60 * 1000,
  EMPTY_CALENDAR_CONFIRMATIONS: 3,
  PDF_FOLDER_NAME: 'Calendários de tarefas',
  WEEKDAYS: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'],
  MONTHS: [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ]
});

/** Cria os comandos no topo da planilha sempre que ela for aberta. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(AGENDA_CALENDAR.MENU)
    .addItem('Atualizar agenda', 'atualizarAgenda')
    .addItem('Gerar PDF', 'gerarPDF')
    .addSeparator()
    .addItem('Diagnosticar atualização', 'diagnosticarAtualizacaoAgenda')
    .addItem('Cancelar atualização travada', 'cancelarAtualizacaoTravada')
    .addToUi();
}

/** Inicia a atualizacao em blocos e agenda as continuacoes automaticamente. */
function atualizarAgenda() {
  const lock = LockService.getScriptLock();
  let runStarted = false;
  let spreadsheet;

  try {
    if (!lock.tryLock(5000)) throw new Error('Já existe uma atualização da agenda em execução. Aguarde.');
    const previousState = loadUpdateState_();
    if (previousState && previousState.version !== AGENDA_CALENDAR.UPDATE_STATE_VERSION) {
      console.log(
        `[AgendaCalendar] runId=${previousState.runId || 'desconhecido'} estado da versão ` +
        `${previousState.version || 'desconhecida'} invalidado pela versão ${AGENDA_CALENDAR.UPDATE_STATE_VERSION}`
      );
      invalidatePreviousUpdate_(previousState, 'stale', 'Estado incompatível com a versão atual.');
    } else if (isUpdateActive_(previousState)) {
      const hasTrigger = hasUpdateContinuationTrigger_(previousState);
      const stale = isUpdateStale_(previousState);
      if (hasTrigger && !stale) {
        throw new Error('Já existe uma atualização em andamento. Aguarde a mensagem de conclusão.');
      }
      console.warn(
        `[AgendaCalendar] runId=${previousState.runId} atualização órfã detectada ` +
        `phase=${previousState.phase} hasTrigger=${hasTrigger} stale=${stale} ` +
        `definitionIndex=${previousState.definitionIndex || 0} created=${previousState.created || 0}`
      );
      invalidatePreviousUpdate_(previousState, 'stale', 'Atualização órfã ou sem progresso detectada.');
    }
    spreadsheet = SpreadsheetApp.getActive();
    const calendarId = namedCellValue_(spreadsheet, AGENDA_CALENDAR.ID_RANGE);
    const definitions = readTaskDefinitions_(spreadsheet);

    if (!calendarId) throw new Error('O intervalo nomeado "ID" está vazio.');
    if (!definitions.length) throw new Error('Nenhuma tarefa válida foi encontrada em "TAREFAS".');

    cancelUpdateTriggers_();
    const runId = Utilities.getUuid();
    const startedAt = new Date().toISOString();
    saveUpdateState_({
      spreadsheetId: spreadsheet.getId(),
      version: AGENDA_CALENDAR.UPDATE_STATE_VERSION,
      runId,
      calendarId: String(calendarId).trim(),
      phase: 'deleting',
      definitionIndex: 0,
      deletionEmptyChecks: 0,
      deleted: 0,
      created: 0,
      total: definitions.length,
      definitionsSignature: definitionsSignature_(definitions),
      startedAt,
      lastHeartbeatAt: startedAt,
      lastProgressAt: startedAt
    });
    runStarted = true;
    console.log(`[AgendaCalendar] runId=${runId} fase=deleting atualização iniciada`);

    const result = processUpdateBatch_();
    const currentState = loadUpdateState_();
    spreadsheet.toast(updateStatusMessage_(currentState || result), 'Agenda', 10);
  } catch (error) {
    if (runStarted) {
      cancelUpdateTriggers_();
      const state = loadUpdateState_();
      if (state) {
        state.phase = 'error';
        state.error = error.message || String(error);
        state.continuationTriggerId = null;
        markProgress_(state);
        saveUpdateState_(state);
        console.error(`[AgendaCalendar] runId=${state.runId} fase=error erro=${state.error}`);
      }
    } else {
      console.error(`[AgendaCalendar] falha antes de iniciar a execução: ${error.message || String(error)}`);
    }
    if (spreadsheet) spreadsheet.toast(error.message || String(error), 'Falha na atualização', 15);
    throw error;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/** Continua uma atualizacao iniciada por atualizarAgenda (acionada por gatilho). */
function continuarAtualizacaoAgenda(event) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(65000)) {
    const state = loadUpdateState_();
    const hasSafetyTrigger = hasAnyUpdateContinuationTrigger_();
    console.log(
      `[AgendaCalendar] callback não obteve lock; outro lote permanece responsável pela continuação ` +
      `runId=${state && state.runId ? state.runId : 'desconhecido'} ` +
      `hasSafetyTrigger=${hasSafetyTrigger} ` +
      `triggerUid=${event && event.triggerUid !== undefined ? event.triggerUid : 'ausente'} ` +
      `triggerUidType=${event && event.triggerUid !== undefined ? typeof event.triggerUid : 'undefined'}`
    );
    if (state && isUpdateActive_(state) && !hasUpdateContinuationTrigger_(state)) {
      scheduleEmergencyContinuation_();
    }
    return;
  }

  try {
    const state = loadUpdateState_();
    if (!state || state.version !== AGENDA_CALENDAR.UPDATE_STATE_VERSION ||
        !['deleting', 'creating'].includes(state.phase)) {
      console.log('[AgendaCalendar] callback sem atualização ativa; encerrando normalmente');
      cancelUpdateTriggers_();
      return;
    }
    console.log(
      `[AgendaCalendar] runId=${state.runId} phase=${state.phase} ` +
      `definitionIndex=${state.definitionIndex || 0} created=${state.created || 0} ` +
      `deleted=${state.deleted || 0} ` +
      `triggerUid=${event && event.triggerUid !== undefined ? event.triggerUid : 'ausente'} ` +
      `triggerUidType=${event && event.triggerUid !== undefined ? typeof event.triggerUid : 'undefined'}`
    );
    state.lastHeartbeatAt = new Date().toISOString();
    saveUpdateState_(state);
    if (isUpdateStale_(state)) {
      throw new Error('Atualização interrompida por falta de progresso real. Use Atualizar agenda novamente.');
    }
    processUpdateBatch_();
  } catch (error) {
    cancelUpdateTriggers_();
    const state = loadUpdateState_();
    if (state) {
      state.error = error.message || String(error);
      state.phase = 'error';
      state.continuationTriggerId = null;
      markProgress_(state);
      saveUpdateState_(state);
      console.error(`[AgendaCalendar] runId=${state.runId} fase=error erro=${state.error}`);
      try {
        SpreadsheetApp.openById(state.spreadsheetId).toast(
          `A atualização foi interrompida: ${state.error}`,
          'Agenda',
          15
        );
      } catch (toastError) {
        console.error(toastError);
      }
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/** Cancela somente o fluxo de sincronizacao; nao altera nenhum evento da agenda. */
function cancelarAtualizacaoTravada() {
  const spreadsheet = SpreadsheetApp.getActive();
  const lock = LockService.getScriptLock();
  let error;

  try {
    if (!lock.tryLock(5000)) {
      throw new Error('Existe um lote em execução neste momento. Aguarde alguns segundos e tente novamente.');
    }
    cancelUpdateTriggers_();
    const state = loadUpdateState_();
    if (state) {
      state.phase = 'cancelled';
      state.error = 'Atualização cancelada manualmente pelo usuário.';
      state.cancelledAt = new Date().toISOString();
      state.continuationTriggerId = null;
      markProgress_(state);
      saveUpdateState_(state);
      console.warn(`[AgendaCalendar] runId=${state.runId || 'desconhecido'} fase=cancelled`);
    }
  } catch (caughtError) {
    error = caughtError;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }

  if (error) {
    spreadsheet.toast(error.message || String(error), 'Não foi possível cancelar', 10);
    throw error;
  }
  spreadsheet.toast(
    'Somente o processo de sincronização foi cancelado; nenhum evento foi alterado.',
    'Atualização cancelada',
    10
  );
}

/** Exibe e registra o estado sem inserir ou excluir eventos. */
function diagnosticarAtualizacaoAgenda() {
  const spreadsheet = SpreadsheetApp.getActive();
  const state = loadUpdateState_();
  const triggers = ScriptApp.getProjectTriggers().filter((trigger) =>
    trigger.getHandlerFunction() === AGENDA_CALENDAR.UPDATE_HANDLER);
  const matchingTrigger = state ? hasUpdateContinuationTrigger_(state) : false;
  const calendarId = namedCellValue_(spreadsheet, AGENDA_CALENDAR.ID_RANGE);
  const calendar = inspectCalendar_(calendarId);
  const diagnostic = {
    updateStateVersion: AGENDA_CALENDAR.UPDATE_STATE_VERSION,
    runId: state && state.runId,
    phase: state && state.phase,
    startedAt: state && state.startedAt,
    updatedAt: state && state.updatedAt,
    lastHeartbeatAt: state && state.lastHeartbeatAt,
    lastProgressAt: state && state.lastProgressAt,
    definitionIndex: state && state.definitionIndex,
    created: state && state.created,
    deleted: state && state.deleted,
    deletionEmptyChecks: state && state.deletionEmptyChecks,
    total: state && state.total,
    continuationTriggerId: state && state.continuationTriggerId,
    matchingTrigger,
    residualTriggerCount: triggers.length - (matchingTrigger ? 1 : 0),
    calendar
  };
  console.log(`[AgendaCalendar] DIAGNÓSTICO ${JSON.stringify(diagnostic)}`);
  calendar.sample.forEach((event) =>
    console.log(`[AgendaCalendar] AMOSTRA ${JSON.stringify(event)}`));
  spreadsheet.toast(
    `Fase: ${diagnostic.phase || 'sem estado'} | ativos: ${calendar.live} | ` +
    `cancelados: ${calendar.cancelled} | trigger válido: ${matchingTrigger ? 'sim' : 'não'}`,
    'Diagnóstico registrado nos logs',
    15
  );
  return diagnostic;
}

function inspectCalendar_(calendarId) {
  const result = { live: 0, cancelled: 0, recurringInstances: 0, recurringMasters: 0, complete: true, sample: [] };
  const started = Date.now();
  let pageToken;

  do {
    if (timeLimitApproaching_(started)) {
      result.complete = false;
      break;
    }
    const options = {
      maxResults: AGENDA_CALENDAR.CALENDAR_LIST_PAGE_SIZE,
      showDeleted: false,
      singleEvents: false
    };
    if (pageToken) options.pageToken = pageToken;
    const response = withCalendarRetry_(
      () => Calendar.Events.list(calendarId, options),
      started,
      'diagnosticar agenda'
    );
    (response.items || []).forEach((event) => {
      if (event.status === 'cancelled') result.cancelled += 1;
      else result.live += 1;
      if (event.recurringEventId) result.recurringInstances += 1;
      if (event.recurrence && event.recurrence.length) result.recurringMasters += 1;
      if (result.sample.length < 10) {
        result.sample.push({
          id: event.id,
          status: event.status,
          summary: event.summary,
          recurringEventId: event.recurringEventId || null,
          hasRecurrence: Boolean(event.recurrence && event.recurrence.length),
          eventType: event.eventType || 'default'
        });
      }
    });
    pageToken = response.nextPageToken;
  } while (pageToken);
  return result;
}

function updateStatusMessage_(state) {
  if (!state) return 'Atualização iniciada.';
  if (state.phase === 'deleting') {
    return `Fase: removendo eventos antigos. Eventos removidos: ${state.deleted || 0}. ` +
      'O processamento continuará automaticamente.';
  }
  if (state.phase === 'creating') {
    return `Fase: criando novos eventos. ${state.created || 0} de ${state.total || 0} criados.`;
  }
  if (state.phase === 'complete') return `Sincronização concluída. ${state.created || 0} de ${state.total || 0} criados.`;
  return `Fase: ${state.phase || 'desconhecida'}.`;
}

function processUpdateBatch_() {
  const state = loadUpdateState_();
  if (!state) throw new Error('Não existe uma atualização de agenda em andamento.');

  state.lastHeartbeatAt = new Date().toISOString();
  saveUpdateState_(state);
  if (isUpdateStale_(state)) {
    throw new Error('Atualização sem progresso real por tempo excessivo.');
  }

  // O proximo gatilho e criado antes das chamadas ao Calendar. Assim, mesmo que
  // esta execucao seja encerrada abruptamente pelo limite do Apps Script, havera
  // uma nova tentativa para retomar o estado salvo.
  scheduleUpdateContinuation_(AGENDA_CALENDAR.SAFETY_TRIGGER_DELAY_MILLISECONDS, state);

  const started = Date.now();
  let operations = 0;

  if (state.phase === 'deleting') {
    const deletion = deleteCalendarEventsBatch_(state.calendarId, started);
    operations = deletion.actuallyDeleted;
    state.deleted = (state.deleted || 0) + deletion.actuallyDeleted;
    if (deletion.actuallyDeleted > 0) markProgress_(state);
    if (deletion.complete && !deletion.hasLiveEvents) {
      const previousChecks = state.deletionEmptyChecks || 0;
      state.deletionEmptyChecks = Math.min(
        previousChecks + 1,
        AGENDA_CALENDAR.EMPTY_CALENDAR_CONFIRMATIONS
      );
      if (state.deletionEmptyChecks > previousChecks) markProgress_(state);
    } else if (deletion.hasLiveEvents) {
      state.deletionEmptyChecks = 0;
    }
    console.log(
      `[AgendaCalendar] runId=${state.runId} phase=deleting pageItems=${deletion.pageItems} ` +
      `cancelledIgnored=${deletion.cancelledIgnored} liveFound=${deletion.liveFound} ` +
      `alreadyGoneIgnored=${deletion.alreadyGoneIgnored} actuallyDeleted=${deletion.actuallyDeleted} ` +
      `deletedTotal=${state.deleted} ` +
      `deletionEmptyChecks=${state.deletionEmptyChecks} lastProgressAt=${state.lastProgressAt}`
    );

    if (state.deletionEmptyChecks >= AGENDA_CALENDAR.EMPTY_CALENDAR_CONFIRMATIONS) {
      // Uma ultima leitura independente impede iniciar a criacao se a API voltar
      // a apresentar um evento durante a janela de consistencia eventual.
      const verification = calendarHasAnyLiveEvent_(state.calendarId, started);
      if (verification.hasLiveEvent) {
        state.deletionEmptyChecks = 0;
      } else if (verification.complete) {
        state.phase = 'creating';
        operations = 0;
        state.creatingStartedAt = new Date().toISOString();
        markProgress_(state);
        console.log(
          `[AgendaCalendar] runId=${state.runId} AGENDA CONFIRMADA SEM EVENTOS ATIVOS; ` +
          'phase deleting -> creating'
        );
      }
    }

    saveUpdateState_(state);
    if (state.phase === 'deleting') {
      scheduleUpdateContinuation_(null, state);
      return updateResult_(state, false);
    }
  }

  const spreadsheet = SpreadsheetApp.openById(state.spreadsheetId);
  const definitions = readTaskDefinitions_(spreadsheet);
  if (definitionsSignature_(definitions) !== state.definitionsSignature) {
    throw new Error('O intervalo "TAREFAS" mudou durante a atualização. Execute "Atualizar agenda" novamente.');
  }

  while (state.definitionIndex < definitions.length && !batchLimitReached_(started, operations)) {
    const definition = definitions[state.definitionIndex];
    if (state.definitionIndex === 0 && state.created === 0) {
      console.log(
        `[AgendaCalendar] runId=${state.runId} INICIANDO PRIMEIRO EVENTO ` +
        `definitionIndex=0 sourceRow=${definition.sourceRow}`
      );
    }
    const eventConfirmed = insertRecurringCalendarEvent_(
      state.calendarId,
      definition,
      deterministicEventId_(state.spreadsheetId, state.runId, definition),
      state.runId,
      started
    );
    if (!eventConfirmed) throw new Error(`Não foi possível confirmar a criação da linha ${definition.sourceRow}.`);
    state.definitionIndex += 1;
    state.created += 1;
    operations += 1;
    markProgress_(state);
    // Persistir apos cada evento torna a retomada segura mesmo em timeout.
    saveUpdateState_(state);
    console.log(
      `[AgendaCalendar] runId=${state.runId} EVENTO CONFIRMADO ` +
      `definitionIndex=${state.definitionIndex} created=${state.created}/${state.total}`
    );
  }

  console.log(
    `[AgendaCalendar] runId=${state.runId} fase=creating definitionIndex=${state.definitionIndex} ` +
    `criados=${state.created}/${state.total}`
  );

  if (state.definitionIndex >= definitions.length) {
    finishUpdate_(state, spreadsheet);
    return updateResult_(state, true);
  }

  saveUpdateState_(state);
  scheduleUpdateContinuation_(null, state);
  return updateResult_(state, false);
}

/** Gera o calendario do mes informado em MÊS, sempre usando o ano atual. */
function gerarPDF() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const month = Number(namedCellValue_(spreadsheet, AGENDA_CALENDAR.MONTH_RANGE));
    if (!Number.isInteger(month) || month < 1 || month > 12) {
      throw new Error('O intervalo nomeado "MÊS" deve conter um número de 1 a 12.');
    }

    const year = Number(Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy'));
    const definitions = readTaskDefinitions_(spreadsheet);
    const entriesByDay = calendarEntries_(definitions, month, year);
    const pdfFile = createCalendarPdf_(entriesByDay, month, year);

    const safeUrl = escapeHtml_(pdfFile.getDownloadUrl());
    const html = HtmlService.createHtmlOutput(
      '<div style="font:14px Arial,sans-serif;padding:18px">' +
      '<h3 style="margin-top:0">PDF gerado</h3>' +
      '<p>O calendário foi salvo no Google Drive.</p>' +
      `<p><a href="${safeUrl}" target="_blank" rel="noopener">Baixar PDF</a></p>` +
      '<button onclick="google.script.host.close()">Fechar</button></div>'
    ).setWidth(360).setHeight(210);
    ui.showModalDialog(html, 'Gerar PDF');
  } catch (error) {
    ui.alert('Não foi possível gerar o PDF', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

function readTaskDefinitions_(spreadsheet) {
  const range = spreadsheet.getRangeByName(AGENDA_CALENDAR.TASKS_RANGE);
  if (!range) throw new Error('Crie o intervalo nomeado "TAREFAS".');
  if (range.getNumColumns() < 5) throw new Error('"TAREFAS" precisa ter pelo menos cinco colunas.');

  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const richTexts = range.getRichTextValues();
  const definitions = [];

  values.forEach((row, index) => {
    const sheetRow = range.getRow() + index;
    const recurrenceText = String(displayValues[index][1] || '').trim();
    const title = String(displayValues[index][0] || '').trim();
    const startDate = coerceDate_(row[2]);
    if (!title && !recurrenceText && !row[2]) return;
    if (!title) {
      console.warn(`[AgendaCalendar] linha ${sheetRow} ignorada: título vazio`);
      return;
    }
    if (!recurrenceText) {
      console.warn(`[AgendaCalendar] linha ${sheetRow} ignorada: recorrência vazia`);
      return;
    }
    if (!startDate) {
      console.warn(`[AgendaCalendar] linha ${sheetRow} ignorada: data inicial inválida`);
      return;
    }

    let recurrence;
    try {
      recurrence = normalizeRecurrence_(recurrenceText, sheetRow);
    } catch (error) {
      console.warn(`[AgendaCalendar] linha ${sheetRow} ignorada: ${error.message || String(error)}`);
      return;
    }

    definitions.push({
      sourceRow: sheetRow,
      title,
      recurrence,
      startDate,
      detail: String(displayValues[index][3] || ''),
      description: richTextAsPlainTextWithUrls_(richTexts[index][4], displayValues[index][4])
    });
  });

  return definitions;
}

function normalizeRecurrence_(value, sheetRow) {
  const normalized = String(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const aliases = {
    semanal: { unit: 'days', amount: 7 },
    quinzenal: { unit: 'days', amount: 14 },
    mensal: { unit: 'months', amount: 1 },
    bimestral: { unit: 'months', amount: 2 },
    trimestral: { unit: 'months', amount: 3 },
    quadrimestral: { unit: 'months', amount: 4 },
    semestral: { unit: 'months', amount: 6 },
    anual: { unit: 'months', amount: 12 }
  };
  if (!aliases[normalized]) throw new Error(`Recorrência inválida na linha ${sheetRow}: ${value}.`);
  return aliases[normalized];
}

function occurrencesFor_(startDate, recurrence) {
  const endExclusive = addMonthsClamped_(startDate, AGENDA_CALENDAR.YEARS_TO_CREATE * 12, startDate.getDate());
  const dates = [];
  let occurrence = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  let index = 0;

  while (occurrence < endExclusive) {
    dates.push(occurrence);
    index += 1;
    occurrence = recurrence.unit === 'days'
      ? addDays_(startDate, recurrence.amount * index)
      : addMonthsClamped_(startDate, recurrence.amount * index, startDate.getDate());
  }
  return dates;
}

function calendarEntries_(definitions, month, year) {
  const entries = {};
  definitions.forEach((definition) => {
    occurrencesFor_(definition.startDate, definition.recurrence).forEach((date) => {
      if (date.getFullYear() !== year || date.getMonth() !== month - 1) return;
      const day = date.getDate();
      if (!entries[day]) entries[day] = [];
      entries[day].push(buildEventTitle_(definition.title, definition.detail));
    });
  });
  return entries;
}

function createCalendarPdf_(entriesByDay, month, year) {
  const title = `Calendário - ${AGENDA_CALENDAR.MONTHS[month - 1]} de ${year}`;
  const document = DocumentApp.create(title);
  const body = document.getBody();
  body.clear();
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const tableData = [AGENDA_CALENDAR.WEEKDAYS.slice()];
  const firstWeekday = new Date(year, month - 1, 1).getDay();
  const daysInMonth = new Date(year, month, 0).getDate();
  let day = 1;
  for (let week = 0; week < 6 && day <= daysInMonth; week += 1) {
    const row = [];
    for (let weekday = 0; weekday < 7; weekday += 1) {
      if ((week === 0 && weekday < firstWeekday) || day > daysInMonth) {
        row.push('');
      } else {
        const tasks = entriesByDay[day] || [];
        row.push([String(day)].concat(tasks.map((task) => `• ${task}`)).join('\n'));
        day += 1;
      }
    }
    tableData.push(row);
  }

  const table = body.appendTable(tableData);
  for (let column = 0; column < 7; column += 1) table.getCell(0, column).setBackgroundColor('#d9eaf7');
  document.saveAndClose();

  const source = DriveApp.getFileById(document.getId());
  const pdf = source.getAs(MimeType.PDF).setName(`${title}.pdf`);
  const folder = getOrCreateFolder_(AGENDA_CALENDAR.PDF_FOLDER_NAME);
  const file = folder.createFile(pdf);
  source.setTrashed(true);
  return file;
}

function deleteCalendarEventsBatch_(calendarId, started) {
  let actuallyDeleted = 0;
  let pageItems = 0;
  let cancelledIgnored = 0;
  let alreadyGoneIgnored = 0;
  let liveFound = 0;
  let pageToken;

  while (!timeLimitApproaching_(started)) {
    const options = {
      maxResults: AGENDA_CALENDAR.CALENDAR_LIST_PAGE_SIZE,
      showDeleted: false,
      singleEvents: false
    };
    if (pageToken) options.pageToken = pageToken;

    const response = withCalendarRetry_(
      () => Calendar.Events.list(calendarId, options),
      started,
      'listar eventos para exclusão'
    );
    const items = response.items || [];
    pageItems += items.length;

    for (let index = 0; index < items.length; index += 1) {
      const event = items[index];
      if (event.status === 'cancelled') {
        cancelledIgnored += 1;
        continue;
      }
      liveFound += 1;
      if (actuallyDeleted >= AGENDA_CALENDAR.DELETE_BATCH_SIZE || timeLimitApproaching_(started)) {
        return {
          actuallyDeleted, pageItems, cancelledIgnored, alreadyGoneIgnored,
          liveFound, hasLiveEvents: true, complete: false
        };
      }
      const removal = removeCalendarEventSafely_(calendarId, event.id, started);
      if (removal.removed) actuallyDeleted += 1;
      if (removal.alreadyGone) alreadyGoneIgnored += 1;
    }

    // Depois de uma mutacao, nao reutilizar o pageToken da colecao anterior.
    if (actuallyDeleted > 0) {
      return {
        actuallyDeleted, pageItems, cancelledIgnored, alreadyGoneIgnored,
        liveFound, hasLiveEvents: true, complete: false
      };
    }
    pageToken = response.nextPageToken;
    if (!pageToken) {
      return {
        actuallyDeleted: 0,
        pageItems,
        cancelledIgnored,
        alreadyGoneIgnored,
        liveFound,
        hasLiveEvents: liveFound > alreadyGoneIgnored,
        complete: true
      };
    }
  }
  return {
    actuallyDeleted, pageItems, cancelledIgnored, alreadyGoneIgnored,
    liveFound, hasLiveEvents: liveFound > alreadyGoneIgnored, complete: false
  };
}

function calendarHasAnyLiveEvent_(calendarId, started) {
  let pageToken;
  let pageItems = 0;
  let cancelledIgnored = 0;
  do {
    if (timeLimitApproaching_(started)) {
      return { hasLiveEvent: false, complete: false, pageItems, cancelledIgnored };
    }
    const options = {
      maxResults: AGENDA_CALENDAR.CALENDAR_LIST_PAGE_SIZE,
      showDeleted: false,
      singleEvents: false
    };
    if (pageToken) options.pageToken = pageToken;
    const response = withCalendarRetry_(
      () => Calendar.Events.list(calendarId, options),
      started,
      'verificar eventos ativos'
    );
    const items = response.items || [];
    pageItems += items.length;
    for (let index = 0; index < items.length; index += 1) {
      if (items[index].status === 'cancelled') cancelledIgnored += 1;
      else return { hasLiveEvent: true, complete: true, pageItems, cancelledIgnored };
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return { hasLiveEvent: false, complete: true, pageItems, cancelledIgnored };
}

function removeCalendarEventSafely_(calendarId, eventId, started) {
  try {
    withCalendarRetry_(
      () => Calendar.Events.remove(calendarId, eventId),
      started,
      `remover evento ${eventId}`
    );
    return { removed: true, alreadyGone: false };
  } catch (error) {
    // A listagem do Calendar pode manter por alguns segundos um evento que ja
    // foi removido. 404/410 significam que o resultado desejado ja foi obtido.
    if (!calendarErrorHasCode_(error, [404, 410])) throw error;
    return { removed: false, alreadyGone: true };
  }
}

function insertRecurringCalendarEvent_(calendarId, definition, eventId, runId, started) {
  const definitionHash = eventDefinitionHash_(definition);
  try {
    withCalendarRetry_(() => Calendar.Events.insert({
      id: eventId,
      summary: buildEventTitle_(definition.title, definition.detail),
      description: definition.description,
      start: { date: calendarDate_(definition.startDate) },
      end: { date: calendarDate_(addDays_(definition.startDate, 1)) },
      recurrence: recurrenceDates_(definition.startDate, definition.recurrence),
      extendedProperties: {
        private: {
          managedBy: 'AgendaCalendar',
          runId,
          sourceRow: String(definition.sourceRow),
          definitionHash
        }
      }
    }, calendarId), started, `criar evento da linha ${definition.sourceRow}`);
    return true;
  } catch (error) {
    if (!calendarErrorHasCode_(error, [409])) throw error;
    const existing = withCalendarRetry_(
      () => Calendar.Events.get(calendarId, eventId),
      started,
      `validar conflito da linha ${definition.sourceRow}`
    );
    if (!eventMatchesCurrentRun_(existing, definition, runId, definitionHash)) {
      throw new Error(`Conflito 409 não idempotente ao criar a linha ${definition.sourceRow}.`);
    }
    console.log(
      `[AgendaCalendar] runId=${runId} 409 validado como repetição legítima ` +
      `sourceRow=${definition.sourceRow}`
    );
    return true;
  }
}

function recurrenceDates_(startDate, recurrence) {
  const additionalDates = occurrencesFor_(startDate, recurrence)
    .slice(1)
    .map((date) => calendarDate_(date).replace(/-/g, ''));
  return additionalDates.length ? [`RDATE;VALUE=DATE:${additionalDates.join(',')}`] : [];
}

function deterministicEventId_(spreadsheetId, runId, definition) {
  const input = [spreadsheetId, runId, definition.sourceRow].join('|');
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8
  );
  const hex = bytes.map((byte) => (`0${((byte + 256) % 256).toString(16)}`).slice(-2)).join('');
  return `a${hex}`;
}

function eventDefinitionHash_(definition) {
  const normalized = {
    sourceRow: definition.sourceRow,
    summary: buildEventTitle_(definition.title, definition.detail),
    description: definition.description,
    start: calendarDate_(definition.startDate),
    end: calendarDate_(addDays_(definition.startDate, 1)),
    recurrence: recurrenceDates_(definition.startDate, definition.recurrence)
  };
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(normalized),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes);
}

function eventMatchesCurrentRun_(event, definition, runId, definitionHash) {
  if (!event || event.status === 'cancelled') return false;
  const privateProperties = event.extendedProperties && event.extendedProperties.private;
  if (!privateProperties || privateProperties.managedBy !== 'AgendaCalendar' ||
      privateProperties.runId !== runId ||
      privateProperties.sourceRow !== String(definition.sourceRow) ||
      privateProperties.definitionHash !== definitionHash) return false;

  return event.summary === buildEventTitle_(definition.title, definition.detail) &&
    String(event.description || '') === String(definition.description || '') &&
    event.start && event.start.date === calendarDate_(definition.startDate) &&
    event.end && event.end.date === calendarDate_(addDays_(definition.startDate, 1)) &&
    JSON.stringify(event.recurrence || []) ===
      JSON.stringify(recurrenceDates_(definition.startDate, definition.recurrence));
}

function definitionsSignature_(definitions) {
  const data = definitions.map((definition) => ({
    sourceRow: definition.sourceRow,
    title: definition.title,
    recurrence: definition.recurrence,
    startDate: calendarDate_(definition.startDate),
    detail: definition.detail,
    description: definition.description
  }));
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(data),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(bytes);
}

function calendarErrorHasCode_(error, codes) {
  const text = [error && error.message, error && error.details, String(error)].join(' ');
  return codes.some((code) => new RegExp(`(^|\\D)${code}(\\D|$)`).test(text));
}

function withCalendarRetry_(operation, started, label) {
  let lastError;
  for (let attempt = 1; attempt <= AGENDA_CALENDAR.CALENDAR_RETRY_ATTEMPTS; attempt += 1) {
    try {
      return operation();
    } catch (error) {
      lastError = error;
      if (!calendarErrorHasCode_(error, [429, 500, 502, 503, 504]) ||
          attempt >= AGENDA_CALENDAR.CALENDAR_RETRY_ATTEMPTS) throw error;
      const delay = 500 * Math.pow(2, attempt - 1);
      if (Date.now() - started + delay >= AGENDA_CALENDAR.MAX_BATCH_MILLISECONDS) throw error;
      console.warn(`[AgendaCalendar] erro transitório em ${label}; tentativa=${attempt} esperaMs=${delay}`);
      Utilities.sleep(delay);
    }
  }
  throw lastError;
}

function timeLimitApproaching_(started) {
  return Date.now() - started >= AGENDA_CALENDAR.MAX_BATCH_MILLISECONDS - 5000;
}

function batchLimitReached_(started, operations) {
  return operations >= AGENDA_CALENDAR.CREATE_BATCH_SIZE || timeLimitApproaching_(started);
}

function updateResult_(state, complete) {
  return { complete, created: state.created, total: state.total };
}

function finishUpdate_(state, spreadsheet) {
  state.phase = 'complete';
  state.completedAt = new Date().toISOString();
  cancelUpdateTriggers_();
  state.continuationTriggerId = null;
  markProgress_(state);
  saveUpdateState_(state);
  console.log(
    `[AgendaCalendar] runId=${state.runId} fase=complete totalExcluído=${state.deleted || 0} ` +
    `criados=${state.created}`
  );
  spreadsheet.toast(`Atualização concluída: ${state.created} eventos criados.`, 'Agenda', 10);
}

function loadUpdateState_() {
  const value = PropertiesService.getScriptProperties().getProperty(AGENDA_CALENDAR.UPDATE_STATE_PROPERTY);
  return value ? JSON.parse(value) : null;
}

function isUpdateActive_(state) {
  if (!state || state.version !== AGENDA_CALENDAR.UPDATE_STATE_VERSION ||
      !['deleting', 'creating'].includes(state.phase)) return false;
  return true;
}

function isUpdateStale_(state) {
  if (!state) return false;
  const lastProgress = new Date(state.lastProgressAt || state.startedAt || 0).getTime();
  const stale = !Number.isFinite(lastProgress) ||
    Date.now() - lastProgress > AGENDA_CALENDAR.MAX_NO_PROGRESS_MILLISECONDS;
  const creatingWithoutProgress = state.phase === 'creating' &&
    Number(state.definitionIndex || 0) === 0 && Number(state.created || 0) === 0 && stale;
  if (creatingWithoutProgress) {
    console.warn(`[AgendaCalendar] runId=${state.runId} creating sem progresso detectado`);
  }
  return stale;
}

function hasUpdateContinuationTrigger_(state) {
  if (!state || !state.continuationTriggerId) return false;
  return ScriptApp.getProjectTriggers().some((trigger) =>
    trigger.getHandlerFunction() === AGENDA_CALENDAR.UPDATE_HANDLER &&
    String(trigger.getUniqueId()) === String(state.continuationTriggerId));
}

function hasAnyUpdateContinuationTrigger_() {
  return ScriptApp.getProjectTriggers().some((trigger) =>
    trigger.getHandlerFunction() === AGENDA_CALENDAR.UPDATE_HANDLER);
}

function invalidatePreviousUpdate_(state, phase, message) {
  cancelUpdateTriggers_();
  state.phase = phase;
  state.error = message;
  state.continuationTriggerId = null;
  state.invalidatedAt = new Date().toISOString();
  markProgress_(state);
  saveUpdateState_(state);
}

function saveUpdateState_(state) {
  state.updatedAt = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(
    AGENDA_CALENDAR.UPDATE_STATE_PROPERTY,
    JSON.stringify(state)
  );
}

function markProgress_(state) {
  state.lastProgressAt = new Date().toISOString();
}

function scheduleUpdateContinuation_(delayMilliseconds, state) {
  cancelUpdateTriggers_();
  const trigger = ScriptApp.newTrigger(AGENDA_CALENDAR.UPDATE_HANDLER)
    .timeBased()
    .after(delayMilliseconds || 60 * 1000)
    .create();
  state.continuationTriggerId = trigger.getUniqueId();
  saveUpdateState_(state);
  console.log(
    `[AgendaCalendar] runId=${state.runId} CONTINUAÇÃO AGENDADA ` +
    `triggerId=${state.continuationTriggerId} delay=${delayMilliseconds || 60 * 1000}`
  );
}

function scheduleEmergencyContinuation_() {
  const trigger = ScriptApp.newTrigger(AGENDA_CALENDAR.UPDATE_HANDLER)
    .timeBased()
    .after(2 * 60 * 1000)
    .create();
  console.warn(`[AgendaCalendar] continuação emergencial agendada triggerId=${trigger.getUniqueId()}`);
}

function cancelUpdateTriggers_() {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === AGENDA_CALENDAR.UPDATE_HANDLER) ScriptApp.deleteTrigger(trigger);
  });
}

function buildEventTitle_(title, detail) {
  const suffix = String(detail || '').trim();
  return suffix ? `${title} - ${suffix}` : title;
}

function richTextAsPlainTextWithUrls_(richText, fallback) {
  if (!richText) return String(fallback || '');
  const runs = richText.getRuns();
  if (!runs.length) return richText.getText() || String(fallback || '');
  return runs.map((run) => {
    const text = run.getText();
    const url = run.getLinkUrl();
    return url ? `${text} (${url})` : text;
  }).join('');
}

function calendarDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function addDays_(date, amount) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + amount);
}

function addMonthsClamped_(date, amount, preferredDay) {
  const first = new Date(date.getFullYear(), date.getMonth() + amount, 1);
  const lastDay = new Date(first.getFullYear(), first.getMonth() + 1, 0).getDate();
  return new Date(first.getFullYear(), first.getMonth(), Math.min(preferredDay, lastDay));
}

function coerceDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  return null;
}

function namedCellValue_(spreadsheet, name) {
  const range = spreadsheet.getRangeByName(name);
  if (!range) throw new Error(`Crie o intervalo nomeado "${name}".`);
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    throw new Error(`O intervalo nomeado "${name}" deve ter uma única célula.`);
  }
  return range.getDisplayValue().trim();
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function escapeHtml_(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
}
