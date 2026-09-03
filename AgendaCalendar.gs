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
  PDF_TEMPLATE_RANGE: 'PDF_TEMPLATE_ID',
  WEEKDAYS: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'],
  MONTHS: [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ]
});

const PDF_CALENDAR = Object.freeze({
  TITLE_FONT_SIZE: 16,
  LEGEND_FONT_SIZE: 8,
  WEEKDAY_FONT_SIZE: 9,
  DAY_NUMBER_FONT_SIZE: 10.5,
  TASK_FONT_SIZE: 8,
  MAX_VISIBLE_TASK_ROWS: 10,
  LEGEND_ROW_HEIGHT: 12,
  TASK_TEXT_HEIGHT: 8.5,
  TASK_UNDERLINE_GAP: 0.5,
  TASK_UNDERLINE_WEIGHT: 1,
  TASK_ROW_GAP: 0.7,
  CELL_PADDING: 2,
  DAY_NUMBER_HEIGHT: 12.5,
  NEUTRAL_ASSIGNEE: 'SEM RESPONSÁVEL',
  NEUTRAL_COLOR: '#777777'
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
    const templateId = pdfTemplateId_(spreadsheet);
    const pdfFile = createCalendarPdf_(entriesByDay, month, year, templateId);

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

function pdfTemplateId_(spreadsheet) {
  const range = spreadsheet.getRangeByName(AGENDA_CALENDAR.PDF_TEMPLATE_RANGE);
  if (!range || range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    throw new Error(
      'Configure o intervalo nomeado "PDF_TEMPLATE_ID" com o ID do template A3 horizontal.'
    );
  }
  const templateId = range.getDisplayValue().trim();
  if (!templateId) {
    throw new Error(
      'Configure o intervalo nomeado "PDF_TEMPLATE_ID" com o ID do template A3 horizontal.'
    );
  }
  return templateId;
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
      const assignee = String(definition.detail || '').trim();
      const displayedAssignee = assignee || PDF_CALENDAR.NEUTRAL_ASSIGNEE;
      entries[day].push({
        title: definition.title,
        assignee: displayedAssignee,
        assigneeKey: normalizeAssigneeKey_(displayedAssignee)
      });
    });
  });
  return entries;
}

function createCalendarPdf_(entriesByDay, month, year, templateId) {
  const title = `Calendário - ${AGENDA_CALENDAR.MONTHS[month - 1]} de ${year}`;
  let presentationFile;
  try {
    let presentation;
    try {
      presentationFile = DriveApp.getFileById(templateId).makeCopy(`${title} - temporário`);
      presentation = SlidesApp.openById(presentationFile.getId());
    } catch (error) {
      throw new Error(
        `Não foi possível usar o template informado em PDF_TEMPLATE_ID. ` +
        `Confirme se o ID pertence a uma apresentação e se há permissão de acesso. ` +
        `Detalhes: ${error.message || String(error)}`
      );
    }
    const slides = presentation.getSlides();
    if (!slides.length) throw new Error('O template de PDF precisa conter pelo menos um slide em branco.');
    const slide = slides[0];
    for (let index = slides.length - 1; index > 0; index -= 1) slides[index].remove();
    slide.getPageElements().forEach((element) => element.remove());
    slide.getBackground().setSolidFill('#FFFFFF');

    const pageWidth = presentation.getPageWidth();
    const pageHeight = presentation.getPageHeight();
    console.log(`[AgendaCalendar] PDF template pageWidth=${pageWidth} pageHeight=${pageHeight}`);
    if (pageWidth <= pageHeight) {
      throw new Error(
        `O template informado em PDF_TEMPLATE_ID deve estar em orientação horizontal. ` +
        `pageWidth=${pageWidth}, pageHeight=${pageHeight}.`
      );
    }
    const firstWeekday = new Date(year, month - 1, 1).getDay();
    const daysInMonth = new Date(year, month, 0).getDate();
    const numberOfWeeks = Math.ceil((firstWeekday + daysInMonth) / 7);
    const assigneeLabels = createAssigneeLabelMap_(entriesByDay);
    const assigneeColorMap = createAssigneeColorMap_(entriesByDay);
    const assignees = Object.keys(assigneeLabels).map((key) => ({ key, label: assigneeLabels[key] }));

    const margin = 18;
    const titleHeight = 24;
    const titleGap = 3;
    const headerHeight = 18;
    const bottomMargin = 14;
    const calendarWidth = pageWidth - margin * 2;
    const legendLayout = calculateLegendLayout_(assignees, calendarWidth);
    const legendHeight = legendLayout.height;
    const calendarTop = margin + titleHeight + titleGap + legendHeight;
    const calendarHeight = pageHeight - calendarTop - bottomMargin;
    const weekHeight = (calendarHeight - headerHeight) / numberOfWeeks;
    const dayWidth = calendarWidth / 7;
    const taskFontSize = calculateCalendarFontSize_();
    validatePdfTemplateLayout_(pageWidth, pageHeight, weekHeight, dayWidth);

    insertSlideText_(slide, `Calendário — ${AGENDA_CALENDAR.MONTHS[month - 1]} de ${year}`,
      margin, margin, calendarWidth, titleHeight, PDF_CALENDAR.TITLE_FONT_SIZE,
      true, SlidesApp.ParagraphAlignment.CENTER);
    drawAssigneeLegend_(slide, assignees, assigneeColorMap, margin, margin + titleHeight,
      calendarWidth, legendLayout);
    drawCalendarGrid_(slide, entriesByDay, assigneeColorMap, firstWeekday, daysInMonth,
      numberOfWeeks, margin, calendarTop, calendarWidth, calendarHeight, headerHeight,
      dayWidth, weekHeight, taskFontSize);

    if (presentation.getSlides().length !== 1) {
      throw new Error('Falha ao garantir que o calendário tenha exatamente um slide.');
    }
    presentation.saveAndClose();
    const pdf = presentationFile.getAs(MimeType.PDF).setName(`${title}.pdf`);
    return getOrCreateFolder_(AGENDA_CALENDAR.PDF_FOLDER_NAME).createFile(pdf);
  } finally {
    if (presentationFile) presentationFile.setTrashed(true);
  }
}

function calculateCalendarFontSize_() {
  return PDF_CALENDAR.TASK_FONT_SIZE;
}

function createAssigneeColorMap_(entriesByDay) {
  const palette = [
    '#2563EB', '#DC2626', '#059669', '#7C3AED', '#D97706', '#0891B2', '#DB2777', '#4F46E5',
    '#65A30D', '#C2410C', '#0F766E', '#9333EA', '#0369A1', '#BE123C', '#4D7C0F', '#A16207'
  ];
  const labels = createAssigneeLabelMap_(entriesByDay);
  const map = {};
  const usedColors = {};
  Object.keys(labels).sort().forEach((key) => {
    if (key === normalizeAssigneeKey_(PDF_CALENDAR.NEUTRAL_ASSIGNEE)) {
      map[key] = PDF_CALENDAR.NEUTRAL_COLOR;
      return;
    }
    const initialIndex = stableTextHash_(key) % palette.length;
    let color;
    for (let offset = 0; offset < palette.length; offset += 1) {
      const candidate = palette[(initialIndex + offset) % palette.length];
      if (!usedColors[candidate]) {
        color = candidate;
        break;
      }
    }
    let attempt = 0;
    while (!color) {
      const hue = (stableTextHash_(key) + attempt * 137.508) % 360;
      const candidate = hslToHex_(hue, 58, 42);
      if (!usedColors[candidate] && candidate !== PDF_CALENDAR.NEUTRAL_COLOR) color = candidate;
      attempt += 1;
    }
    map[key] = color;
    usedColors[color] = true;
  });
  return map;
}

function createAssigneeLabelMap_(entriesByDay) {
  const labels = {};
  Object.keys(entriesByDay).forEach((day) => {
    entriesByDay[day].forEach((task) => {
      const label = String(task.assignee || '').trim() || PDF_CALENDAR.NEUTRAL_ASSIGNEE;
      const key = task.assigneeKey || normalizeAssigneeKey_(label);
      if (!labels[key]) labels[key] = label;
    });
  });
  return labels;
}

function normalizeAssigneeKey_(value) {
  const normalized = String(value || '').trim() || PDF_CALENDAR.NEUTRAL_ASSIGNEE;
  return normalized.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLocaleLowerCase('pt-BR');
}

function stableTextHash_(text) {
  let hash = 2166136261;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function hslToHex_(hue, saturation, lightness) {
  const s = saturation / 100;
  const l = lightness / 100;
  const chroma = (1 - Math.abs(2 * l - 1)) * s;
  const x = chroma * (1 - Math.abs((hue / 60) % 2 - 1));
  const m = l - chroma / 2;
  let rgb;
  if (hue < 60) rgb = [chroma, x, 0];
  else if (hue < 120) rgb = [x, chroma, 0];
  else if (hue < 180) rgb = [0, chroma, x];
  else if (hue < 240) rgb = [0, x, chroma];
  else if (hue < 300) rgb = [x, 0, chroma];
  else rgb = [chroma, 0, x];
  return `#${rgb.map((value) => Math.round((value + m) * 255).toString(16).padStart(2, '0')).join('')}`;
}

function calculateLegendLayout_(assignees, availableWidth) {
  if (!assignees.length) return { fontSize: PDF_CALENDAR.LEGEND_FONT_SIZE, rows: [], height: 4 };
  const rowCount = assignees.length <= 6 ? 1 : 2;
  const rows = packLegendRows_(assignees, availableWidth, rowCount);
  return {
    fontSize: PDF_CALENDAR.LEGEND_FONT_SIZE,
    rows,
    height: rows.length * PDF_CALENDAR.LEGEND_ROW_HEIGHT + 3
  };
}

function packLegendRows_(assignees, availableWidth, rowCount) {
  const rows = [];
  const itemsPerRow = Math.ceil(assignees.length / rowCount);
  for (let start = 0; start < assignees.length; start += itemsPerRow) {
    const rowAssignees = assignees.slice(start, start + itemsPerRow);
    const itemWidth = availableWidth / rowAssignees.length;
    rows.push(rowAssignees.map((assignee) => {
      const markerWidth = 10;
      const markerGap = 4;
      const textWidth = Math.max(1, itemWidth - markerWidth - markerGap - 5);
      return {
        key: assignee.key,
        displayLabel: truncateTaskTitleToWidth_(assignee.label, textWidth, PDF_CALENDAR.LEGEND_FONT_SIZE),
        width: itemWidth,
        markerWidth,
        markerGap
      };
    }));
  }
  return rows;
}

function drawAssigneeLegend_(slide, assignees, colorMap, left, top, width, layout) {
  if (!assignees.length) return;
  layout.rows.forEach((row, rowIndex) => {
    const rowWidth = row.reduce((total, item) => total + item.width, 0);
    let x = left + (width - rowWidth) / 2;
    const y = top + rowIndex * PDF_CALENDAR.LEGEND_ROW_HEIGHT + 2;
    row.forEach((item) => {
      const lineY = y + 5;
      insertSlideLine_(slide, x, lineY, x + item.markerWidth, lineY,
        colorMap[item.key] || PDF_CALENDAR.NEUTRAL_COLOR, PDF_CALENDAR.TASK_UNDERLINE_WEIGHT);
      const textLeft = x + item.markerWidth + item.markerGap;
      insertSlideText_(slide, noWrapText_(item.displayLabel), textLeft, y,
        item.width - item.markerWidth - item.markerGap - 3, 9,
        layout.fontSize, false, SlidesApp.ParagraphAlignment.START);
      x += item.width;
    });
  });
}

function drawCalendarGrid_(slide, entriesByDay, colorMap, firstWeekday, daysInMonth,
    numberOfWeeks, left, top, calendarWidth, calendarHeight, headerHeight,
    dayWidth, weekHeight, taskFontSize) {
  AGENDA_CALENDAR.WEEKDAYS.forEach((weekday, column) => {
    const x = left + column * dayWidth;
    insertSlideRectangle_(slide, x, top, dayWidth, headerHeight, '#F3F4F6', '#CBD5E1');
    insertSlideText_(slide, weekday, x + 2, top + 2, dayWidth - 4, headerHeight - 4,
      PDF_CALENDAR.WEEKDAY_FONT_SIZE, true, SlidesApp.ParagraphAlignment.CENTER);
  });

  for (let week = 0; week < numberOfWeeks; week += 1) {
    for (let column = 0; column < 7; column += 1) {
      const x = left + column * dayWidth;
      const y = top + headerHeight + week * weekHeight;
      insertSlideRectangle_(slide, x, y, dayWidth, weekHeight, '#FFFFFF', '#CBD5E1');
      const day = week * 7 + column - firstWeekday + 1;
      if (day < 1 || day > daysInMonth) continue;
      drawCalendarDay_(slide, day, entriesByDay[day] || [], colorMap,
        x, y, dayWidth, weekHeight, taskFontSize);
    }
  }
}

function drawCalendarDay_(slide, day, tasks, colorMap, left, top, width, height, fontSize) {
  const padding = PDF_CALENDAR.CELL_PADDING;
  const dayNumberHeight = PDF_CALENDAR.DAY_NUMBER_HEIGHT;
  insertSlideText_(slide, String(day), left + padding, top + 1, width - padding * 2,
    dayNumberHeight, PDF_CALENDAR.DAY_NUMBER_FONT_SIZE, true, SlidesApp.ParagraphAlignment.START);

  const layout = calculateDayTaskLayout_(width, height);
  const visibleRows = visibleDayTaskRows_(tasks);
  visibleRows.forEach((rowItem, row) => {
    const taskLeft = left + padding;
    const taskTop = top + dayNumberHeight + row * layout.taskRowHeight;
    const rowBottom = taskTop + PDF_CALENDAR.TASK_TEXT_HEIGHT +
      (rowItem.overflowCount ? 0 : PDF_CALENDAR.TASK_UNDERLINE_GAP) + PDF_CALENDAR.TASK_ROW_GAP;
    if (rowBottom > top + height) {
      throw new Error(`Overflow ao desenhar as tarefas do dia ${day}.`);
    }
    if (rowItem.overflowCount) {
      insertSlideText_(slide, `+${rowItem.overflowCount} tarefas`, taskLeft, taskTop,
        layout.contentWidth, PDF_CALENDAR.TASK_TEXT_HEIGHT, fontSize, true,
        SlidesApp.ParagraphAlignment.START);
      return;
    }
    const task = rowItem.task;
    const displayTitle = truncateTaskTitleToWidth_(task.title, layout.contentWidth, fontSize);
    insertSlideText_(slide, noWrapText_(displayTitle), taskLeft, taskTop, layout.contentWidth,
      PDF_CALENDAR.TASK_TEXT_HEIGHT, fontSize, false, SlidesApp.ParagraphAlignment.START);
    const underlineY = taskTop + PDF_CALENDAR.TASK_TEXT_HEIGHT + PDF_CALENDAR.TASK_UNDERLINE_GAP;
    const underlineWidth = Math.min(layout.contentWidth,
      Math.max(1, estimateSingleLineTextWidth_(displayTitle, fontSize)));
    insertSlideLine_(slide, taskLeft, underlineY, taskLeft + underlineWidth, underlineY,
      colorMap[task.assigneeKey || normalizeAssigneeKey_(task.assignee)] || PDF_CALENDAR.NEUTRAL_COLOR,
      PDF_CALENDAR.TASK_UNDERLINE_WEIGHT);
  });
}

function calculateDayTaskLayout_(cellWidth, cellHeight) {
  const contentWidth = Math.max(1, cellWidth - PDF_CALENDAR.CELL_PADDING * 2);
  const taskRowHeight = PDF_CALENDAR.TASK_TEXT_HEIGHT + PDF_CALENDAR.TASK_UNDERLINE_GAP +
    PDF_CALENDAR.TASK_ROW_GAP;
  const availableTaskHeight = cellHeight - PDF_CALENDAR.DAY_NUMBER_HEIGHT - PDF_CALENDAR.CELL_PADDING;
  const requiredHeight = PDF_CALENDAR.MAX_VISIBLE_TASK_ROWS * taskRowHeight;
  return {
    taskRowHeight,
    maxVisibleRows: PDF_CALENDAR.MAX_VISIBLE_TASK_ROWS,
    availableTaskHeight,
    requiredHeight,
    contentWidth
  };
}

function visibleDayTaskRows_(tasks) {
  if (tasks.length <= PDF_CALENDAR.MAX_VISIBLE_TASK_ROWS) {
    return tasks.map((task) => ({ task }));
  }
  const taskRows = tasks.slice(0, PDF_CALENDAR.MAX_VISIBLE_TASK_ROWS - 1)
    .map((task) => ({ task }));
  taskRows.push({ overflowCount: tasks.length - (PDF_CALENDAR.MAX_VISIBLE_TASK_ROWS - 1) });
  return taskRows;
}

function validatePdfTemplateLayout_(pageWidth, pageHeight, weekHeight, dayWidth) {
  const layout = calculateDayTaskLayout_(dayWidth, weekHeight);
  if (layout.availableTaskHeight + 0.001 < layout.requiredHeight) {
    throw new Error(
      `O template do PDF não possui altura suficiente para ${PDF_CALENDAR.MAX_VISIBLE_TASK_ROWS} ` +
      `tarefas por dia com fonte de ${PDF_CALENDAR.TASK_FONT_SIZE} pt. ` +
      `pageWidth=${pageWidth}, pageHeight=${pageHeight}, weekHeight=${weekHeight}, ` +
      `availableTaskHeight=${layout.availableTaskHeight}, requiredHeight=${layout.requiredHeight}.`
    );
  }
}

function truncateTaskTitleToWidth_(text, availableWidth, fontSize) {
  const normalized = String(text || '').replace(/\s+/g, ' ').trim();
  if (!normalized || estimateSingleLineTextWidth_(normalized, fontSize) <= availableWidth) return normalized;
  const ellipsis = '…';
  if (estimateSingleLineTextWidth_(ellipsis, fontSize) > availableWidth) return ellipsis;
  let low = 0;
  let high = normalized.length;
  while (low < high) {
    const middle = Math.ceil((low + high) / 2);
    const candidate = normalized.slice(0, middle).trimEnd() + ellipsis;
    if (estimateSingleLineTextWidth_(candidate, fontSize) <= availableWidth) low = middle;
    else high = middle - 1;
  }
  return normalized.slice(0, low).trimEnd() + ellipsis;
}

function estimateSingleLineTextWidth_(text, fontSize) {
  return Array.from(String(text || '')).reduce((width, character) => {
    if (/\s/.test(character)) return width + fontSize * 0.28;
    if (/[ilI1\.,'`]/.test(character)) return width + fontSize * 0.26;
    if (/[MW@%]/.test(character)) return width + fontSize * 0.82;
    if (/[A-ZÁÀÂÃÉÊÍÓÔÕÚÇ]/.test(character)) return width + fontSize * 0.61;
    return width + fontSize * 0.52;
  }, 0);
}

function noWrapText_(text) {
  return String(text || '').replace(/\s+/g, ' ').trim().replace(/ /g, '\u00A0');
}

function insertSlideRectangle_(slide, left, top, width, height, fillColor, borderColor) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, width, height);
  shape.getFill().setSolidFill(fillColor);
  shape.getBorder().getLineFill().setSolidFill(borderColor);
  shape.getBorder().setWeight(0.6);
  return shape;
}

function insertSlideText_(slide, text, left, top, width, height, fontSize, bold, alignment) {
  const shape = slide.insertTextBox(String(text || ''), left, top, Math.max(1, width), Math.max(1, height));
  shape.getAutofit().disableAutofit();
  shape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  const textRange = shape.getText();
  textRange.getTextStyle()
    .setFontFamily('Arial')
    .setFontSize(fontSize)
    .setForegroundColor('#1F2937')
    .setBold(Boolean(bold));
  textRange.getParagraphStyle().setParagraphAlignment(alignment);
  return shape;
}

function insertSlideLine_(slide, startX, startY, endX, endY, color, weight) {
  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, startX, startY, endX, endY);
  line.getLineFill().setSolidFill(color);
  line.setWeight(weight);
  return line;
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
