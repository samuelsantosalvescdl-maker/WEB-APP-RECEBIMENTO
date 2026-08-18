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
  UPDATE_STATE_VERSION: 4,
  UPDATE_HANDLER: 'continuarAtualizacaoAgenda',
  UPDATE_STATE_PROPERTY: 'AGENDA_CALENDAR_UPDATE_STATE',
  BATCH_SIZE: 10,
  MAX_BATCH_MILLISECONDS: 60000,
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
    .addToUi();
}

/** Inicia a atualizacao em blocos e agenda as continuacoes automaticamente. */
function atualizarAgenda() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  let runStarted = false;

  try {
    if (!lock.tryLock(5000)) throw new Error('Já existe uma atualização da agenda em execução. Aguarde.');
    const previousState = loadUpdateState_();
    if (isUpdateActive_(previousState)) {
      throw new Error('Já existe uma atualização em andamento. Aguarde a mensagem de conclusão.');
    }
    const spreadsheet = SpreadsheetApp.getActive();
    const calendarId = namedCellValue_(spreadsheet, AGENDA_CALENDAR.ID_RANGE);
    const definitions = readTaskDefinitions_(spreadsheet);

    if (!calendarId) throw new Error('O intervalo nomeado "ID" está vazio.');
    if (!definitions.length) throw new Error('Nenhuma tarefa válida foi encontrada em "TAREFAS".');

    cancelUpdateTriggers_();
    const runId = Utilities.getUuid();
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
      startedAt: new Date().toISOString()
    });
    runStarted = true;
    console.log(`[AgendaCalendar] runId=${runId} fase=deleting atualização iniciada`);

    const result = processUpdateBatch_();
    const message = result.complete
      ? `${result.created} eventos de dia inteiro foram criados.`
      : `Atualização iniciada em blocos (${result.created} de ${result.total} eventos criados). ` +
        'O script continuará automaticamente. Cada evento criado contém sua recorrência completa.';
    ui.alert('Atualização da agenda', message, ui.ButtonSet.OK);
  } catch (error) {
    if (runStarted) {
      cancelUpdateTriggers_();
      const state = loadUpdateState_();
      if (state) {
        state.phase = 'error';
        state.error = error.message || String(error);
        state.continuationTriggerId = null;
        saveUpdateState_(state);
        console.error(`[AgendaCalendar] runId=${state.runId} fase=error erro=${state.error}`);
      }
    } else {
      console.error(`[AgendaCalendar] falha antes de iniciar a execução: ${error.message || String(error)}`);
    }
    ui.alert('Não foi possível atualizar a agenda', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/** Continua uma atualizacao iniciada por atualizarAgenda (acionada por gatilho). */
function continuarAtualizacaoAgenda(event) {
  const expectedState = loadUpdateState_();
  if (event && event.triggerUid && expectedState && expectedState.continuationTriggerId &&
      event.triggerUid !== expectedState.continuationTriggerId) {
    console.log(`[AgendaCalendar] runId=${expectedState.runId} gatilho obsoleto ignorado`);
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    processUpdateBatch_();
  } catch (error) {
    cancelUpdateTriggers_();
    const state = loadUpdateState_();
    if (state) {
      state.error = error.message || String(error);
      state.phase = 'error';
      state.continuationTriggerId = null;
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

function processUpdateBatch_() {
  const state = loadUpdateState_();
  if (!state) throw new Error('Não existe uma atualização de agenda em andamento.');

  // O proximo gatilho e criado antes das chamadas ao Calendar. Assim, mesmo que
  // esta execucao seja encerrada abruptamente pelo limite do Apps Script, havera
  // uma nova tentativa para retomar o estado salvo.
  scheduleUpdateContinuation_(AGENDA_CALENDAR.SAFETY_TRIGGER_DELAY_MILLISECONDS, state);

  const started = Date.now();
  let operations = 0;

  if (state.phase === 'deleting') {
    const deletion = deleteCalendarEventsBatch_(state.calendarId, started);
    operations = deletion.deleted;
    state.deleted = (state.deleted || 0) + deletion.deleted;
    state.deletionEmptyChecks = deletion.empty ? (state.deletionEmptyChecks || 0) + 1 : 0;
    console.log(
      `[AgendaCalendar] runId=${state.runId} fase=deleting excluídosNoBloco=${deletion.deleted} ` +
      `totalExcluído=${state.deleted} confirmaçõesVazias=${state.deletionEmptyChecks}`
    );

    if (state.deletionEmptyChecks >= AGENDA_CALENDAR.EMPTY_CALENDAR_CONFIRMATIONS) {
      // Uma ultima leitura independente impede iniciar a criacao se a API voltar
      // a apresentar um evento durante a janela de consistencia eventual.
      if (calendarHasAnyEvent_(state.calendarId)) {
        state.deletionEmptyChecks = 0;
      } else {
        state.phase = 'creating';
        operations = 0;
        console.log(`[AgendaCalendar] runId=${state.runId} agenda confirmada vazia; fase=creating`);
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
    const eventConfirmed = insertRecurringCalendarEvent_(
      state.calendarId,
      definition,
      deterministicEventId_(state.spreadsheetId, state.runId, definition),
      state.runId
    );
    if (!eventConfirmed) throw new Error(`Não foi possível confirmar a criação da linha ${definition.sourceRow}.`);
    state.definitionIndex += 1;
    state.created += 1;
    operations += 1;
    // Persistir apos cada evento torna a retomada segura mesmo em timeout.
    saveUpdateState_(state);
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
    const recurrenceText = String(displayValues[index][1] || '').trim();
    if (!recurrenceText) return;

    const title = String(displayValues[index][0] || '').trim();
    const startDate = coerceDate_(row[2]);
    // Linhas incompletas podem permanecer no intervalo nomeado sem impedir a
    // atualizacao das demais. Somente linhas com titulo, recorrencia e data
    // inicial validos se tornam eventos.
    if (!title || !startDate) return;

    definitions.push({
      sourceRow: range.getRow() + index,
      title,
      recurrence: normalizeRecurrence_(recurrenceText, range.getRow() + index),
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
  let deleted = 0;
  let pageToken;

  while (!batchLimitReached_(started, deleted)) {
    const options = {
      maxResults: AGENDA_CALENDAR.BATCH_SIZE,
      showDeleted: false,
      singleEvents: false
    };
    if (pageToken) options.pageToken = pageToken;

    const response = Calendar.Events.list(calendarId, options);
    const items = response.items || [];

    // Nunca declarar a agenda vazia enquanto a API ainda indicar outra pagina.
    // Ao encontrar eventos, remove somente este bloco e reinicia a listagem na
    // proxima execucao, evitando reutilizar pageToken depois de mutar a colecao.
    for (let index = 0; index < items.length && !batchLimitReached_(started, deleted); index += 1) {
      removeCalendarEventSafely_(calendarId, items[index].id);
      deleted += 1;
    }

    if (deleted > 0) return { deleted, empty: false };
    pageToken = response.nextPageToken;
    if (!pageToken) return { deleted: 0, empty: true };
  }
  return { deleted, empty: false };
}

function calendarHasAnyEvent_(calendarId) {
  let pageToken;
  do {
    const options = { maxResults: 1, showDeleted: false, singleEvents: false };
    if (pageToken) options.pageToken = pageToken;
    const response = Calendar.Events.list(calendarId, options);
    if (response.items && response.items.length) return true;
    pageToken = response.nextPageToken;
  } while (pageToken);
  return false;
}

function removeCalendarEventSafely_(calendarId, eventId) {
  try {
    Calendar.Events.remove(calendarId, eventId);
  } catch (error) {
    // A listagem do Calendar pode manter por alguns segundos um evento que ja
    // foi removido. 404/410 significam que o resultado desejado ja foi obtido.
    if (!calendarErrorHasCode_(error, [404, 410])) throw error;
  }
}

function insertRecurringCalendarEvent_(calendarId, definition, eventId, runId) {
  const definitionHash = eventDefinitionHash_(definition);
  try {
    Calendar.Events.insert({
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
    }, calendarId);
    return true;
  } catch (error) {
    if (!calendarErrorHasCode_(error, [409])) throw error;
    const existing = Calendar.Events.get(calendarId, eventId);
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

function batchLimitReached_(started, operations) {
  return operations >= AGENDA_CALENDAR.BATCH_SIZE ||
    Date.now() - started >= AGENDA_CALENDAR.MAX_BATCH_MILLISECONDS;
}

function updateResult_(state, complete) {
  return { complete, created: state.created, total: state.total };
}

function finishUpdate_(state, spreadsheet) {
  state.phase = 'complete';
  state.completedAt = new Date().toISOString();
  cancelUpdateTriggers_();
  state.continuationTriggerId = null;
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
  const lastUpdate = new Date(state.updatedAt || state.startedAt || 0).getTime();
  return Date.now() - lastUpdate < 24 * 60 * 60 * 1000;
}

function saveUpdateState_(state) {
  state.updatedAt = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(
    AGENDA_CALENDAR.UPDATE_STATE_PROPERTY,
    JSON.stringify(state)
  );
}

function scheduleUpdateContinuation_(delayMilliseconds, state) {
  cancelUpdateTriggers_();
  const trigger = ScriptApp.newTrigger(AGENDA_CALENDAR.UPDATE_HANDLER)
    .timeBased()
    .after(delayMilliseconds || 60 * 1000)
    .create();
  state.continuationTriggerId = trigger.getUniqueId();
  saveUpdateState_(state);
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
