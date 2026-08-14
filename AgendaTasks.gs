/**
 * Integra a planilha com uma lista do Google Tasks.
 *
 * Intervalos nomeados esperados:
 *   ID      - uma unica celula com o ID da lista do Google Tasks.
 *   TAREFAS - colunas: titulo, recorrencia, data inicial, detalhe, descricao/link.
 *   MÊS     - uma unica celula com o numero do mes (1 a 12).
 */

const AGENDA_TASKS = Object.freeze({
  MENU: 'Agenda',
  ID_RANGE: 'ID',
  TASKS_RANGE: 'TAREFAS',
  MONTH_RANGE: 'MÊS',
  YEARS_TO_CREATE: 2,
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
    .createMenu(AGENDA_TASKS.MENU)
    .addItem('Atualizar agenda', 'atualizarAgenda')
    .addItem('Gerar PDF', 'gerarPDF')
    .addToUi();
}

/** Apaga a lista indicada em ID e a repovoa com ocorrencias para dois anos. */
function atualizarAgenda() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const taskListId = namedCellValue_(spreadsheet, AGENDA_TASKS.ID_RANGE);
    const definitions = readTaskDefinitions_(spreadsheet);

    if (!taskListId) throw new Error('O intervalo nomeado "ID" está vazio.');
    if (!definitions.length) throw new Error('Nenhuma tarefa válida foi encontrada em "TAREFAS".');

    deleteAllTasks_(String(taskListId).trim());

    let created = 0;
    definitions.forEach((definition) => {
      occurrencesFor_(definition.startDate, definition.recurrence).forEach((date) => {
        Tasks.Tasks.insert({
          title: definition.title,
          notes: buildTaskNotes_(definition.detail, definition.description),
          // A API do Tasks conserva somente a parte da data; meio-dia evita mudanca
          // de dia durante a conversao entre os fusos da planilha e da API.
          due: taskDueDate_(date)
        }, String(taskListId).trim());
        created += 1;
      });
    });

    ui.alert('Agenda atualizada', `${created} tarefas foram criadas para os próximos 2 anos.`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Não foi possível atualizar a agenda', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

/** Gera o calendario do mes informado em MÊS, sempre usando o ano atual. */
function gerarPDF() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActive();
    const month = Number(namedCellValue_(spreadsheet, AGENDA_TASKS.MONTH_RANGE));
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
  const range = spreadsheet.getRangeByName(AGENDA_TASKS.TASKS_RANGE);
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
    if (!title) throw new Error(`Título vazio na linha ${range.getRow() + index} de "TAREFAS".`);
    if (!startDate) throw new Error(`Data inicial inválida na linha ${range.getRow() + index} de "TAREFAS".`);

    definitions.push({
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
  const endExclusive = addMonthsClamped_(startDate, AGENDA_TASKS.YEARS_TO_CREATE * 12, startDate.getDate());
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
      entries[day].push(definition.detail ? `${definition.title} — ${definition.detail}` : definition.title);
    });
  });
  return entries;
}

function createCalendarPdf_(entriesByDay, month, year) {
  const title = `Calendário - ${AGENDA_TASKS.MONTHS[month - 1]} de ${year}`;
  const document = DocumentApp.create(title);
  const body = document.getBody();
  body.clear();
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const tableData = [AGENDA_TASKS.WEEKDAYS.slice()];
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
  const folder = getOrCreateFolder_(AGENDA_TASKS.PDF_FOLDER_NAME);
  const file = folder.createFile(pdf);
  source.setTrashed(true);
  return file;
}

function deleteAllTasks_(taskListId) {
  // Sempre consulta a primeira pagina novamente, pois excluir itens invalida a
  // posicao da paginacao. Inclui tarefas concluidas e ocultas.
  while (true) {
    const response = Tasks.Tasks.list(taskListId, { maxResults: 100, showCompleted: true, showHidden: true });
    const items = response.items || [];
    if (!items.length) return;
    items.forEach((task) => Tasks.Tasks.remove(taskListId, task.id));
  }
}

function buildTaskNotes_(detail, description) {
  return `${detail || ''}\n\n${description || ''}`.replace(/\s+$/, '');
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

function taskDueDate_(date) {
  return Utilities.formatDate(new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12), 'UTC', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
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
