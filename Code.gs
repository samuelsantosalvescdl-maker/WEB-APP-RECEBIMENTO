const SHEET_NAMES = {
  ORDERS: 'CONT_FIN',
  ITEMS: 'REC_POR_ITEM',
};

const ORDER_RANGE_MAP = {
  oc: 0, // A
  buyerSelected: 1, // B
  receivedAt: 2, // C
  supplierSelected: 3, // D
  qtyTotal: 4, // E
  valueTotal: 5, // F
  status: 6, // G
  buyerDetailsJson: 8, // I
  supplierDetailsJson: 9, // J
  nf: 10, // K
  nfFrete: 11, // L
  comment: 12, // M
  groupOcs: 13, // N
  groupSentAts: 14, // O
  sentAt: 15, // P
  boleto: 17, // R
};

const ITEM_RANGE_MAP = {
  oc: 0, // A
  buyerSelected: 1, // B
  supplierSelected: 2, // C
  receivedAt: 3, // D
  code: 4, // E
  item: 5, // F
  unit: 6, // G
  qty: 7, // H
  qtyReceived: 8, // I
  unitPrice: 9, // J
  validity: 10, // K
  total: 11, // L
  lineNo: 12, // M
  labels: 13, // N
  obsItem: 15, // P
};

const SPREADSHEET_ID = '1mc3nNSeW6GI2rXudQ30c2bzIlDtccheEdsTG85n_Y4g';
const LABELS_FOLDER_ID = '1UzyIn1fsiVIatfgQQK-GyGIeJI4Z-AFs';
const LAST_LABELS_PDF_PROPERTY = 'LAST_LABELS_PDF_FILE_ID';

function doGet(e) {
  const view = e && e.parameter && e.parameter.view;
  if (view === 'print') {
    return HtmlService.createHtmlOutputFromFile('Print').setTitle('Impressão de Etiquetas');
  }

  if (view === 'pdf') {
    const fileId = e && e.parameter && e.parameter.fileId;
    if (!fileId) {
      return HtmlService.createHtmlOutput('<p>Arquivo não informado.</p>');
    }

    try {
      const file = DriveApp.getFileById(fileId);
      const base64 = Utilities.base64Encode(file.getBlob().getBytes());
      // Mantém resposta simples via data URL; para PDFs muito grandes pode haver limitação no navegador.
      const dataUrl = `data:application/pdf;base64,${base64}`;
      const safeDataUrl = dataUrl.replace(/"/g, '&quot;');
      const html = '<!doctype html><html><head><meta charset="utf-8"><title>PDF</title>'
        + '<style>html,body{margin:0;height:100%;}embed{width:100%;height:100%;}</style></head>'
        + `<body><embed src="${safeDataUrl}" type="application/pdf"></body></html>`;
      return HtmlService.createHtmlOutput(html).setTitle('PDF');
    } catch (error) {
      return HtmlService.createHtmlOutput('<p>PDF não encontrado.</p>');
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Recebimento de Pedidos');
}

function getOrdersWithItems() {
  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);

  const orders = readOrdersFromRange_(ordersRange);
  const itemsByOc = readItemsByOcFromRange_(itemsRange);

  const enriched = orders.map((order) => {
    return Object.assign({}, order, {
      items: itemsByOc[order.oc] || [],
    });
  });

  return enriched.sort((a, b) => {
    const dateA = new Date(a.sentAt || 0).getTime() || 0;
    const dateB = new Date(b.sentAt || 0).getTime() || 0;
    return dateB - dateA;
  });
}

function getBuyerOptions() {
  return getNamedRangeFirstColumnValues_(getSpreadsheet_(), 'EMP_COMP');
}

function getSupplierOptions() {
  return getNamedRangeFirstColumnValues_(getSpreadsheet_(), 'EMP_FORN');
}

function getNamedRangeFirstColumnValues_(spreadsheet, rangeName) {
  const range = getNamedRange_(spreadsheet, rangeName);
  const values = range.offset(0, 0, range.getNumRows(), 1).getDisplayValues();

  const result = values
    .map((row) => row[0])
    .filter((value) => value !== null && value !== undefined && String(value).trim() !== '')
    .map((value) => String(value).trim());

  Logger.log(JSON.stringify({
    fn: 'getNamedRangeFirstColumnValues_',
    rangeName: rangeName,
    a1: range.getA1Notation(),
    sheet: range.getSheet().getName(),
    numRows: range.getNumRows(),
    numCols: range.getNumColumns(),
    resultCount: result.length,
    sample: result.slice(0, 10)
  }));

  return result;
}

function debugDropdownSources() {
  const ss = getSpreadsheet_();
  const empCompRange = getNamedRange_(ss, 'EMP_COMP');
  const empFornRange = getNamedRange_(ss, 'EMP_FORN');

  const buyers = getNamedRangeFirstColumnValues_(ss, 'EMP_COMP');
  const suppliers = getNamedRangeFirstColumnValues_(ss, 'EMP_FORN');

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    empComp: {
      a1: empCompRange.getA1Notation(),
      sheet: empCompRange.getSheet().getName(),
      rows: empCompRange.getNumRows(),
      cols: empCompRange.getNumColumns(),
      count: buyers.length,
      sample: buyers.slice(0, 10)
    },
    empForn: {
      a1: empFornRange.getA1Notation(),
      sheet: empFornRange.getSheet().getName(),
      rows: empFornRange.getNumRows(),
      cols: empFornRange.getNumColumns(),
      count: suppliers.length,
      sample: suppliers.slice(0, 10)
    }
  };
}

function updateOrderHeaderFields(oc, nfValue, nfFreteValue, boletoValue) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const nf = sanitizeText_(nfValue);
  const nfFrete = sanitizeText_(nfFreteValue);
  const boleto = normalizeBoletoValue_(boletoValue);

  const sheet = ordersRange.getSheet();
  const startCol = ordersRange.getColumn();
  sheet.getRange(row, startCol + ORDER_RANGE_MAP.nf).setValue(nf);
  sheet.getRange(row, startCol + ORDER_RANGE_MAP.nfFrete).setValue(nfFrete);
  sheet.getRange(row, startCol + ORDER_RANGE_MAP.boleto).setValue(boleto);

  return { ok: true };
}

function updateOrderComment(oc, commentText) {
  return updateOrderCommentField(oc, commentText);
}

function updateOrderCommentField(oc, commentText) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const sheet = ordersRange.getSheet();
  const startCol = ordersRange.getColumn();
  sheet.getRange(row, startCol + ORDER_RANGE_MAP.comment).setValue(sanitizeText_(commentText));

  return { ok: true };
}

function updateItemFields(oc, lineNo, qtyReceived, validity, labels, rowIndex) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const normalizedValidity = normalizeBrDateInput_(validity);
  if (!normalizedValidity) {
    throw new Error('Informe uma validade no formato dd/MM/aaaa.');
  }

  const qty = parseNumber_(qtyReceived);
  if (!isPositiveNumber_(qty)) {
    throw new Error('Quantidade recebida deve ser maior que zero.');
  }

  const labelsValue = parseNumber_(labels);
  if (!isPositiveInteger_(labelsValue)) {
    throw new Error('Etiquetas deve ser um inteiro positivo.');
  }

  const spreadsheet = getSpreadsheet_();
  const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);
  const row = rowIndex || findItemRowInRange_(itemsRange, oc, lineNo);
  if (!row) {
    throw new Error('Item não encontrado.');
  }

  const sheet = itemsRange.getSheet();
  const startCol = itemsRange.getColumn();
  sheet.getRange(row, startCol + ITEM_RANGE_MAP.qtyReceived).setValue(qty);
  sheet.getRange(row, startCol + ITEM_RANGE_MAP.validity).setValue(normalizedValidity);
  sheet.getRange(row, startCol + ITEM_RANGE_MAP.labels).setValue(labelsValue);

  return { ok: true };
}

function receiveOrder(oc, labelWidthCm, labelHeightCm) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const width = parseNumber_(labelWidthCm);
  const height = parseNumber_(labelHeightCm);
  if (!isPositiveNumber_(width) || !isPositiveNumber_(height)) {
    throw new Error('Informe largura e altura válidas para a etiqueta.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const spreadsheet = getSpreadsheet_();
    const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
    const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);

    const orderRow = findOrderRowByOcInRange_(ordersRange, oc);
    if (!orderRow) {
      throw new Error('Pedido não encontrado.');
    }

    const order = readOrdersFromRange_(ordersRange).find((entry) => String(entry.oc) === String(oc));
    if (!order) {
      throw new Error('Pedido não encontrado.');
    }

    validateOrderHeaderForReceive_(order);

    const items = readItemsByOcFromRange_(itemsRange)[oc] || [];
    if (!items.length) {
      throw new Error('Itens não encontrados.');
    }

    if (items.some((item) => !isItemCompleteWithLabels_(item))) {
      throw new Error('Preencha quantidade recebida, validade válida e etiquetas em todas as linhas.');
    }

    const receivedAt = new Date();
    updateReceiptFieldsInRange_(itemsRange, items, order, receivedAt);

    deleteLastPdf_();
    const pdfResult = generateLabelsPdf_(oc, order, items, width, height);

    const ordersSheet = ordersRange.getSheet();
    const ordersStartCol = ordersRange.getColumn();
    ordersSheet.getRange(orderRow, ordersStartCol + ORDER_RANGE_MAP.status).setValue('RECEBIDO');
    ordersSheet.getRange(orderRow, ordersStartCol + ORDER_RANGE_MAP.receivedAt).setValue(receivedAt);

    return { ok: true, oc, status: 'RECEBIDO', pdfUrl: pdfResult.pdfUrl, pdfFileId: pdfResult.pdfFileId };
  } finally {
    lock.releaseLock();
  }
}

function markReceived(oc) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);

  const order = readOrdersFromRange_(ordersRange).find((entry) => String(entry.oc) === String(oc));
  if (!order) {
    throw new Error('Pedido não encontrado.');
  }

  validateOrderHeaderForReceive_(order);

  const items = readItemsByOcFromRange_(itemsRange)[oc] || [];
  if (!items.length) {
    throw new Error('Itens não encontrados.');
  }

  if (items.some((item) => !isItemCompleteWithLabels_(item))) {
    throw new Error('Preencha todos os campos obrigatórios antes de receber.');
  }

  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const ordersSheet = ordersRange.getSheet();
  const ordersStartCol = ordersRange.getColumn();
  ordersSheet.getRange(row, ordersStartCol + ORDER_RANGE_MAP.status).setValue('RECEBIDO');
  ordersSheet.getRange(row, ordersStartCol + ORDER_RANGE_MAP.receivedAt).setValue(new Date());

  return { ok: true, status: 'RECEBIDO' };
}

function cancelOrderStub(oc) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const ordersSheet = ordersRange.getSheet();
  const ordersStartCol = ordersRange.getColumn();
  const status = 'CANCEL_PENDENTE';
  ordersSheet.getRange(row, ordersStartCol + ORDER_RANGE_MAP.status).setValue(status);

  return { ok: true, oc, status };
}

function getSpreadsheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID não configurado.');
  }

  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw new Error('Falha ao abrir a planilha pelo ID configurado. Verifique SPREADSHEET_ID.');
  }
}

function getNamedRange_(spreadsheet, rangeName) {
  const range = spreadsheet.getRangeByName(rangeName);
  if (!range) {
    throw new Error(`Named range não encontrado: ${rangeName}`);
  }
  return range;
}

function deleteLastPdf_() {
  const props = PropertiesService.getScriptProperties();
  const lastId = props.getProperty(LAST_LABELS_PDF_PROPERTY);
  if (!lastId) {
    return;
  }

  try {
    DriveApp.getFileById(lastId).setTrashed(true);
  } catch (error) {
    // Ignora erro de arquivo ausente.
  }

  props.deleteProperty(LAST_LABELS_PDF_PROPERTY);
}

function generateLabelsPdf_(oc, order, items, labelWidthCm, labelHeightCm) {
  const folder = DriveApp.getFolderById(LABELS_FOLDER_ID);
  const html = buildPdfHtml_(order, items, labelWidthCm, labelHeightCm);
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  const fileName = `Etiquetas_${oc}_${new Date().getTime()}.pdf`;
  const file = folder.createFile(blob).setName(fileName);

  PropertiesService.getScriptProperties().setProperty(LAST_LABELS_PDF_PROPERTY, file.getId());

  return {
    pdfUrl: file.getUrl(),
    pdfFileId: file.getId(),
  };
}

function buildPdfHtml_(order, items, labelWidthCm, labelHeightCm) {
  const gapCm = 0.2;
  const pageWidth = (3 * labelWidthCm) + (2 * gapCm);
  const pageHeight = labelHeightCm;
  const labels = buildLabelsQueue_(order, items);
  const pages = [];

  for (let i = 0; i < labels.length; i += 3) {
    const slice = labels.slice(i, i + 3);
    while (slice.length < 3) {
      slice.push(null);
    }

    const labelsHtml = slice.map((entry) => {
      if (!entry) {
        return '<div class="label label-empty"></div>';
      }

      return `
        <div class="label">
          <div class="label-line"><strong>OC:</strong> ${escapeHtml_(entry.oc)}</div>
          <div class="label-line"><strong>Fornecedor:</strong> ${escapeHtml_(entry.supplier)}</div>
          <div class="label-line"><strong>Item:</strong> ${escapeHtml_(entry.item)}</div>
          <div class="label-line"><strong>Código:</strong> ${escapeHtml_(entry.code)}</div>
          <div class="label-line"><strong>Endereço:</strong> ${escapeHtml_(entry.obsItem)}</div>
          <div class="label-line"><strong>Validade:</strong> ${escapeHtml_(entry.validity)}</div>
        </div>
      `;
    });

    pages.push(`<div class="page">${labelsHtml.join('')}</div>`);
  }

  return `
    <html>
      <head>
        <style>
          @page { size: ${pageWidth}cm ${pageHeight}cm; margin: 0; }
          body { margin: 0; font-family: Arial, sans-serif; }
          .page {
            width: ${pageWidth}cm;
            height: ${pageHeight}cm;
            page-break-after: always;
            display: grid;
            grid-template-columns: repeat(3, ${labelWidthCm}cm);
            column-gap: ${gapCm}cm;
            align-items: stretch;
          }
          .label {
            width: ${labelWidthCm}cm;
            height: ${labelHeightCm}cm;
            border: 1px solid #000;
            padding: 0.2cm;
            box-sizing: border-box;
            font-size: 10px;
            display: flex;
            flex-direction: column;
            justify-content: center;
          }
          .label-empty { border-color: transparent; }
          .label-line { line-height: 1.2; }
        </style>
      </head>
      <body>
        ${pages.join('')}
      </body>
    </html>
  `;
}

function buildLabelsQueue_(order, items) {
  const queue = [];
  items.forEach((item) => {
    const labelsCount = Number(item.labels) || 0;
    for (let i = 0; i < labelsCount; i += 1) {
      queue.push({
        oc: order.oc,
        supplier: order.supplierSelected || '',
        item: item.item || '',
        code: item.code || '',
        validity: item.validity || '',
        obsItem: item.obsItem || '',
      });
    }
  });
  return queue;
}

function findOrderRowByOcInRange_(range, oc) {
  const values = range.getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][ORDER_RANGE_MAP.oc]) === String(oc)) {
      return range.getRow() + i;
    }
  }
  return null;
}

function findItemRowInRange_(range, oc, lineNo) {
  const values = range.getValues();
  for (let i = 0; i < values.length; i += 1) {
    const rowOc = String(values[i][ITEM_RANGE_MAP.oc]);
    const rowLine = Number(values[i][ITEM_RANGE_MAP.lineNo]);
    if (rowOc === String(oc) && rowLine === Number(lineNo)) {
      return range.getRow() + i;
    }
  }
  return null;
}

function readOrdersFromRange_(range) {
  const values = range.getValues();
  return values
    .filter((row) => {
      const ocValue = row[ORDER_RANGE_MAP.oc];
      return ocValue !== null && ocValue !== undefined && String(ocValue).trim() !== '';
    })
    .map((row) => ({
      oc: row[ORDER_RANGE_MAP.oc],
      status: row[ORDER_RANGE_MAP.status],
      sentAt: formatDateValue_(row[ORDER_RANGE_MAP.sentAt]),
      buyerSelected: row[ORDER_RANGE_MAP.buyerSelected],
      supplierSelected: row[ORDER_RANGE_MAP.supplierSelected],
      qtyTotal: row[ORDER_RANGE_MAP.qtyTotal],
      valueTotal: row[ORDER_RANGE_MAP.valueTotal],
      buyerDetails: parseJsonSafe_(row[ORDER_RANGE_MAP.buyerDetailsJson]),
      supplierDetails: parseJsonSafe_(row[ORDER_RANGE_MAP.supplierDetailsJson]),
      receivedAt: formatDateValue_(row[ORDER_RANGE_MAP.receivedAt]),
      nf: sanitizeText_(row[ORDER_RANGE_MAP.nf]),
      nfFrete: sanitizeText_(row[ORDER_RANGE_MAP.nfFrete]),
      comment: sanitizeText_(row[ORDER_RANGE_MAP.comment]),
      groupOcs: sanitizeText_(row[ORDER_RANGE_MAP.groupOcs]),
      groupSentAts: sanitizeText_(row[ORDER_RANGE_MAP.groupSentAts]),
      boleto: sanitizeText_(row[ORDER_RANGE_MAP.boleto]),
    }));
}

function readItemsByOcFromRange_(range) {
  const values = range.getValues();
  return values.reduce((acc, row, index) => {
    const oc = row[ITEM_RANGE_MAP.oc];
    if (!oc) {
      return acc;
    }
    if (!acc[oc]) {
      acc[oc] = [];
    }

    const normalizedValidity = normalizeBrDateInput_(row[ITEM_RANGE_MAP.validity]);
    acc[oc].push({
      oc,
      lineNo: row[ITEM_RANGE_MAP.lineNo],
      rowIndex: range.getRow() + index,
      code: row[ITEM_RANGE_MAP.code],
      item: row[ITEM_RANGE_MAP.item],
      unit: row[ITEM_RANGE_MAP.unit],
      qty: row[ITEM_RANGE_MAP.qty],
      unitPrice: row[ITEM_RANGE_MAP.unitPrice],
      total: row[ITEM_RANGE_MAP.total] || (row[ITEM_RANGE_MAP.qty] * row[ITEM_RANGE_MAP.unitPrice]),
      qtyReceived: row[ITEM_RANGE_MAP.qtyReceived],
      validity: normalizedValidity || formatDateOnly_(row[ITEM_RANGE_MAP.validity]),
      labels: row[ITEM_RANGE_MAP.labels],
      buyerSelected: row[ITEM_RANGE_MAP.buyerSelected],
      supplierSelected: row[ITEM_RANGE_MAP.supplierSelected],
      receivedAt: row[ITEM_RANGE_MAP.receivedAt],
      obsItem: row[ITEM_RANGE_MAP.obsItem],
    });
    return acc;
  }, {});
}

function updateReceiptFieldsInRange_(range, items, order, receivedAt) {
  const sheet = range.getSheet();
  const startCol = range.getColumn();
  const startRow = range.getRow();
  const values = range.getValues();

  // Mantém Date real para compatibilidade com filtros/ordenações na planilha.
  values.forEach((row, index) => {
    if (String(row[ITEM_RANGE_MAP.oc]) !== String(order.oc)) {
      return;
    }
    const rowNumber = startRow + index;
    sheet.getRange(rowNumber, startCol + ITEM_RANGE_MAP.buyerSelected).setValue(order.buyerSelected || '');
    sheet.getRange(rowNumber, startCol + ITEM_RANGE_MAP.supplierSelected).setValue(order.supplierSelected || '');
    sheet.getRange(rowNumber, startCol + ITEM_RANGE_MAP.receivedAt).setValue(receivedAt);
  });
}

function validateOrderHeaderForReceive_(order) {
  const nf = sanitizeText_(order.nf);
  const nfFrete = sanitizeText_(order.nfFrete);
  const boleto = normalizeBoletoValue_(order.boleto);

  if (!nf) {
    throw new Error('Preencha NF ou marque como não veio.');
  }
  if (!nfFrete) {
    throw new Error('Preencha NF Frete ou marque como não veio.');
  }
  if (!boleto) {
    throw new Error('Selecione o status do boleto (VEIO ou NÃO VEIO).');
  }
}

function isItemCompleteWithLabels_(item) {
  const qtyReceived = parseNumber_(item.qtyReceived);
  const labels = parseNumber_(item.labels);
  const normalizedValidity = normalizeBrDateInput_(item.validity);
  return isPositiveNumber_(qtyReceived)
    && isPositiveInteger_(labels)
    && !!normalizedValidity
    && isValidDateString_(normalizedValidity);
}

function normalizeBrDateInput_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  if (value === null || value === undefined) {
    return '';
  }

  const clean = String(value).trim();
  if (!clean) {
    return '';
  }

  const normalized = clean.replace(/[-.]/g, '/');
  const parts = normalized.split('/');
  if (parts.length !== 3) {
    return '';
  }

  const day = Number(parts[0]);
  const month = Number(parts[1]);
  const year = Number(parts[2]);

  if (!Number.isInteger(day) || !Number.isInteger(month) || !Number.isInteger(year)) {
    return '';
  }

  const dateText = `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(year).padStart(4, '0')}`;
  return isValidBrDateStrict_(dateText) ? dateText : '';
}

function isValidBrDateStrict_(value) {
  if (typeof value !== 'string') {
    return false;
  }

  const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) {
    return false;
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }

  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === (month - 1) && date.getDate() === day;
}

function isValidDateString_(value) {
  return !!normalizeBrDateInput_(value);
}

function parseNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return NaN;
  }
  if (typeof value === 'number') {
    return value;
  }
  return Number(String(value).replace(',', '.'));
}

function isPositiveNumber_(value) {
  return typeof value === 'number' && !Number.isNaN(value) && value > 0;
}

function isPositiveInteger_(value) {
  return Number.isInteger(value) && value > 0;
}

function normalizeBoletoValue_(value) {
  const text = sanitizeText_(value).toUpperCase();
  if (!text) {
    return '';
  }
  if (text === 'VEIO') {
    return 'VEIO';
  }
  if (text === 'NÃO VEIO' || text === 'NAO VEIO') {
    return 'NÃO VEIO';
  }
  throw new Error('Boleto inválido. Use VEIO ou NÃO VEIO.');
}

function sanitizeText_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function escapeHtml_(value) {
  return String(value || '').replace(/[&<>"']/g, (char) => {
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
    };
    return map[char] || char;
  });
}

function parseJsonSafe_(value) {
  if (!value) {
    return [];
  }
  try {
    return JSON.parse(value);
  } catch (error) {
    return [];
  }
}

function formatDateValue_(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (!value) {
    return '';
  }
  return String(value);
}

function formatDateOnly_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return sanitizeText_(value);
}
