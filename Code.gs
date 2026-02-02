const API_KEY = '741852963';

const SHEET_NAMES = {
  ORDERS: 'ORDERS',
  ITEMS: 'ITEMS',
};

const SPREADSHEET_ID = '1mc3nNSeW6GI2rXudQ30c2bzIlDtccheEdsTG85n_Y4g';
const LABELS_FOLDER_ID = '1UzyIn1fsiVIatfgQQK-GyGIeJI4Z-AFs';
const WEBAPP_C_URL = 'URL_DO_WEBAPP_C_EXEC_AQUI'; // placeholder
const WEBAPP_C_API_KEY = 'SUA_API_KEY_C_AQUI'; // placeholder
const LAST_LABELS_PDF_PROPERTY = 'LAST_LABELS_PDF_FILE_ID';

const ORDER_HEADERS = [
  'oc',
  'status',
  'sentAt',
  'buyerSelected',
  'supplierSelected',
  'qtyTotal',
  'valueTotal',
  'buyerDetailsJson',
  'supplierDetailsJson',
  'receivedAt',
];

const ITEM_HEADERS = [
  'oc',
  'lineNo',
  'code',
  'item',
  'unit',
  'qty',
  'unitPrice',
  'total',
  'qtyReceived',
  'validity',
  'labels',
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Recebimento de Pedidos');
}

function doPost(e) {
  try {
    const parameterKey = (e && e.parameter && e.parameter.apiKey) || '';

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: 'Payload vazio.' });
    }

    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return jsonResponse_({ ok: false, error: 'Payload JSON inválido.' });
    }

    const apiKey = parameterKey || (payload && payload.apiKey);

    if (!apiKey || apiKey !== API_KEY) {
      return jsonResponse_({ ok: false, error: 'API key inválida.' });
    }

    const validationError = validatePayload_(payload);
    if (validationError) {
      return jsonResponse_({ ok: false, error: validationError });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const spreadsheet = getSpreadsheet_();
      const ordersSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ORDERS, ORDER_HEADERS);
      const itemsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ITEMS, ITEM_HEADERS);

      const existingRow = findOrderRowByOc_(ordersSheet, payload.oc);
      if (existingRow) {
        return jsonResponse_({ ok: true, duplicate: true });
      }

      const orderRow = buildOrderRow_(payload);
      ordersSheet.appendRow(orderRow);

      const itemRows = buildItemRows_(payload.oc, payload.items);
      if (itemRows.length) {
        itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemRows.length, itemRows[0].length)
          .setValues(itemRows);
      }

      return jsonResponse_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || 'Erro inesperado.' });
  }
}

function getOrdersWithItems() {
  const spreadsheet = getSpreadsheet_();
  const ordersSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ORDERS, ORDER_HEADERS);
  const itemsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ITEMS, ITEM_HEADERS);

  const orders = readOrders_(ordersSheet);
  const itemsByOc = readItemsByOc_(itemsSheet);

  const enriched = orders.map((order) => {
    return Object.assign({}, order, {
      items: itemsByOc[order.oc] || [],
    });
  });

  return enriched.sort((a, b) => {
    const dateA = a.sentAt ? new Date(a.sentAt).getTime() : 0;
    const dateB = b.sentAt ? new Date(b.sentAt).getTime() : 0;
    return dateB - dateA;
  });
}

function getBuyerOptions() {
  const spreadsheet = getSpreadsheet_();
  return getNamedRangeOptions_(spreadsheet, 'EMP_COMP');
}

function getSupplierOptions() {
  const spreadsheet = getSpreadsheet_();
  return getNamedRangeOptions_(spreadsheet, 'EMP_FORN');
}

function updateItemFields(oc, lineNo, qtyReceived, validity, labels) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const itemsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ITEMS, ITEM_HEADERS);
  const row = findItemRow_(itemsSheet, oc, lineNo);
  if (!row) {
    throw new Error('Item não encontrado.');
  }

  itemsSheet.getRange(row, 9, 1, 3).setValues([[
    qtyReceived,
    validity,
    labels,
  ]]);

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
    const ordersSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ORDERS, ORDER_HEADERS);
    const itemsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ITEMS, ITEM_HEADERS);

    const orderRow = findOrderRowByOc_(ordersSheet, oc);
    if (!orderRow) {
      throw new Error('Pedido não encontrado.');
    }

    const order = readOrders_(ordersSheet).find((entry) => String(entry.oc) === String(oc));
    if (!order) {
      throw new Error('Pedido não encontrado.');
    }

    const items = readItemsByOc_(itemsSheet)[oc] || [];
    if (!items.length) {
      throw new Error('Itens não encontrados.');
    }

    const invalid = items.some((item) => !isItemCompleteWithLabels_(item));
    if (invalid) {
      throw new Error('Preencha validade, etiquetas e quantidade recebida em todas as linhas.');
    }

    const receivedAt = new Date();
    const recRange = spreadsheet.getRangeByName('REC_POR_ITEM');
    if (!recRange) {
      throw new Error('Named range REC_POR_ITEM não encontrado.');
    }

    const recValues = recRange.getValues();
    const insertIndex = findNextEmptyIndex_(recValues);
    if (insertIndex === -1 || insertIndex + items.length > recValues.length) {
      throw new Error('Sem espaço disponível no REC_POR_ITEM para inserir todos os itens.');
    }

    const recRows = buildRecPorItemRows_(order, items, receivedAt);
    const startRow = recRange.getRow() + insertIndex;
    const startCol = recRange.getColumn();
    spreadsheet.getRange(startRow, startCol, recRows.length, recRows[0].length)
      .setValues(recRows);

    deleteLastPdf_();
    const pdfResult = generateLabelsPdf_(oc, order, items, width, height);

    sendToWebAppC_(order, items, pdfResult);

    ordersSheet.getRange(orderRow, 2).setValue('RECEBIDO');
    ordersSheet.getRange(orderRow, 10).setValue(receivedAt);

    return { ok: true, oc, pdfUrl: pdfResult.pdfUrl };
  } finally {
    lock.releaseLock();
  }
}

function markReceived(oc) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ORDERS, ORDER_HEADERS);
  const itemsSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ITEMS, ITEM_HEADERS);

  const items = readItemsByOc_(itemsSheet)[oc] || [];
  if (!items.length) {
    throw new Error('Itens não encontrados.');
  }

  const incomplete = items.some((item) => !isItemComplete_(item));
  if (incomplete) {
    throw new Error('Preencha todos os campos obrigatórios antes de receber.');
  }

  const row = findOrderRowByOc_(ordersSheet, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  ordersSheet.getRange(row, 2).setValue('RECEBIDO');
  ordersSheet.getRange(row, 10).setValue(new Date());

  return { ok: true };
}

function cancelOrderStub(oc) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ORDERS, ORDER_HEADERS);
  const row = findOrderRowByOc_(ordersSheet, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  ordersSheet.getRange(row, 2).setValue('CANCEL_PENDENTE');
  return { ok: true };
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

function getNamedRangeOptions_(spreadsheet, rangeName) {
  const range = spreadsheet.getRangeByName(rangeName);
  if (!range) {
    throw new Error('Named range não encontrado: ' + rangeName);
  }
  const values = range.getValues();
  const options = values.map((row) => row[0]).filter((value) => value !== null && value !== '');
  return options.map((value) => String(value));
}

function findNextEmptyIndex_(values) {
  let lastFilled = -1;
  for (let i = 0; i < values.length; i += 1) {
    if (values[i][0] !== null && values[i][0] !== '') {
      lastFilled = i;
    }
  }
  return lastFilled + 1;
}

function buildRecPorItemRows_(order, items, receivedAt) {
  const timestamp = formatDateTime_(receivedAt);
  return items.map((item) => [
    order.oc,
    order.buyerSelected || '',
    order.supplierSelected || '',
    timestamp,
    item.code || '',
    item.item || '',
    item.unit || '',
    item.qty || 0,
    Number(item.qtyReceived) || 0,
    item.unitPrice || 0,
    item.validity || '',
  ]);
}

function deleteLastPdf_() {
  const props = PropertiesService.getScriptProperties();
  const lastId = props.getProperty(LAST_LABELS_PDF_PROPERTY);
  if (!lastId) {
    return;
  }
  try {
    const file = DriveApp.getFileById(lastId);
    file.setTrashed(true);
  } catch (error) {
    // Ignore missing file errors.
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
  const pages = items.map((item) => {
    const labelsCount = Number(item.labels) || 0;
    const pageCount = Math.ceil(labelsCount / 3);
    const pagesHtml = [];
    for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
      const labelsHtml = [];
      for (let slot = 0; slot < 3; slot += 1) {
        const labelIndex = pageIndex * 3 + slot;
        if (labelIndex < labelsCount) {
          labelsHtml.push(`
            <div class="label">
              <div class="label-line"><strong>OC:</strong> ${escapeHtml_(order.oc)}</div>
              <div class="label-line"><strong>Fornecedor:</strong> ${escapeHtml_(order.supplierSelected || '')}</div>
              <div class="label-line"><strong>Item:</strong> ${escapeHtml_(item.item || '')}</div>
              <div class="label-line"><strong>Código:</strong> ${escapeHtml_(item.code || '')}</div>
              <div class="label-line"><strong>Validade:</strong> ${escapeHtml_(item.validity || '')}</div>
            </div>
          `);
        } else {
          labelsHtml.push('<div class="label label-empty"></div>');
        }
      }
      pagesHtml.push(`
        <div class="page">
          ${labelsHtml.join('')}
        </div>
      `);
    }
    return pagesHtml.join('');
  });

  return `
    <html>
      <head>
        <style>
          @page { margin: 0.5cm; }
          body { margin: 0; font-family: Arial, sans-serif; }
          .page {
            width: 100%;
            page-break-after: always;
            display: flex;
            flex-direction: column;
            gap: 0.5cm;
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
          .label-empty {
            border-color: transparent;
          }
          .label-line {
            line-height: 1.2;
          }
        </style>
      </head>
      <body>
        ${pages.join('')}
      </body>
    </html>
  `;
}

function sendToWebAppC_(order, items, pdfResult) {
  const payload = {
    apiVersion: 1,
    oc: order.oc,
    sentAt: new Date().toISOString(),
    buyerSelected: order.buyerSelected || '',
    supplierSelected: order.supplierSelected || '',
    items: items.map((item) => ({
      lineNo: item.lineNo,
      code: item.code || '',
      item: item.item || '',
      unit: item.unit || '',
      qty: item.qty || 0,
      unitPrice: item.unitPrice || 0,
      total: item.total || 0,
      qtyReceived: Number(item.qtyReceived) || 0,
      validity: item.validity || '',
      labels: Number(item.labels) || 0,
    })),
    pdfFileId: pdfResult.pdfFileId,
    pdfUrl: pdfResult.pdfUrl,
  };

  const response = UrlFetchApp.fetch(WEBAPP_C_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-API-KEY': WEBAPP_C_API_KEY,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  if (status >= 400) {
    throw new Error('Falha ao enviar para o WebApp C.');
  }

  let parsed;
  try {
    parsed = JSON.parse(response.getContentText());
  } catch (error) {
    return;
  }

  if (parsed && parsed.ok === false) {
    throw new Error('WebApp C retornou erro ao receber o pedido.');
  }
}

function formatDateTime_(value) {
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.appendRow(headers);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  return sheet;
}

function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    return 'Payload inválido.';
  }
  if (!payload.oc) {
    return 'OC é obrigatória.';
  }
  if (!Array.isArray(payload.items) || payload.items.length === 0) {
    return 'Itens são obrigatórios.';
  }

  const invalidItem = payload.items.find((item) => {
    return !isNumber_(item.qty) || !isNumber_(item.unitPrice) || !isNumber_(item.total);
  });
  if (invalidItem) {
    return 'Campos numéricos inválidos nos itens.';
  }

  return '';
}

function buildOrderRow_(payload) {
  const totals = payload.totals || {};
  const buyer = payload.buyer || {};
  const supplier = payload.supplier || {};
  const sentAt = payload.sentAt ? new Date(payload.sentAt) : new Date();

  return [
    payload.oc,
    'PENDENTE',
    sentAt,
    buyer.selected || '',
    supplier.selected || '',
    totals.qtyTotal || 0,
    totals.valueTotal || 0,
    JSON.stringify(buyer.details || []),
    JSON.stringify(supplier.details || []),
    '',
  ];
}

function buildItemRows_(oc, items) {
  return items.map((item, index) => {
    return [
      oc,
      index + 1,
      item.code || '',
      item.item || '',
      item.unit || '',
      item.qty || 0,
      item.unitPrice || 0,
      item.total || 0,
      '',
      '',
      '',
    ];
  });
}

function findOrderRowByOc_(sheet, oc) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0]) === String(oc)) {
      return i + 2;
    }
  }
  return null;
}

function findItemRow_(sheet, oc, lineNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return null;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i += 1) {
    const rowOc = String(values[i][0]);
    const rowLine = Number(values[i][1]);
    if (rowOc === String(oc) && rowLine === Number(lineNo)) {
      return i + 2;
    }
  }
  return null;
}

function readOrders_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, ORDER_HEADERS.length).getValues();
  return values.map((row) => {
    return {
      oc: row[0],
      status: row[1],
      sentAt: formatDateValue_(row[2]),
      buyerSelected: row[3],
      supplierSelected: row[4],
      qtyTotal: row[5],
      valueTotal: row[6],
      buyerDetails: parseJsonSafe_(row[7]),
      supplierDetails: parseJsonSafe_(row[8]),
      receivedAt: formatDateValue_(row[9]),
    };
  });
}

function readItemsByOc_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }

  const values = sheet.getRange(2, 1, lastRow - 1, ITEM_HEADERS.length).getValues();
  return values.reduce((acc, row) => {
    const oc = row[0];
    if (!acc[oc]) {
      acc[oc] = [];
    }
    acc[oc].push({
      oc,
      lineNo: row[1],
      code: row[2],
      item: row[3],
      unit: row[4],
      qty: row[5],
      unitPrice: row[6],
      total: row[7],
      qtyReceived: row[8],
      validity: row[9],
      labels: row[10],
    });
    return acc;
  }, {});
}

function isItemComplete_(item) {
  if (item.qtyReceived === '' || item.qtyReceived === null || item.qtyReceived === undefined) {
    return false;
  }
  const qtyReceived = Number(item.qtyReceived);
  if (!isNumber_(qtyReceived) || qtyReceived <= 0) {
    return false;
  }
  if (!item.validity || String(item.validity).trim() === '') {
    return false;
  }
  if (!item.labels || String(item.labels).trim() === '') {
    return false;
  }
  return true;
}

function isNumber_(value) {
  return typeof value === 'number' && !Number.isNaN(value);
}

function parseNumber_(value) {
  if (value === null || value === undefined || value === '') {
    return NaN;
  }
  if (typeof value === 'number') {
    return value;
  }
  const normalized = String(value).replace(',', '.');
  return Number(normalized);
}

function isPositiveNumber_(value) {
  return typeof value === 'number' && !Number.isNaN(value) && value > 0;
}

function isPositiveInteger_(value) {
  return Number.isInteger(value) && value > 0;
}

function isValidDateString_(value) {
  return typeof value === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(value.trim());
}

function isItemCompleteWithLabels_(item) {
  const qtyReceived = parseNumber_(item.qtyReceived);
  const labels = parseNumber_(item.labels);
  return isPositiveNumber_(qtyReceived)
    && isPositiveInteger_(labels)
    && isValidDateString_(String(item.validity || ''));
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
  return value;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
