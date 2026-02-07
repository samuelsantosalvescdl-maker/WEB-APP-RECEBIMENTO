const API_KEY = '741852963';

const SHEET_NAMES = {
  ORDERS: 'CONT_FIN',
  ITEMS: 'REC_POR_ITEM',
};

const ORDER_COLUMN_LETTERS = {
  oc: 'A',
  status: 'G',
  sentAt: 'H',
  buyerSelected: 'B',
  supplierSelected: 'D',
  qtyTotal: 'E',
  valueTotal: 'F',
  nf: 'K',
  nfFrete: 'L',
  buyerDetailsJson: 'I',
  supplierDetailsJson: 'J',
  receivedAt: 'O',
};

const LEGACY_ORDER_DETAIL_COLUMNS = ['M', 'N'];

const ITEM_COLUMN_LETTERS = {
  oc: 'A',
  lineNo: 'M',
  code: 'E',
  item: 'F',
  unit: 'G',
  qty: 'H',
  unitPrice: 'J',
  total: 'L',
  qtyReceived: 'I',
  validity: 'K',
  labels: 'N',
  buyerSelected: 'B',
  supplierSelected: 'C',
  receivedAt: 'D',
};

const LEGACY_ITEM_BUYER_SUPPLIER_COLUMNS = {
  buyerSelected: 'L',
  supplierSelected: 'M',
};

const SPREADSHEET_ID = '1mc3nNSeW6GI2rXudQ30c2bzIlDtccheEdsTG85n_Y4g';
const LABELS_FOLDER_ID = '1UzyIn1fsiVIatfgQQK-GyGIeJI4Z-AFs';
const WEBAPP_C_URL = 'https://script.google.com/macros/s/AKfycbyohvLNZUxc1Kdyg0N5dr4lxgA9pXMbzEUwy2dLWF_P5IHfeEpyPnkjnKGAvOQfk1Y3/exec';
const WEBAPP_C_API_KEY = '369258147';
const LAST_LABELS_PDF_PROPERTY = 'LAST_LABELS_PDF_FILE_ID';

const ORDER_HEADERS = [
  'oc',
  'status',
  'sentAt',
  'buyerSelected',
  'supplierSelected',
  'qtyTotal',
  'valueTotal',
  'nf',
  'nfFrete',
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
  'buyerSelected',
  'supplierSelected',
  'receivedAt',
];

function doGet(e) {
  const view = e && e.parameter && e.parameter.view;
  if (view === 'print') {
    return HtmlService.createHtmlOutputFromFile('Print')
      .setTitle('Impressão de Etiquetas');
  }
  if (view === 'pdf') {
    const fileId = e && e.parameter && e.parameter.fileId;
    if (!fileId) {
      return ContentService.createTextOutput('Arquivo não informado.');
    }
    try {
      const file = DriveApp.getFileById(fileId);
      return file.getBlob().setContentType('application/pdf');
    } catch (error) {
      return ContentService.createTextOutput('PDF não encontrado.');
    }
  }
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
    const apiKeyNormalized = String(apiKey || '').trim();
    const expectedKey = String(API_KEY || '').trim();

    if (!apiKeyNormalized || apiKeyNormalized !== expectedKey) {
      return jsonResponse_({ ok: false, error: 'API key inválida.' });
    }

    const validationError = validatePayload_(payload);
    if (validationError) {
      Logger.log(`Validação falhou: ${validationError}`);
      return jsonResponse_({ ok: false, error: validationError });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const normalized = normalizeBuyerSupplierFromPayload_(payload);
      const spreadsheet = getSpreadsheet_();
      const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
      const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);

      const existingRow = findOrderRowByOcInRange_(ordersRange, payload.oc);
      if (existingRow) {
        Logger.log(`Pedido duplicado ignorado: ${payload.oc}`);
        return jsonResponse_({ ok: true, duplicate: true });
      }

      Logger.log(`Recebendo pedido: oc=${payload.oc} itens=${payload.items.length} comprador=${normalized.buyerSelected} fornecedor=${normalized.supplierSelected}`);
      const orderRow = buildOrderRow_(payload, normalized);
      writeOrderRowsToRange_(ordersRange, [orderRow]);

      const itemRows = buildItemRows_(payload.oc, payload.items, normalized.buyerSelected, normalized.supplierSelected);
      if (itemRows.length) {
        writeItemRowsToRange_(itemsRange, itemRows);
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
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);

  const orders = readOrdersFromRange_(ordersRange);
  const itemsByOc = readItemsByOcFromRange_(itemsRange);

  const enriched = orders.map((order) => {
    const items = itemsByOc[order.oc] || [];
    const buyerDetailFallback = order.buyerDetails && order.buyerDetails.length ? order.buyerDetails[0] : '';
    const supplierDetailFallback = order.supplierDetails && order.supplierDetails.length ? order.supplierDetails[0] : '';
    const fallbackBuyer = order.buyerSelected
      || (items[0] && items[0].buyerSelected)
      || buyerDetailFallback
      || '';
    const fallbackSupplier = order.supplierSelected
      || (items[0] && items[0].supplierSelected)
      || supplierDetailFallback
      || '';
    return Object.assign({}, order, {
      buyerSelected: fallbackBuyer,
      supplierSelected: fallbackSupplier,
      items,
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

function updateItemFields(oc, lineNo, qtyReceived, validity, labels, rowIndex) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);
  const row = rowIndex || findItemRowInRange_(itemsRange, oc, lineNo);
  if (!row) {
    throw new Error('Item não encontrado.');
  }

  const indexMap = getItemIndexMap_(itemsRange);
  const startCol = itemsRange.getColumn();
  const sheet = itemsRange.getSheet();
  sheet.getRange(row, startCol + indexMap.qtyReceived, 1, 1).setValue(qtyReceived);
  sheet.getRange(row, startCol + indexMap.validity, 1, 1).setValue(validity);
  sheet.getRange(row, startCol + indexMap.labels, 1, 1).setValue(labels);

  return { ok: true };
}

function updateOrderHeaderFields(oc, nfValue, nfFreteValue) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const indexMap = getOrderIndexMap_(ordersRange);
  const sheet = ordersRange.getSheet();
  const startCol = ordersRange.getColumn();
  sheet.getRange(row, startCol + indexMap.nf, 1, 1).setValue(nfValue || '');
  sheet.getRange(row, startCol + indexMap.nfFrete, 1, 1).setValue(nfFreteValue || '');

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
    if (String(order.status || '').trim().toLowerCase() === 'cancelado') {
      throw new Error('Pedido cancelado.');
    }

    const items = readItemsByOcFromRange_(itemsRange)[oc] || [];
    if (!items.length) {
      throw new Error('Itens não encontrados.');
    }

    const invalid = items.some((item) => !isItemCompleteWithLabels_(item));
    if (invalid) {
      throw new Error('Preencha validade, etiquetas e quantidade recebida em todas as linhas.');
    }

    const receivedAt = new Date();
    const recRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);
    updateReceiptFieldsInRange_(recRange, items, order, receivedAt);

    deleteLastPdf_();
    const pdfResult = generateLabelsPdf_(oc, order, items, width, height);

    sendToWebAppC_(order, items, pdfResult);

    const ordersSheet = ordersRange.getSheet();
    const ordersStartCol = ordersRange.getColumn();
    const orderIndexMap = getOrderIndexMap_(ordersRange);
    ordersSheet.getRange(orderRow, ordersStartCol + orderIndexMap.status).setValue('RECEBIDO');
    ordersSheet.getRange(orderRow, ordersStartCol + orderIndexMap.receivedAt).setValue(receivedAt);

    return { ok: true, oc, pdfUrl: pdfResult.pdfUrl, pdfFileId: pdfResult.pdfFileId };
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

  const items = readItemsByOcFromRange_(itemsRange)[oc] || [];
  if (!items.length) {
    throw new Error('Itens não encontrados.');
  }

  const incomplete = items.some((item) => !isItemComplete_(item));
  if (incomplete) {
    throw new Error('Preencha todos os campos obrigatórios antes de receber.');
  }

  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const ordersSheet = ordersRange.getSheet();
  const ordersStartCol = ordersRange.getColumn();
  const orderIndexMap = getOrderIndexMap_(ordersRange);
  ordersSheet.getRange(row, ordersStartCol + orderIndexMap.status).setValue('RECEBIDO');
  ordersSheet.getRange(row, ordersStartCol + orderIndexMap.receivedAt).setValue(new Date());

  return { ok: true };
}

function cancelOrderStub(oc) {
  if (!oc) {
    throw new Error('OC inválida.');
  }

  const spreadsheet = getSpreadsheet_();
  const ordersRange = getNamedRange_(spreadsheet, SHEET_NAMES.ORDERS);
  const itemsRange = getNamedRange_(spreadsheet, SHEET_NAMES.ITEMS);
  const row = findOrderRowByOcInRange_(ordersRange, oc);
  if (!row) {
    throw new Error('Pedido não encontrado.');
  }

  const ordersSheet = ordersRange.getSheet();
  const ordersStartCol = ordersRange.getColumn();
  const orderIndexMap = getOrderIndexMap_(ordersRange);
  ordersSheet.getRange(row, ordersStartCol + orderIndexMap.status).setValue('cancelado');

  markItemsCancelledInRange_(itemsRange, oc);
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

function getNamedRange_(spreadsheet, rangeName) {
  const range = spreadsheet.getRangeByName(rangeName);
  if (!range) {
    throw new Error('Named range não encontrado: ' + rangeName);
  }
  return range;
}

function getNamedRangeOptions_(spreadsheet, rangeName) {
  const range = getNamedRange_(spreadsheet, rangeName);
  const values = range.getValues();
  const options = values.map((row) => row[0]).filter((value) => value !== null && value !== '');
  return options.map((value) => String(value));
}

function getOrderIndexMap_(range) {
  return buildRangeIndexMap_(range, ORDER_COLUMN_LETTERS);
}

function getItemIndexMap_(range) {
  return buildRangeIndexMap_(range, ITEM_COLUMN_LETTERS);
}

function buildRangeIndexMap_(range, columnLettersByField) {
  const map = {};
  Object.keys(columnLettersByField).forEach((field) => {
    const absCol = colLetterToAbsIndex_(columnLettersByField[field]);
    map[field] = absToRelIndex0_(absCol, range);
  });
  return map;
}

function colLetterToAbsIndex_(letter) {
  const normalized = String(letter || '').trim().toUpperCase();
  let result = 0;
  for (let i = 0; i < normalized.length; i += 1) {
    const code = normalized.charCodeAt(i);
    if (code < 65 || code > 90) {
      throw new Error(`Coluna inválida: ${letter}`);
    }
    result = (result * 26) + (code - 64);
  }
  return result;
}

function absToRelIndex0_(absCol, range) {
  const rel0 = absCol - range.getColumn();
  if (rel0 < 0 || rel0 >= range.getNumColumns()) {
    throw new Error('Coluna fora do intervalo do named range.');
  }
  return rel0;
}

function findNextEmptyIndex_(values, colIndex) {
  const column = Number.isInteger(colIndex) ? colIndex : 0;
  for (let i = 0; i < values.length; i += 1) {
    const value = values[i][column];
    if (value === null || value === undefined || String(value).trim() === '') {
      return i;
    }
  }
  return values.length;
}

function buildRecPorItemRows_(order, items, receivedAt) {
  const timestamp = formatDateTime_(receivedAt);
  return items.map((item) => ({
    oc: order.oc,
    lineNo: item.lineNo || '',
    code: item.code || '',
    item: item.item || '',
    unit: item.unit || '',
    qty: item.qty || 0,
    unitPrice: item.unitPrice || 0,
    total: item.total || 0,
    qtyReceived: Number(item.qtyReceived) || 0,
    validity: item.validity || '',
    labels: Number(item.labels) || 0,
    buyerSelected: order.buyerSelected || '',
    supplierSelected: order.supplierSelected || '',
    receivedAt: timestamp,
  }));
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
          <div class="label-line"><strong>Validade:</strong> ${escapeHtml_(entry.validity)}</div>
        </div>
      `;
    });
    pages.push(`
      <div class="page">
        ${labelsHtml.join('')}
      </div>
    `);
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
    apiKey: WEBAPP_C_API_KEY,
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

  const url = WEBAPP_C_URL
    + (WEBAPP_C_URL.includes('?') ? '&' : '?')
    + 'apiKey=' + encodeURIComponent(WEBAPP_C_API_KEY);

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const text = response.getContentText();
  let parsed = null;

  try {
    parsed = JSON.parse(text);
  } catch (error) {
    parsed = null;
  }

  if (parsed && (parsed.ok === true || parsed.duplicate === true)) {
    return { ok: true, duplicate: parsed.duplicate === true };
  }

  if (parsed && parsed.ok === false) {
    throw new Error(buildWebAppCError_(status, parsed.error, text));
  }

  if (status >= 400) {
    if (parsed && parsed.error) {
      throw new Error(buildWebAppCError_(status, parsed.error, text));
    }
    throw new Error(buildWebAppCError_(status, null, text));
  }

  if (!parsed) {
    throw new Error(buildWebAppCError_(status, null, text));
  }
}

function testWebAppCConnection_() {
  const url = WEBAPP_C_URL
    + (WEBAPP_C_URL.includes('?') ? '&' : '?')
    + 'ping=1';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
  });
  const status = response.getResponseCode();
  const text = response.getContentText();
  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    parsed = null;
  }
  if (!parsed || parsed.ok !== true) {
    throw new Error(buildWebAppCError_(status, parsed && parsed.error, text));
  }
  const info = {
    scriptId: parsed.scriptId || '',
    expectedKeyLast3: parsed.expectedKeyLast3 || '',
  };
  Logger.log(`WebApp C OK: scriptId=${info.scriptId} keyLast3=${info.expectedKeyLast3}`);
  return info;
}

function buildWebAppCError_(status, message, bodyText) {
  const snippet = String(bodyText || '').slice(0, 200);
  const safeMessage = message ? String(message) : 'Erro ao enviar para o WebApp C.';
  return `WebApp C: ${safeMessage} | HTTP ${status} | respSnippet: ${snippet}`;
}

function formatDateTime_(value) {
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
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
      });
    }
  });
  return queue;
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
    return !isNumeric_(item.qty) || !isNumeric_(item.unitPrice) || !isNumeric_(item.total);
  });
  if (invalidItem) {
    return 'Campos numéricos inválidos nos itens.';
  }

  return '';
}

function buildOrderRow_(payload, normalized) {
  const totals = payload.totals || {};
  const sentAt = payload.sentAt ? new Date(payload.sentAt) : new Date();

  return {
    oc: payload.oc,
    status: 'PENDENTE',
    sentAt,
    buyerSelected: normalized.buyerSelected || '',
    supplierSelected: normalized.supplierSelected || '',
    qtyTotal: totals.qtyTotal || 0,
    valueTotal: totals.valueTotal || 0,
    nf: '',
    nfFrete: '',
    buyerDetailsJson: JSON.stringify(normalized.buyerDetails || []),
    supplierDetailsJson: JSON.stringify(normalized.supplierDetails || []),
    receivedAt: '',
  };
}

function buildItemRows_(oc, items, buyerSelected, supplierSelected) {
  return items.map((item, index) => ({
    oc,
    lineNo: index + 1,
    code: item.code || '',
    item: item.item || '',
    unit: item.unit || '',
    qty: item.qty || 0,
    unitPrice: item.unitPrice || 0,
    total: item.total || 0,
    qtyReceived: '',
    validity: '',
    labels: '',
    buyerSelected: buyerSelected || '',
    supplierSelected: supplierSelected || '',
    receivedAt: '',
  }));
}

function findOrderRowByOcInRange_(range, oc) {
  const values = range.getValues();
  const indexMap = getOrderIndexMap_(range);
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][indexMap.oc]) === String(oc)) {
      return range.getRow() + i;
    }
  }
  return null;
}

function findItemRowInRange_(range, oc, lineNo) {
  const values = range.getValues();
  const indexMap = getItemIndexMap_(range);
  for (let i = 0; i < values.length; i += 1) {
    const rowOc = String(values[i][indexMap.oc]);
    const rowLine = Number(values[i][indexMap.lineNo]);
    if (rowOc === String(oc) && rowLine === Number(lineNo)) {
      return range.getRow() + i;
    }
  }
  return null;
}

function readOrdersFromRange_(range) {
  const values = range.getValues();
  const indexMap = getOrderIndexMap_(range);
  const legacyDetails = getLegacyOrderDetailIndexes_(range, indexMap);
  return values
    .filter((row) => {
      const ocValue = row[indexMap.oc];
      return ocValue !== null && ocValue !== undefined && String(ocValue).trim() !== '';
    })
    .map((row) => ({
      oc: row[indexMap.oc],
      status: row[indexMap.status],
      sentAt: formatDateValue_(row[indexMap.sentAt]),
      buyerSelected: row[indexMap.buyerSelected],
      supplierSelected: row[indexMap.supplierSelected],
      qtyTotal: row[indexMap.qtyTotal],
      valueTotal: row[indexMap.valueTotal],
      nf: row[indexMap.nf],
      nfFrete: row[indexMap.nfFrete],
      buyerDetails: parseOrderDetails_(row, indexMap.buyerDetailsJson, legacyDetails[0]),
      supplierDetails: parseOrderDetails_(row, indexMap.supplierDetailsJson, legacyDetails[1]),
      receivedAt: formatDateValue_(row[indexMap.receivedAt]),
    }));
}

function readItemsByOcFromRange_(range) {
  const values = range.getValues();
  const indexMap = getItemIndexMap_(range);
  const legacyIndexMap = getLegacyItemBuyerSupplierIndexes_(range);
  return values.reduce((acc, row, index) => {
    const oc = row[indexMap.oc];
    if (!oc) {
      return acc;
    }
    if (!acc[oc]) {
      acc[oc] = [];
    }
    const rawBuyer = row[indexMap.buyerSelected];
    const rawSupplier = row[indexMap.supplierSelected];
    const legacyBuyer = legacyIndexMap.buyerSelected !== null ? row[legacyIndexMap.buyerSelected] : '';
    const legacySupplier = legacyIndexMap.supplierSelected !== null ? row[legacyIndexMap.supplierSelected] : '';
    const buyerSelected = rawBuyer || legacyBuyer || '';
    const supplierSelected = rawSupplier || legacySupplier || '';
    acc[oc].push({
      oc,
      lineNo: row[indexMap.lineNo],
      rowIndex: range.getRow() + index,
      code: row[indexMap.code],
      item: row[indexMap.item],
      unit: row[indexMap.unit],
      qty: row[indexMap.qty],
      unitPrice: row[indexMap.unitPrice],
      total: row[indexMap.total] || (row[indexMap.qty] * row[indexMap.unitPrice]),
      qtyReceived: row[indexMap.qtyReceived],
      validity: formatDateOnly_(row[indexMap.validity]),
      labels: row[indexMap.labels],
      buyerSelected,
      supplierSelected,
      receivedAt: row[indexMap.receivedAt],
    });
    return acc;
  }, {});
}

function writeOrderRowsToRange_(range, orderRows) {
  if (!orderRows.length) {
    return;
  }
  const values = range.getValues();
  const indexMap = getOrderIndexMap_(range);
  const legacyIndexes = getLegacyOrderDetailIndexes_(range, indexMap);
  const insertIndex = findNextEmptyIndex_(values, indexMap.oc);
  if (insertIndex === -1 || insertIndex + orderRows.length > values.length) {
    throw new Error('Sem espaço disponível no intervalo CONT_FIN para inserir pedidos.');
  }
  const output = orderRows.map((row) => mapOrderRowToRange_(row, values[0].length, indexMap, legacyIndexes));
  const sheet = range.getSheet();
  const startRow = range.getRow() + insertIndex;
  const startCol = range.getColumn();
  sheet.getRange(startRow, startCol, output.length, output[0].length).setValues(output);
}

function writeItemRowsToRange_(range, itemRows) {
  if (!itemRows.length) {
    return;
  }
  const values = range.getValues();
  const indexMap = getItemIndexMap_(range);
  const insertIndex = findNextEmptyIndex_(values, indexMap.oc);
  if (insertIndex === -1 || insertIndex + itemRows.length > values.length) {
    throw new Error('Sem espaço disponível no intervalo REC_POR_ITEM para inserir itens.');
  }
  const output = itemRows.map((row) => mapItemRowToRange_(row, values[0].length, indexMap));
  const sheet = range.getSheet();
  const startRow = range.getRow() + insertIndex;
  const startCol = range.getColumn();
  sheet.getRange(startRow, startCol, output.length, output[0].length).setValues(output);
}

function updateReceiptFieldsInRange_(range, items, order, receivedAt) {
  const sheet = range.getSheet();
  const startCol = range.getColumn();
  const startRow = range.getRow();
  const values = range.getValues();
  const indexMap = getItemIndexMap_(range);
  const receivedText = formatDateTime_(receivedAt);
  values.forEach((row, index) => {
    if (String(row[indexMap.oc]) !== String(order.oc)) {
      return;
    }
    const rowNumber = startRow + index;
    sheet.getRange(rowNumber, startCol + indexMap.buyerSelected).setValue(order.buyerSelected || '');
    sheet.getRange(rowNumber, startCol + indexMap.supplierSelected).setValue(order.supplierSelected || '');
    sheet.getRange(rowNumber, startCol + indexMap.receivedAt).setValue(receivedText);
  });
}

function markItemsCancelledInRange_(range, oc) {
  const sheet = range.getSheet();
  const startCol = range.getColumn();
  const startRow = range.getRow();
  const values = range.getValues();
  const indexMap = getItemIndexMap_(range);
  values.forEach((row, index) => {
    if (String(row[indexMap.oc]) !== String(oc)) {
      return;
    }
    const rowNumber = startRow + index;
    sheet.getRange(rowNumber, startCol + indexMap.receivedAt).setValue('cancelado');
  });
}

function mapOrderRowToRange_(row, totalCols, indexMap, legacyIndexes) {
  const output = Array(totalCols).fill('');
  output[indexMap.oc] = row.oc;
  output[indexMap.status] = row.status;
  output[indexMap.sentAt] = row.sentAt;
  output[indexMap.buyerSelected] = row.buyerSelected;
  output[indexMap.supplierSelected] = row.supplierSelected;
  output[indexMap.qtyTotal] = row.qtyTotal;
  output[indexMap.valueTotal] = row.valueTotal;
  output[indexMap.nf] = row.nf;
  output[indexMap.nfFrete] = row.nfFrete;
  output[indexMap.buyerDetailsJson] = row.buyerDetailsJson;
  output[indexMap.supplierDetailsJson] = row.supplierDetailsJson;
  output[indexMap.receivedAt] = row.receivedAt;
  clearLegacyOrderDetailColumns_(output, legacyIndexes);
  return output;
}

function getLegacyOrderDetailIndexes_(range, indexMap) {
  const usedIndexes = new Set(Object.values(indexMap));
  const legacyIndexes = [];
  LEGACY_ORDER_DETAIL_COLUMNS.forEach((letter) => {
    const absCol = colLetterToAbsIndex_(letter);
    try {
      const rel0 = absToRelIndex0_(absCol, range);
      if (!usedIndexes.has(rel0)) {
        legacyIndexes.push(rel0);
      }
    } catch (error) {
      // ignore columns outside the named range
    }
  });
  return legacyIndexes;
}

function clearLegacyOrderDetailColumns_(output, legacyIndexes) {
  if (!legacyIndexes || !legacyIndexes.length) {
    return;
  }
  legacyIndexes.forEach((index) => {
    output[index] = '';
  });
}

function getLegacyItemBuyerSupplierIndexes_(range) {
  const result = { buyerSelected: null, supplierSelected: null };
  Object.keys(LEGACY_ITEM_BUYER_SUPPLIER_COLUMNS).forEach((field) => {
    const absCol = colLetterToAbsIndex_(LEGACY_ITEM_BUYER_SUPPLIER_COLUMNS[field]);
    try {
      result[field] = absToRelIndex0_(absCol, range);
    } catch (error) {
      result[field] = null;
    }
  });
  return result;
}

function mapItemRowToRange_(row, totalCols, indexMap) {
  const output = Array(totalCols).fill('');
  output[indexMap.oc] = row.oc;
  output[indexMap.lineNo] = row.lineNo;
  output[indexMap.code] = row.code;
  output[indexMap.item] = row.item;
  output[indexMap.unit] = row.unit;
  output[indexMap.qty] = row.qty;
  output[indexMap.unitPrice] = row.unitPrice;
  output[indexMap.total] = row.total;
  output[indexMap.qtyReceived] = row.qtyReceived;
  output[indexMap.validity] = row.validity;
  output[indexMap.labels] = row.labels;
  output[indexMap.buyerSelected] = row.buyerSelected;
  output[indexMap.supplierSelected] = row.supplierSelected;
  output[indexMap.receivedAt] = row.receivedAt;
  return output;
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

function isNumeric_(value) {
  if (value === null || value === undefined || value === '') {
    return false;
  }
  const numberValue = Number(String(value).replace(',', '.'));
  return !Number.isNaN(numberValue);
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
  if (value instanceof Date) {
    return true;
  }
  if (typeof value !== 'string') {
    return false;
  }
  const normalized = value.trim().replace(/[-.]/g, '/');
  return /^\d{2}\/\d{2}\/\d{4}$/.test(normalized);
}

function isItemCompleteWithLabels_(item) {
  const qtyReceived = parseNumber_(item.qtyReceived);
  const labels = parseNumber_(item.labels);
  return isPositiveNumber_(qtyReceived)
    && isPositiveInteger_(labels)
    && isValidDateString_(item.validity || '');
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
    return normalizeDetailsArray_(JSON.parse(value));
  } catch (error) {
    return [String(value)];
  }
}

function normalizeDetailsArray_(parsed) {
  if (Array.isArray(parsed)) {
    return parsed;
  }
  if (parsed && typeof parsed === 'object') {
    return Object.values(parsed);
  }
  if (parsed === null || parsed === undefined || parsed === '') {
    return [];
  }
  return [String(parsed)];
}

function parseOrderDetails_(row, primaryIndex, legacyIndex) {
  const primaryValue = primaryIndex !== null && primaryIndex !== undefined ? row[primaryIndex] : '';
  if (primaryValue) {
    return parseJsonSafe_(primaryValue);
  }
  if (legacyIndex !== null && legacyIndex !== undefined) {
    const legacyValue = row[legacyIndex];
    return parseJsonSafe_(legacyValue);
  }
  return [];
}

function parseOrderDetails_(row, primaryIndex, legacyIndex) {
  const primaryValue = primaryIndex !== null && primaryIndex !== undefined ? row[primaryIndex] : '';
  if (primaryValue) {
    return parseJsonSafe_(primaryValue);
  }
  if (legacyIndex !== null && legacyIndex !== undefined) {
    const legacyValue = row[legacyIndex];
    return parseJsonSafe_(legacyValue);
  }
  return [];
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

function formatDateOnly_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  if (!value) {
    return '';
  }
  return String(value);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeBuyerSupplierFromPayload_(payload) {
  const buyerSelected =
    (payload && payload.buyer && payload.buyer.selected)
    || (payload && payload.buyerSelected)
    || '';
  const supplierSelected =
    (payload && payload.supplier && payload.supplier.selected)
    || (payload && payload.supplierSelected)
    || '';
  const buyerDetails =
    (payload && payload.buyer && payload.buyer.details)
    || (payload && payload.buyerDetails)
    || [];
  const supplierDetails =
    (payload && payload.supplier && payload.supplier.details)
    || (payload && payload.supplierDetails)
    || [];

  return {
    buyerSelected: String(buyerSelected || '').trim(),
    supplierSelected: String(supplierSelected || '').trim(),
    buyerDetails: Array.isArray(buyerDetails) ? buyerDetails : [],
    supplierDetails: Array.isArray(supplierDetails) ? supplierDetails : [],
  };
}
