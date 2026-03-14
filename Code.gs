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
  sentAt: 15, // P
  nf: 10, // K
  nfFrete: 11, // L
  comment: 12, // M
  groupOcs: 13, // N
  groupSentAts: 14, // O
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
    return HtmlService.createHtmlOutputFromFile('Print')
      .setTitle('Impressão de Etiquetas');
  }
  if (view === 'pdf') {
    const fileId = e && e.parameter && e.parameter.fileId;
    if (!fileId) {
      return HtmlService.createHtmlOutput('<p>Arquivo não informado.</p>');
    }
    try {
      const file = DriveApp.getFileById(fileId);
      const base64 = Utilities.base64Encode(file.getBlob().getBytes());
      const dataUrl = 'data:application/pdf;base64,' + base64;
      const safeDataUrl = dataUrl.replace(/"/g, '&quot;');
      const html = '<!doctype html><html><head><meta charset="utf-8"><title>PDF</title>'
        + '<style>html,body{margin:0;height:100%;}embed{width:100%;height:100%;}</style></head>'
        + '<body><embed src="' + safeDataUrl + '" type="application/pdf"></body></html>';
      return HtmlService.createHtmlOutput(html).setTitle('PDF');
    } catch (error) {
      return HtmlService.createHtmlOutput('<p>PDF não encontrado.</p>');
    }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Recebimento de Pedidos');
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

  const startCol = itemsRange.getColumn();
  const sheet = itemsRange.getSheet();
  sheet.getRange(row, startCol + ITEM_RANGE_MAP.qtyReceived, 1, 1).setValue(qtyReceived);
  sheet.getRange(row, startCol + ITEM_RANGE_MAP.validity, 1, 1).setValue(validity);
  sheet.getRange(row, startCol + ITEM_RANGE_MAP.labels, 1, 1).setValue(labels);

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


    const ordersSheet = ordersRange.getSheet();
    const ordersStartCol = ordersRange.getColumn();
    ordersSheet.getRange(orderRow, ordersStartCol + ORDER_RANGE_MAP.status).setValue('RECEBIDO');
    ordersSheet.getRange(orderRow, ordersStartCol + ORDER_RANGE_MAP.receivedAt).setValue(receivedAt);

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
  ordersSheet.getRange(row, ordersStartCol + ORDER_RANGE_MAP.status).setValue('RECEBIDO');
  ordersSheet.getRange(row, ordersStartCol + ORDER_RANGE_MAP.receivedAt).setValue(new Date());

  return { ok: true };
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
  ordersSheet.getRange(row, ordersStartCol + ORDER_RANGE_MAP.status).setValue('CANCEL_PENDENTE');
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
          <div class="label-line"><strong>Endereço:</strong> ${escapeHtml_(entry.obsItem)}</div>
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
      validity: formatDateOnly_(row[ITEM_RANGE_MAP.validity]),
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
  const receivedText = formatDateTime_(receivedAt);
  values.forEach((row, index) => {
    if (String(row[ITEM_RANGE_MAP.oc]) !== String(order.oc)) {
      return;
    }
    const rowNumber = startRow + index;
    sheet.getRange(rowNumber, startCol + ITEM_RANGE_MAP.buyerSelected).setValue(order.buyerSelected || '');
    sheet.getRange(rowNumber, startCol + ITEM_RANGE_MAP.supplierSelected).setValue(order.supplierSelected || '');
    sheet.getRange(rowNumber, startCol + ITEM_RANGE_MAP.receivedAt).setValue(receivedText);
  });
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

function formatDateOnly_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  if (!value) {
    return '';
  }
  return String(value);
}

