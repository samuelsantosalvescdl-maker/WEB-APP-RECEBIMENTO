const API_KEY = '741852963';

const SHEET_NAMES = {
  ORDERS: 'ORDERS',
  ITEMS: 'ITEMS',
};

const SPREADSHEET_ID = '1mc3nNSeW6GI2rXudQ30c2bzIlDtccheEdsTG85n_Y4g';

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
  if (SPREADSHEET_ID && SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE') {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (error) {
      throw new Error('Falha ao abrir a planilha pelo ID configurado. Verifique SPREADSHEET_ID.');
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (spreadsheet) {
    return spreadsheet;
  }

  throw new Error('Nenhuma planilha ativa associada ao projeto. Configure SPREADSHEET_ID.');
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
