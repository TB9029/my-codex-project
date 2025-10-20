/**
 * 在庫チェック（GS1スキャン）システム
 * Google Apps Script バックエンド実装
 */

const SPREADSHEET_ID_KEY = 'INVENTORY_SPREADSHEET_ID';
const TIMEZONE = 'Asia/Tokyo';
const MASTER_SHEET_NAME = 'master';
const SCAN_SHEET_NAME = 'scans';
const CONSTANTS_SHEET_NAME = 'constants';
const STATUS_EXPIRED = '🚨';
const STATUS_LT1M = '⚠️';
const STATUS_LT2M = '✅';
const STATUS_OK = 'OK';

const MASTER_HEADERS = [
  'id',
  '商品名',
  'GS1コード',
  '予備GS1コード1',
  '予備GS1コード2',
  '定数',
  '単位',
  'QTY',
  'タグ',
  '作成日',
  '更新日',
  '有効'
];

const SCAN_HEADERS = [
  'timestamp',
  'raw',
  'gtin',
  'expiry',
  'lot',
  'serial',
  'マスタid',
  '判定',
  '備考',
  'ユーザー'
];

const CONSTANTS_HEADERS = [
  'id',
  '商品名',
  '単位',
  'QTY',
  '定数',
  '1月',
  '2月',
  '3月',
  '4月',
  '5月',
  '6月',
  '7月',
  '8月',
  '9月',
  '10月',
  '11月',
  '12月',
  '最終更新'
];

const MONTH_NAMES = [
  '1月',
  '2月',
  '3月',
  '4月',
  '5月',
  '6月',
  '7月',
  '8月',
  '9月',
  '10月',
  '11月',
  '12月'
];

function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appVersion = 'v1.0.0';
  return template
    .evaluate()
    .setTitle('在庫チェック')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setSpreadsheetId(id) {
  if (!id) {
    throw new Error('スプレッドシートIDが指定されていません。');
  }
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_KEY, id);
  return { success: true, id: id };
}

function getSpreadsheetId() {
  return getSpreadsheetId_();
}

function getSpreadsheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty(
    SPREADSHEET_ID_KEY
  );
  if (!id) {
    throw new Error(
      'スクリプトプロパティにスプレッドシートIDが設定されていません。setSpreadsheetId(id) を実行してください。'
    );
  }
  return id;
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(getSpreadsheetId_());
}

function getSheet_(sheetName, headers) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`スプレッドシートにシート「${sheetName}」が見つかりません。`);
  }
  ensureHeaders_(sheet, headers);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsUpdate = headers.some((header, idx) => firstRow[idx] !== header);
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function formatDate_(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function formatDateShort_(date) {
  if (!date) return '';
  return Utilities.formatDate(new Date(date), TIMEZONE, 'yyyy-MM-dd');
}

function parseNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const num = Number(value);
  if (Number.isNaN(num)) {
    throw new Error(`数値に変換できません: ${value}`);
  }
  return num;
}

function loadMasters_() {
  const sheet = getSheet_(MASTER_SHEET_NAME, MASTER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { records: [], byGtin: {} };
  }
  const values = sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).getValues();
  const records = values
    .filter(row => row[0])
    .map(row => ({
      id: row[0],
      name: row[1],
      gs1: row[2],
      alt1: row[3],
      alt2: row[4],
      minimum: Number(row[5]) || 0,
      unit: row[6] || '個',
      qty: Number(row[7]) || 1,
      tags: (row[8] || '')
        .split(',')
        .map(t => t.trim())
        .filter(Boolean),
      createdAt: row[9],
      updatedAt: row[10],
      active: row[11] === true || row[11] === 'TRUE' || row[11] === 'true'
    }));

  const byGtin = {};
  records.forEach(record => {
    [record.gs1, record.alt1, record.alt2]
      .filter(Boolean)
      .forEach(code => {
        byGtin[code] = record;
      });
  });
  return { records, byGtin };
}

function loadScans_(limit) {
  const sheet = getSheet_(SCAN_SHEET_NAME, SCAN_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const startRow = Math.max(2, lastRow - (limit || 50) + 1);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, SCAN_HEADERS.length).getValues();
  return values
    .filter(row => row[0])
    .map(row => ({
      timestamp: row[0],
      raw: row[1],
      gtin: row[2],
      expiry: row[3],
      lot: row[4],
      serial: row[5],
      masterId: row[6],
      status: row[7],
      notes: row[8],
      user: row[9]
    }))
    .reverse();
}

function loadConstants_() {
  const sheet = getSheet_(CONSTANTS_SHEET_NAME, CONSTANTS_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  const values = sheet
    .getRange(2, 1, lastRow - 1, CONSTANTS_HEADERS.length)
    .getValues();
  return values
    .filter(row => row[0])
    .map(row => {
      const months = {};
      MONTH_NAMES.forEach((monthName, idx) => {
        months[monthName] = Number(row[5 + idx]) || 0;
      });
      return {
        id: row[0],
        name: row[1],
        unit: row[2],
        qty: Number(row[3]) || 1,
        minimum: Number(row[4]) || 0,
        months: months,
        updatedAt: row[17]
      };
    });
}

function getDashboardSnapshot() {
  const masters = loadMasters_();
  const scans = loadScans_(100);
  const counters = {
    expired: 0,
    lt1m: 0,
    lt2m: 0
  };
  const today = today_();
  scans.forEach(entry => {
    if (!entry.expiry) return;
    const status = determineExpiryStatus_(entry.expiry, today);
    if (status.code === 'expired') counters.expired += 1;
    if (status.code === 'lt1m') counters.lt1m += 1;
    if (status.code === 'lt2m') counters.lt2m += 1;
  });

  const searchIndex = masters.records.map(record => ({
    id: record.id,
    name: record.name,
    tags: record.tags,
    gs1: [record.gs1, record.alt1, record.alt2].filter(Boolean)
  }));

  return {
    nearExpiry: {
      expired: counters.expired,
      lt1m: counters.lt1m,
      lt2m: counters.lt2m
    },
    latestScans: scans.slice(0, 10),
    searchIndex: searchIndex
  };
}

function searchMaster(payload) {
  const query = (payload && payload.query) ? payload.query.trim() : '';
  const tag = payload && payload.tag ? payload.tag.trim() : '';
  const { records } = loadMasters_();
  const normalizedQuery = query.toLowerCase();
  const result = records.filter(record => {
    if (tag && !record.tags.includes(tag)) {
      return false;
    }
    if (!normalizedQuery) {
      return true;
    }
    const combined = [
      record.name,
      record.gs1,
      record.alt1,
      record.alt2,
      record.tags.join(',')
    ]
      .filter(Boolean)
      .join(' ')
      .toLowerCase();
    return combined.includes(normalizedQuery);
  });
  return { items: result };
}

function createMaster(payload) {
  validateMasterPayload_(payload, false);
  const sheet = getSheet_(MASTER_SHEET_NAME, MASTER_HEADERS);
  const id = Utilities.getUuid();
  const now = new Date();
  const record = [
    id,
    payload.name,
    payload.gs1,
    payload.alt1 || '',
    payload.alt2 || '',
    Number(payload.minimum) || 0,
    payload.unit,
    Number(payload.qty) || 1,
    (payload.tags || []).join(', '),
    formatDate_(now),
    formatDate_(now),
    payload.active === false ? false : true
  ];
  sheet.appendRow(record);
  ensureConstantsRow_(id, payload);
  return { id: id };
}

function updateMaster(payload) {
  validateMasterPayload_(payload, true);
  const sheet = getSheet_(MASTER_SHEET_NAME, MASTER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('マスタが登録されていません。');
  }
  const data = sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).getValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) {
    throw new Error('指定されたIDのマスタが見つかりません。');
  }
  const now = new Date();
  const tags = (payload.tags || []).join(', ');
  const record = [
    payload.id,
    payload.name,
    payload.gs1,
    payload.alt1 || '',
    payload.alt2 || '',
    Number(payload.minimum) || 0,
    payload.unit,
    Number(payload.qty) || 1,
    tags,
    data[rowIndex - 2][9] || formatDate_(now),
    formatDate_(now),
    payload.active === false ? false : true
  ];
  sheet.getRange(rowIndex, 1, 1, MASTER_HEADERS.length).setValues([record]);
  ensureConstantsRow_(payload.id, payload);
  return { id: payload.id };
}

function validateMasterPayload_(payload, requireId) {
  if (!payload) {
    throw new Error('マスタ情報が渡されていません。');
  }
  if (requireId && !payload.id) {
    throw new Error('ID が指定されていません。');
  }
  if (!payload.name) {
    throw new Error('商品名は必須です。');
  }
  if (!payload.gs1) {
    throw new Error('GS1コードは必須です。');
  }
  ['minimum', 'qty'].forEach(field => {
    if (payload[field] === undefined || payload[field] === null || payload[field] === '') {
      throw new Error(`${field} は必須です。`);
    }
    if (Number.isNaN(Number(payload[field]))) {
      throw new Error(`${field} には数値を指定してください。`);
    }
  });
  if (!payload.unit) {
    throw new Error('単位を選択してください。');
  }
}

function ensureConstantsRow_(id, payload) {
  const sheet = getSheet_(CONSTANTS_SHEET_NAME, CONSTANTS_HEADERS);
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, CONSTANTS_HEADERS.length).getValues()
    : [];
  let rowIndex = -1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === id) {
      rowIndex = i + 2;
      break;
    }
  }
  const now = formatDate_(new Date());
  if (rowIndex === -1) {
    const months = MONTH_NAMES.map(() => 0);
    const record = [
      id,
      payload.name,
      payload.unit,
      Number(payload.qty) || 1,
      Number(payload.minimum) || 0,
      ...months,
      now
    ];
    sheet.appendRow(record);
  } else {
    const range = sheet.getRange(rowIndex, 1, 1, CONSTANTS_HEADERS.length);
    const current = range.getValues()[0];
    current[1] = payload.name;
    current[2] = payload.unit;
    current[3] = Number(payload.qty) || 1;
    current[4] = Number(payload.minimum) || 0;
    current[17] = now;
    range.setValues([current]);
  }
}

function deleteMaster(payload) {
  if (!payload || !payload.id) {
    throw new Error('削除対象のIDが指定されていません。');
  }
  const sheet = getSheet_(MASTER_SHEET_NAME, MASTER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('マスタが存在しません。');
  }
  const data = sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  throw new Error('指定されたマスタが存在しません。');
}

function importMastersCsv(csvText) {
  if (!csvText) {
    throw new Error('CSVテキストが空です。');
  }
  const rows = Utilities.parseCsv(csvText, ',');
  if (!rows || rows.length === 0) {
    throw new Error('CSVの内容を読み取れませんでした。');
  }
  const header = rows[0];
  const headerText = header.join('');
  MASTER_HEADERS.forEach((headerLabel, index) => {
    if (!headerText.includes(headerLabel) && header[index] !== headerLabel) {
      throw new Error('CSVヘッダが想定と異なります。');
    }
  });
  const sheet = getSheet_(MASTER_SHEET_NAME, MASTER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).clearContent();
  }
  const payload = rows.slice(1).filter(row => row[0]);
  if (payload.length) {
    sheet.getRange(2, 1, payload.length, MASTER_HEADERS.length).setValues(payload);
  }
  return { imported: payload.length };
}

function exportMastersCsv() {
  const sheet = getSheet_(MASTER_SHEET_NAME, MASTER_HEADERS);
  const range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), MASTER_HEADERS.length);
  const values = range.getValues();
  const csv = values.map(row => row.map(value => value === null ? '' : value).join(',')).join('\n');
  return { csv: csv };
}

function listMasters() {
  const { records } = loadMasters_();
  return { items: records };
}

function recordScan(payload) {
  if (!payload || !payload.raw) {
    throw new Error('読み取りデータが空です。');
  }
  const now = new Date();
  const sheet = getSheet_(SCAN_SHEET_NAME, SCAN_HEADERS);
  let parsed;
  try {
    parsed = parseGs1_(payload.raw);
  } catch (error) {
    sheet.appendRow([
      formatDate_(now),
      payload.raw,
      '',
      '',
      '',
      '',
      '',
      STATUS_EXPIRED,
      error.message,
      Session.getActiveUser().getEmail()
    ]);
    throw error;
  }
  const masters = loadMasters_();
  let matched = null;
  if (parsed.gtin) {
    matched = masters.byGtin[parsed.gtin] || null;
  }
  const status = determineExpiryStatus_(parsed.expiryIso);
  const row = [
    formatDate_(now),
    payload.raw,
    parsed.gtin || '',
    parsed.expiryIso || '',
    parsed.lot || '',
    parsed.serial || '',
    matched ? matched.id : '',
    status.icon,
    matched ? '' : 'マスタ未登録',
    Session.getActiveUser().getEmail()
  ];
  sheet.appendRow(row);
  return {
    parsed: parsed,
    status: status,
    master: matched,
    timestamp: formatDate_(now)
  };
}

function determineExpiryStatus_(expiryIso, today) {
  if (!expiryIso) {
    return {
      code: 'no-data',
      label: '期限情報なし',
      icon: 'ℹ️',
      color: 'neutral'
    };
  }
  const baseDate = today || today_();
  const expiryDate = new Date(expiryIso + 'T00:00:00');
  const diffMs = expiryDate.getTime() - baseDate.getTime();
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  if (diffDays < 0) {
    return {
      code: 'expired',
      label: '期限切れ',
      icon: STATUS_EXPIRED,
      color: 'danger'
    };
  }
  if (diffDays < 30) {
    return {
      code: 'lt1m',
      label: '1か月未満',
      icon: STATUS_LT1M,
      color: 'warning'
    };
  }
  if (diffDays < 60) {
    return {
      code: 'lt2m',
      label: '2か月未満',
      icon: STATUS_LT2M,
      color: 'caution'
    };
  }
  return {
    code: 'ok',
    label: '十分な期限',
    icon: STATUS_OK,
    color: 'success'
  };
}

function today_() {
  const now = new Date();
  const tokyo = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  return new Date(tokyo + 'T00:00:00');
}

function parseGs1_(raw) {
  if (!raw) {
    throw new Error('GS1コードが空です。');
  }
  let data = String(raw).trim();
  const fnc1 = String.fromCharCode(29);
  const segments = [];
  if (data.includes('(')) {
    const regex = /\((\d{2})\)([^()]*)/g;
    let match;
    while ((match = regex.exec(data)) !== null) {
      segments.push({ ai: match[1], value: match[2] });
    }
  } else {
    const parts = data.split(fnc1);
    parts.forEach((part, index) => {
      if (!part) return;
      const ai = part.substring(0, 2);
      let value = part.substring(2);
      if (index === 0) {
        // 固定長フィールドを順番に処理
        let cursor = 0;
        while (cursor < part.length) {
          const currentAi = part.substring(cursor, cursor + 2);
          cursor += 2;
          if (currentAi === '01') {
            value = part.substring(cursor, cursor + 14);
            cursor += 14;
            segments.push({ ai: '01', value: value });
          } else if (currentAi === '17') {
            value = part.substring(cursor, cursor + 6);
            cursor += 6;
            segments.push({ ai: '17', value: value });
          } else if (currentAi === '10' || currentAi === '21') {
            value = part.substring(cursor);
            cursor = part.length;
            segments.push({ ai: currentAi, value: value });
          } else {
            // 未対応 AI
            break;
          }
        }
      } else {
        segments.push({ ai: ai, value: value });
      }
    });
  }

  const result = {
    raw: raw,
    gtin: '',
    expiryIso: '',
    lot: '',
    serial: ''
  };
  segments.forEach(segment => {
    switch (segment.ai) {
      case '01':
        if (segment.value.length !== 14) {
          throw new Error('GTIN(01) は14桁である必要があります。');
        }
        result.gtin = segment.value;
        break;
      case '17':
        result.expiryIso = convertExpiry_(segment.value);
        break;
      case '10':
        result.lot = segment.value;
        break;
      case '21':
        result.serial = segment.value;
        break;
      default:
        break;
    }
  });
  return result;
}

function convertExpiry_(value) {
  if (!/^\d{6}$/.test(value)) {
    throw new Error('(17) 期限は YYMMDD の6桁である必要があります。');
  }
  const yy = Number(value.substring(0, 2));
  const mm = Number(value.substring(2, 4));
  const dd = Number(value.substring(4, 6));
  if (mm < 1 || mm > 12) {
    throw new Error('(17) 期限の月が不正です。');
  }
  const year = 2000 + yy;
  const date = new Date(Date.UTC(year, mm - 1, dd));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== mm - 1 || date.getUTCDate() !== dd) {
    throw new Error('(17) 期限の日付が不正です。');
  }
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

function listConstants() {
  const constants = loadConstants_();
  const masters = loadMasters_();
  const masterMap = {};
  masters.records.forEach(record => {
    masterMap[record.id] = record;
  });
  const rows = constants.map(row => {
    const master = masterMap[row.id];
    const baseline = (row.minimum || 0) * (row.qty || 1);
    const currentMonthName = MONTH_NAMES[new Date().getMonth()];
    const currentValue = Number(row.months[currentMonthName]) || 0;
    let color = 'success';
    if (currentValue < baseline) {
      color = 'danger';
    } else if (currentValue === baseline) {
      color = 'warning';
    }
    return {
      id: row.id,
      name: row.name,
      unit: row.unit,
      qty: row.qty,
      minimum: row.minimum,
      months: row.months,
      updatedAt: row.updatedAt,
      baseline: baseline,
      currentMonth: currentMonthName,
      currentValue: currentValue,
      color: color,
      active: master ? master.active : true
    };
  });
  return { items: rows };
}

function updateConstant(payload) {
  if (!payload || !payload.id) {
    throw new Error('更新対象のIDが指定されていません。');
  }
  const sheet = getSheet_(CONSTANTS_SHEET_NAME, CONSTANTS_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    throw new Error('商品定数が未登録です。');
  }
  const data = sheet.getRange(2, 1, lastRow - 1, CONSTANTS_HEADERS.length).getValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === payload.id) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex === -1) {
    throw new Error('商品定数が見つかりません。');
  }
  const row = data[rowIndex - 2];
  if (payload.month && MONTH_NAMES.includes(payload.month)) {
    const monthIndex = MONTH_NAMES.indexOf(payload.month);
    row[5 + monthIndex] = Number(payload.value) || 0;
  }
  row[17] = formatDate_(new Date());
  sheet.getRange(rowIndex, 1, 1, CONSTANTS_HEADERS.length).setValues([row]);
  return { success: true };
}

function exportConstantsCsv() {
  const sheet = getSheet_(CONSTANTS_SHEET_NAME, CONSTANTS_HEADERS);
  const range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), CONSTANTS_HEADERS.length);
  const values = range.getValues();
  const csv = values.map(row => row.map(value => value === null ? '' : value).join(',')).join('\n');
  return { csv: csv };
}

function importConstantsCsv(csvText) {
  if (!csvText) {
    throw new Error('CSVテキストが空です。');
  }
  const rows = Utilities.parseCsv(csvText, ',');
  if (!rows || rows.length === 0) {
    throw new Error('CSVの内容を読み取れませんでした。');
  }
  const sheet = getSheet_(CONSTANTS_SHEET_NAME, CONSTANTS_HEADERS);
  const header = rows[0];
  if (header.length !== CONSTANTS_HEADERS.length) {
    throw new Error('CSVヘッダが一致しません。');
  }
  const payload = rows.slice(1).filter(row => row[0]);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, CONSTANTS_HEADERS.length).clearContent();
  }
  if (payload.length) {
    sheet.getRange(2, 1, payload.length, CONSTANTS_HEADERS.length).setValues(payload);
  }
  return { imported: payload.length };
}

function exportDeficitCsv() {
  const constants = listConstants().items;
  const deficit = constants.filter(row => row.color === 'danger');
  const header = ['id', '商品名', '基準', '現在値', '単位', '月'];
  const rows = deficit.map(item => [
    item.id,
    item.name,
    item.baseline,
    item.currentValue,
    item.unit,
    item.currentMonth
  ]);
  const csvRows = [header, ...rows];
  const csv = csvRows.map(row => row.join(',')).join('\n');
  return { csv: csv };
}

function appendDemoScans(count) {
  const masters = loadMasters_();
  const sheet = getSheet_(SCAN_SHEET_NAME, SCAN_HEADERS);
  const now = new Date();
  const rows = [];
  for (let i = 0; i < count; i++) {
    const master = masters.records[i % masters.records.length];
    if (!master) break;
    const expiry = new Date(now.getTime() + (i - 2) * 7 * 24 * 60 * 60 * 1000);
    const expiryIso = formatDateShort_(expiry);
    const status = determineExpiryStatus_(expiryIso);
    rows.push([
      formatDate_(new Date(now.getTime() - i * 60 * 60 * 1000)),
      `(01)${master.gs1}`,
      master.gs1,
      expiryIso,
      `LOT-${i + 1}`,
      `SER-${i + 1}`,
      master.id,
      status.icon,
      '',
      Session.getActiveUser().getEmail()
    ]);
  }
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SCAN_HEADERS.length).setValues(rows);
  }
  return { inserted: rows.length };
}

function getAppBootstrap() {
  return {
    dashboard: getDashboardSnapshot(),
    masters: listMasters(),
    constants: listConstants()
  };
}
