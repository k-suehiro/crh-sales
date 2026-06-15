const SHEET_NAME = '課題管理';
const LIST_SHEET_NAME = 'リスト';
const EMPLOYEE_DB_SPREADSHEET_ID = '1CBhRtlFao3walyQvAPGo2Me17rVVeyMtWalv8FVdHKE';
const EMPLOYEE_BRANCH = '浜松事業所';

// 課題管理シートの列構成（社内ヒアリング用）
// A:No. B:登録日時 C:営業担当者名 D:所属部署 E:営業歴（年）
// F:役職 G:（未使用） H:担当エリア I:課題カテゴリ J:困っている内容
// K:発生頻度 L:深刻度 M:改善したい度合い N:いまの対処方法 O:月あたり影響時間
// P:ツールAI改善余地 Q:社内打開見込み R:メモ S:ステータス T:メールアドレス

const FIELD_TO_HEADER = {
  category: '課題カテゴリ',
  frequency: '発生頻度',
  severity: '深刻度',
  intent: '改善したい度合い',
  aiPossible: 'ツールAI改善余地',
  proposalPotential: '社内打開見込み',
  status: 'ステータス'
};

const TARGET_HEADER_TO_COL = {
  '課題カテゴリ': 9,
  '発生頻度': 11,
  '深刻度': 12,
  '改善したい度合い': 13,
  'ツールAI改善余地': 16,
  '社内打開見込み': 17,
  'ステータス': 19
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('浜松事業所 営業課題ヒアリング')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDropdownOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LIST_SHEET_NAME);

  if (!sheet) {
    throw new Error(`シート「${LIST_SHEET_NAME}」が見つかりません。`);
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0];

  const result = {};

  Object.keys(FIELD_TO_HEADER).forEach(fieldName => {
    const headerName = FIELD_TO_HEADER[fieldName];
    const colIndex = headers.indexOf(headerName);

    if (colIndex === -1) {
      result[fieldName] = [];
      return;
    }

    result[fieldName] = values
      .slice(1)
      .map(row => row[colIndex])
      .filter(value => value !== '' && value !== null);
  });

  return result;
}

function getHamamatsuEmployees() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'hamamatsu_employees_v2';
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const employees = loadHamamatsuEmployees_();
  cache.put(cacheKey, JSON.stringify(employees), 21600);
  return employees;
}

function loadHamamatsuEmployees_() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_DB_SPREADSHEET_ID);
  const sheet = ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  const headers = values[0].map(header => String(header).trim());
  const branchCol = headers.indexOf('branch');
  const kanjiCol = headers.indexOf('kanji');
  const mailCol = headers.indexOf('mail');
  const divisionCol = headers.indexOf('division');
  const locationCol = headers.indexOf('location');
  const positionCol = headers.indexOf('position');

  if (branchCol === -1 || kanjiCol === -1) {
    throw new Error('社員DBに branch または kanji 列が見つかりません。');
  }

  const kanaCol = findEmployeeKanaColumnIndex_(headers);

  return values
    .slice(1)
    .filter(row => String(row[branchCol]).trim() === EMPLOYEE_BRANCH)
    .map(row => ({
      kanji: String(row[kanjiCol] || '').trim(),
      mail: mailCol !== -1 ? String(row[mailCol] || '').trim() : '',
      kana: kanaCol !== -1 ? String(row[kanaCol] || '').trim() : '',
      division: divisionCol !== -1 ? String(row[divisionCol] || '').trim() : '',
      location: locationCol !== -1 ? String(row[locationCol] || '').trim() : '',
      position: positionCol !== -1 ? String(row[positionCol] || '').trim() : ''
    }))
    .filter(employee => employee.kanji !== '')
    .sort((a, b) => a.kanji.localeCompare(b.kanji, 'ja'));
}

function findEmployeeKanaColumnIndex_(headers) {
  const kanaColumnNames = ['kana', 'hiragana', 'hira', 'yomi', 'reading', 'furigana', 'name_kana', 'kana_name', 'かな', 'ひらがな'];

  for (let i = 0; i < kanaColumnNames.length; i++) {
    const colIndex = headers.indexOf(kanaColumnNames[i]);
    if (colIndex !== -1) {
      return colIndex;
    }
  }

  return -1;
}

function submitForm(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);
  }

  const nextNo = getNextNo_(sheet);

  const row = [
    nextNo,
    new Date(),
    data.companyName || '',
    data.industry || '',
    data.employeeCount || '',
    data.position || '',
    '',
    data.area || '',
    data.category || '',
    data.issue || '',
    data.frequency || '',
    data.severity || '',
    data.intent || '',
    data.currentMethod || '',
    data.lossHours || '',
    data.aiPossible || '',
    data.proposalPotential || '',
    data.memo || '',
    data.status || '記録済',
    data.salesEmail || ''
  ];

  sheet.appendRow(row);

  const newRowNumber = sheet.getLastRow();
  applyDropdownsToRowFromList_(sheet, newRowNumber);

  return {
    success: true,
    message: '登録しました',
    no: nextNo
  };
}

function applyDropdownsToRowFromList_(targetSheet, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName(LIST_SHEET_NAME);

  if (!listSheet) {
    throw new Error(`シート「${LIST_SHEET_NAME}」が見つかりません。`);
  }

  const listValues = listSheet.getDataRange().getValues();
  const listHeaders = listValues[0];

  Object.keys(TARGET_HEADER_TO_COL).forEach(headerName => {
    const listColIndex = listHeaders.indexOf(headerName);

    if (listColIndex === -1) return;

    const listColumn = listColIndex + 1;
    const lastRow = getLastDataRowInColumn_(listSheet, listColumn);

    if (lastRow < 2) return;

    const sourceRange = listSheet.getRange(2, listColumn, lastRow - 1, 1);

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(false)
      .build();

    targetSheet
      .getRange(rowNumber, TARGET_HEADER_TO_COL[headerName])
      .setDataValidation(rule);
  });
}

function getLastDataRowInColumn_(sheet, column) {
  const values = sheet
    .getRange(1, column, sheet.getLastRow(), 1)
    .getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) {
      return i + 1;
    }
  }

  return 1;
}

function getNextNo_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return 1;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  const numbers = values
    .map(value => Number(value))
    .filter(value => !isNaN(value));

  return numbers.length > 0 ? Math.max(...numbers) + 1 : 1;
}


















// LPの問い合わせ
const ENTRY_SHEET_NAME = '協力申込';

const ENTRY_HEADERS = [
  'No.',
  '登録日時',
  'お名前',
  '会社名・屋号',
  'メールアドレス',
  '電話番号',
  '業種',
  '地域',
  '希望方法',
  '相談内容',
  'ステータス',
  'メモ'
];

function submitEntry(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRY_SHEET_NAME);

  if (!sheet) {
    throw new Error(`シート「${ENTRY_SHEET_NAME}」が見つかりません。`);
  }

  setupEntryHeaderIfNeeded_(sheet);

  const nextNo = getNextNo_(sheet);

  const row = [
    nextNo,
    new Date(),
    data.name || '',
    data.company || '',
    data.email || '',
    data.phone || '',
    data.industry || '',
    data.area || '',
    data.method || '',
    data.message || '',
    '未対応',
    ''
  ];

  sheet.appendRow(row);

  return {
    success: true,
    message: '送信しました',
    no: nextNo
  };
}

function setupEntryHeaderIfNeeded_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, ENTRY_HEADERS.length).getValues()[0];
  const hasHeader = firstRow.some(value => value !== '');

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, ENTRY_HEADERS.length).setValues([ENTRY_HEADERS]);
    sheet.getRange(1, 1, 1, ENTRY_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#111827')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}