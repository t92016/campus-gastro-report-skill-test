/**
 * ============================================================
 * (skill測試)忠義國小 — 校園集體腸胃不適症狀通報系統
 * 完整 Apps Script 程式碼（單一檔案版）
 * ============================================================
 *
 * 使用方式：
 * 1. 開啟試算表 → 擴充功能 → Apps Script
 * 2. 把編輯器裡原本的內容「全部刪除」，貼上這整份程式碼
 * 3. 儲存（Ctrl+S）
 * 4. 依照下方「要執行的函式」清單，依序執行
 *
 * ★ createReportForm 只能執行一次，表單建立後請勿重複執行
 */


// ============================================================
// 【設定區】所有共用常數
// ============================================================

const SHEET_MAIN = '彙整總表';
const SHEET_DEID = '去識別化總覽';
const GRADE_LIST = ['幼兒園', '1年級', '2年級', '3年級', '4年級', '5年級', '6年級'];
const FORM_ID = '1vlhVcoqAOj9LJP2o3u6hvrXeTte2xBygkisEzdUfAGc';

const SCHOOL_LOGO_BASE64 = ''; // 測試用暫不提供 logo

// 紫色暖色系主題色
const THEME = {
  primary: '#9b7ebd',
  primaryDark: '#7b5ea7',
  primaryLight: '#c4b0db',
  bannerGradient: 'linear-gradient(135deg, #9b7ebd, #7b5ea7)',
  bg: '#f7f3fb',
  cardBg: '#fffdfd',
  altRow: '#faf7fd',
  border: '#d5c8e8',
  textMuted: '#8a7b96',
  badge: '#7b5ea7',
  buttonText: '#fff',
  secondaryBg: '#e8dff5',
  secondaryText: '#5a4a6a',
  accent: '#c0645c' // 保留：標題列底色
};

// ============================================================
// 彙整總表欄位「依標題文字動態尋找」
// ============================================================
const HEADER_NAMES = {
  TIMESTAMP: '時間戳記',
  GRADE: '年級',
  CLASS: '班級',
  SEAT: '座號',
  NAME: '姓名',
  GENDER: '性別',
  CONTACT: '家人稱謂及緊急聯絡電話',
  SYMPTOM: '身體症狀',
  HEIGHT: '身高',
  WEIGHT: '體重',
  CASE_ID: '案件編號',
  SORT_KEY: '排序值',
  TRIAGE: '檢傷初判',
  LOCATION: '目前所在位置',
  HOSPITAL: '就醫醫院',
  ESCORT: '護送教師',
  NOTE: '備註',
  STATUS: '狀態',
  CLOSED: '結案',
  UPDATED_AT: '最後更新時間',
  UPDATED_BY: '最後更新人員'
};

const REQUIRED_HEADER_KEYS = [
  'GRADE', 'CLASS', 'SEAT', 'NAME',
  'CASE_ID', 'SORT_KEY', 'STATUS', 'CLOSED', 'UPDATED_AT', 'UPDATED_BY'
];

function getMainColMap_(sheet, throwOnMissing) {
  if (throwOnMissing === undefined) throwOnMissing = true;
  const lastCol = sheet.getLastColumn();
  const headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const textToCol = {};
  headerRow.forEach(function (text, idx) {
    const t = String(text || '').trim();
    if (t) textToCol[t] = idx + 1;
  });
  const map = {};
  Object.keys(HEADER_NAMES).forEach(function (key) {
    const headerText = HEADER_NAMES[key];
    map[key] = textToCol[headerText] || null;
  });
  map.LAST_COL = lastCol;
  if (throwOnMissing) {
    const missing = REQUIRED_HEADER_KEYS.filter(function (key) { return !map[key]; });
    if (missing.length) {
      const missingText = missing.map(function (k) { return HEADER_NAMES[k]; }).join('、');
      throw new Error(
        '「' + sheet.getName() + '」找不到必要欄位標題：[' + missingText + ']。' +
        '可能是 Google 表單或試算表標題列被誤改/誤刪，請檢查標題列文字是否跟系統預期的一致。'
      );
    }
  }
  return map;
}

// 選項清單
const OPTIONS = {
  TRIAGE: ['休息觀察區', '送醫區'],
  LOCATION: [
    '休息觀察區(1F國際文教中心)', '送醫區(1F楊陳包藝文中心)',
    '已回教室', '回檢傷組重新判斷', '家長接回'
  ],
  HOSPITAL: [
    '家長接回', '蘆洲世足盃醫院', '蘆洲衛生所', '新北市立聯合醫院(三重)', '三重衛生所',
    '三重中興醫院', '三重宏仁醫院', '淡水馬偕醫院', '士林新光醫院',
    '中山馬偕醫院', '大同中興醫院', '中正臺大醫院', '北投振興醫院',
    '北投榮民醫院', '新莊新仁醫院', '新莊臺北醫院', '新莊新泰醫院',
    '林口長庚醫院', '松山長庚醫院', '已回教室'
  ]
};

const GRADE_SHEET_FIELDS = [
  'CASE_ID', 'CLASS', 'SEAT', 'NAME', 'GENDER', 'CONTACT', 'SYMPTOM',
  'TRIAGE', 'LOCATION', 'HOSPITAL', 'ESCORT', 'NOTE',
  'STATUS', 'CLOSED', 'UPDATED_AT', 'UPDATED_BY', 'SORT_KEY'
];

const GRADE_EDITABLE_FIELDS = [
  'CLASS', 'SEAT', 'NAME', 'GENDER', 'CONTACT', 'SYMPTOM',
  'TRIAGE', 'LOCATION', 'HOSPITAL', 'ESCORT', 'NOTE'
];

// 三組通報設定（紫色暖色系配色）
const GROUP_CONFIG = {
  triage: {
    label: '檢傷組', propKey: 'PWD_TRIAGE',
    themeColor: '#c495a8', themeColorLight: '#f7edf1',
    updateFieldLabel: '檢傷初判', options: OPTIONS.TRIAGE
  },
  rest: {
    label: '休息組', propKey: 'PWD_REST',
    themeColor: '#a0b0c4', themeColorLight: '#eceff5',
    updateFieldLabel: '目前所在位置', options: OPTIONS.LOCATION
  },
  hospital: {
    label: '送醫組', propKey: 'PWD_HOSPITAL',
    themeColor: '#b8a47e', themeColorLight: '#efe9db',
    updateFieldLabel: '就醫醫院', options: OPTIONS.HOSPITAL
  }
};

const ADMIN_PROP_KEY = 'PWD_ADMIN';

function gradeToNumber_(gradeText) {
  const map = {
    '幼兒園': 0, '1年級': 1, '2年級': 2, '3年級': 3,
    '4年級': 4, '5年級': 5, '6年級': 6
  };
  return (gradeText in map) ? map[gradeText] : 99;
}


// ============================================================
// ★【步驟 A：請執行一次】建立 Google 表單（含性別題目）
// ============================================================
function createReportForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.create('(技能測試)忠義國小校園出現集體腸胃不適症狀通報表');

  form.setDescription(
    '1.全校師生若於午餐後發生疑似食物中毒情形，請填寫表單通報個案資料，並即時告知學務處。\n' +
    '2.每個個案請個別填寫一張表單。\n' +
    '3.導師請填1~7題後提交，其餘內容由相關各組人員填寫。'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);
  form.setProgressBar(false);
  form.setShowLinkToRespondAgain(true);

  form.addListItem().setTitle('年級')
    .setChoiceValues(['1年級', '2年級', '3年級', '4年級', '5年級', '6年級', '幼兒園'])
    .setRequired(true);

  const classChoices = [];
  for (let i = 1; i <= 17; i++) classChoices.push(String(i));
  form.addListItem().setTitle('班級').setChoiceValues(classChoices).setRequired(true);

  const seatChoices = [];
  for (let i = 1; i <= 34; i++) seatChoices.push(String(i));
  form.addListItem().setTitle('座號').setChoiceValues(seatChoices).setRequired(true);

  form.addTextItem().setTitle('姓名').setRequired(true);

  form.addListItem()
    .setTitle('性別')
    .setChoiceValues(['男', '女'])
    .setRequired(true);

  form.addTextItem().setTitle('家人稱謂及緊急聯絡電話')
    .setHelpText('例如：媽媽 0911333666').setRequired(true);

  form.addCheckboxItem().setTitle('身體症狀')
    .setChoiceValues(['腹瀉', '嘔吐', '噁心', '腹痛', '發燒', '頭痛', '紅疹'])
    .showOtherOption(true).setRequired(true);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('表單網址：' + form.getPublishedUrl());
  Logger.log('表單編輯網址：' + form.getEditUrl());
  Logger.log('表單 ID：' + form.getId());
}


// ============================================================
// ★【追加欄位：請執行一次】在已建立的表單中新增「身高」「體重」兩題
function addHeightWeightQuestions() {
  const form = FormApp.openById(FORM_ID);

  // 防呆：如果已有身高或體重題目，略過不重複新增
  const existingTitles = form.getItems().map(function (item) { return item.getTitle(); });
  if (existingTitles.indexOf('身高') !== -1 && existingTitles.indexOf('體重') !== -1) {
    Logger.log('身高、體重題目已存在，略過新增。如需調整位置請先執行 removeDuplicateHeightWeight。');
    return;
  }

  form.addTextItem()
    .setTitle('身高')
    .setHelpText('單位：公分，請輸入數字')
    .setRequired(true);

  form.addTextItem()
    .setTitle('體重')
    .setHelpText('單位：公斤，請輸入數字')
    .setRequired(true);

  const items = form.getItems();
  const heightItem = items[items.length - 2];
  const weightItem = items[items.length - 1];
  // 移動到「性別」之後（題目順序：年級、班級、座號、姓名、性別、身高、體重、聯絡電話、身體症狀）
  form.moveItem(heightItem, 5);
  form.moveItem(weightItem, 6);

  Logger.log('已新增「身高」「體重」兩題，並移動到性別之後。');
  Logger.log('⚠ 新欄位在試算表仍會出現在最後兩欄（Google 表單行為，不影響系統）。');
}

// ★【出現重複身高體重時請執行】刪除多餘的身高體重題目，只保留各一題
function removeDuplicateHeightWeight() {
  const form = FormApp.openById(FORM_ID);
  const items = form.getItems();

  // 找出第一題「身高」和第一題「體重」的索引（保留這些，刪除其餘的）
  let firstHeightIdx = -1, firstWeightIdx = -1;
  for (let i = 0; i < items.length; i++) {
    const t = items[i].getTitle();
    if (t === '身高' && firstHeightIdx === -1) firstHeightIdx = i;
    if (t === '體重' && firstWeightIdx === -1) firstWeightIdx = i;
  }

  let deleted = 0;
  // 從後往前刪（避免索引偏移），跳過第一題身高和第一題體重
  for (let i = items.length - 1; i >= 0; i--) {
    const t = items[i].getTitle();
    if (t === '身高' && i !== firstHeightIdx) { form.deleteItem(i); deleted++; }
    if (t === '體重' && i !== firstWeightIdx) { form.deleteItem(i); deleted++; }
  }

  Logger.log('已刪除 ' + deleted + ' 個重複的身高/體重題目。');
}

// ★【身高體重資料沒出現在試算表時請執行】還原正確欄位
function restoreHeightWeightColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const lastCol = mainSheet.getLastColumn();

  // 1. 先取消隱藏所有欄位
  mainSheet.showColumns(1, lastCol);

  // 2. 讀取標題列
  const headerRow = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 3. 找出所有「身高」「體重」的欄位位置（從後往前找，最後面的是接收表單資料的）
  let realHeightCol = -1, realWeightCol = -1;
  for (let i = lastCol - 1; i >= 0; i--) {
    const t = String(headerRow[i] || '').trim();
    if (t === '身高' && realHeightCol === -1) realHeightCol = i + 1;
    if (t === '體重' && realWeightCol === -1) realWeightCol = i + 1;
  }

  // 4. 把不是最後一個的、或是被改名為 _重複_已隱藏 的都隱藏
  let fixed = 0;
  for (let i = 0; i < lastCol; i++) {
    const t = String(headerRow[i] || '').trim();
    const col = i + 1;
    if (t === '身高' && col !== realHeightCol) {
      mainSheet.getRange(1, col).setValue('_重複_已隱藏');
      mainSheet.hideColumns(col);
      fixed++;
    } else if (t === '_重複_已隱藏' && col !== realHeightCol && col !== realWeightCol) {
      mainSheet.hideColumns(col);
    } else if (t === '體重' && col !== realWeightCol) {
      mainSheet.getRange(1, col).setValue('_重複_已隱藏');
      mainSheet.hideColumns(col);
      fixed++;
    } else if (t === '_重複_已隱藏') {
      // 把之前錯改名的那欄恢復成正確標題
      if (col === realHeightCol) {
        mainSheet.getRange(1, col).setValue('身高');
        fixed++;
      } else if (col === realWeightCol) {
        mainSheet.getRange(1, col).setValue('體重');
        fixed++;
      }
    }
  }

  Logger.log('已修正身高體重欄位，共修復 ' + fixed + ' 處。表單新送出的資料將會出現在正確欄位。');
}

// ★【試算表出現重複欄位標題時請執行】刪除彙整總表中重複的欄位（保留最後一欄——接收表單資料的那欄）
function cleanupDuplicateSheetColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const lastCol = mainSheet.getLastColumn();
  if (lastCol < 2) return;

  const headerRow = mainSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 找出每個標題「最後一次出現」的欄位（Google 表單永遠寫到最後面）
  const lastOccurrence = {};
  headerRow.forEach(function (text, idx) {
    const t = String(text || '').trim();
    if (t && t !== '_重複_已隱藏') lastOccurrence[t] = idx + 1;
  });

  // 把不是「最後一次出現」的重複欄位隱藏
  const duplicates = [];
  const firstSeen = {};
  headerRow.forEach(function (text, idx) {
    const t = String(text || '').trim();
    if (!t || t === '_重複_已隱藏') return;
    if (firstSeen[t] === undefined) {
      firstSeen[t] = idx + 1;
    } else if (idx + 1 !== lastOccurrence[t]) {
      duplicates.push(idx + 1);
    }
  });

  if (duplicates.length === 0) {
    Logger.log('沒有重複的欄位標題，試算表結構正常。');
    return;
  }

  duplicates.forEach(function (col) {
    mainSheet.getRange(1, col).setValue('_重複_已隱藏');
    mainSheet.hideColumns(col);
    Logger.log('已隱藏重複欄位：第 ' + col + ' 欄');
  });

  Logger.log('清理完成，共隱藏 ' + duplicates.length + ' 個重複欄位。');
}

// ★【步驟 B：請執行】初始化試算表結構
// ============================================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureFormLinkedToThisSpreadsheet_(ss);
  const mainSheet = findOrRenameFormResponseSheet_(ss);

  const existingMap = getMainColMap_(mainSheet, false);
  const startCol = existingMap.CASE_ID || (mainSheet.getLastColumn() + 1);

  const headers = [
    '案件編號', '排序值', '檢傷初判', '目前所在位置', '就醫醫院',
    '護送教師', '備註', '狀態', '結案', '最後更新時間', '最後更新人員'
  ];
  mainSheet.getRange(1, startCol, 1, headers.length).setValues([headers]);
  mainSheet.getRange(1, 1, 1, startCol + headers.length - 1)
    .setFontWeight('bold').setBackground(THEME.accent).setFontColor('#ffffff');
  mainSheet.setFrozenRows(1);

  GRADE_LIST.forEach(function (gradeName) {
    getOrCreateSheet_(ss, gradeName, buildGradeSheetHeaders_());
  });

  const deidSheet = getOrCreateSheet_(ss, SHEET_DEID, [
    '班級-座號', '姓名', '性別', '狀態', '休息觀察', '送醫醫院', '家長帶離醫院', '出院時間', '排序值'
  ]);
  protectSheetWithWarning_(deidSheet, '此工作表為「去識別化總覽」，資料由系統自動產生（姓名已去識別化），請勿手動編輯。');

  const COL = getMainColMap_(mainSheet);
  protectRangeWithWarning_(
    mainSheet.getRange(1, COL.CASE_ID, mainSheet.getMaxRows(), 2),
    '「案件編號」「排序值」為系統自動產生/計算欄位，請勿手動修改。'
  );
  protectRangeWithWarning_(
    mainSheet.getRange(1, COL.STATUS, mainSheet.getMaxRows(), 4),
    '「狀態」「結案」「最後更新時間」「最後更新人員」為系統自動計算欄位，請勿手動修改。'
  );

  GRADE_LIST.forEach(function (gradeName) {
    const gradeSheet = ss.getSheetByName(gradeName);
    protectRangeWithWarning_(
      gradeSheet.getRange(1, 1, gradeSheet.getMaxRows(), 1),
      '「案件編號」為系統自動產生欄位，請勿手動修改。'
    );
    protectRangeWithWarning_(
      gradeSheet.getRange(1, 13, gradeSheet.getMaxRows(), 5),
      '「狀態」「結案」「最後更新時間」「最後更新人員」「排序值」為系統自動計算欄位，請勿手動修改。'
    );
  });

  SpreadsheetApp.getUi().alert('初始化完成！請檢查「彙整總表」欄位與各年級/去識別化工作表是否已建立。');
}

function protectSheetWithWarning_(sheet, message) {
  removeExistingProtectionsWithDescription_(sheet, message);
  const protection = sheet.protect().setDescription(message);
  protection.setWarningOnly(true);
}

function protectRangeWithWarning_(range, message) {
  removeExistingProtectionsWithDescription_(range.getSheet(), message);
  const protection = range.protect().setDescription(message);
  protection.setWarningOnly(true);
}

function removeExistingProtectionsWithDescription_(sheet, message) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET));
  protections.forEach(function (p) {
    if (p.getDescription() === message) p.remove();
  });
}

function ensureFormLinkedToThisSpreadsheet_(ss) {
  if (!FORM_ID || FORM_ID === 'REPLACE_WITH_YOUR_FORM_ID') {
    Logger.log('FORM_ID 尚未設定（仍是預設值），跳過表單連結檢查。請執行 createReportForm 取得表單 ID 後，更新 FORM_ID 常數再重新執行 setupSpreadsheet。');
    return;
  }
  const form = FormApp.openById(FORM_ID);
  const destId = form.getDestinationId();
  if (destId !== ss.getId()) {
    Logger.log('表單回覆目的地(' + destId + ')與本試算表(' + ss.getId() + ')不同，正在自動重新連結...');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    Utilities.sleep(2000);
  }
}

function findOrRenameFormResponseSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_MAIN);
  if (sheet) return sheet;
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(function (s) { return s.getName(); });
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name.indexOf('表單回應') !== -1 || name.indexOf('表單回覆') !== -1) {
      sheets[i].setName(SHEET_MAIN);
      return sheets[i];
    }
  }
  throw new Error(
    '找不到表單回應工作表。目前工作表有：[' + sheetNames.join(', ') + ']。' +
    '請確認您是在「表單要連結的那一份」試算表裡執行本程式碼。'
  );
}

function getOrCreateSheet_(ss, name, headerRow) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  sheet.getRange(1, 1, 1, headerRow.length)
    .setFontWeight('bold').setBackground(THEME.accent).setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  return sheet;
}

function buildGradeSheetHeaders_() {
  return [
    '案件編號', '班級', '座號', '姓名', '性別', '聯絡電話', '身體症狀',
    '檢傷初判', '目前所在位置', '就醫醫院', '護送教師', '備註',
    '狀態', '結案', '最後更新時間', '最後更新人員', '排序值'
  ];
}


// ============================================================
// ★【步驟 C：請執行】安裝自動觸發器
// ============================================================
function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['onFormSubmitHandler', 'onEditHandler', 'onChangeHandler'].forEach(function (fnName) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
    });
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(ss).onChange().create();
  SpreadsheetApp.getUi().alert(
    '已安裝觸發器：\n' +
    '1) 表單送出時 → 自動處理新資料\n' +
    '2) 手動編輯「彙整總表」時 → 自動同步\n' +
    '3) 試算表結構變更時 → 自動同步\n\n' +
    '請注意：只有編輯「彙整總表」才會觸發同步。各年級分頁與去識別化總覽是自動產生的檢視表，' +
    '請不要直接在那些分頁手動修改。'
  );
}

function onChangeHandler(e) { syncDerivedSheets(); }

function onEditHandler(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (sheetName === SHEET_MAIN) { syncDerivedSheets(); return; }
  if (GRADE_LIST.indexOf(sheetName) !== -1) { propagateGradeEditToMain_(sheet, e.range); return; }
}

function propagateGradeEditToMain_(gradeSheet, editedRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const startRow = editedRange.getRow();
  const numRows = editedRange.getNumRows();
  const startCol = editedRange.getColumn();
  const numCols = editedRange.getNumColumns();

  for (let i = 0; i < numRows; i++) {
    const gRow = startRow + i;
    if (gRow === 1) continue;
    const caseId = gradeSheet.getRange(gRow, 1).getValue();
    if (!caseId) continue;
    const mainRow = findMainRowByCaseId_(mainSheet, caseId, COL);
    if (!mainRow) continue;

    let rowChanged = false;
    for (let j = 0; j < numCols; j++) {
      const gCol = startCol + j;
      const fieldKey = GRADE_SHEET_FIELDS[gCol - 1];
      if (!fieldKey) continue;
      if (GRADE_EDITABLE_FIELDS.indexOf(fieldKey) === -1) continue;
      const newValue = gradeSheet.getRange(gRow, gCol).getValue();
      const mainCol = COL[fieldKey];
      mainSheet.getRange(mainRow, mainCol).setValue(newValue);
      rowChanged = true;
    }

    if (rowChanged) {
      mainSheet.getRange(mainRow, COL.UPDATED_AT).setValue(new Date());
      mainSheet.getRange(mainRow, COL.UPDATED_BY).setValue('人工修改(' + gradeSheet.getName() + ')');
    }
  }
  syncDerivedSheets();
}

function findMainRowByCaseId_(mainSheet, caseId, colMap) {
  const COL = colMap || getMainColMap_(mainSheet);
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = mainSheet.getRange(2, COL.CASE_ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === caseId) return i + 2;
  }
  return null;
}


// ============================================================
// 【表單送出時自動執行】觸發器呼叫
// ============================================================
function onFormSubmitHandler(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(sheet);
  const row = e.range.getRow();
  const grade = sheet.getRange(row, COL.GRADE).getValue();
  const classNo = sheet.getRange(row, COL.CLASS).getValue();
  const seatNo = sheet.getRange(row, COL.SEAT).getValue();
  const now = new Date();
  const caseId = generateCaseId_(sheet, now, COL);
  const sortKey = gradeToNumber_(grade) * 10000 + Number(classNo) * 100 + Number(seatNo);
  sheet.getRange(row, COL.CASE_ID).setValue(caseId);
  sheet.getRange(row, COL.SORT_KEY).setValue(sortKey);
  sheet.getRange(row, COL.STATUS).setValue('檢傷中');
  sheet.getRange(row, COL.CLOSED).setValue('否');
  sheet.getRange(row, COL.UPDATED_AT).setValue(now);
  sheet.getRange(row, COL.UPDATED_BY).setValue('導師通報');
  syncDerivedSheets();
}


// ============================================================
// 【全站資料同步】
// ============================================================
function syncDerivedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return;

  const range = mainSheet.getRange(2, 1, lastRow - 1, COL.LAST_COL);
  const values = range.getValues();

  const updatedValues = values.map(function (r) {
    const grade = r[COL.GRADE - 1];
    const classNo = r[COL.CLASS - 1];
    const seatNo = r[COL.SEAT - 1];
    if (grade && classNo !== '' && seatNo !== '') {
      r[COL.SORT_KEY - 1] = gradeToNumber_(grade) * 10000 + Number(classNo) * 100 + Number(seatNo);
    }
    const triage = r[COL.TRIAGE - 1];
    const location = r[COL.LOCATION - 1];
    const hospital = r[COL.HOSPITAL - 1];
    const status = computeStatus_(triage, location, hospital);
    r[COL.STATUS - 1] = status.text;
    r[COL.CLOSED - 1] = status.closed ? '是' : '否';
    return r;
  });

  const sorted = updatedValues.slice().sort(function (a, b) {
    return (a[COL.SORT_KEY - 1] || 0) - (b[COL.SORT_KEY - 1] || 0);
  });

  range.setValues(sorted);

  GRADE_LIST.forEach(function (gradeName) {
    const gradeRows = sorted.filter(function (r) { return r[COL.GRADE - 1] === gradeName; });
    writeGradeSheet_(ss, gradeName, gradeRows, COL);
  });

  writeDeidSheet_(ss, sorted, COL);
}

function computeStatus_(triage, location, hospital) {
  if (hospital === '已回教室' || location === '已回教室') {
    return { text: '已回教室', closed: true };
  }
  if (hospital === '家長接回' || location === '家長接回') {
    return { text: '家長接回', closed: true };
  }
  if (hospital) return { text: '送醫中(已抵達' + hospital + ')', closed: false };
  if (triage === '送醫區') return { text: '送醫中', closed: false };
  if (location) {
    if (location.indexOf('休息觀察區') === 0) return { text: '休息', closed: false };
    if (location === '回檢傷組重新判斷') return { text: '檢傷中', closed: false };
    if (location.indexOf('送醫區') === 0) return { text: '送醫中', closed: false };
    return { text: '休息', closed: false };
  }
  if (triage === '休息觀察區') return { text: '休息', closed: false };
  return { text: '檢傷中', closed: false };
}

function generateCaseId_(sheet, dateObj, colMap) {
  const COL = colMap || getMainColMap_(sheet);
  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
  const lastRow = sheet.getLastRow();
  let countToday = 0;
  if (lastRow >= 2) {
    const ids = sheet.getRange(2, COL.CASE_ID, lastRow - 1, 1).getValues();
    ids.forEach(function (r) {
      if (String(r[0]).indexOf(dateStr) === 0) countToday++;
    });
  }
  const seq = String(countToday + 1).padStart(3, '0');
  return dateStr + '-' + seq;
}

function deidentifyName_(name) {
  if (!name) return '';
  const s = String(name).trim();
  if (s.length <= 1) return s;
  if (s.length === 2) return s.charAt(0) + '○';
  const middle = '○'.repeat(s.length - 2);
  return s.charAt(0) + middle + s.charAt(s.length - 1);
}

function formatClassSeatCode_(grade, classNo, seatNo) {
  const g = gradeToNumber_(grade);
  const c = String(classNo).padStart(2, '0');
  const s = String(seatNo).padStart(2, '0');
  return String(g) + c + '-' + s;
}

function writeGradeSheet_(ss, gradeName, rows, COL) {
  const sheet = ss.getSheetByName(gradeName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (rows.length === 0) return;
  const output = rows.map(function (r) {
    return [
      r[COL.CASE_ID - 1], r[COL.CLASS - 1], r[COL.SEAT - 1], r[COL.NAME - 1], r[COL.GENDER - 1],
      r[COL.CONTACT - 1], r[COL.SYMPTOM - 1], r[COL.TRIAGE - 1], r[COL.LOCATION - 1],
      r[COL.HOSPITAL - 1], r[COL.ESCORT - 1], r[COL.NOTE - 1], r[COL.STATUS - 1],
      r[COL.CLOSED - 1], r[COL.UPDATED_AT - 1], r[COL.UPDATED_BY - 1], r[COL.SORT_KEY - 1]
    ];
  });
  sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}

function writeDeidSheet_(ss, sortedRows, COL) {
  const sheet = ss.getSheetByName(SHEET_DEID);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  if (sortedRows.length === 0) return;
  const tz = Session.getScriptTimeZone();
  const output = sortedRows.map(function (r) {
    const classSeatCode = formatClassSeatCode_(r[COL.GRADE - 1], r[COL.CLASS - 1], r[COL.SEAT - 1]);
    const deidName = deidentifyName_(r[COL.NAME - 1]);
    const status = r[COL.STATUS - 1];
    const location = r[COL.LOCATION - 1] || '';
    const hospital = r[COL.HOSPITAL - 1] || '';
    const parentPickup = (hospital === '家長接回' || location === '家長接回') ? '是' : '';
    const dischargeTime = (status === '已回教室' || status === '家長接回') ? (r[COL.UPDATED_AT - 1] ? Utilities.formatDate(new Date(r[COL.UPDATED_AT - 1]), tz, 'yyyy/MM/dd HH:mm') : '') : '';
    return [
      classSeatCode, deidName, r[COL.GENDER - 1] || '', status,
      location, hospital, parentPickup, dischargeTime, r[COL.SORT_KEY - 1]
    ];
  });
  sheet.getRange(2, 1, output.length, output[0].length).setValues(output);
}


// ============================================================
// 【網頁應用程式】doGet 路由
// ============================================================
function doGet(e) {
  const page = ((e && e.parameter && e.parameter.page) || '').toLowerCase();
  if (page === 'admin') {
    return HtmlService.createHtmlOutput(buildAdminHtml_())
      .setTitle('管理者後台')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (GROUP_CONFIG[page]) {
    return HtmlService.createHtmlOutput(buildGroupPortalHtml_(page))
      .setTitle(GROUP_CONFIG[page].label + '通報')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutput(buildHomeHtml_())
    .setTitle('校園出現集體腸胃不適症狀事件通報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// 【後端函式】由前端 google.script.run 呼叫
// ============================================================

function verifyGroupPassword(group, pwd) {
  const cfg = GROUP_CONFIG[group];
  if (!cfg) return false;
  const stored = PropertiesService.getScriptProperties().getProperty(cfg.propKey);
  if (!stored) return false;
  return String(pwd) === String(stored);
}

function searchCases(group, pwd, keyword) {
  if (!verifyGroupPassword(group, pwd)) {
    throw new Error('密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, COL.LAST_COL).getValues();
  const kw = String(keyword || '').trim();
  const matched = values.filter(function (r) {
    const classSeat = r[COL.CLASS - 1] + '-' + r[COL.SEAT - 1];
    if (!kw) return r[COL.CLOSED - 1] !== '是';
    return (
      classSeat.indexOf(kw) !== -1 ||
      String(r[COL.NAME - 1]).indexOf(kw) !== -1 ||
      String(r[COL.CASE_ID - 1]).indexOf(kw) !== -1
    );
  });
  return matched.slice(0, 50).map(function (r) {
    return {
      caseId: r[COL.CASE_ID - 1],
      grade: r[COL.GRADE - 1],
      classNo: r[COL.CLASS - 1],
      seatNo: r[COL.SEAT - 1],
      name: r[COL.NAME - 1],
      gender: r[COL.GENDER - 1],
      symptom: r[COL.SYMPTOM - 1],
      triage: r[COL.TRIAGE - 1],
      location: r[COL.LOCATION - 1],
      hospital: r[COL.HOSPITAL - 1],
      escort: r[COL.ESCORT - 1],
      note: r[COL.NOTE - 1],
      status: r[COL.STATUS - 1],
      closed: r[COL.CLOSED - 1]
    };
  });
}

function updateCase(group, pwd, caseId, payload) {
  if (!verifyGroupPassword(group, pwd)) {
    throw new Error('密碼錯誤或登入已失效，請重新整理頁面再登入一次。');
  }
  const cfg = GROUP_CONFIG[group];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const row = findMainRowByCaseId_(mainSheet, caseId, COL);
  if (!row) throw new Error('找不到案件編號：' + caseId);
  if (group === 'triage') {
    mainSheet.getRange(row, COL.TRIAGE).setValue(payload.value || '');
  } else if (group === 'rest') {
    mainSheet.getRange(row, COL.LOCATION).setValue(payload.value || '');
  } else if (group === 'hospital') {
    mainSheet.getRange(row, COL.HOSPITAL).setValue(payload.value || '');
    mainSheet.getRange(row, COL.ESCORT).setValue(payload.escort || '');
  }
  if (payload.note !== undefined) {
    mainSheet.getRange(row, COL.NOTE).setValue(payload.note);
  }
  mainSheet.getRange(row, COL.UPDATED_AT).setValue(new Date());
  mainSheet.getRange(row, COL.UPDATED_BY).setValue(cfg.label);
  syncDerivedSheets();
  return { success: true };
}

function verifyAdminPassword(pwd) {
  const stored = PropertiesService.getScriptProperties().getProperty(ADMIN_PROP_KEY);
  if (!stored) return false;
  return String(pwd) === String(stored);
}

function adminListAllCases(pwd) {
  if (!verifyAdminPassword(pwd)) {
    throw new Error('管理者密碼錯誤或登入已失效。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, COL.LAST_COL).getValues();
  return values.map(function (r) {
    return {
      caseId: r[COL.CASE_ID - 1],
      grade: r[COL.GRADE - 1],
      classNo: r[COL.CLASS - 1],
      seatNo: r[COL.SEAT - 1],
      name: r[COL.NAME - 1],
      gender: r[COL.GENDER - 1],
      contact: r[COL.CONTACT - 1],
      symptom: r[COL.SYMPTOM - 1],
      triage: r[COL.TRIAGE - 1],
      location: r[COL.LOCATION - 1],
      hospital: r[COL.HOSPITAL - 1],
      escort: r[COL.ESCORT - 1],
      note: r[COL.NOTE - 1],
      status: r[COL.STATUS - 1],
      closed: r[COL.CLOSED - 1]
    };
  });
}

function adminUpdateCase(pwd, caseId, fields) {
  if (!verifyAdminPassword(pwd)) throw new Error('管理者密碼錯誤。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const row = findMainRowByCaseId_(mainSheet, caseId, COL);
  if (!row) throw new Error('找不到案件編號：' + caseId);
  const fieldColMap = {
    grade: COL.GRADE, classNo: COL.CLASS, seatNo: COL.SEAT, name: COL.NAME,
    gender: COL.GENDER, contact: COL.CONTACT, symptom: COL.SYMPTOM,
    triage: COL.TRIAGE, location: COL.LOCATION, hospital: COL.HOSPITAL,
    escort: COL.ESCORT, note: COL.NOTE,
    height: COL.HEIGHT, weight: COL.WEIGHT
  };
  Object.keys(fieldColMap).forEach(function (key) {
    if (fields[key] !== undefined) {
      mainSheet.getRange(row, fieldColMap[key]).setValue(fields[key]);
    }
  });
  mainSheet.getRange(row, COL.UPDATED_AT).setValue(new Date());
  mainSheet.getRange(row, COL.UPDATED_BY).setValue('管理者');
  syncDerivedSheets();
  return { success: true };
}

function adminAddCase(pwd, fields) {
  if (!verifyAdminPassword(pwd)) throw new Error('管理者密碼錯誤。');
  if (!fields.grade || !fields.classNo || !fields.seatNo || !fields.name) {
    throw new Error('年級、班級、座號、姓名為必填欄位。');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const now = new Date();
  const caseId = generateCaseId_(mainSheet, now, COL);
  const sortKey = gradeToNumber_(fields.grade) * 10000 + Number(fields.classNo) * 100 + Number(fields.seatNo);
  const newRow = new Array(COL.LAST_COL).fill('');
  newRow[COL.TIMESTAMP - 1] = now;
  newRow[COL.GRADE - 1] = fields.grade;
  newRow[COL.CLASS - 1] = fields.classNo;
  newRow[COL.SEAT - 1] = fields.seatNo;
  newRow[COL.NAME - 1] = fields.name;
  newRow[COL.GENDER - 1] = fields.gender || '';
  newRow[COL.CONTACT - 1] = fields.contact || '';
  newRow[COL.SYMPTOM - 1] = fields.symptom || '';
  newRow[COL.HEIGHT - 1] = fields.height || '';
  newRow[COL.WEIGHT - 1] = fields.weight || '';
  newRow[COL.CASE_ID - 1] = caseId;
  newRow[COL.SORT_KEY - 1] = sortKey;
  newRow[COL.STATUS - 1] = '檢傷中';
  newRow[COL.CLOSED - 1] = '否';
  newRow[COL.UPDATED_AT - 1] = now;
  newRow[COL.UPDATED_BY - 1] = '管理者新增';
  mainSheet.appendRow(newRow);
  syncDerivedSheets();
  return { success: true, caseId: caseId };
}

/**
 * 清理用：刪除彙整總表中所有案件編號為空白的列（幽靈資料），
 * 以及同名重複的資料（保留第一筆）。
 * 操作完成後會自動全站同步。
 */
function cleanupEmptyAndDupCases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const COL = getMainColMap_(mainSheet);
  const lastRow = mainSheet.getLastRow();
  if (lastRow < 2) return { deleted: 0 };

  // 從最後一列往上刪（避免刪除後列號偏移）
  const deleted = [];
  const seenNames = {};

  for (let r = lastRow; r >= 2; r--) {
    const caseId = mainSheet.getRange(r, COL.CASE_ID).getValue();
    const name = String(mainSheet.getRange(r, COL.NAME).getValue() || '').trim();

    if (!caseId) {
      // 案件編號空白 → 刪除
      mainSheet.deleteRow(r);
      deleted.push({ row: r, name: name || '(empty)', reason: 'no_caseId' });
    } else if (name) {
      // 檢查是否為同名重複
      if (seenNames[name]) {
        mainSheet.deleteRow(r);
        deleted.push({ row: r, name: name, caseId: caseId, reason: 'duplicate' });
      } else {
        seenNames[name] = true;
      }
    }
  }

  syncDerivedSheets();
  return { deleted: deleted.length, details: deleted };
}

function adminDeleteCase(pwd, caseId) {
  if (!verifyAdminPassword(pwd)) throw new Error('管理者密碼錯誤。');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_MAIN);
  const row = findMainRowByCaseId_(mainSheet, caseId);
  if (!row) throw new Error('找不到案件編號：' + caseId);
  mainSheet.deleteRow(row);
  syncDerivedSheets();
  return { success: true };
}

function getHomeTableData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DEID);
  const lastRow = sheet.getLastRow();
  const tz = Session.getScriptTimeZone();
  if (lastRow < 2) {
    return { count: 0, rows: [], updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss') };
  }
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const rows = values.map(function (r) {
    return {
      classSeat: r[0],
      name: r[1],
      gender: r[2],
      status: r[3],
      location: r[4],
      hospital: r[5],
      parentPickup: r[6],
      dischargeTime: r[7] ? Utilities.formatDate(new Date(r[7]), tz, 'yyyy/MM/dd HH:mm') : ''
    };
  });
  return {
    count: rows.length,
    rows: rows,
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')
  };
}


// ============================================================
// 【HTML 樣板】首頁（紫色暖色系）
// ============================================================
function buildHomeHtml_() {
  return '' +
'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">' +
'<style>' +
'  * { box-sizing: border-box; }' +
'  body { margin:0; font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif; background:' + THEME.bg + '; color:#3d3546; }' +
'  .banner { background:' + THEME.bannerGradient + '; padding:24px 32px; display:flex; align-items:center; gap:20px; color:#fff; }' +
'  .banner .logo { width:64px; height:64px; border-radius:50%; background:#fff; display:flex; align-items:center; justify-content:center; font-size:28px; flex-shrink:0; box-shadow:0 2px 6px rgba(0,0,0,.15); }' +
'  .banner h1 { margin:0; font-size:1.5rem; }' +
'  .banner p { margin:4px 0 0; font-size:1rem; opacity:.9; }' +
'  .container { max-width:1100px; margin:24px auto; padding:0 16px; }' +
'  .summary-bar { display:flex; align-items:center; gap:12px; justify-content:center; margin-bottom:16px; flex-wrap:wrap; }' +
'  .summary-badge { background:' + THEME.badge + '; color:#fff; padding:10px 20px; border-radius:20px; font-weight:bold; }' +
'  .updated-at { font-size:.85rem; color:' + THEME.textMuted + '; }' +
'  table { width:100%; border-collapse:collapse; background:' + THEME.cardBg + '; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(120,80,140,.1); }' +
'  th { background:#9b7ebd; color:#fff; padding:10px 8px; font-size:.9rem; }' +
'  td { padding:9px 8px; text-align:center; border-bottom:1px solid ' + THEME.border + '; font-size:.9rem; }' +
'  tr:nth-child(even) { background:' + THEME.altRow + '; }' +
'  tr:last-child td { border-bottom:none; }' +
'  .status-tag { padding:3px 10px; border-radius:12px; color:#fff; font-size:.8rem; display:inline-block; }' +
'  .st-檢傷中 { background:#e0a458; } .st-休息 { background:#8fa9b9; } .st-送醫中 { background:#c98a5c; }' +
'  .st-已回教室 { background:#94998c; } .st-家長接回 { background:#6a8fa3; }' +
'  .empty { text-align:center; padding:40px; color:#9a91a8; }' +
'  .search-bar { display:flex; align-items:center; gap:10px; justify-content:center; margin-bottom:16px; flex-wrap:wrap; }' +
'  .search-bar input[type=text], .search-bar select { padding:8px 12px; border-radius:8px; border:1px solid ' + THEME.border + '; font-size:.9rem; font-family:inherit; }' +
'  .search-bar input[type=text] { width:220px; }' +
'  .search-bar button { padding:8px 16px; border-radius:8px; border:none; background:' + THEME.primary + '; color:#fff; font-weight:bold; cursor:pointer; font-size:.9rem; }' +
'  .search-bar button.secondary { background:' + THEME.secondaryBg + '; color:' + THEME.secondaryText + '; }' +
'  .refresh-icon { width:34px; height:34px; border-radius:50%; border:1px solid ' + THEME.border + '; background:#fff; cursor:pointer; font-size:1rem; display:flex; align-items:center; justify-content:center; }' +
'  .refresh-icon:hover { background:' + THEME.altRow + '; }' +
'  @media screen and (max-width: 768px) {' +
'    html, body { overflow-x: hidden !important; width: 100vw !important; margin: 0; padding: 0; }' +
'    .container { width: 100% !important; max-width: 100vw !important; padding: 10px !important; box-sizing: border-box !important; }' +
'    #view-search table, #view-search thead, #view-search tbody, #view-search tr, #view-search td { display: flex !important; flex-direction: column !important; width: 100% !important; max-width: 100% !important; box-sizing: border-box !important; min-width: 0 !important; }' +
'    #view-search thead { display: none !important; }' +
'    #view-search tr { display: block !important; border: 1px solid ' + THEME.border + '; border-radius: 12px; margin-bottom: 18px !important; background-color: #fff; box-shadow: 0 4px 10px rgba(100,60,140,0.08); padding: 6px !important; overflow: hidden !important; }' +
'    #view-search td { display: flex !important; flex-direction: row !important; justify-content: flex-start !important; align-items: flex-start !important; text-align: left !important; border: none !important; border-bottom: 1px solid ' + THEME.border + ' !important; padding: 10px 10px 10px 120px !important; position: relative; min-height: 42px; word-break: break-all !important; white-space: normal !important; font-size: 0.92em; color: #444; line-height: 1.4; }' +
'    #view-search td:last-child { border-bottom: none !important; }' +
'    #view-search td:before { content: attr(data-label); position: absolute; left: 10px; font-weight: bold; color: ' + THEME.badge + '; font-size: 0.85em; text-align: left; white-space: nowrap; }' +
'    #view-search tr:last-child { margin-bottom: 100px !important; }' +
'    #view-search td:nth-of-type(1):before { content: "班級-座號"; }' +
'    #view-search td:nth-of-type(2):before { content: "姓名"; }' +
'    #view-search td:nth-of-type(3):before { content: "性別"; }' +
'    #view-search td:nth-of-type(4):before { content: "狀態"; }' +
'    #view-search td:nth-of-type(5):before { content: "休息觀察"; }' +
'    #view-search td:nth-of-type(6):before { content: "送醫醫院"; }' +
'    #view-search td:nth-of-type(7):before { content: "家長帶離"; }' +
'    #view-search td:nth-of-type(8):before { content: "出院時間"; }' +
'    .banner { padding: 16px 20px; }' +
'    .banner h1 { font-size: 1.2rem; }' +
'    .banner .logo { width: 48px; height: 48px; font-size: 22px; }' +
'  }' +
'</style></head><body>' +
'  <div class="banner">' +
'    <div class="logo">' + (SCHOOL_LOGO_BASE64 ? ('<img src="data:image/png;base64,' + SCHOOL_LOGO_BASE64 + '" style="width:100%;height:100%;object-fit:contain;padding:6px;">') : '<span style="font-size:28px;">🏫</span>') + '</div>' +
'    <div><h1>(技能測試)忠義國小</h1><p>校園出現集體腸胃不適症狀事件通報</p></div>' +
'  </div>' +
'  <div class="container">' +
'    <div class="summary-bar">' +
'      <span class="summary-badge">學生醫療資料總覽 <span id="count">0</span> 人</span>' +
'      <span class="updated-at">最後更新：<span id="updatedAt">--</span></span>' +
'    </div>' +
'    <div class="search-bar">' +
'      <button class="refresh-icon" onclick="refresh()" title="重新整理">🔄</button>' +
'      <input type="text" id="searchInput" placeholder="搜尋姓名或班級-座號..." onkeydown="if(event.key===\\\'Enter\\\')doSearch()">' +
'      <select id="statusSelect">' +
'        <option value="">所有狀態</option>' +
'        <option value="檢傷中">檢傷中</option>' +
'        <option value="休息">休息</option>' +
'        <option value="送醫中">送醫中</option>' +
'        <option value="已回教室">已回教室</option>' +
'        <option value="家長接回">家長接回</option>' +
'      </select>' +
'      <button onclick="doSearch()">查詢</button>' +
'      <button class="secondary" onclick="showAll()">顯示全部</button>' +
'    </div>' +
'    <div id="tableWrap"><div class="empty">載入中...</div></div>' +
'  </div>' +
'<script>' +
'var homeData = { count: 0, rows: [], updatedAt: "--" };' +
'function statusClass(s){ return "status-tag st-" + s.replace(/\\(.*\\)/,""); }' +
'function escHtml(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }' +
'function passesFilter(r, kw, statusFilter){' +
'  var matchKw = !kw || String(r.classSeat).indexOf(kw) !== -1 || String(r.name||"").indexOf(kw) !== -1;' +
'  var matchStatus = !statusFilter ||' +
'    String(r.status||"").indexOf(statusFilter) === 0 ||' +
'    String(r.location||"").indexOf(statusFilter) === 0 ||' +
'    String(r.hospital||"").indexOf(statusFilter) === 0;' +
'  return matchKw && matchStatus;' +
'}' +
'function renderTable(){' +
'  var kw = document.getElementById("searchInput").value.trim();' +
'  var statusFilter = document.getElementById("statusSelect").value;' +
'  var rows = homeData.rows.filter(function(r){ return passesFilter(r, kw, statusFilter); });' +
'  var wrap = document.getElementById("tableWrap");' +
'  if (!rows.length) { wrap.innerHTML = "<div class=\\"empty\\">查無符合條件的資料</div>"; return; }' +
'  var html = "<div id=\\"view-search\\"><table><thead><tr><th>班級-座號</th><th>姓名</th><th>性別</th><th>狀態</th>" +' +
'    "<th>休息觀察</th><th>送醫醫院</th><th>家長帶離</th><th>出院時間</th></tr></thead><tbody>";' +
'  rows.forEach(function(r){' +
'    html += "<tr><td data-label=\\"班級-座號\\">"+escHtml(r.classSeat)+"</td><td data-label=\\"姓名\\">"+escHtml(r.name)+"</td><td data-label=\\"性別\\">"+escHtml(r.gender||"")+"</td>" +' +
'      "<td data-label=\\"狀態\\"><span class=\\""+statusClass(r.status)+"\\">"+escHtml(r.status)+"</span></td>" +' +
'      "<td data-label=\\"休息觀察\\">"+escHtml(r.location||"")+"</td><td data-label=\\"送醫醫院\\">"+escHtml(r.hospital||"")+"</td><td data-label=\\"家長帶離\\">"+escHtml(r.parentPickup||"")+"</td>" +' +
'      "<td data-label=\\"出院時間\\">"+escHtml(r.dischargeTime||"")+"</td></tr>";' +
'  });' +
'  html += "</tbody></table></div>";' +
'  wrap.innerHTML = html;' +
'}' +
'function doSearch(){ renderTable(); }' +
'function showAll(){' +
'  document.getElementById("searchInput").value = "";' +
'  document.getElementById("statusSelect").value = "";' +
'  renderTable();' +
'}' +
'function refresh(){' +
'  google.script.run.withSuccessHandler(function(data){' +
'    homeData = data;' +
'    document.getElementById("count").textContent = data.count;' +
'    document.getElementById("updatedAt").textContent = data.updatedAt;' +
'    renderTable();' +
'  }).getHomeTableData();' +
'}' +
'refresh();' +
'setInterval(refresh, 15000);' +
'</script>' +
'</body></html>';
}


// ============================================================
// 【HTML 樣板】三組通報共用頁面
// ============================================================
function buildGroupPortalHtml_(group) {
  const cfg = GROUP_CONFIG[group];
  const optionsHtml = cfg.options.map(function (o) {
    return '<option value="' + o + '">' + o + '</option>';
  }).join('');

  return '' +
'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">' +
'<style>' +
'  * { box-sizing: border-box; }' +
'  body { margin:0; font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif; background:' + cfg.themeColorLight + '; color:#3d3230; min-height:100vh; }' +
'  .topbar { background:' + cfg.themeColor + '; color:#3d332f; padding:16px 24px; font-size:1.2rem; font-weight:bold; }' +
'  .container { max-width:720px; margin:32px auto; padding:0 16px; }' +
'  .card { background:#fff; border-radius:12px; padding:24px; box-shadow:0 4px 16px rgba(0,0,0,.08); margin-bottom:16px; }' +
'  input, select, button, textarea { font-size:1rem; padding:10px 12px; border-radius:8px; border:1px solid #ddd; width:100%; margin-top:6px; font-family:inherit; }' +
'  label { font-weight:bold; font-size:.9rem; color:#665; }' +
'  button { background:' + cfg.themeColor + '; color:#3d332f; border:none; font-weight:bold; cursor:pointer; margin-top:12px; }' +
'  button:hover { opacity:.9; }' +
'  .msg { padding:10px; border-radius:8px; margin-top:10px; font-size:.9rem; }' +
'  .msg.error { background:#fdecea; color:#c0392b; }' +
'  .msg.ok { background:#eafaf1; color:#27ae60; }' +
'  .case-item { border:1px solid #eee; border-radius:8px; padding:12px; margin-bottom:10px; cursor:pointer; }' +
'  .case-item:hover { background:#faf9f7; }' +
'  .case-item .title { font-weight:bold; }' +
'  .case-item .status { float:right; font-size:.85rem; color:#888; }' +
'  #searchResults { max-height:400px; overflow-y:auto; }' +
'  #updatePanel { display:none; }' +
'  .hidden { display:none; }' +
'</style></head><body>' +
'  <div class="topbar">' + cfg.label + '通報系統</div>' +
'  <div class="container">' +
'    <div class="card" id="loginCard">' +
'      <label>請輸入' + cfg.label + '通報密碼</label>' +
'      <input type="password" id="pwdInput" placeholder="請輸入密碼">' +
'      <button onclick="doLogin()">登入</button>' +
'      <div id="loginMsg"></div>' +
'    </div>' +
'    <div class="card hidden" id="searchCard">' +
'      <label>搜尋學生（班級-座號 / 姓名 / 案件編號，留空列出所有未結案案件）</label>' +
'      <input type="text" id="kwInput" placeholder="例如：3-15 或 王小明">' +
'      <button onclick="doSearch()">搜尋</button>' +
'      <div id="searchMsg"></div>' +
'      <div id="searchResults"></div>' +
'    </div>' +
'    <div class="card" id="updatePanel">' +
'      <div id="updateTitle" style="font-weight:bold; margin-bottom:10px;"></div>' +
'      <label>' + cfg.updateFieldLabel + '</label>' +
'      <select id="valueSelect"><option value="">（請選擇）</option>' + optionsHtml + '</select>' +
       (group === 'hospital' ?
'      <label>護送教師</label><input type="text" id="escortInput" placeholder="護送教師姓名">' : '') +
'      <label>備註</label><textarea id="noteInput" rows="3" placeholder="選填"></textarea>' +
'      <button onclick="doUpdate()">送出更新</button>' +
'      <button style="background:#eee;" onclick="closeUpdatePanel()">取消</button>' +
'      <div id="updateMsg"></div>' +
'    </div>' +
'  </div>' +
'<script>' +
'var GROUP = "' + group + '";' +
'var PWD = "";' +
'var currentCaseId = "";' +
'function showMsg(elId, text, ok){' +
'  var el = document.getElementById(elId);' +
'  el.innerHTML = "<div class=\\"msg " + (ok ? "ok" : "error") + "\\">" + text + "</div>";' +
'}' +
'function doLogin(){' +
'  var pwd = document.getElementById("pwdInput").value;' +
'  if (!pwd) { showMsg("loginMsg", "請輸入密碼", false); return; }' +
'  google.script.run.withSuccessHandler(function(ok){' +
'    if (ok) {' +
'      PWD = pwd;' +
'      document.getElementById("loginCard").classList.add("hidden");' +
'      document.getElementById("searchCard").classList.remove("hidden");' +
'      doSearch();' +
'    } else { showMsg("loginMsg", "密碼錯誤，請重新輸入", false); }' +
'  }).verifyGroupPassword(GROUP, pwd);' +
'}' +
'function doSearch(){' +
'  var kw = document.getElementById("kwInput").value;' +
'  google.script.run.withSuccessHandler(function(list){' +
'    var html = "";' +
'    list.forEach(function(c){' +
'      html += "<div class=\\"case-item\\" onclick=\\"selectCase(\\\'"+c.caseId+"\\\', \\\'"+c.name+"\\\')\\">" +' +
'        "<span class=\\"title\\">"+c.caseId+" "+c.name+" ("+c.grade+" "+c.classNo+"班"+c.seatNo+"號)</span>" +' +
'        "<span class=\\"status\\">"+c.status+"</span>" +' +
'        "</div>";' +
'    });' +
'    if (!list.length) html = "<div class=\\"msg ok\\">查無符合條件的案件</div>";' +
'    document.getElementById("searchResults").innerHTML = html;' +
'  }).withFailureHandler(function(err){ showMsg("searchMsg", err.message, false); })' +
'    .searchCases(GROUP, PWD, kw);' +
'}' +
'function selectCase(caseId, name){' +
'  currentCaseId = caseId;' +
'  document.getElementById("updateTitle").textContent = "更新案件：" + caseId + " " + name;' +
'  document.getElementById("updatePanel").style.display = "block";' +
'  document.getElementById("valueSelect").value = "";' +
'  document.getElementById("noteInput").value = "";' +
'  var escEl = document.getElementById("escortInput"); if (escEl) escEl.value = "";' +
'}' +
'function closeUpdatePanel(){' +
'  document.getElementById("updatePanel").style.display = "none";' +
'  currentCaseId = "";' +
'}' +
'function doUpdate(){' +
'  if (!currentCaseId) { showMsg("updateMsg", "請先選擇要更新的案件", false); return; }' +
'  var payload = { value: document.getElementById("valueSelect").value, note: document.getElementById("noteInput").value };' +
'  var escEl = document.getElementById("escortInput"); if (escEl) payload.escort = escEl.value;' +
'  google.script.run.withSuccessHandler(function(){' +
'    showMsg("updateMsg", "更新成功！", true);' +
'    closeUpdatePanel();' +
'    doSearch();' +
'  }).withFailureHandler(function(err){ showMsg("updateMsg", err.message, false); })' +
'    .updateCase(GROUP, PWD, currentCaseId, payload);' +
'}' +
'</script>' +
'</body></html>';
}


// ============================================================
// 【HTML 樣板】管理者後台
// ============================================================
function buildAdminHtml_() {
  return '' +
'<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">' +
'<style>' +
'  * { box-sizing: border-box; }' +
'  body { margin:0; font-family:"Noto Sans TC","Microsoft JhengHei",sans-serif; background:#f3f0f7; color:#3d3546; }' +
'  .topbar { background:#7b5ea7; color:#fff; padding:16px 24px; font-size:1.2rem; font-weight:bold; }' +
'  .container { max-width:960px; margin:24px auto; padding:0 16px; }' +
'  .card { background:#fff; border-radius:12px; padding:24px; box-shadow:0 4px 16px rgba(0,0,0,.08); margin-bottom:16px; }' +
'  input, select, button, textarea { font-size:1rem; padding:10px 12px; border-radius:8px; border:1px solid #ddd; width:100%; margin-top:6px; font-family:inherit; }' +
'  label { font-weight:bold; font-size:.9rem; color:#665; }' +
'  button { background:#9b7ebd; color:#fff; border:none; font-weight:bold; cursor:pointer; margin-top:12px; }' +
'  button:hover { opacity:.9; }' +
'  button.danger { background:#d4776e; }' +
'  .msg { padding:10px; border-radius:8px; margin-top:10px; font-size:.9rem; }' +
'  .msg.error { background:#fdecea; color:#c0392b; }' +
'  .msg.ok { background:#eafaf1; color:#27ae60; }' +
'  table { width:100%; border-collapse:collapse; margin-top:12px; }' +
'  th { background:#9b7ebd; color:#fff; padding:8px 6px; font-size:.8rem; }' +
'  td { padding:8px 6px; border-bottom:1px solid #e8dff5; font-size:.8rem; text-align:center; }' +
'  tr:nth-child(even) { background:#faf7fd; }' +
'  .hidden { display:none; }' +
'  .btn-sm { font-size:.75rem; padding:4px 10px; margin:2px; width:auto; }' +
'</style></head><body>' +
'  <div class="topbar">管理者後台</div>' +
'  <div class="container">' +
'    <div class="card" id="loginCard">' +
'      <label>管理者密碼</label>' +
'      <input type="password" id="pwdInput" placeholder="請輸入管理者密碼">' +
'      <button onclick="doLogin()">登入</button>' +
'      <div id="loginMsg"></div>' +
'    </div>' +
'    <div class="card hidden" id="mainCard">' +
'      <div style="display:flex; justify-content:space-between; align-items:center;">' +
'        <span style="font-weight:bold;">案件列表</span>' +
'        <button style="width:auto;" onclick="openAddForm()">+ 新增案件</button>' +
'      </div>' +
'      <div id="listMsg"></div>' +
'      <div id="listWrap"></div>' +
'    </div>' +
'    <div class="card hidden" id="editCard">' +
'      <div style="font-weight:bold; margin-bottom:10px;" id="editTitle"></div>' +
'      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">' +
'        <div><label>年級</label><select id="f_grade"><option value="1年級">1年級</option><option value="2年級">2年級</option><option value="3年級">3年級</option><option value="4年級">4年級</option><option value="5年級">5年級</option><option value="6年級">6年級</option><option value="幼兒園">幼兒園</option></select></div>' +
'        <div><label>班級</label><input type="text" id="f_class" placeholder="數字"></div>' +
'        <div><label>座號</label><input type="text" id="f_seat" placeholder="數字"></div>' +
'        <div><label>姓名</label><input type="text" id="f_name"></div>' +
'        <div><label>性別</label><input type="text" id="f_gender" placeholder="男/女"></div>' +
'        <div><label>聯絡電話</label><input type="text" id="f_contact"></div>' +
'        <div><label>身高</label><input type="text" id="f_height" placeholder="公分"></div>' +
'        <div><label>體重</label><input type="text" id="f_weight" placeholder="公斤"></div>' +
'      </div>' +
'      <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-top:10px;">' +
'        <div><label>身體症狀</label><input type="text" id="f_symptom"></div>' +
'        <div><label>檢傷初判</label><select id="f_triage"><option value="">（無）</option><option value="休息觀察區">休息觀察區</option><option value="送醫區">送醫區</option></select></div>' +
'        <div><label>目前所在位置</label><input type="text" id="f_location"></div>' +
'        <div><label>就醫醫院</label><input type="text" id="f_hospital"></div>' +
'        <div><label>護送教師</label><input type="text" id="f_escort"></div>' +
'        <div><label>備註</label><textarea id="f_note" rows="2"></textarea></div>' +
'      </div>' +
'      <div style="margin-top:12px;">' +
'        <button onclick="submitEdit()">儲存</button>' +
'        <button style="background:#eee; color:#333;" onclick="closeEditForm()">取消</button>' +
'        <div id="editMsg"></div>' +
'      </div>' +
'    </div>' +
'  </div>' +
'<script>' +
'var PWD = "";' +
'var editMode = "edit";' +
'var editingCaseId = "";' +
'function showMsg(elId, text, ok){' +
'  var el = document.getElementById(elId);' +
'  el.innerHTML = "<div class=\\"msg " + (ok ? "ok" : "error") + "\\">" + text + "</div>";' +
'}' +
'function doLogin(){' +
'  var pwd = document.getElementById("pwdInput").value;' +
'  if (!pwd) { showMsg("loginMsg", "請輸入密碼", false); return; }' +
'  google.script.run.withSuccessHandler(function(ok){' +
'    if (ok) {' +
'      PWD = pwd;' +
'      document.getElementById("loginCard").classList.add("hidden");' +
'      document.getElementById("mainCard").classList.remove("hidden");' +
'      loadList();' +
'    } else { showMsg("loginMsg", "密碼錯誤", false); }' +
'  }).verifyAdminPassword(pwd);' +
'}' +
'function loadList(){' +
'  google.script.run.withSuccessHandler(function(list){' +
'    var html = "<table><thead><tr><th>案件編號</th><th>年級</th><th>班</th><th>座號</th><th>姓名</th><th>性別</th><th>狀態</th><th>結案</th><th>操作</th></tr></thead><tbody>";' +
'    list.forEach(function(c){' +
'      html += "<tr><td>"+c.caseId+"</td><td>"+c.grade+"</td><td>"+c.classNo+"</td><td>"+c.seatNo+"</td>" +' +
'        "<td>"+c.name+"</td><td>"+c.gender+"</td><td>"+c.status+"</td><td>"+c.closed+"</td>" +' +
'        "<td><button class=\\"btn-sm\\" onclick=\\"openEditForm(\\\'"+c.caseId+"\\\')\\" id=\\"btn-"+c.caseId+"\\">編輯</button>" +' +
'        "<button class=\\"btn-sm danger\\" onclick=\\"doDelete({caseId:\\\'"+c.caseId+"\\\',name:\\\'"+c.name+"\\\'})\\">刪除</button></td></tr>";' +
'    });' +
'    if (!list.length) html += "<tr><td colspan=\\"9\\">尚無案件資料</td></tr>";' +
'    html += "</tbody></table>";' +
'    document.getElementById("listWrap").innerHTML = html;' +
'  }).withFailureHandler(function(err){ showMsg("listMsg", err.message, false); })' +
'    .adminListAllCases(PWD);' +
'}' +
'function openAddForm(){' +
'  editMode = "add";' +
'  editingCaseId = "";' +
'  document.getElementById("editTitle").textContent = "新增案件";' +
'  ["f_grade","f_class","f_seat","f_name","f_gender","f_contact","f_height","f_weight","f_symptom","f_triage","f_location","f_hospital","f_escort","f_note"].forEach(function(id){ document.getElementById(id).value = ""; });' +
'  document.getElementById("editCard").classList.remove("hidden");' +
'}' +
'function openEditForm(caseId){' +
'  editMode = "edit";' +
'  editingCaseId = caseId;' +
'  document.getElementById("editTitle").textContent = "編輯案件：" + caseId;' +
'  document.getElementById("editCard").classList.remove("hidden");' +
'  google.script.run.withSuccessHandler(function(list){' +
'    var c = list.find(function(x){ return x.caseId === caseId; });' +
'    if (c) {' +
'      document.getElementById("f_grade").value = c.grade || "";' +
'      document.getElementById("f_class").value = c.classNo || "";' +
'      document.getElementById("f_seat").value = c.seatNo || "";' +
'      document.getElementById("f_name").value = c.name || "";' +
'      document.getElementById("f_gender").value = c.gender || "";' +
'      document.getElementById("f_contact").value = c.contact || "";' +
'      document.getElementById("f_height").value = c.height || "";' +
'      document.getElementById("f_weight").value = c.weight || "";' +
'      document.getElementById("f_symptom").value = c.symptom || "";' +
'      document.getElementById("f_triage").value = c.triage || "";' +
'      document.getElementById("f_location").value = c.location || "";' +
'      document.getElementById("f_hospital").value = c.hospital || "";' +
'      document.getElementById("f_escort").value = c.escort || "";' +
'      document.getElementById("f_note").value = c.note || "";' +
'    }' +
'  }).adminListAllCases(PWD);' +
'}' +
'function closeEditForm(){' +
'  document.getElementById("editCard").classList.add("hidden");' +
'  editingCaseId = "";' +
'}' +
'function collectFields(){' +
'  return {' +
'    grade: document.getElementById("f_grade").value,' +
'    classNo: document.getElementById("f_class").value,' +
'    seatNo: document.getElementById("f_seat").value,' +
'    name: document.getElementById("f_name").value,' +
'    gender: document.getElementById("f_gender").value,' +
'    contact: document.getElementById("f_contact").value,' +
'    height: document.getElementById("f_height").value,' +
'    weight: document.getElementById("f_weight").value,' +
'    symptom: document.getElementById("f_symptom").value,' +
'    triage: document.getElementById("f_triage").value,' +
'    location: document.getElementById("f_location").value,' +
'    hospital: document.getElementById("f_hospital").value,' +
'    escort: document.getElementById("f_escort").value,' +
'    note: document.getElementById("f_note").value' +
'  };' +
'}' +
'function submitEdit(){' +
'  var fields = collectFields();' +
'  if (editMode === "add") {' +
'    google.script.run.withSuccessHandler(function(){' +
'      showMsg("editMsg", "新增成功！", true);' +
'      closeEditForm();' +
'      loadList();' +
'    }).withFailureHandler(function(err){ showMsg("editMsg", "新增失敗："+err.message, false); })' +
'      .adminAddCase(PWD, fields);' +
'  } else {' +
'    google.script.run.withSuccessHandler(function(){' +
'      showMsg("editMsg", "更新成功！", true);' +
'      closeEditForm();' +
'      loadList();' +
'    }).withFailureHandler(function(err){ showMsg("editMsg", "更新失敗："+err.message, false); })' +
'      .adminUpdateCase(PWD, editingCaseId, fields);' +
'  }' +
'}' +
'function doDelete(c){' +
'  if (!confirm("確定要刪除「" + c.name + "（" + c.caseId + "）」這筆案件嗎？此動作無法復原！")) return;' +
'  google.script.run.withSuccessHandler(function(){' +
'    showMsg("listMsg", "已刪除", true);' +
'    loadList();' +
'  }).withFailureHandler(function(err){ showMsg("listMsg", "刪除失敗："+err.message, false); })' +
'    .adminDeleteCase(PWD, c.caseId);' +
'}' +
'</script>' +
'</body></html>';
}
