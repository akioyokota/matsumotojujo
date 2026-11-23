/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー')
    .addItem('翌日のシートを作成', 'createSheetForNextDay')
    .addItem('抽出元ファイルを登録', 'registerSourceSpreadsheet')
    .addToUi();
}

// 抽出元（アサイン表）を探すときの基準ファイル。
// このファイルと同じフォルダから、月ごとのアサイン表を自動で探します。
const REFERENCE_SOURCE_URL = "https://docs.google.com/spreadsheets/d/1bTKGyMFzNVTsaEDvLnbDtFioEdyW5pag8cLO7NZPWGI/edit?gid=1385730121#gid=1385730121";

// 月ごとの抽出元URLを覚えておくスクリプトプロパティのキー接頭辞。
const SOURCE_URL_PREFIX = 'SOURCE_URL_';

/**
 * 日付シートを探す共通関数。
 * 「8/4」「08/04」だけでなく、「8/4し」「8/8し(後半〆)」「8/12仮」のように
 * 日付の後ろに文字が付いたシート名にも一致します。
 * ※「8/1」で「8/10」を誤って拾わないよう、日付の直後に数字が続く名前は除外しています。
 */
function findDateSheet(sheets, month, day) {
  const pattern = new RegExp('^0?' + month + '\\/0?' + day + '(?![0-9])');
  return sheets.find(s => pattern.test(s.getName().trim()));
}

/**
 * 日付から「2026-08」形式の月キーを作ります。
 */
function getMonthKey(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

function openSpreadsheetByUrl(url) {
  try {
    return SpreadsheetApp.openByUrl(url);
  } catch (e) {
    return null;
  }
}

/**
 * 指定した日付の月に対応するアサイン表を返します。
 * 登録済みURL → 同じフォルダ内の自動検索、の順に探し、
 * 見つかったURLは次回以降のためにスクリプトプロパティへ保存します。
 */
function getSourceSpreadsheetFor(date, cache) {
  const monthKey = getMonthKey(date);
  if (cache[monthKey]) return cache[monthKey];

  const properties = PropertiesService.getScriptProperties();
  const savedUrl = properties.getProperty(SOURCE_URL_PREFIX + monthKey);
  let spreadsheet = savedUrl ? openSpreadsheetByUrl(savedUrl) : null;

  if (!spreadsheet) {
    spreadsheet = searchSourceSpreadsheet(date, monthKey);
    properties.setProperty(SOURCE_URL_PREFIX + monthKey, spreadsheet.getUrl());
  }

  cache[monthKey] = spreadsheet;
  return spreadsheet;
}

/**
 * 基準ファイルと同じフォルダから、その月のアサイン表をファイル名で探します。
 * 1つに絞り込めないときは、対処方法を添えて例外を投げます。
 */
function searchSourceSpreadsheet(date, monthKey) {
  const reference = openSpreadsheetByUrl(REFERENCE_SOURCE_URL);
  if (!reference) {
    throw new Error(`基準となる抽出元ファイルが開けません。メニューの「抽出元ファイルを登録」から ${monthKey} のURLを登録してください。`);
  }

  const parents = DriveApp.getFileById(reference.getId()).getParents();
  if (!parents.hasNext()) {
    throw new Error(`抽出元ファイルの入っているフォルダが特定できません。メニューの「抽出元ファイルを登録」から ${monthKey} のURLを登録してください。`);
  }

  const month = date.getMonth() + 1;
  const year = date.getFullYear();
  const monthPattern = new RegExp('(^|[^0-9])0?' + month + '\\s*月');
  const candidates = [];

  const files = parents.next().getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (!monthPattern.test(name)) continue;
    // ファイル名に年が入っている場合は、その年のものだけを候補にする
    const yearInName = name.match(/20[0-9]{2}/);
    if (yearInName && Number(yearInName[0]) !== year) continue;
    candidates.push(file);
  }

  if (candidates.length === 1) {
    return SpreadsheetApp.openById(candidates[0].getId());
  }

  const detail = candidates.length === 0
    ? '該当するファイルが見つかりませんでした。'
    : `候補が複数見つかりました（${candidates.map(f => f.getName()).join('、')}）。`;
  throw new Error(`${monthKey} のアサイン表を自動で特定できません。${detail}メニューの「抽出元ファイルを登録」からURLを登録してください。`);
}

/**
 * 月ごとの抽出元URLを手動で登録します（カスタムメニューから実行）。
 */
function registerSourceSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1, 1);

  const monthResponse = ui.prompt(
    '抽出元ファイルの登録',
    `対象の月を yyyy-MM 形式で入力してください（例：${getMonthKey(nextMonth)}）`,
    ui.ButtonSet.OK_CANCEL
  );
  if (monthResponse.getSelectedButton() !== ui.Button.OK) return;

  const monthKey = monthResponse.getResponseText().trim();
  if (!/^[0-9]{4}-[0-9]{2}$/.test(monthKey)) {
    ui.alert('月の形式が正しくありません。yyyy-MM 形式で入力してください。');
    return;
  }

  const urlResponse = ui.prompt(
    '抽出元ファイルの登録',
    `${monthKey} のアサイン表のURLを貼り付けてください`,
    ui.ButtonSet.OK_CANCEL
  );
  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;

  const spreadsheet = openSpreadsheetByUrl(urlResponse.getResponseText().trim());
  if (!spreadsheet) {
    ui.alert('そのURLのスプレッドシートを開けません。URLと共有設定を確認してください。');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(SOURCE_URL_PREFIX + monthKey, spreadsheet.getUrl());
  ui.alert(`${monthKey} の抽出元を「${spreadsheet.getName()}」に設定しました。`);
}

/**
 * 翌日分の清掃管理シートを自動で作成・設定します。
 */
function createSheetForNextDay() {
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // =================================================================
  // 💡【重要】日付がズレてエラーになる場合は、下に 'yyyy-MM-dd' を設定してください。
  //    空文字のままなら実行日を使います。
  // =================================================================
  const forceDate = '';

  const todayDate = forceDate ? new Date(forceDate + 'T12:00:00') : new Date();
  const nextDate = new Date(todayDate.getTime());
  nextDate.setDate(nextDate.getDate() + 1);

  console.log("基準にする日時:", todayDate.toString());

  const todayMonth = todayDate.getMonth() + 1;
  const todayDay = todayDate.getDate();
  const nextMonth = nextDate.getMonth() + 1;
  const nextDay = nextDate.getDate();

  const todaySheetName = todayMonth + '/' + todayDay;
  const nextDaySheetName = nextMonth + '/' + nextDay;

  console.log(`探索する当日シート名: 「${todaySheetName}」で始まるシート`);
  console.log(`探索する翌日シート名: 「${nextDaySheetName}」で始まるシート`);

  // =================================================================
  // 1. 抽出元スプレッドシートの準備
  // ※ アサイン表は月ごとにファイルが分かれているため、月末に翌月分を作るときは
  //    当日と翌日で別々のファイルを参照します。
  // =================================================================

  const sourceCache = {};
  let sourceForToday, sourceForNextDay;
  try {
    sourceForToday = getSourceSpreadsheetFor(todayDate, sourceCache);
    sourceForNextDay = getSourceSpreadsheetFor(nextDate, sourceCache);
  } catch (e) {
    targetSpreadsheet.toast(e.message, "エラー", 30);
    return;
  }

  console.log(`当日の抽出元: ${sourceForToday.getName()} / 翌日の抽出元: ${sourceForNextDay.getName()}`);

  // 翌日のシートを抽出元から取得（「8/4し」等の接尾辞付きにも一致）
  const sourceSheetNextDay = findDateSheet(sourceForNextDay.getSheets(), nextMonth, nextDay);

  // 当日のシートを抽出元から取得
  const sourceSheetToday = findDateSheet(sourceForToday.getSheets(), todayMonth, todayDay);

  if (!sourceSheetNextDay) {
    targetSpreadsheet.toast(`抽出元（${sourceForNextDay.getName()}）に「${nextDaySheetName}」で始まるシートが見つかりません。`, "エラー", 30);
    return;
  }
  if (!sourceSheetToday) {
    targetSpreadsheet.toast(`抽出元（${sourceForToday.getName()}）に当日の「${todaySheetName}」で始まるシートが見つかりません。`, "エラー", 30);
    return;
  }

  // 操作元（コピー元）スプレッドシートのテンプレート（今日のシート）を取得
  const templateSheet = findDateSheet(targetSpreadsheet.getSheets(), todayMonth, todayDay);

  if (!templateSheet) {
    targetSpreadsheet.toast(`このスプレッドシート内に、テンプレートとなる当日のシート（${todaySheetName}）が見つかりません。前日分のシートを手動で複製して「${todaySheetName}」にリネームしてから再実行してください。`, "エラー", 10);
    return;
  }

  if (findDateSheet(targetSpreadsheet.getSheets(), nextMonth, nextDay)) {
    targetSpreadsheet.toast(`シート「${nextDaySheetName}」はすでに作成済みです。`, "確認", 10);
    return;
  }

  // =================================================================
  // 2. データの取得 (26部屋分：4 + 22)
  // =================================================================

  const reservationS = sourceSheetNextDay.getRange("D3:D6").getValues().concat(sourceSheetNextDay.getRange("D27:D48").getValues());
  const reservationS_V = sourceSheetNextDay.getRange("S3:V6").getValues().concat(sourceSheetNextDay.getRange("S27:V48").getValues());
  const memberStatusAToday = sourceSheetToday.getRange("A3:A6").getValues().concat(sourceSheetToday.getRange("A27:A48").getValues());

  const sValuesToday = templateSheet.getRange("T3:T28").getValues();
  const d_mValuesToday = templateSheet.getRange("D3:M28").getValues();
  const o_pValuesToday = templateSheet.getRange("O3:P28").getValues();

  // =================================================================
  // 3. 新しいシートの作成とロジック計算 (26回ループ)
  // =================================================================

  const newSheet = templateSheet.copyTo(targetSpreadsheet).setName(nextDaySheetName);
  let bValues = [], cValues = [], d_mValues = [], o_pValues = [];

  for (let i = 0; i < 26; i++) {
    const sToday = String(sValuesToday[i][0]).trim();
    const sNew = String(reservationS[i][0]).trim();

    let aValue = "";
    if (memberStatusAToday[i] && memberStatusAToday[i][0] !== undefined) {
      aValue = String(memberStatusAToday[i][0]).trim().toUpperCase();
    }
    const isMarriott = (aValue === "M" || aValue === "Ｍ");

    let bValue, cValue;
    let new_d_m_row, new_o_p_row;

    if (sToday === "") {
      bValue = '×';
    } else {
      if (sNew !== "") {
        const match = sNew.match(/(\d+)\/(\d+)/);
        if (match) {
          const current = parseInt(match[1], 10);
          bValue = (current > 1) ? '連泊' : 'IN';
        } else { bValue = 'IN'; }
      } else { bValue = 'IN'; }
    }

    if (isMarriott && bValue === 'IN') {
      bValue = 'IN(M)';
    }

    if (sNew !== "") {
      const match = sNew.match(/(\d+)\/(\d+)/);
      if (match) {
        const current = parseInt(match[1], 10), total = parseInt(match[2], 10);
        cValue = (current === 1) ? '◯' : ((current === total) ? ((total % 3 === 0) ? '連泊A' : '連泊×') : '連泊×');
      } else { cValue = '◯'; }
    } else { cValue = '×'; }

    const { d_m, o_p } = determineCleaningTasks(bValue, cValue);

    if (sToday === "") {
      new_d_m_row = d_mValuesToday[i];
      new_o_p_row = o_p;
    } else {
      new_d_m_row = d_m;
      new_o_p_row = o_p;
    }

    bValues.push([bValue]);
    cValues.push([cValue]);
    d_mValues.push(new_d_m_row);
    o_pValues.push(new_o_p_row);
  }

  // =================================================================
  // 4. 書き出し (3行目から28行目まで)
  // =================================================================

  newSheet.getRange("T3:T28").setValues(reservationS);
  newSheet.getRange("U3:X28").setValues(reservationS_V);

  newSheet.getRange("B3:B28").setValues(bValues);
  newSheet.getRange("C3:C28").setValues(cValues);
  newSheet.getRange("D3:M28").setValues(d_mValues);
  newSheet.getRange("O3:P28").setValues(o_pValues);

  setupFixedRanges(newSheet);
  moveSheetToFront(newSheet);
  targetSpreadsheet.toast(`シート「${nextDaySheetName}」を作成しました。`);
}

/**
 * 手動編集トリガー (26部屋対応)
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  if (row < 3 || row > 28) return;

  if (col === 2) {
    const bValue = range.getValue();
    const cValue = sheet.getRange(row, 3).getValue();
    if (bValue === "IN" || bValue === "IN(M)" || cValue === "連泊A") {
      const { d_m } = determineCleaningTasks(bValue, cValue);
      sheet.getRange(row, 4, 1, 10).setValues([d_m]);
    }
  } else if (col === 3) {
    const cValue = range.getValue();
    const bValue = sheet.getRange(row, 2).getValue();
    const { d_m, o_p } = determineCleaningTasks(bValue, cValue);
    if (["ｱｰﾘｰ", "◯", "IN17", "×"].includes(cValue)) {
      sheet.getRange(row, 15, 1, 2).setValues([o_p]);
    } else {
      sheet.getRange(row, 4, 1, 10).setValues([d_m]);
      sheet.getRange(row, 15, 1, 2).setValues([o_p]);
    }
  }
}

/**
 * 清掃判定ロジック
 */
function determineCleaningTasks(bValue, cValue) {
  let d_m = Array(10).fill('');
  let o_p = Array(2).fill('');
  if (bValue === "IN" || bValue === "IN(M)") d_m.fill('未');

  switch (cValue) {
    case "連泊A":
      d_m.fill('未');
      o_p = ['', ''];
      break;
    case "連泊×":
      o_p = ['', ''];
      break;
    case "ｱｰﾘｰ":
    case "IN17":
    case "◯":
      o_p = ['未', ''];
      break;
    case "×":
      o_p = ['', '未'];
      break;
  }
  return { d_m, o_p };
}

function moveSheetToFront(sheet) {
  const ss = sheet.getParent();
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}

function setupFixedRanges(sheet) {
  const rangesToSetMikan = [
    'E31:E38', 'H44:H62', 'O44:O63',
    'Q3:Q6', 'Q7:Q9', 'Q10:Q13', 'Q14:Q17', 'Q18:Q19', 'Q20'
  ];
  if (rangesToSetMikan.length > 0) {
    sheet.getRangeList(rangesToSetMikan).setValue('未');
  }
}
