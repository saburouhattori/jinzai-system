// =========================================
// システム全体の設定値・共通関数・UI
// =========================================


/**
 * シートを取得する共通関数
 * @param {string} sheetName シート名
 * @return {Sheet} シートオブジェクト
 */
function getMasterSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.error(sheetName + " が見つかりません。シート名を確認してください。");
  }
  return sheet;
}

/**
 * ヘッダー名から列番号を取得
 * @param {Sheet} sheet 対象シート
 * @return {Object} ヘッダー名と列番号のマップ
 */
function getMasterColumnMap(sheet) {
  if (!sheet) return {};
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    // スペースや改行を詰めて正規化
    const cleanHeader = String(h).replace(/\n/g, '').replace(/\s/g, '').trim();
    if (cleanHeader) map[cleanHeader] = i + 1;
  });
  return map;
}

/**
 * HTMLファイル読み込み用
 */
function include(filename, mode) {
  const template = HtmlService.createTemplateFromFile(filename);
  template.mode = mode;
  // フォームのモード（NEW/EDIT等）を渡す
  return template.evaluate().getContent();
}


// =========================================
// メニューとUIの表示処理
// =========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // メニュー①：登録者マスタ管理
  ui.createMenu('1. 登録者マスタ管理')
    .addItem('1. 候補者の新規登録', 'showSidebarNew')
    .addItem('2. 登録者データの修正・更新', 'showSidebarEdit')
    .addSeparator()
    .addItem('3. 登録者データの削除', 'showSidebarDelete')
    .addToUi();

  // メニュー②：案件・採用管理
  ui.createMenu('2. 案件・採用管理')
    .addItem('1. 案件の新規登録', 'showSidebarJobNew')
    .addItem('2. 簡易リスト作成', 'showSidebarList')
    .addItem('3. 履歴書作成', 'rirekisyo')
    .addItem('4. 案件情報の修正・削除', 'showSidebarJobEdit')
    .addSeparator()
    .addItem('5. 面接結果の登録', 'showSidebarHire')
    .addItem('6. 面接結果の修正・更新', 'showSidebarHireEdit')
    .addSeparator()
    .addItem('7. COE申請情報の登録・更新 (COE Info)', 'showSidebarAddInfo')
    .addItem('8. 入国日・入職日の一括登録', 'showSidebarEntryDate')
    .addToUi();

  // メニュー③：管理者用（マスタ・同期）
  ui.createMenu('3. 管理者用（マスタ・同期）')
    .addItem('1. 事業者の新規登録', 'showSidebarCompany')
    .addSeparator()
    .addItem('2. 採用・未採用リストの同期', 'runSyncListSheets')
    .addItem('3. 支払い管理シートへ同期（Funtoco）', 'runSyncPaymentManagement')
    .addToUi();
}


/**
 * サイドバー/ダイアログ表示の共通処理
 * @param {string} mode フォームの動作モード
 * @param {string} title ダイアログのタイトル
 */
function showMainSidebar(mode, title) {
  const html = HtmlService.createTemplateFromFile('MainSidebar');
  html.mode = mode;
  const output = html.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(output, title);
}

// 各メニューから呼び出される関数
function showSidebarNew()       { showMainSidebar('NEW',        '候補者の新規登録'); }
function showSidebarEdit()      { showMainSidebar('EDIT',       '登録者データの修正・更新'); }
function showSidebarDelete()    { showMainSidebar('DELETE',     '登録者データの削除'); }

function showSidebarJobNew()    { showMainSidebar('JOB',        '案件の新規登録'); }
function showSidebarList()      { showMainSidebar('LIST',       '簡易リスト作成'); }
function showSidebarJobEdit()   { showMainSidebar('JOB_EDIT',   '案件情報の修正・削除'); }
function showSidebarHire()      { showMainSidebar('HIRE',       '面接結果の登録'); }
function showSidebarHireEdit()  { showMainSidebar('HIRE_EDIT',  '面接結果の修正・更新'); }
function showSidebarAddInfo()   { showMainSidebar('ADDINFO',    'COE申請情報の登録・更新 / COE Application Info'); }
function showSidebarEntryDate() { showMainSidebar('ENTRY_DATE', '入国日・入職日の一括登録'); }

function showSidebarCompany()   { showMainSidebar('COMPANY',    '事業者の新規登録'); }

/**
 * リスト同期の実行と結果表示
 */
function runSyncListSheets() {
  const msg = syncListSheets();
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 支払い管理への同期実行と結果表示
 */
function runSyncPaymentManagement() {
  const msg = syncToPaymentManagement();
  SpreadsheetApp.getUi().alert(msg);
}