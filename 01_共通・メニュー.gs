// =========================================
// システム全体の設定値・共通関数・UI
// =========================================

/**
 * シートを取得する共通関数
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
 */
function getMasterColumnMap(sheet) {
  if (!sheet) return {};
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const cleanHeader = String(h).replace(/\n/g, '').replace(/\s/g, '').trim();
    if (cleanHeader) map[cleanHeader] = i + 1;
  });
  return map;
}

/**
 * HTMLファイル読み込み用（★修正：安全な読み込み方式）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =========================================
// メニューとUIの表示処理
// =========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu( '人材事業メニュー' )
    .addItem( '候補者登録' , 'showSidebarNew')
    .addItem( 'データ更新' , 'showSidebarEdit')
    .addItem( '採用者情報登録' , 'showSidebarAddInfo')
    .addItem( 'コメント登録' , 'showSidebarComment')
    .addSeparator()
    .addItem( '登録者削除' , 'showSidebarDelete')
    .addSeparator()
    .addItem( '事業者マスタ登録' , 'showSidebarCompany')
    .addSeparator()
    .addItem( '履歴書出力' , 'rirekisyo') 
    .addItem( '簡易リスト出力' , 'showSidebarList') 
    .addSeparator()
    .addItem('採用・未採用リストの同期', 'runSyncListSheets')
    .addToUi();
    
  ui.createMenu( '案件・採用管理' )
    .addItem( '案件登録' , 'showSidebarJobNew')
    .addItem( '案件更新/削除' , 'showSidebarJobEdit')
    .addItem( '採用者登録' , 'showSidebarHire') 
    .addToUi();
}

/**
 * サイドバー/ダイアログ表示の共通処理
 */
function showMainSidebar(mode, title, prefillName) {
  // ★修正：テンプレートを正しく評価（evaluate）し、HTML構造を崩さずに表示します
  const template = HtmlService.createTemplateFromFile('MainSidebar');
  template.mode = mode || "";
  template.prefillName = prefillName || "";
  
  const output = template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(output, title);
}

function showSidebarNew()     { showMainSidebar('NEW',  '候補者登録' ); }
function showSidebarEdit()    { showMainSidebar('EDIT',  'データ更新' ); }
function showSidebarAddInfo() { showMainSidebar('ADDINFO',  '採用者情報登録' ); }
function showSidebarComment() { showMainSidebar('COMMENT',  'コメント登録' ); }
function showSidebarCompany(prefillName) { showMainSidebar('COMPANY',  '事業者マスタ登録', typeof prefillName === 'string' ? prefillName : "" ); }
function showSidebarJobNew()  { showMainSidebar('JOB',  '案件登録' ); }
function showSidebarJobEdit() { showMainSidebar('JOB_EDIT', '案件更新/削除' ); }
function showSidebarDelete()  { showMainSidebar('DELETE', '登録者削除'); }
function showSidebarHire()    { showMainSidebar('HIRE', '採用者登録'); }
function showSidebarList()    { showMainSidebar('LIST', '簡易リスト出力'); }

function runSyncListSheets() {
  const msg = syncListSheets();
  SpreadsheetApp.getUi().alert(msg);
}