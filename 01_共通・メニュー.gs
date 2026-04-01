// =========================================
// システム全体の設定値・共通関数・UI
// =========================================

const URL_UNADOPTED = "https://docs.google.com/spreadsheets/d/1vwBBwQNvTrZ0jBa1-ZfYmYdEZG6YBwEQeZ8PJ9vkrmQ/edit?gid=1414821006#gid=1414821006&fvid=331083492";
const URL_ADOPTED   = "https://docs.google.com/spreadsheets/d/1vwBBwQNvTrZ0jBa1-ZfYmYdEZG6YBwEQeZ8PJ9vkrmQ/edit?gid=1414821006#gid=1414821006&fvid=1493453362";

function getMasterSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.error(sheetName + " が見つかりません。シート名を確認してください。");
  }
  return sheet;
}

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

function include(filename, mode) {
  const template = HtmlService.createTemplateFromFile(filename);
  template.mode = mode; 
  return template.evaluate().getContent();
}

// =========================================
// メニューとUIの表示処理
// =========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu( '人材事業メニュー' )
    .addItem( '【新規】候補者登録' , 'showSidebarNew')
    .addItem( '【修正】データ更新' , 'showSidebarEdit')
    .addItem( '【追加】採用者情報登録' , 'showSidebarAddInfo')
    .addItem( '【コメント登録】' , 'showSidebarComment')
    .addSeparator()
    .addItem( '【削除】登録者削除' , 'showSidebarDelete')
    .addSeparator()
    .addItem( '【事業者】マスタ登録' , 'showSidebarCompany')
    .addSeparator()
    .addItem( '【作成】履歴書出力' , 'rirekisyo') 
    .addItem( '【作成】簡易リスト出力' , 'showSidebarList') 
    .addSeparator()
    .addItem('【更新】採用・未採用リストの同期', 'runSyncListSheets') // ★追加
    .addSubMenu(ui.createMenu('【表示】リスト絞り込み')
      .addItem('未採用者リストを開く', 'openFilterUnadopted')
      .addItem('採用者リストを開く', 'openFilterAdopted')
    )
    .addToUi();

  ui.createMenu( '案件・採用管理' )
    .addItem( '【新規】案件登録' , 'showSidebarJobNew')
    .addItem( '【修正】案件更新/削除' , 'showSidebarJobEdit')
    .addItem( '【登録】採用者登録' , 'showSidebarHire') 
    .addToUi();
}

function showMainSidebar(mode, title) {
  const html = HtmlService.createTemplateFromFile('MainSidebar');
  html.mode = mode;
  const output = html.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(output, title);
}

function showSidebarNew()     { showMainSidebar('NEW',  '【新規】候補者登録' ); }
function showSidebarEdit()    { showMainSidebar('EDIT',  '【修正】データ更新' ); }
function showSidebarAddInfo() { showMainSidebar('ADDINFO',  '【追加】採用者情報登録' ); }
function showSidebarComment() { showMainSidebar('COMMENT',  '【コメント登録】' ); }
function showSidebarCompany() { showMainSidebar('COMPANY',  '【事業者】マスタ登録' ); }
function showSidebarJobNew()  { showMainSidebar('JOB',  '【新規】案件登録' ); }
function showSidebarJobEdit() { showMainSidebar('JOB_EDIT', '【修正】案件更新/削除' ); }
function showSidebarDelete()  { showMainSidebar('DELETE', '【削除】登録者削除'); }
function showSidebarHire()    { showMainSidebar('HIRE', '【登録】採用者登録'); }
function showSidebarList()    { showMainSidebar('LIST', '【作成】簡易リスト出力'); }

// ★追加：同期メニュー用関数
function runSyncListSheets() {
  const msg = syncListSheets();
  SpreadsheetApp.getUi().alert(msg);
}

function openFilterUnadopted() { showLinkDialog(URL_UNADOPTED, '未採用者リスト'); }
function openFilterAdopted()   { showLinkDialog(URL_ADOPTED, '採用者リスト'); }
function showLinkDialog(url, title) {
  const html = `<div style="text-align:center;padding:20px;"><a href="${url}" target="_blank" style="padding:12px;background:#1a73e8;color:white;text-decoration:none;border-radius:4px;">${title}を開く</a></div>`;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(320).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}