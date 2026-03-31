// 【重要】外部マスタ管理スプレッドシートのID（マスタ移植が完了するまで使用します）
const MASTER_SS_ID = '1cq4h6yI0on-bm_MlqUlUi6MMXBlFIXyTCZpzcTZlCMw';

// ⬇️ フィルタ表示呼び出し用のURL設定 ⬇️
const URL_UNADOPTED = "https://docs.google.com/spreadsheets/d/1vwBBwQNvTrZ0jBa1-ZfYmYdEZG6YBwEQeZ8PJ9vkrmQ/edit?gid=1414821006#gid=1414821006&fvid=331083492";
const URL_ADOPTED   = "https://docs.google.com/spreadsheets/d/1vwBBwQNvTrZ0jBa1-ZfYmYdEZG6YBwEQeZ8PJ9vkrmQ/edit?gid=1414821006#gid=1414821006&fvid=1493453362";

/**
 * 【共通】マスタのシートを取得する関数（自シートを優先し、なければ外部IDを見に行く）
 */
function getMasterSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    try {
      sheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(sheetName);
    } catch(e) {
      console.error(sheetName + " が見つかりません: " + e.message);
    }
  }
  return sheet;
}

/**
 * 【共通】マスタのヘッダー名から列番号を取得
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
    .addItem( '【新規】案件登録' , 'showSidebarJobNew')
    .addItem( '【登録】採用者登録' , 'showSidebarHire') 
    .addSeparator()
    .addSubMenu(ui.createMenu('【表示】リスト絞り込み')
      .addItem('未採用者リストを開く', 'openFilterUnadopted')
      .addItem('採用者リストを開く', 'openFilterAdopted')
    )
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
function showSidebarDelete()  { showMainSidebar('DELETE', '【削除】登録者削除'); }
function showSidebarHire()    { showMainSidebar('HIRE', '【登録】採用者登録'); }
function showSidebarList()    { showMainSidebar('LIST', '【作成】簡易リスト出力'); }

function updateAddInfoRow(formData) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) return "エラー：登録者マスタが見つかりません。";
    const col = getMasterColumnMap(sheet);
    const row = Number(formData.row);
    const mapping = {
      agent: '所属送り出し機関', offerDate: '内定日', birthCity: '出生地（都市名）',
      addressDetail: '住所詳細', passportNum: 'パスポート番号', passportExp: 'パスポート有効期限',
      job: '職業', traineeExp: '技能実習の経験の有無', traineeCert: '技能実習修了書の有無',
      crime: '犯罪歴の有無', applyCount: '在留資格交付申請の回数', rejectCount: '不許可となった在留資格交付申請の回数',
      overseasExp: '海外への出入国歴の有無', travelCount: '出入国の回数', lastInDate: '直近の入国日', lastOutDate: '直近の出国日',
      relName2: '日本在住の親族情報親族の名前', relRelation2: '日本在住の親族情報続柄',
      relBirth2: '日本在住の親族情報親族の生年月日', relCountry2: '日本在住の親族情報親族の国籍・地域',
      relLive2: '日本在住の親族情報親族との同居予定の有無', relWork2: '日本在住の親族情報親族の勤務先・通学先',
      relCard2: '日本在住の親族情報親族の在留カード番号', memo: '備考・メモ'
    };
    for (let key in mapping) {
      const h = mapping[key].replace(/\s/g, '');
      if (col[h] && formData[key] !== undefined) sheet.getRange(row, col[h]).setValue(formData[key]);
    }
    return `追加情報の登録が完了しました。`;
  } catch (e) { return "エラー: " + e.message; }
}

function getAgentList() {
  const sheet = getMasterSheet('送り出し機関マスタ');
  return sheet ? [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getSchoolList() {
  const sheet = getMasterSheet('日本語学校マスタ');
  return sheet ? [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getCandidateDict() {
  const sheet = getMasterSheet('登録者マスタ');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const dict = {};
  data.forEach(row => { if (row[0]) dict[String(row[0]).trim()] = String(row[1]); });
  return dict;
}

function openFilterUnadopted() { showLinkDialog(URL_UNADOPTED, '未採用者リスト'); }
function openFilterAdopted()   { showLinkDialog(URL_ADOPTED, '採用者リスト'); }

function showLinkDialog(url, title) {
  const html = `<div style="text-align:center;padding:20px;"><a href="${url}" target="_blank" style="padding:12px;background:#1a73e8;color:white;text-decoration:none;border-radius:4px;">${title}を開く</a></div>`;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(320).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}