// =========================================
// 各種リスト・辞書取得処理（プルダウンやID検索用）
// =========================================

function getSchoolList() {
  const sheet = getMasterSheet('日本語学校マスタ');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
}

function getCompanyList() {
  const sheet = getMasterSheet('事業者マスタ');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat().filter(String);
}

function getAgentList() {
  const sheet = getMasterSheet('送り出し機関マスタ');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
}

function getCandidateDict() {
  const sheet = getMasterSheet('登録者マスタ');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const dict = {};
  data.forEach(row => {
    if (row[0]) dict[String(row[0])] = row[1];
  });
  return dict;
}

function getJobDict() {
  const sheet = getMasterSheet('案件管理');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const dict = {};
  data.forEach(row => {
    if (row[0]) dict[String(row[0])] = row[3];
  });
  return dict;
}

/**
 * ★UI初期化時の通信を1回にまとめるための関数（通信渋滞・遅延の解消）
 */
function getInitialData(mode) {
  const data = {};
  
  if (mode === 'NEW' || mode === 'EDIT') {
    data.schools = getSchoolList();
  }
  
  if (mode === 'JOB' || mode === 'JOB_EDIT') {
    data.companies = getCompanyList();
  }
  
  if (mode === 'ADDINFO') {
    data.agents = getAgentList();
  }
  
  data.candidateDict = getCandidateDict();
  
  if (mode === 'HIRE') {
    data.jobDict = getJobDict();
  }
  
  return data;
}

// =========================================
// その他ユーティリティ・外部連携処理
// =========================================

function searchDriveFiles(query) {
  if (!query) return [];
  const files = DriveApp.searchFiles("title contains '" + query.replace(/'/g, "\\'") + "' and trashed = false");
  const result = [];
  let count = 0;
  while (files.hasNext() && count < 20) {
    const f = files.next();
    result.push({ name: f.getName(), url: f.getUrl() });
    count++;
  }
  return result;
}