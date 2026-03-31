/**
 * 採用者一括登録処理
 */
function registerHire(jobId, candIds) {
  try {
    const localSs = SpreadsheetApp.getActiveSpreadsheet();
    const jobSheet = localSs.getSheetByName('案件管理');
    const jobData = jobSheet.getDataRange().getValues();
    let jobRow = -1, companyName = "", existingHiredStr = "", interviewDateRaw = "";
    
    for (let i = 1; i < jobData.length; i++) {
      if (String(jobData[i][0]).trim() === String(jobId).trim()) {
        jobRow = i + 1; companyName = jobData[i][3];
        existingHiredStr = jobData[i][5] || ""; interviewDateRaw = jobData[i][6] || "";
        break;
      }
    }
    if (jobRow === -1) return "エラー: 案件IDが見つかりません。";

    const interviewDateStr = interviewDateRaw instanceof Date ? Utilities.formatDate(interviewDateRaw, "JST", "yyyy/MM/dd") : String(interviewDateRaw || "日付不明");
    const masterSheet = getMasterSheet('登録者マスタ');
    const candData = masterSheet.getDataRange().getValues();
    const col = getMasterColumnMap(masterSheet);

    const hiredNames = [];
    for (let i = 1; i < candData.length; i++) {
      const cId = String(candData[i][0] || "").trim();
      if (candIds.indexOf(cId) !== -1) {
        const row = i + 1;
        const addText = "【採用】" + interviewDateStr + "：" + companyName;
        if (col['面接履歴']) {
          const currentHistory = candData[i][col['面接履歴']-1] || "";
          masterSheet.getRange(row, col['面接履歴']).setValue(currentHistory ? currentHistory + "\n" + addText : addText);
        }
        if (col['ステータス']) masterSheet.getRange(row, col['ステータス']).setValue("採用");
        if (col['採用事業者']) masterSheet.getRange(row, col['採用事業者']).setValue(companyName);
        hiredNames.push(cId + "-" + candData[i][1]);
      }
    }

    let currentHiredArr = existingHiredStr ? String(existingHiredStr).split('\n') : [];
    hiredNames.forEach(name => { if (currentHiredArr.indexOf(name) === -1) currentHiredArr.push(name); });
    jobSheet.getRange(jobRow, 6).setValue(currentHiredArr.join('\n')); 
    jobSheet.getRange(jobRow, 2).setValue("入国準備"); 

    return `採用登録完了: ${hiredNames.length}名`;
  } catch (e) { return "エラー: " + e.message; }
}

function getJobCandidates(jobId) {
  const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
  const jobData = jobSheet.getDataRange().getValues();
  for (let i = 1; i < jobData.length; i++) {
    if (String(jobData[i][0]).trim() === String(jobId).trim()) {
      const candidatesStr = String(jobData[i][4] || "");
      return candidatesStr.split(/\r?\n|,/).map(c => {
        const str = c.trim();
        if (!str) return null;
        const match = str.match(/^([a-zA-Z]+-\d+)/);
        return { id: match ? match[1] : str.split('-')[0].trim(), display: str };
      }).filter(x => x);
    }
  }
  return [];
}

// ★修正：新しい項目を受け取ってシートに保存
function addJob(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
  const lastRow = sheet.getLastRow();
  let nextId = "JOB-0001";
  
  if (lastRow >= 2) {
    const lastId = String(sheet.getRange(lastRow, 1).getValue());
    const match = lastId.match(/\d+/);
    if (match) {
      const nextNum = parseInt(match[0], 10) + 1;
      nextId = "JOB-" + nextNum.toString().padStart(4, '0');
    }
  }
  
  const candDict = getCandidateDict();
  const candList = formData.candidates.map(id => id + "-" + (candDict[id] || "不明")).join('\n');
  const dateStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  // A:案件ID, B:ステータス, C:案件登録日, D:事業者名, E:候補者名, F:面接日, G:採用者氏名, H:関連ファイル, I:備考・メモ
  sheet.appendRow([
    nextId, 
    "面接待ち", 
    dateStr, 
    formData.company, 
    candList, 
    formData.interviewDate, 
    "", 
    formData.relatedFile, 
    formData.memo
  ]);
  
  return `案件登録完了: ${nextId}`;
}

function generateSimpleList(candIds) {
  try {
    const masterSheet = getMasterSheet('登録者マスタ');
    const masterData = masterSheet.getDataRange().getValues();
    const col = getMasterColumnMap(masterSheet);
    const listSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('簡易リスト');
    listSheet.getRange('B2:L51').clearContent();

    const result = [];
    const formulas = [];
    candIds.forEach(id => {
      let rowData = null;
      const sid = String(id).trim().toUpperCase();
      for (let i = 1; i < masterData.length; i++) { if (String(masterData[i][0]).trim().toUpperCase() === sid) { rowData = masterData[i]; break; } }
      if (rowData) {
        const getVal = (name) => col[name.replace(/\s/g, '')] ? rowData[col[name.replace(/\s/g, '')]-1] : "";
        result.push([getVal('名前'), getVal('フリガナ'), getVal('満年齢'), getVal('性別'), getVal('学歴＞学校名'), getVal('学歴＞状況'), 
                     getVal('特定技能要件＞JLPTレベル') || "×", getVal('特定技能要件＞JFTBasicレベル') || "×", getVal('その他の日本語能力試験'), id]);
        formulas.push(['=IFERROR(VLOOKUP(L' + (result.length + 1) + ', \'候補者写真\'!$A:$B, 2, FALSE), "")']);
      }
    });
    if (result.length) {
      listSheet.getRange(2, 3, result.length, 10).setValues(result);
      listSheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
    }
    return `${result.length}名の簡易リストを作成しました。`;
  } catch(e) { return "エラー: " + e.message; }
}

function getCompanyList() {
  const sheet = getMasterSheet('事業者マスタ');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[1]).filter(n => n);
}

function checkCompanyDuplicate(name) {
  const list = getCompanyList();
  return list.indexOf(name.trim()) !== -1;
}

function addCompany(formData) {
  const sheet = getMasterSheet('事業者マスタ');
  if (!sheet) return "エラー：事業者マスタが見つかりません。";
  const lastRow = sheet.getLastRow();
  let nextId = "CO-0001";
  if (lastRow >= 2) {
    const lastId = String(sheet.getRange(lastRow, 1).getValue());
    const nextNum = (parseInt(lastId.match(/\d+/)[0], 10) + 1);
    nextId = "CO-" + nextNum.toString().padStart(4, '0');
  }
  sheet.appendRow([nextId, formData.name, formData.yomi, formData.address, formData.url, formData.note]);
  return `事業者登録完了: ${nextId}`;
}

function getJobDict() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('案件管理');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    const dict = {};
    for (let i = 1; i < data.length; i++) {
      const jobId = String(data[i][0]).trim();
      const company = String(data[i][3]).trim();
      if (jobId) { dict[jobId] = company; }
    }
    return dict;
  } catch(e) { return {}; }
}