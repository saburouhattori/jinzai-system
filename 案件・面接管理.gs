/**
 * 採用者一括登録処理（ヘッダー名ベース）
 */
function registerHire(jobId, candIds) {
  try {
    const jobSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
    const jobData = jobSheet.getDataRange().getValues();
    let jobRow = -1, companyName = "", existingHiredStr = "", interviewDateRaw = "";
    for (let i = 1; i < jobData.length; i++) {
      if (jobData[i][0] === jobId) {
        jobRow = i + 1; companyName = jobData[i][3];
        existingHiredStr = jobData[i][5] || ""; interviewDateRaw = jobData[i][6] || "";
        break;
      }
    }
    if (jobRow === -1) return "エラー: 案件IDが見つかりません。";

    const interviewDateStr = interviewDateRaw instanceof Date ? Utilities.formatDate(interviewDateRaw, "JST", "yyyy/MM/dd") : String(interviewDateRaw || "日付不明");
    const masterSheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('登録者マスタ');
    const candData = masterSheet.getDataRange().getValues();
    const col = getMasterColumnMap(masterSheet);

    const hiredNames = [];
    for (let i = 1; i < candData.length; i++) {
      const cId = candData[i][0];
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
    if (jobData[i][0] === jobId) {
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

function addJob(company, candIds) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
  const lastRow = sheet.getLastRow();
  let nextId = "JOB-0001";
  if (lastRow >= 2) {
    const lastId = sheet.getRange(lastRow, 1).getValue().toString();
    const nextNum = (parseInt(lastId.match(/\d+/)[0], 10) + 1);
    nextId = "JOB-" + nextNum.toString().padStart(4, '0');
  }
  const candDict = getCandidateDict();
  const candList = candIds.map(id => id + "-" + (candDict[id] || "不明")).join('\n');
  const dateStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  sheet.appendRow([nextId, "面接待ち", dateStr, company, candList]);
  return `案件登録完了: ${nextId}`;
}

function generateSimpleList(candIds) {
  try {
    const masterSheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('登録者マスタ');
    const masterData = masterSheet.getDataRange().getValues();
    const col = getMasterColumnMap(masterSheet);
    const listSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('簡易リスト');
    listSheet.getRange('B2:L51').clearContent();

    const result = [];
    const formulas = [];
    candIds.forEach(id => {
      let rowData = null;
      for (let i = 1; i < masterData.length; i++) { if (String(masterData[i][0]) === id) { rowData = masterData[i]; break; } }
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
  const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
  const sheet = masterSs.getSheetByName('事業者マスタ');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[1]).filter(n => n);
}

function checkCompanyDuplicate(name) {
  const list = getCompanyList();
  return list.indexOf(name) !== -1;
}

function addCompany(formData) {
  const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
  const sheet = masterSs.getSheetByName('事業者マスタ');
  const lastRow = sheet.getLastRow();
  let nextId = "CO-0001";
  if (lastRow >= 2) {
    const lastId = sheet.getRange(lastRow, 1).getValue().toString();
    const nextNum = (parseInt(lastId.match(/\d+/)[0], 10) + 1);
    nextId = "CO-" + nextNum.toString().padStart(4, '0');
  }
  sheet.appendRow([nextId, formData.name, formData.yomi, formData.address, formData.url, formData.note]);
  return `事業者登録完了: ${nextId}`;
}