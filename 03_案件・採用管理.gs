// =========================================
// 案件登録・修正・採用者処理
// =========================================

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
  
  sheet.appendRow([
    nextId, "未着手", dateStr, formData.company, candList, 
    formData.interviewDate, "", formData.relatedFile, formData.memo
  ]);
  
  const newRow = sheet.getLastRow();
  
  if (formData.relatedFile && formData.relatedFile.startsWith("http")) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetId = sheet.getSheetId();
      const req = {
        "updateCells": {
          "range": { "sheetId": sheetId, "startRowIndex": newRow - 1, "endRowIndex": newRow, "startColumnIndex": 7, "endColumnIndex": 8 },
          "rows": [{ "values": [{ "userEnteredValue": { "stringValue": "@" }, "chipRuns": [{ "startIndex": 0, "chip": { "richLinkProperties": { "uri": formData.relatedFile } } }] }] }],
          "fields": "userEnteredValue,chipRuns"
        }
      };
      const res = UrlFetchApp.fetch(`https://sheets.googleapis.com/v4/spreadsheets/${ss.getId()}:batchUpdate`, {
        method: "post", contentType: "application/json",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ requests: [req] }), muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) throw new Error(res.getContentText());
    } catch(e) {
      const displayLabel = formData.relatedFileName ? formData.relatedFileName : "関連ファイル";
      const richText = SpreadsheetApp.newRichTextValue().setText(displayLabel).setLinkUrl(formData.relatedFile).build();
      sheet.getRange(newRow, 8).setRichTextValue(richText);
    }
  }
  return `案件登録完了: ${nextId}`;
}

/**
 * ★追加：案件IDから詳細データを取得する
 */
function getJobDetails(jobId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
  const data = sheet.getDataRange().getValues();
  const targetId = String(jobId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === targetId) {
      const row = data[i];
      
      // 関連ファイルのURL（スマートチップ等のリンク）を抽出
      let fileUrl = row[7] || "";
      const richText = sheet.getRange(i + 1, 8).getRichTextValue();
      if (richText && richText.getLinkUrl()) {
         fileUrl = richText.getLinkUrl();
      }

      return {
        row: i + 1,
        id: row[0],
        status: row[1],
        company: row[3],
        candidates: row[4],
        interviewDate: row[5] instanceof Date ? Utilities.formatDate(row[5], "JST", "yyyy-MM-dd") : row[5],
        relatedFile: fileUrl,
        memo: row[8]
      };
    }
  }
  return null;
}

/**
 * ★追加：案件データの更新
 */
function updateJob(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
  const row = Number(formData.row);
  if (!row) return "エラー：対象の行が見つかりません。";

  const candDict = getCandidateDict();
  const candList = formData.candidates.map(id => id + "-" + (candDict[id] || "不明")).join('\n');

  sheet.getRange(row, 2).setValue(formData.status);
  sheet.getRange(row, 4).setValue(formData.company);
  sheet.getRange(row, 5).setValue(candList);
  sheet.getRange(row, 6).setValue(formData.interviewDate);
  sheet.getRange(row, 9).setValue(formData.memo);

  if (formData.relatedFile && formData.relatedFile.startsWith("http")) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetId = sheet.getSheetId();
      const req = {
        "updateCells": {
          "range": { "sheetId": sheetId, "startRowIndex": row - 1, "endRowIndex": row, "startColumnIndex": 7, "endColumnIndex": 8 },
          "rows": [{ "values": [{ "userEnteredValue": { "stringValue": "@" }, "chipRuns": [{ "startIndex": 0, "chip": { "richLinkProperties": { "uri": formData.relatedFile } } }] }] }],
          "fields": "userEnteredValue,chipRuns"
        }
      };
      const res = UrlFetchApp.fetch(`https://sheets.googleapis.com/v4/spreadsheets/${ss.getId()}:batchUpdate`, {
        method: "post", contentType: "application/json",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ requests: [req] }), muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) throw new Error(res.getContentText());
    } catch(e) {
      const displayLabel = formData.relatedFileName ? formData.relatedFileName : "関連ファイル";
      const richText = SpreadsheetApp.newRichTextValue().setText(displayLabel).setLinkUrl(formData.relatedFile).build();
      sheet.getRange(row, 8).setRichTextValue(richText);
    }
  } else {
    sheet.getRange(row, 8).setValue(formData.relatedFile || "");
  }

  return `案件 ${formData.id} の情報を更新しました。`;
}

/**
 * ★追加：案件の削除
 */
function deleteJobRow(jobId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('案件管理');
  const data = sheet.getDataRange().getValues();
  const targetId = String(jobId).trim().toUpperCase();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toUpperCase() === targetId) {
      sheet.deleteRow(i + 1);
      return `案件 ${targetId} を削除しました。`;
    }
  }
  return "エラー: 指定された案件が見つかりません。";
}

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