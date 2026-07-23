// =========================================
// 外部連携（支払い管理への同期）
// =========================================

const EXTERNAL_SS_ID_FUNTOCO = "1Yo6Oz3iK6OlWjzl7BVUWeElO4__mPjJST3Jaaiys9yw";

function syncToPaymentManagement() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('案件管理');
    if (!sourceSheet) throw new Error("「案件管理」シートが見つかりません。");

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const targetSheet = targetSS.getSheetByName("支払い管理");
    if (!targetSheet) throw new Error("外部シートに「支払い管理」が見つかりません。");

    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceMap = getMasterColumnMap(sourceSheet);
    const targetData = targetSheet.getDataRange().getValues();
    const targetMap = getMasterColumnMap(targetSheet);

    if (sourceData.length < 2) return "同期対象の案件がありません。";

    const tJobIdx = targetMap['案件ID'] - 1;
    const tIdIdx = targetMap['登録者ID'] - 1;

    let appendCount = 0;
    let skipCount = 0;
    const targetKeys = {};
    const existingCandidateMap = new Map();
    let footerRowIndex = targetData.length + 1;

    if (targetData.length > 1) {
      for (let i = 1; i < targetData.length; i++) {
        const jId = String(targetData[i][tJobIdx] || "").trim();
        const cId = String(targetData[i][tIdIdx] || "").trim();
        
        if (!jId && !cId) {
          footerRowIndex = i + 1;
          break;
        }
        if (jId && cId) {
          const key = jId + "_" + cId;
          targetKeys[key] = i; 
        }
        if (cId && cId !== "採用者なし") {
          if (!existingCandidateMap.has(cId)) {
            existingCandidateMap.set(cId, []);
          }
          if (jId && !existingCandidateMap.get(cId).includes(jId)) {
            existingCandidateMap.get(cId).push(jId);
          }
        }
      }
    }

    const syncRecords = [];
    
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      if (!jobID) continue;

      const hiredText = sourceMap['採用者名'] ? String(row[sourceMap['採用者名'] - 1] || "").trim() : "";
      if (!hiredText) continue;

      const companyNameCell = sourceMap['事業者名'] ? String(row[sourceMap['事業者名'] - 1] || "") : "";
      const defaultCompany = companyNameCell.split(/\r?\n/)[0];
      const fieldName = sourceMap['技能分野'] ? row[sourceMap['技能分野'] - 1] : "";
      
      let interviewDate = sourceMap['面接日'] ? row[sourceMap['面接日'] - 1] : "";
      if (typeof interviewDate === "string") {
        interviewDate = interviewDate.trim();
      }
      
      const hiredList = hiredText.split(/\r?\n/).filter(line => line.trim() !== "");
      let currentCompany = defaultCompany;

      for (const line of hiredList) {
        if (line.startsWith('【') && line.endsWith('】')) {
          currentCompany = line.slice(1, -1).trim();
          continue;
        }

        let candidateID = "";
        let candidateName = "";
        
        if (line.trim() === "採用者なし") {
          candidateID = "採用者なし";
          candidateName = "採用者なし";
        } else {
          const match = line.match(/^(SD-\d+)(?:-(.*))?$/);
          if (match) {
            candidateID = match[1].trim();
            candidateName = match[2] ? match[2].trim() : "";
          } else {
            continue; 
          }
        }

        if(candidateID) {
           syncRecords.push({
             jobID: jobID,
             candidateID: candidateID,
             companyName: currentCompany,
             fieldName: fieldName,
             candidateName: candidateName,
             interviewDate: interviewDate
           });
        }
      }
    }

    const numCols = targetMap['備考'] ? targetMap['備考'] : (targetSheet.getLastColumn() || Object.keys(targetMap).length);
    const warnings = new Set();
    const newRowsToAppend = [];

    for (const record of syncRecords) {
      const key = record.jobID + "_" + record.candidateID;
      
      if (targetKeys[key] !== undefined) {
        skipCount++;
        continue; 
      }

      const vals = {};
      vals['案件ID'] = record.jobID;
      vals['登録者ID'] = record.candidateID;
      vals['事業者名'] = record.companyName;
      vals['技能分野'] = record.fieldName;
      vals['名前'] = record.candidateName;
      vals['面接日'] = record.interviewDate;

      if (record.candidateID !== "採用者なし" && existingCandidateMap.has(record.candidateID)) {
         const oldJobs = existingCandidateMap.get(record.candidateID).join(", ");
         warnings.add(`・${record.candidateID} ${record.candidateName} (既存案件ID: ${oldJobs})`);
      }

      const newRowValues = new Array(numCols).fill("");
      for (let headerName in vals) {
        if (targetMap[headerName] !== undefined && vals[headerName] !== undefined) {
          newRowValues[targetMap[headerName] - 1] = vals[headerName];
        }
      }
      newRowsToAppend.push(newRowValues);
      appendCount++;
    }

    if (newRowsToAppend.length > 0) {
      targetSheet.insertRowsBefore(footerRowIndex, newRowsToAppend.length);
      targetSheet.getRange(footerRowIndex, 1, newRowsToAppend.length, numCols).setValues(newRowsToAppend);
    }

    let resultMessage = `支払い管理への同期が完了しました。\n新規追加: ${appendCount}件\nスキップ（既存）: ${skipCount}件`;
    if (warnings.size > 0) {
      resultMessage += `\n\n【重複警告】\n以下の登録者は、別案件IDで既に登録されています。\n`;
      resultMessage += Array.from(warnings).join("\n");
    }

    return resultMessage;

  } catch (e) {
    throw new Error("外部同期エラー: " + e.message);
  }
}