// =========================================
// 外部連携（支払い管理への同期）
// =========================================

const EXTERNAL_SS_ID_FUNTOCO = "1Yo6Oz3iK6OlWjzl7BVUWeElO4__mPjJST3Jaaiys9yw";

function syncToPaymentManagement() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('案件管理');
    if (!sourceSheet) throw new Error("「案件管理」シートが見つかりません。");

    const masterSheet = ss.getSheetByName('登録者マスタ');
    if (!masterSheet) throw new Error("「登録者マスタ」シートが見つかりません。");

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const targetSheet = targetSS.getSheetByName("支払い管理");
    if (!targetSheet) throw new Error("外部シートに「支払い管理」が見つかりません。");

    // ▼ 両方のスプレッドシートのタイムゾーンを取得（時差ズレ防止の要）
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const targetTZ = targetSS.getSpreadsheetTimeZone();

    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceMap = getMasterColumnMap(sourceSheet);

    const masterData = masterSheet.getDataRange().getValues();
    const masterMap = getMasterColumnMap(masterSheet);

    const targetData = targetSheet.getDataRange().getValues();
    const targetMap = getMasterColumnMap(targetSheet);

    if (sourceData.length < 2) return "同期対象の案件がありません。";

    // ▼ 日付フォーマット用ヘルパー関数（NaNやundefined等のゴミデータを排除）
    const formatDateStr = (val, tz) => {
      if (!val) return "";
      if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy/MM/dd");
      const s = String(val).trim().replace(/-/g, '/');
      if (s === "NaN" || s === "undefined" || s === "Invalid Date") return "";
      return s;
    };

    // ▼▼ マスタから 内定日・入国日・入職日の抽出（登録者マスタがSSoT） ▼▼
    const candidateInfo = {};
    if (masterMap['登録者ID']) {
      const idIdx = masterMap['登録者ID'] - 1;
      const offerIdx = masterMap['内定日'] ? masterMap['内定日'] - 1 : -1;
      const entryIdx = masterMap['入国日'] ? masterMap['入国日'] - 1 : -1;
      const workIdx = masterMap['入職日'] ? masterMap['入職日'] - 1 : -1;

      for (let i = 1; i < masterData.length; i++) {
        const cid = String(masterData[i][idIdx]).trim();
        if (!cid) continue;

        let oDate = offerIdx !== -1 ? formatDateStr(masterData[i][offerIdx], sourceTZ) : "";
        let eDate = entryIdx !== -1 ? formatDateStr(masterData[i][entryIdx], sourceTZ) : "";
        let wDate = workIdx !== -1 ? formatDateStr(masterData[i][workIdx], sourceTZ) : "";

        candidateInfo[cid] = { offerDate: oDate, entryDate: eDate, workDate: wDate };
      }
    }

    const tJobIdx = targetMap['案件ID'] - 1;
    const tIdIdx = targetMap['登録者ID'] - 1;

    let appendCount = 0;
    let updateCount = 0;
    const targetKeys = {};
    const existingCandidateMap = new Map();
    let footerRowIndex = targetData.length + 1;

    // 既存データのマッピング
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
    
    // 案件管理からのデータ抽出
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      if (!jobID) continue;

      const detailsText = sourceMap['面接結果詳細'] ? String(row[sourceMap['面接結果詳細'] - 1] || "").trim() : "";
      if (!detailsText) continue;

      const companyNameCell = sourceMap['事業者名'] ? String(row[sourceMap['事業者名'] - 1] || "") : "";
      const defaultCompany = companyNameCell.split(/\r?\n/)[0];
      const fieldName = sourceMap['技能分野'] ? row[sourceMap['技能分野'] - 1] : "";
      
      const interviewDate = sourceMap['面接日'] ? formatDateStr(row[sourceMap['面接日'] - 1], sourceTZ) : "";
      
      const dLines = detailsText.split(/\r?\n/).filter(line => line.trim() !== "");
      let currentCompany = defaultCompany;
      let hasHired = false;

      for (const line of dLines) {
        if (line.startsWith('【') && line.endsWith('】')) {
          currentCompany = line.slice(1, -1).trim();
          continue;
        }

        if (line.includes('（採用）')) {
          let candidateID = "";
          let candidateName = "";
          const idMatch = line.match(/^(SD-\d+)/);
          
          if (idMatch) {
             candidateID = idMatch[1].trim();
             const nMatch = line.match(/^SD-\d+-(.*?)(?:（.*?）)$/);
             candidateName = nMatch ? nMatch[1].trim() : "";
             hasHired = true;

             let oDate = "";
             let eDate = "";
             let wDate = "";
             if (candidateID !== "採用者なし" && candidateInfo[candidateID]) {
               oDate = candidateInfo[candidateID].offerDate;
               eDate = candidateInfo[candidateID].entryDate;
               wDate = candidateInfo[candidateID].workDate;
             }

             syncRecords.push({
               jobID: jobID,
               candidateID: candidateID,
               companyName: currentCompany,
               fieldName: fieldName,
               candidateName: candidateName,
               interviewDate: interviewDate,
               offerDate: oDate,
               entryDate: eDate,
               workDate: wDate
             });
          }
        }
      }
      
      if (!hasHired) {
         syncRecords.push({
           jobID: jobID,
           candidateID: "採用者なし",
           companyName: defaultCompany,
           fieldName: fieldName,
           candidateName: "採用者なし",
           interviewDate: interviewDate,
           offerDate: "",
           entryDate: "",
           workDate: ""
         });
      }
    }

    // 列数の取得を柔軟に変更
    const numCols = targetSheet.getLastColumn() || Object.keys(targetMap).length;
    const warnings = new Set();
    const newRowsToAppend = [];
    
    // 削除対象の特定用セット
    const syncKeys = new Set(syncRecords.map(r => r.jobID + "_" + r.candidateID));

    // 書き込み・更新処理
    for (const record of syncRecords) {
      const key = record.jobID + "_" + record.candidateID;
      
      // ▼▼ 既存データの場合は、差分がある基本情報のみを更新 ▼▼
      if (targetKeys[key] !== undefined) {
        const tIdx = targetKeys[key];
        let hasChanges = false;
        
        const checkAndUpdate = (colName, newVal) => {
          if (targetMap[colName]) {
            const cIdx = targetMap[colName] - 1;
            const oldVal = targetData[tIdx][cIdx];
            
            const oldStr = formatDateStr(oldVal, targetTZ);
            const newStr = formatDateStr(newVal, targetTZ);
            
            if (oldStr !== newStr) {
              targetSheet.getRange(tIdx + 1, cIdx + 1).setValue(newStr);
              if (['面接日', '内定日', '入国日', '入職日'].includes(colName)) {
                targetSheet.getRange(tIdx + 1, cIdx + 1).setNumberFormat('yyyy/MM/dd');
              }
              hasChanges = true;
            }
          }
        };

        checkAndUpdate('事業者名', record.companyName);
        checkAndUpdate('技能分野', record.fieldName);
        checkAndUpdate('名前', record.candidateName);
        checkAndUpdate('面接日', record.interviewDate);
        checkAndUpdate('内定日', record.offerDate);
        checkAndUpdate('入国日', record.entryDate);
        checkAndUpdate('入職日', record.workDate);
        
        if (hasChanges) updateCount++;
        continue; 
      }

      // ▼▼ 新規データの場合は配列にストックして後で一括追加 ▼▼
      const vals = {};
      vals['チェッカー'] = false; 
      vals['案件ID'] = record.jobID;
      vals['登録者ID'] = record.candidateID;
      vals['事業者名'] = record.companyName;
      vals['技能分野'] = record.fieldName;
      vals['名前'] = record.candidateName;
      vals['面接日'] = record.interviewDate;
      vals['内定日'] = record.offerDate; 
      vals['入国日'] = record.entryDate;
      vals['入職日'] = record.workDate;

      if (record.candidateID !== "採用者なし" && existingCandidateMap.has(record.candidateID)) {
         const oldJobs = existingCandidateMap.get(record.candidateID).join(", ");
         warnings.add(`・${record.candidateID} ${record.candidateName} (既存案件ID: ${oldJobs})`);
      }

      const newRowValues = new Array(numCols).fill("");
      newRowValues[0] = false; 

      for (let headerName in vals) {
        if (targetMap[headerName] !== undefined && vals[headerName] !== undefined) {
          newRowValues[targetMap[headerName] - 1] = vals[headerName];
        }
      }
      newRowsToAppend.push(newRowValues);
      appendCount++;
    }

    // ▼▼ 削除対象の特定 ▼▼
    const rowsToDelete = [];
    for (const key in targetKeys) {
      if (!syncKeys.has(key)) {
        rowsToDelete.push(targetKeys[key] + 1); 
      }
    }

    // 新規行の一括追加
    if (newRowsToAppend.length > 0) {
      targetSheet.insertRowsBefore(footerRowIndex, newRowsToAppend.length);
      targetSheet.getRange(footerRowIndex, 1, newRowsToAppend.length, numCols).setValues(newRowsToAppend);
      
      const setFormatIfExist = (colName) => {
        if (targetMap[colName]) {
          targetSheet.getRange(footerRowIndex, targetMap[colName], newRowsToAppend.length, 1).setNumberFormat('yyyy/MM/dd');
        }
      };
      setFormatIfExist('面接日');
      setFormatIfExist('内定日');
      setFormatIfExist('入国日');
      setFormatIfExist('入職日');
    }

    // 行の削除
    rowsToDelete.sort((a, b) => b - a);
    for (const rowNum of rowsToDelete) {
      targetSheet.deleteRow(rowNum);
    }

    let resultMessage = `支払い管理への同期が完了しました。\n新規追加: ${appendCount}件\n情報更新: ${updateCount}件\n削除: ${rowsToDelete.length}件\n（変更なしスキップ: ${syncRecords.length - appendCount - updateCount}件）`;
    if (warnings.size > 0) {
      resultMessage += `\n\n【重複警告】\n以下の登録者は、別案件IDで既に登録されています。\n`;
      resultMessage += Array.from(warnings).join("\n");
    }

    return resultMessage;

  } catch (e) {
    throw new Error("外部同期エラー: " + e.message);
  }
}