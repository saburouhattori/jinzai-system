// =========================================
// 外部連携（支払い管理への同期）
// =========================================

// 外部連携スプレッドシートID（Funtoco支払い管理用）
const EXTERNAL_SS_ID_FUNTOCO = "1Yo6Oz3iK6OlWjzl7BVUWeElO4__mPjJST3Jaaiys9yw";

/**
 * 案件管理シートから外部の「支払い管理」シートへデータを同期する
 */
function syncToPaymentManagement() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('案件管理');
    if (!sourceSheet) throw new Error("「案件管理」シートが見つかりません。");

    // 外部スプレッドシートを開く
    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const targetSheet = targetSS.getSheetByName("支払い管理");
    if (!targetSheet) throw new Error("外部シートに「支払い管理」が見つかりません。");

    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceMap = getMasterColumnMap(sourceSheet);
    const targetData = targetSheet.getDataRange().getValues();
    const targetMap = getMasterColumnMap(targetSheet);

    if (sourceData.length < 2) return "同期対象の案件がありません。";

    // Funtoco側のキー列インデックス
    const tJobIdx = targetMap['案件ID'] - 1;
    const tIdIdx = targetMap['登録者ID'] - 1;

    let updateCount = 0;
    let appendCount = 0;

    // 支払い管理側の既存キーをマップ化（重複チェック用）
    const targetKeys = {};
    const existingCandidateMap = new Map(); // 登録者ID -> [案件IDの配列] (重複チェック用)

    if (targetData.length > 1) {
      for (let i = 1; i < targetData.length; i++) {
        const jId = String(targetData[i][tJobIdx] || "").trim();
        const cId = String(targetData[i][tIdIdx] || "").trim();
        
        if (jId && cId) {
          const key = jId + "_" + cId;
          targetKeys[key] = i; // 配列のインデックス（0始まり）として保持する
        }
        
        // 重複警告用に、登録者IDに紐づく案件IDを記録しておく
        if (cId) {
          if (!existingCandidateMap.has(cId)) {
            existingCandidateMap.set(cId, []);
          }
          if (jId && !existingCandidateMap.get(cId).includes(jId)) {
            existingCandidateMap.get(cId).push(jId);
          }
        }
      }
    }

    // 案件管理から転記用データを生成する
    const syncRecords = [];
    
    // 案件管理をループして採用者リストを展開
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      if (!jobID) continue;

      const hiredText = sourceMap['採用者名'] ? String(row[sourceMap['採用者名'] - 1] || "").trim() : "";
      if (!hiredText) continue;

      const companyName = sourceMap['事業者名'] ? row[sourceMap['事業者名'] - 1] : "";
      const fieldName = sourceMap['技能分野'] ? row[sourceMap['技能分野'] - 1] : "";
      
      // ★面接日の取得と日付フォーマットの安全な処理
      let interviewDate = sourceMap['面接日'] ? row[sourceMap['面接日'] - 1] : "";
      if (interviewDate instanceof Date) {
        interviewDate = Utilities.formatDate(interviewDate, "JST", "yyyy/MM/dd");
      } else if (interviewDate) {
        interviewDate = String(interviewDate).trim();
      }
      
      const hiredList = hiredText.split(/\r?\n/).filter(line => line.trim() !== "");
      for (const line of hiredList) {
        const match = line.match(/^(SD-\d+)-(.*)$/);
        let candidateID = "";
        let candidateName = "";
        
        if (match) {
          candidateID = match[1].trim();
          candidateName = match[2].trim();
        } else {
          candidateID = line.trim();
        }

        if(candidateID) {
           syncRecords.push({
             jobID: jobID,
             candidateID: candidateID,
             companyName: companyName,
             fieldName: fieldName,
             candidateName: candidateName,
             interviewDate: interviewDate // ★レコードに面接日を追加
           });
        }
      }
    }

    const numCols = targetSheet.getLastColumn() || Object.keys(targetMap).length;
    const warnings = new Set();
    
    let hasUpdates = false;
    const newRowsToAppend = [];

    // メモリ上で配列データを更新・追加
    for (const record of syncRecords) {
      const key = record.jobID + "_" + record.candidateID;
      
      const vals = {};
      vals['案件ID'] = record.jobID;
      vals['登録者ID'] = record.candidateID;
      vals['事業者名'] = record.companyName;
      vals['技能分野'] = record.fieldName;
      vals['名前'] = record.candidateName;
      vals['面接日'] = record.interviewDate; // ★Valsに面接日を設定

      if (targetKeys[key] !== undefined) {
        // メモリ上の targetData を更新
        const rowIdx = targetKeys[key];
        for (let headerName in vals) {
          if (targetMap[headerName] !== undefined && vals[headerName] !== undefined) {
            targetData[rowIdx][targetMap[headerName] - 1] = vals[headerName];
            hasUpdates = true;
          }
        }
        updateCount++;
      } else {
        // 新規追加用の配列にストック
        if (existingCandidateMap.has(record.candidateID)) {
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
    }

    // 最後に1回だけシートへ一括書き込み
    if (hasUpdates) {
      targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
    }
    if (newRowsToAppend.length > 0) {
      targetSheet.getRange(targetData.length + 1, 1, newRowsToAppend.length, numCols).setValues(newRowsToAppend);
    }

    // ----------------------------------------------------
    // 案件ID順に並べ替え (ソート)
    // ----------------------------------------------------
    const finalLastRow = targetData.length + newRowsToAppend.length;
    const finalLastCol = numCols;
    
    if (finalLastRow >= 2 && targetMap['案件ID']) {
      const dataRange = targetSheet.getRange(2, 1, finalLastRow - 1, finalLastCol);
      dataRange.sort({column: targetMap['案件ID'], ascending: true});
    }

    let resultMessage = `支払い管理への同期が完了しました。\n新規追加: ${appendCount}件\n情報更新: ${updateCount}件\n※「面接日」も含めて案件ID順に並べ替えました。`;
    if (warnings.size > 0) {
      resultMessage += `\n\n【⚠️警告】\n以下の登録者は別の案件IDで既に登録されていましたが、新たに重複して書き込まれました。\n不要なデータが残っていないか「支払い管理」シートを確認してください。\n`;
      resultMessage += Array.from(warnings).join("\n");
    }

    return resultMessage;

  } catch (e) {
    console.error("syncToPaymentManagement error: ", e);
    throw new Error("外部同期エラー: " + e.message);
  }
}