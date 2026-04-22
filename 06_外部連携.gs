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
    if (targetData.length > 1) {
      for (let i = 1; i < targetData.length; i++) {
        const key = String(targetData[i][tJobIdx] || "").trim() + "_" + String(targetData[i][tIdIdx] || "").trim();
        targetKeys[key] = i + 1; // 行番号を保持
      }
    }

    // 案件管理から転記用データを生成する
    const syncRecords = [];
    
    // 案件管理をループして採用者リストを展開
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      if (!jobID) continue; // 案件IDがない行はスキップ

      const hiredText = sourceMap['採用者名'] ? String(row[sourceMap['採用者名'] - 1] || "").trim() : "";
      if (!hiredText) continue; // 採用者がいない場合はスキップ

      const companyName = sourceMap['事業者名'] ? row[sourceMap['事業者名'] - 1] : "";
      const fieldName = sourceMap['技能分野'] ? row[sourceMap['技能分野'] - 1] : "";
      const interviewDate = sourceMap['面接日'] ? row[sourceMap['面接日'] - 1] : ""; // 内定日として利用

      // "SD-0064-HNIN EI HLAING" などをパースして複数人の配列にする
      const hiredList = hiredText.split(/\r?\n/).filter(line => line.trim() !== "");
      for (const line of hiredList) {
        const match = line.match(/^(SD-\d+)-(.*)$/);
        let candidateID = "";
        let candidateName = "";
        
        if (match) {
          candidateID = match[1].trim();
          candidateName = match[2].trim();
        } else {
          // フォーマット外の場合、そのままIDとして扱うなどのフェールセーフ
          candidateID = line.trim();
        }

        if(candidateID) {
           syncRecords.push({
             jobID: jobID,
             candidateID: candidateID,
             companyName: companyName,
             fieldName: fieldName,
             candidateName: candidateName,
             interviewDate: interviewDate
           });
        }
      }
    }

    // 展開したリストを支払い管理へ同期
    const numCols = targetSheet.getLastColumn() || Object.keys(targetMap).length;

    for (const record of syncRecords) {
      const key = record.jobID + "_" + record.candidateID;

      // 転記するデータのマッピング作成
      const vals = {};
      vals['案件ID'] = record.jobID;
      vals['登録者ID'] = record.candidateID;
      vals['事業者名'] = record.companyName;
      vals['技能分野'] = record.fieldName;
      vals['内定日'] = record.interviewDate;
      vals['名前'] = record.candidateName;

      if (targetKeys[key]) {
        // 更新処理（Funtoco側で入力する「金額」等は上書きしない）
        const rowNum = targetKeys[key];
        for (let headerName in vals) {
          if (targetMap[headerName] !== undefined && vals[headerName] !== undefined) {
            targetSheet.getRange(rowNum, targetMap[headerName]).setValue(vals[headerName]);
          }
        }
        updateCount++;
      } else {
        // 新規追記処理
        const newRowValues = new Array(numCols).fill("");
        for (let headerName in vals) {
          if (targetMap[headerName] !== undefined && vals[headerName] !== undefined) {
            newRowValues[targetMap[headerName] - 1] = vals[headerName];
          }
        }
        targetSheet.appendRow(newRowValues);
        appendCount++;
      }
    }

    // ----------------------------------------------------
    // 案件ID順に並べ替え (ソート)
    // ----------------------------------------------------
    const finalLastRow = targetSheet.getLastRow();
    const finalLastCol = targetSheet.getLastColumn();
    // 2行以上データがあり、案件IDの列が存在する場合のみソートを実行
    if (finalLastRow >= 2 && targetMap['案件ID']) {
      const dataRange = targetSheet.getRange(2, 1, finalLastRow - 1, finalLastCol);
      dataRange.sort({column: targetMap['案件ID'], ascending: true});
    }

    return `支払い管理への同期が完了しました。\n新規追加: ${appendCount}件\n情報更新: ${updateCount}件\n※案件ID順に並べ替えました。`;

  } catch (e) {
    console.error("syncToPaymentManagement error: ", e);
    throw new Error("外部同期エラー: " + e.message);
  }
}