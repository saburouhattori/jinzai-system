// =========================================
// 外部連携（支払い管理への同期）
// =========================================

// 外部連携スプレッドシートID（Funtoco支払い管理用）
const EXTERNAL_SS_ID_FUNTOCO = "1Yo6Oz3iK6OlWjzl7BVUWeElO4__mPjJST3Jaaiys9yw";


/**
 * 採用者一覧から外部の「支払い管理」シートへデータを同期する
 */
function syncToPaymentManagement() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('採用者一覧');
    if (!sourceSheet) throw new Error("「採用者一覧」シートが見つかりません。");

    // 外部スプレッドシートを開く
    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const targetSheet = targetSS.getSheetByName("支払い管理");
    if (!targetSheet) throw new Error("外部シートに「支払い管理」が見つかりません。");

    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceMap = getMasterColumnMap(sourceSheet);
    
    const targetData = targetSheet.getDataRange().getValues();
    const targetMap = getMasterColumnMap(targetSheet);

    if (sourceData.length < 2) return "同期対象の採用者がいません。";

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

    // 採用者一覧をループして同期
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      const candidateID = sourceMap['登録者ID'] ? String(row[sourceMap['登録者ID'] - 1] || "").trim() : "";
      
      if (!jobID || !candidateID) continue; // IDが揃っていない行はスキップ

      const key = jobID + "_" + candidateID;

      // 転記するデータのマッピング作成
      const vals = {};
      vals['案件ID'] = jobID;
      vals['登録者ID'] = candidateID;
      if (sourceMap['採用事業者名']) vals['採用事業者名'] = row[sourceMap['採用事業者名'] - 1];
      if (sourceMap['技能分野']) vals['技能分野'] = row[sourceMap['技能分野'] - 1];
      if (sourceMap['内定日']) vals['内定日'] = row[sourceMap['内定日'] - 1];
      if (sourceMap['名前']) vals['名前'] = row[sourceMap['名前'] - 1];
      if (sourceMap['直近の入国日']) vals['入国日'] = row[sourceMap['直近の入国日'] - 1];
      if (sourceMap['備考・メモ']) vals['備考'] = row[sourceMap['備考・メモ'] - 1];

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
        const numCols = targetSheet.getLastColumn() || Object.keys(targetMap).length;
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
      // 2行目以降のデータ領域全体を取得
      const dataRange = targetSheet.getRange(2, 1, finalLastRow - 1, finalLastCol);
      // 案件IDの列番号（1始まり）を基準に昇順(ascending: true)でソート
      dataRange.sort({column: targetMap['案件ID'], ascending: true});
    }

    return `支払い管理への同期が完了しました。\n新規追加: ${appendCount}件\n情報更新: ${updateCount}件\n※案件ID順に並べ替えました。`;

  } catch (e) {
    console.error("syncToPaymentManagement error: ", e);
    throw new Error("外部同期エラー: " + e.message);
  }
}