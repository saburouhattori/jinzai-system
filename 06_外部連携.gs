// =========================================
// 外部連携（支払い管理への同期）
// =========================================

// 外部連携スプレッドシートID（Funtoco支払い管理用）
const EXTERNAL_SS_ID_FUNTOCO = "1Yo6Oz3iK6OlWjzl7BVUWeElO4__mPjJST3Jaaiys9yw";

/**
 * 案件管理シートから外部の「支払い管理」シートへ未登録の新規データのみを安全に同期する
 * スマートテーブルのフッター（集計行）を自動検知し、その直前にテーブル行を挿入して美しく格納する
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

    let appendCount = 0;
    let skipCount = 0;

    // 支払い管理側の既存キーをマップ化（重複・存在チェック用）
    const targetKeys = {};
    const existingCandidateMap = new Map(); // 登録者ID -> [案件IDの配列] (重複チェック用)

    // ★修正1: 罫線だけの空行やフッターに騙されないよう、「真のテーブルフッター行（空行）」を自動判定する
    let footerRowIndex = targetData.length + 1; // デフォルトは最終行の次

    if (targetData.length > 1) {
      for (let i = 1; i < targetData.length; i++) {
        const jId = String(targetData[i][tJobIdx] || "").trim();
        const cId = String(targetData[i][tIdIdx] || "").trim();
        
        // 案件IDも登録者IDも空の行を見つけたら、そこを「フッター行」として記録してブレイク
        if (!jId && !cId) {
          footerRowIndex = i + 1; // 1-basedの行番号
          break;
        }

        if (jId && cId) {
          const key = jId + "_" + cId;
          targetKeys[key] = i; // 既存の存在を示すインデックスを保持
        }
        
        // 重複警告用に、登録者IDに紐づく案件IDを記録（「採用者なし」は除外）
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

    // 案件管理から転記用データを抽出
    const syncRecords = [];
    
    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      if (!jobID) continue;

      const hiredText = sourceMap['採用者名'] ? String(row[sourceMap['採用者名'] - 1] || "").trim() : "";
      // 空欄のみスキップし、「採用者なし」は通す
      if (!hiredText) continue;

      const companyName = sourceMap['事業者名'] ? row[sourceMap['事業者名'] - 1] : "";
      const fieldName = sourceMap['技能分野'] ? row[sourceMap['技能分野'] - 1] : "";
      
      // 面接日はスマートテーブルの日付書式と連動させるため、Dateオブジェクトのまま（または元の値のまま）扱う
      let interviewDate = sourceMap['面接日'] ? row[sourceMap['面接日'] - 1] : "";
      if (typeof interviewDate === "string") {
        interviewDate = interviewDate.trim();
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
          const rawId = line.trim();
          if (rawId.startsWith("SD-")) {
            candidateID = rawId;
          } else if (rawId === "採用者なし") {
            candidateID = "採用者なし";
            candidateName = "採用者なし";
          } else {
            continue; 
          }
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

    const numCols = targetSheet.getLastColumn() || Object.keys(targetMap).length;
    const warnings = new Set();
    const newRowsToAppend = [];

    // メモリ上で判定を行い、未登録データのみを抽出
    for (const record of syncRecords) {
      const key = record.jobID + "_" + record.candidateID;
      
      // すでに Funtoco 側に「案件ID＋登録者ID」のペアが存在する場合は完全にスキップ
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

      // 過去に別案件で登録がある場合は警告をストック（「採用者なし」は警告対象外）
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

    // 完全に「新しく追加された採用者データ」のみをシートへ一括書き込み
    if (newRowsToAppend.length > 0) {
      // ★修正2: フッター行の「真上」に、足りない件数分の行を一括挿入する
      // これによりスマートテーブルが壊れず内部に自動拡張され、既存のデザインや交互背景色が完璧に連動します。
      targetSheet.insertRowsBefore(footerRowIndex, newRowsToAppend.length);

      // 確保されたテーブル内の新しい空行（開始位置は元のfooterRowIndex）にデータを流し込む
      targetSheet.getRange(footerRowIndex, 1, newRowsToAppend.length, numCols).setValues(newRowsToAppend);
      
      // ※スマートテーブル自体の持つソート・集計の崩壊を防ぐため、GASによる強制並び替え（.sort()）は安全のために撤廃しています。
    }

    let resultMessage = `支払い管理への同期が完了しました。\n新規追加: ${appendCount}件\nスキップ（既存）: ${skipCount}件\n\n※スマートテーブルのフッターを自動で押し下げ、枠内へ美しくデータが格納されました。`;
    if (warnings.size > 0) {
      resultMessage += `\n\n【⚠️重複警告】\n以下の登録者は、今回の案件とは別の案件IDで既に過去に登録されています。今回の新しい案件情報も重複して追記されましたので、問題がないか「支払い管理」シートをご確認ください。\n`;
      resultMessage += Array.from(warnings).join("\n");
    }

    return resultMessage;

  } catch (e) {
    console.error("syncToPaymentManagement error: ", e);
    throw new Error("外部同期エラー: " + e.message);
  }
}