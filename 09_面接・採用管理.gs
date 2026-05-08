// =========================================
// 面接および採用に関する操作
// =========================================

/**
 * 案件に紐づく候補者情報のリスト取得
 */
function getJobCandidates(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");
    if (!details.interviewDate) throw new Error("面接日が設定されていません。\n先に「案件更新/削除」から面接日を登録してください。");
    if (!details.candidates) return [];

    const candDict = getCandidateDict(); 
    const ids = details.candidates.split(/\r?\n/).filter(id => id.trim());
    return ids.map(id => {
      const cleanId = id.split('-').slice(0, 2).join('-').trim();
      return { id: cleanId, display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId]})` : id };
    }).filter(c => c.id);
  } catch(e) { throw new Error(e.message); }
}

/**
 * 採用登録（＋面接履歴の自動追記機能）
 */
function registerHire(jobId, hiredIds) {
  try {
    const sheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!sheet || !mSheet) throw new Error("シートへのアクセスに失敗しました。");

    const mCol = getMasterColumnMap(mSheet);
    const data = sheet.getDataRange().getValues();
    const mData = mSheet.getDataRange().getValues();
    
    let companyName = "", rawInterviewDate = "", allCandidatesRaw = "", targetJobRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(jobId).trim()) {
        companyName = String(data[i][3]).trim();
        allCandidatesRaw = String(data[i][5]);
        rawInterviewDate = data[i][6];
        targetJobRow = i + 1;
        break;
      }
    }
    if (!companyName) throw new Error("案件が見つかりません。");
    if (!rawInterviewDate) throw new Error("面接日が設定されていません。\n先に「案件更新/削除」から面接日を登録してください。");

    let formattedDate = "日付不明";
    if (rawInterviewDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawInterviewDate, "JST", "yyyy/MM/dd");
    } else if (rawInterviewDate) {
      formattedDate = String(rawInterviewDate).replace(/[年月]/g, '/').replace(/日/g, '');
    }

    const candDict = getCandidateDict();
    const allCandidateIds = allCandidatesRaw.split(/\r?\n/).map(line => line.split('-').slice(0, 2).join('-').trim()).filter(id => id !== "");
    const hiredIdSet = new Set(hiredIds.map(id => String(id).trim()));
    
    allCandidateIds.forEach(candId => {
      const isHired = hiredIdSet.has(candId);
      const resultText = isHired ? "（採用）" : "（不採用）";
      const newHistoryLine = `${formattedDate}：${companyName}${resultText}`;

      for (let j = 1; j < mData.length; j++) {
        if (String(mData[j][0]).trim() === candId) {
          const rowIdx = j + 1;
          if (isHired) {
            if (mCol['ステータス']) mSheet.getRange(rowIdx, mCol['ステータス']).setValue('採用');
            if (mCol['採用事業者']) mSheet.getRange(rowIdx, mCol['採用事業者']).setValue(companyName);
          }
          if (mCol['面接履歴']) {
            const historyCell = mSheet.getRange(rowIdx, mCol['面接履歴']);
            const currentHistory = String(historyCell.getValue() || "").trim();
            historyCell.setValue(currentHistory ? currentHistory + "\n" + newHistoryLine : newHistoryLine);
          }
          break;
        }
      }
    });

    let hiredNamesText = "採用者なし";
    if (hiredIds.length > 0) {
      hiredNamesText = hiredIds.map(id => candDict[id] ? `${id}-${candDict[id]}` : id).join('\n');
    }

    sheet.getRange(targetJobRow, 8).setValue(hiredNamesText);
    sheet.getRange(targetJobRow, 2).setValue(hiredIds.length > 0 ? '入国準備' : '終了');

    if (hiredIds.length > 0) return `${hiredIds.length} 名の面接結果（ステータス：入国準備）、および対象候補者全員の「面接履歴」への追記が完了しました。`;
    return `「採用者なし」として案件を終了し、対象候補者全員の「面接履歴」への追記が完了しました。`;
  } catch(e) { throw new Error(e.message); }
}