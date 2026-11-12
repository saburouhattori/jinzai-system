// =========================================
// データ保守・一括再構築用スクリプト
// 「案件管理」を正本とした登録者マスタの一括再構築（面接履歴・ステータス・名前表記統一・メモ分離）
// =========================================

function rebuildMasterFromJobSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("登録者マスタ");
  const jobSheet = ss.getSheetByName("案件管理");
  if (!masterSheet || !jobSheet)
    throw new Error(
      "「登録者マスタ」または「案件管理」シートが見つかりません。"
    );

  // 1. 候補者名の表記統一・(2)除去、および名前欄の改行メモの「備考・メモ」列への移動
  const nameFixMap = {
    "SD-0142": "THANDAR LIN",
    "SD-0966": "HNAUNG THARAPHI OO",
    "SD-1212": "EI SHWE SIN",
    "SD-1248": "SAUNG HNIN PHYU",
    "SD-1251": "AYE SU MOE",
    "SD-1269": "NYEIN NYEIN AYE",
    "SD-1270": "HNIN EI PHYU",
    "SD-1281": "AYE AYE NYEIN",
    "SD-1285": "MAY ZIN PHYO",
  };

  const mCol = getMasterColumnMap(masterSheet);
  const mData = masterSheet.getDataRange().getValues();
  const idIdx = mCol["登録者ID"] - 1;
  const nameIdx = mCol["名前"] - 1;
  const memoIdx = mCol["備考・メモ"] ? mCol["備考・メモ"] - 1 : -1;

  let memoMovedCount = 0;

  for (let i = 1; i < mData.length; i++) {
    const cid = String(mData[i][idIdx] || "")
      .trim()
      .toUpperCase();
    if (!cid) continue;

    let rawName = String(mData[i][nameIdx] || "");

    // ① 名前セルに改行が含まれている場合：1行目を名前、2行目以降を「備考・メモ」へ移動
    if (rawName.includes("\n") || rawName.includes("\r")) {
      const parts = rawName
        .split(/\r?\n/)
        .map((s) => s.trim())
        .filter((s) => s !== "");
      if (parts.length >= 2) {
        const cleanName = parts[0];
        const extractedNote = parts.slice(1).join("\n");

        masterSheet.getRange(i + 1, nameIdx + 1).setValue(cleanName);
        mData[i][nameIdx] = cleanName;
        rawName = cleanName;

        if (memoIdx !== -1) {
          const currentMemo = String(mData[i][memoIdx] || "").trim();
          let newMemo = currentMemo;
          if (!currentMemo) {
            newMemo = extractedNote;
          } else if (!currentMemo.includes(extractedNote)) {
            newMemo = currentMemo + "\n" + extractedNote;
          }
          if (newMemo !== currentMemo) {
            masterSheet.getRange(i + 1, memoIdx + 1).setValue(newMemo);
            mData[i][memoIdx] = newMemo;
          }
        }
        updateJobCandidateName(cid, cleanName);
        memoMovedCount++;
      }
    }

    // ② 個別の名前表記統一・(2)除去
    if (nameFixMap[cid]) {
      const targetName = nameFixMap[cid];
      if (String(mData[i][nameIdx] || "").trim() !== targetName) {
        masterSheet.getRange(i + 1, nameIdx + 1).setValue(targetName);
        mData[i][nameIdx] = targetName;
      }
      updateJobCandidateName(cid, targetName);
    }
  }

  // 2. 案件管理データの読み込みと各候補者の面接履歴・採用結果の集計
  const jobData = jobSheet.getDataRange().getValues();
  const jCol = getMasterColumnMap(jobSheet);
  const jIdIdx = jCol["案件ID"] - 1;
  const jCompIdx = jCol["事業者名"] - 1;
  const jDateIdx = jCol["面接日"] - 1;
  const jDetailsIdx = jCol["面接結果詳細"] - 1;

  const toSlashDate = (val) => {
    if (!val) return "日付不明";
    if (val instanceof Date)
      return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
    const s = String(val)
      .trim()
      .replace(/[年月\-]/g, "/")
      .replace(/日/g, "");
    const parts = s.split("/");
    if (parts.length === 3) {
      return `${parts[0]}/${parts[1].padStart(2, "0")}/${parts[2].padStart(
        2,
        "0"
      )}`;
    }
    return s || "日付不明";
  };

  const knownStatuses = [
    "採用",
    "不採用",
    "内定辞退（候補者都合）",
    "内定取消（事業者都合）",
    "面接結果の取消",
  ];
  const historyMap = new Map(); // cid -> Array of { sortDate, jobId, line }

  for (let i = 1; i < jobData.length; i++) {
    const row = jobData[i];
    const jobId = String(row[jIdIdx] || "").trim();
    const detailsText = String(row[jDetailsIdx] || "").trim();
    if (!jobId || !detailsText) continue;

    const compText = String(row[jCompIdx] || "")
      .normalize("NFC")
      .trim();
    const companies = compText
      .split(/\r?\n/)
      .map((c) => c.trim())
      .filter((c) => c);
    const defaultJoinedCompany = companies.join("・");
    let currentCompany = companies.length > 0 ? companies[0] : "";

    const formattedDate = toSlashDate(row[jDateIdx]);
    const sortDateKey =
      formattedDate === "日付不明" ? "0000/00/00" : formattedDate;

    const lines = detailsText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l);
    for (const line of lines) {
      if (line.startsWith("【") && line.endsWith("】")) {
        currentCompany = line.slice(1, -1).normalize("NFC").trim();
        continue;
      }

      const idMatch = line.match(/^(SD-\d+)/i);
      if (!idMatch) continue;
      const cid = idMatch[1].toUpperCase();

      let parsedResult = null;
      for (const st of knownStatuses) {
        if (line.endsWith(`（${st}）`)) {
          parsedResult = st;
          break;
        }
      }
      if (!parsedResult || parsedResult === "面接結果の取消") continue;

      let histComp = currentCompany;
      let suffix = "";
      if (parsedResult === "採用") {
        histComp = currentCompany;
        suffix = "（採用）";
      } else if (parsedResult === "不採用") {
        histComp = defaultJoinedCompany;
        suffix = "（不採用）";
      } else if (parsedResult === "内定辞退（候補者都合）") {
        histComp = currentCompany;
        suffix = "（内定辞退）";
      } else if (parsedResult === "内定取消（事業者都合）") {
        histComp = currentCompany;
        suffix = "（事業者都合取消）";
      }

      const histLine = `${formattedDate}：${histComp}${suffix}`;
      if (!historyMap.has(cid)) historyMap.set(cid, []);
      historyMap.get(cid).push({
        sortDate: sortDateKey,
        jobId: jobId,
        line: histLine,
      });
    }
  }

  // 3. 登録者マスタの「面接履歴」列を一括更新
  const historyColIdx = mCol["面接履歴"] - 1;
  const historyUpdates = [];
  let historyUpdatedCount = 0;

  for (let i = 1; i < mData.length; i++) {
    const cid = String(mData[i][idIdx] || "")
      .trim()
      .toUpperCase();
    const oldHist = String(mData[i][historyColIdx] || "").trim();

    let newHist = "";
    if (cid && historyMap.has(cid)) {
      const items = historyMap.get(cid);
      items.sort((a, b) => {
        if (a.sortDate !== b.sortDate)
          return a.sortDate.localeCompare(b.sortDate);
        return a.jobId.localeCompare(b.jobId);
      });
      // 完全一致の重複行を除外しつつ結合
      const uniqueLines = [...new Set(items.map((item) => item.line))];
      newHist = uniqueLines.join("\n");
    }

    if (oldHist !== newHist) {
      historyUpdatedCount++;
    }
    historyUpdates.push([newHist]);
  }

  if (historyUpdates.length > 0) {
    masterSheet
      .getRange(2, historyColIdx + 1, historyUpdates.length, 1)
      .setValues(historyUpdates);
  }

  // 4. ステータス・採用事業者・内定日の一括同期および採用者一覧・未採用者一覧の更新
  updateCandidateLists(true);

  SpreadsheetApp.getUi().alert(
    `案件管理からの一括反映が完了しました。\n` +
      `・名前欄からの日本語メモ分離・備考移動：${memoMovedCount} 名\n` +
      `・面接履歴の再構築・更新：${historyUpdatedCount} 名\n` +
      `・ステータス／採用事業者／内定日／各種リストの同期完了`
  );
}
