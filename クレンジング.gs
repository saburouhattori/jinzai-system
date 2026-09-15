function cleanUpInterviewHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('登録者マスタ');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const col = {};
  headers.forEach((h, i) => col[String(h).replace(/\n/g, '').replace(/\s/g, '')] = i);
  
  const historyIdx = col['面接履歴'];
  const statusIdx = col['ステータス'];
  const hiredCompIdx = col['採用事業者'];
  
  if (historyIdx === undefined || statusIdx === undefined || hiredCompIdx === undefined) {
    SpreadsheetApp.getUi().alert("必要な列（面接履歴、ステータス、採用事業者）が見つかりません。");
    return;
  }

  let updateCount = 0;
  
  // 書き戻し用の配列（AQ列: 面接履歴 のみ）
  const historyColumnData = [];
  historyColumnData.push([headers[historyIdx]]); // 1行目のヘッダー

  for (let i = 1; i < data.length; i++) {
    const rawHistory = String(data[i][historyIdx] || "");
    const status = String(data[i][statusIdx] || "").trim();
    // スペースを除去して比較しやすくする
    const hiredComp = String(data[i][hiredCompIdx] || "").trim().replace(/\s/g, '');

    if (!rawHistory.trim()) {
      historyColumnData.push([""]);
      continue;
    }

    const lines = rawHistory.split(/\r?\n/).map(l => l.trim()).filter(l => l !== "");
    const parsedLines = [];
    
    lines.forEach(line => {
      // 「不明」「-」などの無効行はスキップ
      if (line === "不明" || line === "-" || line === "無し") return;

      // 1. 新フォーマット判定 (例: 2022/10/21：株式会社〇〇（不採用）)
      const newFmtMatch = line.match(/^(\d{4}\/\d{1,2}\/\d{1,2})：(.*?)（(.*?)）$/);
      if (newFmtMatch) {
        parsedLines.push({
          date: newFmtMatch[1],
          company: newFmtMatch[2].trim(),
          result: newFmtMatch[3].trim(),
          isNew: true,
          raw: line
        });
        return;
      }

      // 2. 旧フォーマットの整理
      // 行頭の丸数字などを削除
      let cleanLine = line.replace(/^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳\d\.\s・]+/, '').trim();
      if (!cleanLine) return;

      // 日付の抽出 (例: 株式会社〇〇（2022年10月21日）)
      let dateStr = "日付不明";
      let compName = cleanLine;
      
      const dateMatch = cleanLine.match(/(.*?)[（(](20\d{2})年(\d{1,2})月(?:(\d{1,2})日)?[)）]/);
      if (dateMatch) {
        compName = dateMatch[1].trim();
        const year = dateMatch[2];
        const month = dateMatch[3].padStart(2, '0');
        const day = dateMatch[4] ? dateMatch[4].padStart(2, '0') : '01';
        dateStr = `${year}/${month}/${day}`;
      }

      // 結果の判定
      let resultStr = "不採用";
      if (status === "採用") {
        const normCompName = compName.replace(/\s/g, '');
        // 採用事業者にこの企業名が含まれていれば「採用」とする
        if (normCompName && hiredComp.includes(normCompName)) {
          resultStr = "採用";
        }
      }

      parsedLines.push({
        date: dateStr,
        company: compName,
        result: resultStr,
        isNew: false,
        raw: `${dateStr}：${compName}（${resultStr}）`
      });
    });

    // 3. 重複の除去（新フォーマット優先）
    const finalLines = [];
    const seenCompanies = new Set();
    
    // まず新フォーマットを登録
    parsedLines.filter(p => p.isNew).forEach(p => {
      finalLines.push(p.raw);
      seenCompanies.add(p.company.replace(/\s/g, ''));
    });

    // 旧フォーマットを登録（まだ追加されていない企業のみ）
    parsedLines.filter(p => !p.isNew).forEach(p => {
      const normComp = p.company.replace(/\s/g, '');
      if (!seenCompanies.has(normComp)) {
        finalLines.push(p.raw);
        seenCompanies.add(normComp);
      }
    });

    const newHistoryText = finalLines.join('\n');
    historyColumnData.push([newHistoryText]);

    if (newHistoryText !== rawHistory) {
      updateCount++;
    }
  }

  // シートに変更を反映（面接履歴のAQ列のみを一括で上書き）
  if (updateCount > 0) {
    sheet.getRange(1, historyIdx + 1, historyColumnData.length, 1).setValues(historyColumnData);
  }
  
  SpreadsheetApp.getUi().alert(`${updateCount} 件の候補者の面接履歴を整理・統合しました！`);
}