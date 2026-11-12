// =========================================================
// 外部連携：Funtoco「COE info」シートへの同期・転記実行 (13b_外部COE同期実行.gs)
// =========================================================

/**
 * 案件管理シートから現在「採用」となっている全候補者の案件情報をマップで取得するヘルパー関数
 * @return {Map} key: SD-XXXX -> { jobId, company, skillField, offerDate }
 */
function getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ) {
  const jobData = jobSheet.getDataRange().getValues();
  const jCol = getMasterColumnMap(jobSheet);
  const hiredMap = new Map();

  if (jobData.length < 2 || !jCol['案件ID'] || !jCol['面接結果詳細']) return hiredMap;

  const jIdIdx = jCol['案件ID'] - 1;
  const jCompIdx = jCol['事業者名'] ? jCol['事業者名'] - 1 : -1;
  const jSkillIdx = jCol['技能分野'] ? jCol['技能分野'] - 1 : -1;
  const jOfferIdx = jCol['内定日'] ? jCol['内定日'] - 1 : -1;
  const jDetailsIdx = jCol['面接結果詳細'] - 1;

  const formatSlashDate = (val) => {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, sourceTZ, "yyyy/MM/dd");
    const s = String(val).trim().replace(/[年月\-]/g, '/').replace(/日/g, '');
    const parts = s.split('/');
    if (parts.length === 3) {
      return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
    }
    return s;
  };

  for (let i = 1; i < jobData.length; i++) {
    const row = jobData[i];
    const jobId = String(row[jIdIdx] || "").trim();
    const detailsText = String(row[jDetailsIdx] || "").trim();
    if (!jobId || !detailsText) continue;

    const compCell = jCompIdx !== -1 ? String(row[jCompIdx] || "").normalize('NFC') : "";
    const defaultComp = compCell.split(/\r?\n/)[0].trim();
    let currentComp = defaultComp;
    const skillFieldStr = jSkillIdx !== -1 ? String(row[jSkillIdx] || "").trim() : "";
    const offerDateStr = jOfferIdx !== -1 ? formatSlashDate(row[jOfferIdx]) : "";

    const lines = detailsText.split(/\r?\n/).map(l => l.trim()).filter(l => l !== "");
    for (const line of lines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentComp = line.slice(1, -1).normalize('NFC').trim();
        continue;
      }
      if (line.includes('（採用）')) {
        const m = line.match(/^(SD-\d+)/i);
        if (m) {
          const cid = m[1].toUpperCase();
          hiredMap.set(cid, {
            jobId: jobId,
            company: currentComp,
            skillField: skillFieldStr,
            offerDate: offerDateStr
          });
        }
      }
    }
  }
  return hiredMap;
}

/**
 * 特定の候補者（複数可）を外部シート「COE info」へ即時同期（Upsert）する関数
 * @return {Array<number>} 更新または追加されたシート上の行番号の配列
 */
function syncCandidatesToCoeInfo(candidateIds, removeCandidateIds = []) {
  try {
    const targetIds = (Array.isArray(candidateIds) ? candidateIds : [candidateIds])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(id => id !== "");
    const removeIds = (Array.isArray(removeCandidateIds) ? removeCandidateIds : [removeCandidateIds])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(id => id !== "");

    if (targetIds.length === 0 && removeIds.length === 0) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const masterSheet = ss.getSheetByName('登録者マスタ');
    const jobSheet = ss.getSheetByName('案件管理');
    if (!masterSheet || !jobSheet) return [];

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const coeSheet = targetSS.getSheetByName(EXTERNAL_SHEET_COE_INFO);
    if (!coeSheet) {
      console.warn(`外部シートに「${EXTERNAL_SHEET_COE_INFO}」が見つかりません。`);
      return [];
    }

    const layout = getCoeSheetLayout_(coeSheet);
    const { lastCol, colKeys, candIdColIdx, jobIdColIdx } = layout;

    const hiredJobMap = getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ);
    const masterData = masterSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(masterSheet);
    const idIdx = mCol['登録者ID'] - 1;

    const masterRowMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      const cid = String(masterData[i][idIdx] || "").trim().toUpperCase();
      if (cid) masterRowMap.set(cid, masterData[i]);
    }

    // 1. 取消・辞退となった候補者の行があれば先に削除（降順）
    if (removeIds.length > 0 && coeSheet.getLastRow() >= 2) {
      const removeSet = new Set(removeIds);
      const coeData = coeSheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let r = 1; r < coeData.length; r++) {
        const existingCid = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
        if (existingCid && removeSet.has(existingCid) && !hiredJobMap.has(existingCid)) {
          rowsToDelete.push(r + 1);
        }
      }
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach(rowNum => coeSheet.deleteRow(rowNum));
      if (rowsToDelete.length > 0) SpreadsheetApp.flush();
    }

    if (targetIds.length === 0) return [];

    // 2. 現在の COE info シートの既存行と最終データ行（スマートテーブルのフッター直前）を特定
    const existingRowMap = new Map(); // cid -> 1-based row number
    let lastValidRowIndex = 1;
    const lastSheetRow = coeSheet.getLastRow();

    if (lastSheetRow >= 2) {
      const keyColsData = coeSheet.getRange(2, 1, lastSheetRow - 1, Math.max(candIdColIdx, jobIdColIdx) + 1).getValues();
      for (let i = 0; i < keyColsData.length; i++) {
        const jId = String(keyColsData[i][jobIdColIdx] || "").trim();
        const cId = String(keyColsData[i][candIdColIdx] || "").trim().toUpperCase();
        if (jId || cId) {
          lastValidRowIndex = i + 2;
        }
        if (cId && !existingRowMap.has(cId)) {
          existingRowMap.set(cId, i + 2);
        }
      }
    }

    const rowsToAppend = [];
    const affectedRows = [];

    for (const cid of targetIds) {
      const mRow = masterRowMap.get(cid);
      if (!mRow) continue;

      const jobInfo = hiredJobMap.get(cid) || null;
      const rowArray = buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys);

      if (existingRowMap.has(cid)) {
        const targetRowNum = existingRowMap.get(cid);
        coeSheet.getRange(targetRowNum, 1, 1, lastCol).setValues([rowArray]);
        affectedRows.push(targetRowNum);
      } else {
        rowsToAppend.push(rowArray);
      }
    }

    if (rowsToAppend.length > 0) {
      const insertRowIndex = lastValidRowIndex + 1;
      if (insertRowIndex <= coeSheet.getMaxRows()) {
        coeSheet.insertRowsBefore(insertRowIndex, rowsToAppend.length);
      } else {
        coeSheet.insertRowsAfter(coeSheet.getMaxRows(), rowsToAppend.length);
      }
      coeSheet.getRange(insertRowIndex, 1, rowsToAppend.length, lastCol).setValues(rowsToAppend);
      for (let k = 0; k < rowsToAppend.length; k++) {
        affectedRows.push(insertRowIndex + k);
      }
    }

    SpreadsheetApp.flush();
    return affectedRows;
  } catch (e) {
    console.error("COE info シートへの自動同期エラー: " + e.message);
    return [];
  }
}

/**
 * 管理者メニュー用：全採用者のCOE情報を外部シート「COE info」へ一括同期する関数
 * ・日付フォーマット揺れを吸収する正規化比較と、更新・追加された対象行の詳細表示を搭載
 */
function syncAllToCoeInfoSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const masterSheet = ss.getSheetByName('登録者マスタ');
    const jobSheet = ss.getSheetByName('案件管理');
    if (!masterSheet || !jobSheet) return "エラー：「登録者マスタ」または「案件管理」シートが見つかりません。";

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const coeSheet = targetSS.getSheetByName(EXTERNAL_SHEET_COE_INFO);
    if (!coeSheet) return `エラー：外部シートに「${EXTERNAL_SHEET_COE_INFO}」シートが見つかりません。`;

    const layout = getCoeSheetLayout_(coeSheet);
    const { lastCol, colKeys, colLabels, candIdColIdx, jobIdColIdx } = layout;
    const nameColIdx = colKeys.indexOf('name');

    const hiredJobMap = getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ);
    const masterData = masterSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(masterSheet);
    const idIdx = mCol['登録者ID'] - 1;

    const masterRowMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      const cid = String(masterData[i][idIdx] || "").trim().toUpperCase();
      if (cid) masterRowMap.set(cid, masterData[i]);
    }

    // 案件ID順・登録者ID順でソートして同期対象リストを作成
    const sortedEntries = Array.from(hiredJobMap.entries()).sort((a, b) => {
      const jA = a[1].jobId || "";
      const jB = b[1].jobId || "";
      if (jA !== jB) return jA.localeCompare(jB);
      return a[0].localeCompare(b[0]);
    });

    // 1. 既存シートの行を取得し、採用取り消し・辞退となった行を先に削除（降順）
    let coeData = coeSheet.getDataRange().getValues();
    const targetTZ = targetSS.getSpreadsheetTimeZone();
    const rowsToDelete = [];

    for (let r = 1; r < coeData.length; r++) {
      const existingJid = String(coeData[r][jobIdColIdx] || "").trim();
      const existingCid = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
      if (!existingJid && !existingCid) continue; // フッターや空行はスキップ
      if (existingCid && !hiredJobMap.has(existingCid)) {
        rowsToDelete.push(r + 1);
      }
    }
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(rowNum => coeSheet.deleteRow(rowNum));
    if (rowsToDelete.length > 0) {
      SpreadsheetApp.flush();
      coeData = coeSheet.getDataRange().getValues();
    }

    // 2. 削除後の最新データで既存行マップ（coeData上のインデックス）と最終データ行を特定
    const existingIdxMap = new Map(); // cid -> r (1 〜 numDataRows)
    let lastValidRowIndex = 1;
    for (let r = 1; r < coeData.length; r++) {
      const jId = String(coeData[r][jobIdColIdx] || "").trim();
      const cId = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
      if (jId || cId) {
        lastValidRowIndex = r + 1;
      }
      if (cId && !existingIdxMap.has(cId)) {
        existingIdxMap.set(cId, r);
      }
    }

    const numDataRows = lastValidRowIndex - 1;
    const colBuffers = [];
    const colDirty = new Array(lastCol).fill(false);

    if (numDataRows > 0) {
      for (let c = 0; c < lastCol; c++) {
        const buf = [];
        for (let r = 1; r <= numDataRows; r++) {
          buf.push([normalizeCoeCompareVal_(coeData[r][c], targetTZ)]);
        }
        colBuffers.push(buf);
      }
    }

    let updateCount = 0;
    let appendCount = 0;
    const rowsToAppend = [];
    const updatedRowDetails = [];
    const appendedRowDetails = [];

    for (const [cid, jobInfo] of sortedEntries) {
      const mRow = masterRowMap.get(cid);
      if (!mRow) continue;
      const newRowArr = buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys);

      if (existingIdxMap.has(cid)) {
        const rIdx = existingIdxMap.get(cid); // 1 〜 numDataRows
        if (rIdx >= 1 && rIdx <= numDataRows) {
          const bufIdx = rIdx - 1;
          const sheetRowNum = rIdx + 1;
          const changedColNames = [];

          for (let c = 0; c < lastCol; c++) {
            const oldNorm = colBuffers[c][bufIdx][0];
            const newNorm = normalizeCoeCompareVal_(newRowArr[c], sourceTZ);
            if (oldNorm !== newNorm) {
              colBuffers[c][bufIdx][0] = newRowArr[c] !== undefined ? newRowArr[c] : "";
              colDirty[c] = true;
              changedColNames.push(colLabels[c] || `列${c + 1}`);
            }
          }
          if (changedColNames.length > 0) {
            updateCount++;
            const candName = nameColIdx !== -1 ? newRowArr[nameColIdx] : "";
            updatedRowDetails.push(`・${sheetRowNum}行目: ${cid} ${candName} [${changedColNames.join(', ')}]`);
          }
        }
      } else {
        rowsToAppend.push({ cid: cid, rowArr: newRowArr });
        appendCount++;
      }
    }

    // 3. 変更があった列のみ、1列につき1回の setValues で一括更新
    if (numDataRows > 0) {
      for (let c = 0; c < lastCol; c++) {
        if (colDirty[c]) {
          coeSheet.getRange(2, c + 1, numDataRows, 1).setValues(colBuffers[c]);
        }
      }
    }

    // 4. 新規行の一括追加
    if (rowsToAppend.length > 0) {
      const insertRowIndex = lastValidRowIndex + 1;
      if (insertRowIndex <= coeSheet.getMaxRows()) {
        coeSheet.insertRowsBefore(insertRowIndex, rowsToAppend.length);
      } else {
        coeSheet.insertRowsAfter(coeSheet.getMaxRows(), rowsToAppend.length);
      }
      const appendMatrix = rowsToAppend.map(item => item.rowArr);
      coeSheet.getRange(insertRowIndex, 1, rowsToAppend.length, lastCol).setValues(appendMatrix);

      rowsToAppend.forEach((item, idx) => {
        const candName = nameColIdx !== -1 ? item.rowArr[nameColIdx] : "";
        appendedRowDetails.push(`・${insertRowIndex + idx}行目: ${item.cid} ${candName}`);
      });
    }

    SpreadsheetApp.flush();

    const unchangedCount = sortedEntries.length - appendCount - updateCount;
    let msg = `「COE info」シートへの同期が完了しました。\n新規追加: ${appendCount}件\n情報更新: ${updateCount}件\n削除: ${rowsToDelete.length}件\n（変更なしスキップ: ${unchangedCount}件）`;

    if (appendedRowDetails.length > 0) {
      const showAppend = appendedRowDetails.slice(0, 15);
      msg += `\n\n【新規追加の対象行】\n` + showAppend.join("\n");
      if (appendedRowDetails.length > 15) msg += `\n...他 ${appendedRowDetails.length - 15} 件`;
    }

    if (updatedRowDetails.length > 0) {
      const showUpdate = updatedRowDetails.slice(0, 15);
      msg += `\n\n【情報更新の対象行・変更項目】\n` + showUpdate.join("\n");
      if (updatedRowDetails.length > 15) msg += `\n...他 ${updatedRowDetails.length - 15} 件`;
    }

    return msg;
  } catch (e) {
    return "COE info 同期エラー: " + e.message;
  }
}