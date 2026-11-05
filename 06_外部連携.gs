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

    // 両方のスプレッドシートのタイムゾーンを取得（時差ズレ防止）
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const targetTZ = targetSS.getSpreadsheetTimeZone();

    const sourceData = sourceSheet.getDataRange().getValues();
    const sourceMap = getMasterColumnMap(sourceSheet);

    const masterData = masterSheet.getDataRange().getValues();
    const masterMap = getMasterColumnMap(masterSheet);

    const targetMap = getMasterColumnMap(targetSheet);
    if (!targetMap['案件ID'] || !targetMap['登録者ID']) {
      throw new Error("外部シート「支払い管理」に「案件ID」または「登録者ID」列が見つかりません。");
    }

    if (sourceData.length < 2) return "同期対象の案件がありません。";

    // 日付フォーマット用ヘルパー関数（NaNやundefined等のゴミデータを排除）
    const formatDateStr = (val, tz) => {
      if (!val) return "";
      if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy/MM/dd");
      const s = String(val).trim().replace(/[年月\-]/g, '/').replace(/日/g, '');
      if (s === "NaN" || s === "undefined" || s === "Invalid Date") return "";
      const parts = s.split('/');
      if (parts.length === 3) {
        return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
      }
      return s;
    };

    // 1. 登録者マスタから 内定日・入国日・入職日の抽出
    const candidateInfo = {};
    if (masterMap['登録者ID']) {
      const idIdx = masterMap['登録者ID'] - 1;
      const offerIdx = masterMap['内定日'] ? masterMap['内定日'] - 1 : -1;
      const entryIdx = masterMap['入国日'] ? masterMap['入国日'] - 1 : -1;
      const workIdx = masterMap['入職日'] ? masterMap['入職日'] - 1 : -1;

      for (let i = 1; i < masterData.length; i++) {
        const cid = String(masterData[i][idIdx]).trim().toUpperCase();
        if (!cid) continue;

        const oDate = offerIdx !== -1 ? formatDateStr(masterData[i][offerIdx], sourceTZ) : "";
        const eDate = entryIdx !== -1 ? formatDateStr(masterData[i][entryIdx], sourceTZ) : "";
        const wDate = workIdx !== -1 ? formatDateStr(masterData[i][workIdx], sourceTZ) : "";

        candidateInfo[cid] = { offerDate: oDate, entryDate: eDate, workDate: wDate };
      }
    }

    // 2. 案件管理から同期対象レコード（正本）の抽出
    const syncRecords = [];
    const syncRecordMap = new Map(); // key: "JOB-XXXX_SD-XXXX" -> record

    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const jobID = sourceMap['案件ID'] ? String(row[sourceMap['案件ID'] - 1] || "").trim() : "";
      if (!jobID) continue;

      const detailsText = sourceMap['面接結果詳細'] ? String(row[sourceMap['面接結果詳細'] - 1] || "").trim() : "";
      if (!detailsText) continue;

      const companyNameCell = sourceMap['事業者名'] ? String(row[sourceMap['事業者名'] - 1] || "") : "";
      const defaultCompany = companyNameCell.split(/\r?\n/)[0].trim();
      const fieldName = sourceMap['技能分野'] ? String(row[sourceMap['技能分野'] - 1] || "").trim() : "";

      const interviewDate = sourceMap['面接日'] ? formatDateStr(row[sourceMap['面接日'] - 1], sourceTZ) : "";
      const jobOfferDate = sourceMap['内定日'] ? formatDateStr(row[sourceMap['内定日'] - 1], sourceTZ) : "";

      const dLines = detailsText.split(/\r?\n/).filter(line => line.trim() !== "");
      let currentCompany = defaultCompany;
      let hasHired = false;

      for (const line of dLines) {
        const trimmedLine = line.trim();
        if (trimmedLine.startsWith('【') && trimmedLine.endsWith('】')) {
          currentCompany = trimmedLine.slice(1, -1).trim();
          continue;
        }

        if (trimmedLine.includes('（採用）')) {
          const idMatch = trimmedLine.match(/^(SD-\d+)/i);
          if (idMatch) {
            const candidateID = idMatch[1].trim().toUpperCase();
            const nMatch = trimmedLine.match(/^SD-\d+-(.*?)(?:（.*?）)$/i);
            const candidateName = nMatch ? nMatch[1].trim() : "";
            hasHired = true;

            let oDate = jobOfferDate;
            let eDate = "";
            let wDate = "";
            if (candidateID !== "採用者なし" && candidateInfo[candidateID]) {
              if (!oDate) oDate = candidateInfo[candidateID].offerDate;
              eDate = candidateInfo[candidateID].entryDate;
              wDate = candidateInfo[candidateID].workDate;
            }

            const rec = {
              jobID: jobID,
              candidateID: candidateID,
              companyName: currentCompany,
              fieldName: fieldName,
              candidateName: candidateName,
              interviewDate: interviewDate,
              offerDate: oDate,
              entryDate: eDate,
              workDate: wDate
            };
            const key = `${jobID}_${candidateID}`;
            if (!syncRecordMap.has(key)) {
              syncRecords.push(rec);
              syncRecordMap.set(key, rec);
            }
          }
        }
      }

      if (!hasHired) {
        const rec = {
          jobID: jobID,
          candidateID: "採用者なし",
          companyName: defaultCompany,
          fieldName: fieldName,
          candidateName: "採用者なし",
          interviewDate: interviewDate,
          offerDate: "",
          entryDate: "",
          workDate: ""
        };
        const key = `${jobID}_採用者なし`;
        if (!syncRecordMap.has(key)) {
          syncRecords.push(rec);
          syncRecordMap.set(key, rec);
        }
      }
    }

    // 3. 支払い管理シートの現在データを取得し、まず「不要行の削除」を先に実行する（行ズレ防止）
    const tJobIdx = targetMap['案件ID'] - 1;
    const tIdIdx = targetMap['登録者ID'] - 1;

    let targetData = targetSheet.getDataRange().getValues();
    const rowsToDelete = [];

    for (let i = 1; i < targetData.length; i++) {
      const jId = String(targetData[i][tJobIdx] || "").trim();
      const cId = String(targetData[i][tIdIdx] || "").trim();
      if (!jId && !cId) continue; // フッターや空行はスキップ

      const key = `${jId}_${cId}`;
      if (!syncRecordMap.has(key)) {
        rowsToDelete.push(i + 1); // シート上の行番号（1始まり）
      }
    }

    // 下の行から上に向かって（降順）削除を実行
    rowsToDelete.sort((a, b) => b - a);
    for (const rowNum of rowsToDelete) {
      targetSheet.deleteRow(rowNum);
    }
    if (rowsToDelete.length > 0) {
      SpreadsheetApp.flush();
      targetData = targetSheet.getDataRange().getValues(); // 削除後の最新行配置を再取得
    }

    // 4. 削除後の最新データで既存行のマッピングと最終データ行（フッター位置）を特定
    const targetKeys = {};
    const existingCandidateMap = new Map();
    let lastValidRowIndex = 1; // ヘッダー行（1行目）

    for (let i = 1; i < targetData.length; i++) {
      const jId = String(targetData[i][tJobIdx] || "").trim();
      const cId = String(targetData[i][tIdIdx] || "").trim();

      if (jId || cId) {
        lastValidRowIndex = i + 1; // データが存在する最後の行番号（1始まり）
      }
      if (jId && cId) {
        const key = `${jId}_${cId}`;
        // 万一外部シート側に同一キーが重複している場合は最初の行を採用
        if (targetKeys[key] === undefined) {
          targetKeys[key] = i; // 0始まりの配列インデックス
        }
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

    // フッター行（新規挿入位置）は「データが存在する最後の行のすぐ下」
    const footerRowIndex = lastValidRowIndex + 1;

    // 5. 既存行の差分チェックと一括更新（B列「ステータス」やL〜N列「金額・入金日・備考」には一切触れない）
    const updateCols = ['事業者名', '技能分野', '名前', '面接日', '内定日', '入国日', '入職日'];
    const dateCols = ['面接日', '内定日', '入国日', '入職日'];

    let updateCount = 0;
    let appendCount = 0;
    const warnings = new Set();
    const newRowsToAppend = [];

    // 列ごとの書き込み用バッファ（変更があった列だけを一括書き込みする）
    const numDataRows = lastValidRowIndex - 1;
    const colBuffers = {};
    const colDirty = {};

    if (numDataRows > 0) {
      updateCols.forEach(colName => {
        if (targetMap[colName]) {
          const cIdx = targetMap[colName] - 1;
          colBuffers[colName] = [];
          for (let r = 1; r <= numDataRows; r++) {
            const rawVal = targetData[r][cIdx];
            if (dateCols.includes(colName)) {
              colBuffers[colName].push([formatDateStr(rawVal, targetTZ)]);
            } else {
              colBuffers[colName].push([rawVal !== null && rawVal !== undefined ? String(rawVal).trim() : ""]);
            }
          }
          colDirty[colName] = false;
        }
      });
    }

    for (const record of syncRecords) {
      const key = `${record.jobID}_${record.candidateID}`;

      if (targetKeys[key] !== undefined) {
        const tIdx = targetKeys[key]; // 1 〜 numDataRows
        if (tIdx >= 1 && tIdx <= numDataRows) {
          const bufIdx = tIdx - 1;
          let rowChanged = false;

          const newValuesMap = {
            '事業者名': String(record.companyName || "").trim(),
            '技能分野': String(record.fieldName || "").trim(),
            '名前': String(record.candidateName || "").trim(),
            '面接日': formatDateStr(record.interviewDate, targetTZ),
            '内定日': formatDateStr(record.offerDate, targetTZ),
            '入国日': formatDateStr(record.entryDate, targetTZ),
            '入職日': formatDateStr(record.workDate, targetTZ)
          };

          updateCols.forEach(colName => {
            if (colBuffers[colName]) {
              const oldStr = String(colBuffers[colName][bufIdx][0] || "").trim();
              const newStr = newValuesMap[colName];
              if (oldStr !== newStr) {
                colBuffers[colName][bufIdx][0] = newStr;
                colDirty[colName] = true;
                rowChanged = true;
              }
            }
          });

          if (rowChanged) updateCount++;
        }
        continue;
      }

      // 新規追加データの作成
      if (record.candidateID !== "採用者なし" && existingCandidateMap.has(record.candidateID)) {
        const oldJobs = existingCandidateMap.get(record.candidateID).join(", ");
        warnings.add(`・${record.candidateID} ${record.candidateName} (既存案件ID: ${oldJobs})`);
      }

      newRowsToAppend.push(record);
      appendCount++;
    }

    // 変更があった列のみ、1列につき1回の setValues で一括更新（B列ステータスやL〜N列は完全にノータッチ）
    if (numDataRows > 0) {
      updateCols.forEach(colName => {
        if (colDirty[colName] && targetMap[colName]) {
          const colNum = targetMap[colName];
          const range = targetSheet.getRange(2, colNum, numDataRows, 1);
          range.setValues(colBuffers[colName]);
          if (dateCols.includes(colName)) {
            range.setNumberFormat('yyyy/MM/dd');
          }
        }
      });
    }

    // 6. 新規行の一括追加（A列チェッカー と C列〜K列のみ書き込み、B列ステータスやL〜N列はスマートテーブルの初期値を保護）
    if (newRowsToAppend.length > 0) {
      targetSheet.insertRowsBefore(footerRowIndex, newRowsToAppend.length);

      // A列（チェッカー）に false を一括セット
      if (targetMap['チェッカー']) {
        const checkerVals = newRowsToAppend.map(() => [false]);
        targetSheet.getRange(footerRowIndex, targetMap['チェッカー'], newRowsToAppend.length, 1).setValues(checkerVals);
      }

      // C列〜K列（案件ID, 事業者名, 技能分野, 登録者ID, 名前, 面接日, 内定日, 入国日, 入職日）が連続しているか確認して一括セット
      const syncColNames = ['案件ID', '事業者名', '技能分野', '登録者ID', '名前', '面接日', '内定日', '入国日', '入職日'];
      syncColNames.forEach(colName => {
        if (targetMap[colName]) {
          const colNum = targetMap[colName];
          const colVals = newRowsToAppend.map(rec => {
            if (colName === '案件ID') return [rec.jobID];
            if (colName === '事業者名') return [rec.companyName];
            if (colName === '技能分野') return [rec.fieldName];
            if (colName === '登録者ID') return [rec.candidateID];
            if (colName === '名前') return [rec.candidateName];
            if (colName === '面接日') return [rec.interviewDate];
            if (colName === '内定日') return [rec.offerDate];
            if (colName === '入国日') return [rec.entryDate];
            if (colName === '入職日') return [rec.workDate];
            return [""];
          });
          const range = targetSheet.getRange(footerRowIndex, colNum, newRowsToAppend.length, 1);
          range.setValues(colVals);
          if (dateCols.includes(colName)) {
            range.setNumberFormat('yyyy/MM/dd');
          }
        }
      });
    }

    SpreadsheetApp.flush();

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