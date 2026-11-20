// =========================================
// 入国・入職日の一括登録処理 (ENTRY_DATE) および一時データ移行処理 (TEMP_MIGRATE) (08b_入国入職日管理.gs)
// =========================================

function getHiredCandidatesForEntry(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");

    const resultDetailsStr = details.resultDetails || "";
    if (!resultDetailsStr.trim()) {
      return { candidates: [], offerDate: details.offerDate || "", company: details.company || "", status: details.status || "" };
    }

    const companies = details.company ? details.company.split(/\r?\n/).filter(c => c.trim()) : [];
    const defaultCompany = companies.length > 0 ? companies[0] : "";
    let currentComp = defaultCompany;

    const hiredList = [];
    const detailLines = resultDetailsStr.split(/\r?\n/).filter(l => l.trim() !== "");

    for (const line of detailLines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentComp = line.slice(1, -1).trim();
      } else if (line.includes('（採用）')) {
        const idMatch = line.match(/^(SD-\d+)/);
        if (idMatch) {
          hiredList.push({
            id: idMatch[1].trim(),
            company: companies.length > 1 ? currentComp : ""
          });
        }
      }
    }

    if (hiredList.length === 0) {
      return { candidates: [], offerDate: details.offerDate || "", company: details.company || "", status: details.status || "" };
    }

    const mSheet = getMasterSheet('登録者マスタ');
    const mData = mSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(mSheet);
    const idIdx = mCol['登録者ID'] ? mCol['登録者ID'] - 1 : 0;
    const nameIdx = mCol['名前'] ? mCol['名前'] - 1 : -1;
    const entryIdx = mCol['入国日'] ? mCol['入国日'] - 1 : -1;
    const workIdx = mCol['入職日'] ? mCol['入職日'] - 1 : -1;

    const toIsoDate = (val) => {
      if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy-MM-dd");
      if (typeof val === 'string' && val.trim()) {
        const s = val.trim().replace(/[年月\/]/g, '-').replace(/日/g, '');
        const parts = s.split('-');
        if (parts.length === 3) {
          return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
        }
        return s;
      }
      return '';
    };

    const masterMap = new Map();
    for (let i = 1; i < mData.length; i++) {
      const cid = String(mData[i][idIdx] || "").trim().toUpperCase();
      if (!cid) continue;
      masterMap.set(cid, {
        name: nameIdx !== -1 ? String(mData[i][nameIdx] || "").trim() : "",
        entryDate: entryIdx !== -1 ? toIsoDate(mData[i][entryIdx]) : "",
        workDate: workIdx !== -1 ? toIsoDate(mData[i][workIdx]) : ""
      });
    }

    const candidates = hiredList.map(item => {
      const mInfo = masterMap.get(item.id.toUpperCase()) || { name: "", entryDate: "", workDate: "" };
      return {
        id: item.id,
        display: mInfo.name ? `${item.id} (${mInfo.name})` : item.id,
        company: item.company,
        entryDate: mInfo.entryDate,
        workDate: mInfo.workDate
      };
    });

    return {
      candidates: candidates,
      offerDate: details.offerDate || "",
      company: companies.join('・'),
      status: details.status || ""
    };
  } catch (e) {
    throw new Error(e.message);
  }
}

function updateBatchEntryDates(jobId, records) {
  try {
    if (!Array.isArray(records) || records.length === 0) {
      throw new Error("更新対象の候補者データがありません。");
    }

    const mSheet = getMasterSheet('登録者マスタ');
    if (!mSheet) throw new Error("「登録者マスタ」シートが見つかりません。");

    const mCol = getMasterColumnMap(mSheet);
    const idCol = mCol['登録者ID'];
    const entryCol = mCol['入国日'];
    const workCol = mCol['入職日'];

    if (!idCol || !entryCol || !workCol) {
      throw new Error("登録者マスタに「登録者ID」「入国日」「入職日」のいずれかの列が見つかりません。");
    }

    const mData = mSheet.getDataRange().getValues();
    const recordMap = new Map();
    records.forEach(r => {
      if (r.id) recordMap.set(String(r.id).trim().toUpperCase(), r);
    });

    let updateCount = 0;

    for (let i = 1; i < mData.length; i++) {
      const cid = String(mData[i][idCol - 1] || "").trim().toUpperCase();
      if (recordMap.has(cid)) {
        const rec = recordMap.get(cid);
        const rowNum = i + 1;

        if (rec.entryDate) {
          const eDateObj = new Date(rec.entryDate.replace(/-/g, '/'));
          mSheet.getRange(rowNum, entryCol).setValue(eDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
        } else {
          mSheet.getRange(rowNum, entryCol).clearContent();
        }

        if (rec.workDate) {
          const wDateObj = new Date(rec.workDate.replace(/-/g, '/'));
          mSheet.getRange(rowNum, workCol).setValue(wDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
        } else {
          mSheet.getRange(rowNum, workCol).clearContent();
        }

        updateCount++;
      }
    }

    return `案件 ${jobId} の採用者 ${updateCount} 名の入国日・入職日を更新しました。`;
  } catch (e) {
    throw new Error("一括更新エラー: " + e.message);
  }
}

// =========================================================
// 【一時作業用】内定日・入国日・入職日の一括移行・登録用関数 (TEMP_MIGRATE)
// =========================================================

/**
 * 採用者が存在する案件一覧（未入力項目の有無フラグ付き）を取得する関数
 */
function getTempMigrationJobList() {
  try {
    const jSheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!jSheet || !mSheet) return [];

    const jData = jSheet.getDataRange().getValues();
    const jCol = getMasterColumnMap(jSheet);
    const mData = mSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(mSheet);

    const idIdx = mCol['登録者ID'] ? mCol['登録者ID'] - 1 : 0;
    const entryIdx = mCol['入国日'] ? mCol['入国日'] - 1 : -1;
    const workIdx = mCol['入職日'] ? mCol['入職日'] - 1 : -1;

    const candDateMap = new Map();
    for (let i = 1; i < mData.length; i++) {
      const cid = String(mData[i][idIdx] || "").trim().toUpperCase();
      if (!cid) continue;
      candDateMap.set(cid, {
        hasEntry: entryIdx !== -1 && mData[i][entryIdx] !== "" && mData[i][entryIdx] !== null,
        hasWork: workIdx !== -1 && mData[i][workIdx] !== "" && mData[i][workIdx] !== null
      });
    }

    const jIdIdx = jCol['案件ID'] ? jCol['案件ID'] - 1 : 0;
    const jCompIdx = jCol['事業者名'] ? jCol['事業者名'] - 1 : -1;
    const jOfferIdx = jCol['内定日'] ? jCol['内定日'] - 1 : -1;
    const jDetailIdx = jCol['面接結果詳細'] ? jCol['面接結果詳細'] - 1 : -1;

    const list = [];
    for (let i = 1; i < jData.length; i++) {
      const jobId = String(jData[i][jIdIdx] || "").trim();
      const detailStr = jDetailIdx !== -1 ? String(jData[i][jDetailIdx] || "").trim() : "";
      if (!jobId || !detailStr.includes('（採用）')) continue;

      const compStr = jCompIdx !== -1 ? String(jData[i][jCompIdx] || "").replace(/\r?\n/g, '・').trim() : "";
      const hasOffer = jOfferIdx !== -1 && jData[i][jOfferIdx] !== "" && jData[i][jOfferIdx] !== null;

      const lines = detailStr.split(/\r?\n/);
      const hiredIds = [];
      lines.forEach(l => {
        if (l.includes('（採用）')) {
          const m = l.match(/^(SD-\d+)/i);
          if (m) hiredIds.push(m[1].toUpperCase());
        }
      });

      if (hiredIds.length === 0) continue;

      let allEntry = true;
      let allWork = true;
      hiredIds.forEach(cid => {
        const info = candDateMap.get(cid);
        if (!info || !info.hasEntry) allEntry = false;
        if (!info || !info.hasWork) allWork = false;
      });

      const missingTags = [];
      if (!hasOffer) missingTags.push("内定日未");
      if (!allEntry) missingTags.push("入国日未");
      if (!allWork) missingTags.push("入職日未");

      const statusBadge = missingTags.length > 0 ? `【${missingTags.join('・')}】` : "【入力済】";
      list.push({
        jobId: jobId,
        hasMissing: missingTags.length > 0,
        label: `${jobId} ｜ ${compStr} （採用${hiredIds.length}名） ${statusBadge}`
      });
    }

    return list;
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * 一時作業用：指定案件の内定日・入国日・入職日を一括保存する関数
 */
function updateTempMigrationDates(jobId, offerDateStr, records) {
  try {
    if (!jobId) throw new Error("案件IDが指定されていません。");
    if (!Array.isArray(records) || records.length === 0) {
      throw new Error("更新対象の候補者データがありません。");
    }

    const jSheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!jSheet || !mSheet) throw new Error("シートが見つかりません。");

    const jCol = getMasterColumnMap(jSheet);
    const mCol = getMasterColumnMap(mSheet);

    const jIdCol = jCol['案件ID'];
    const jOfferCol = jCol['内定日'];
    if (!jIdCol || !jOfferCol) throw new Error("案件管理シートに「案件ID」または「内定日」列が見つかりません。");

    const mIdCol = mCol['登録者ID'];
    const mOfferCol = mCol['内定日'];
    const mEntryCol = mCol['入国日'];
    const mWorkCol = mCol['入職日'];
    if (!mIdCol || !mEntryCol || !mWorkCol) {
      throw new Error("登録者マスタに必須列（登録者ID・入国日・入職日）が見つかりません。");
    }

    const offerDateObj = offerDateStr ? new Date(offerDateStr.replace(/-/g, '/')) : null;

    // 1. 案件管理シートの「内定日」を更新（ステータス等は一切変更しない）
    const jData = jSheet.getDataRange().getValues();
    const targetJobId = String(jobId).trim().toUpperCase();
    for (let i = 1; i < jData.length; i++) {
      if (String(jData[i][jIdCol - 1] || "").trim().toUpperCase() === targetJobId) {
        const jRowNum = i + 1;
        if (offerDateObj) {
          jSheet.getRange(jRowNum, jOfferCol).setValue(offerDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
        } else {
          jSheet.getRange(jRowNum, jOfferCol).clearContent();
        }
        break;
      }
    }

    // 2. 登録者マスタの「内定日」「入国日」「入職日」を更新
    const mData = mSheet.getDataRange().getValues();
    const recordMap = new Map();
    records.forEach(r => {
      if (r.id) recordMap.set(String(r.id).trim().toUpperCase(), r);
    });

    let updateCount = 0;
    const updatedIds = [];

    for (let i = 1; i < mData.length; i++) {
      const cid = String(mData[i][mIdCol - 1] || "").trim().toUpperCase();
      if (recordMap.has(cid)) {
        const rec = recordMap.get(cid);
        const rowNum = i + 1;

        if (mOfferCol) {
          if (offerDateObj) {
            mSheet.getRange(rowNum, mOfferCol).setValue(offerDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
          } else {
            mSheet.getRange(rowNum, mOfferCol).clearContent();
          }
        }

        if (rec.entryDate) {
          const eDateObj = new Date(rec.entryDate.replace(/-/g, '/'));
          mSheet.getRange(rowNum, mEntryCol).setValue(eDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
        } else {
          mSheet.getRange(rowNum, mEntryCol).clearContent();
        }

        if (rec.workDate) {
          const wDateObj = new Date(rec.workDate.replace(/-/g, '/'));
          mSheet.getRange(rowNum, mWorkCol).setValue(wDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
        } else {
          mSheet.getRange(rowNum, mWorkCol).clearContent();
        }

        updateCount++;
        updatedIds.push(cid);
      }
    }

    SpreadsheetApp.flush();

    // 3. 外部シート「COE info」の内定日も最新状態に連動更新
    if (updatedIds.length > 0 && typeof syncCandidatesToCoeInfo === 'function') {
      syncCandidatesToCoeInfo(updatedIds);
    }

    return `案件 ${jobId} の内定日、および採用者 ${updateCount} 名の入国日・入職日を一括保存しました。`;
  } catch (e) {
    throw new Error("一括保存エラー: " + e.message);
  }
}