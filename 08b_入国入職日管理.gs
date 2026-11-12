// =========================================
// 入国・入職日の一括登録処理 (ENTRY_DATE)
// =========================================

function getHiredCandidatesForEntry(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");

    const resultDetailsStr = details.resultDetails || "";
    if (!resultDetailsStr.trim()) {
      return { candidates: [] };
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
      return { candidates: [] };
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

    return { candidates: candidates };
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