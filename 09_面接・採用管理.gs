// =========================================
// 面接および採用に関する操作
// =========================================

// ---------------------------------------------------------
// 1. 面接結果登録 (改修前: HIRE) 用
// ---------------------------------------------------------

function getJobCandidates(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");
    if (!details.interviewDate) throw new Error("面接日が設定されていません。\n先に「案件更新/削除」から面接日を登録してください。");
    
    // ▼ バックエンドの強固なブロック（既に結果がある場合は弾く）
    if (details.hireNames && details.hireNames.trim() !== "") {
        throw new Error("この案件は既に面接結果が登録されています。\n「面接結果の修正・更新」メニューを使用してください。");
    }
    
    const companies = details.company ? details.company.split(/\r?\n/).filter(c => c.trim()) : [];
    const ids = details.candidates ? details.candidates.split(/\r?\n/).filter(id => id.trim()) : [];

    const candDict = getCandidateDict(); 
    const candidates = ids.map(id => {
      const cleanId = id.split('-').slice(0, 2).join('-').trim();
      return { 
        id: cleanId, 
        display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId].name})` : id, 
        name: candDict[cleanId] ? candDict[cleanId].name : "" 
      };
    }).filter(c => c.id);

    return { candidates: candidates, companies: companies };
  } catch(e) { throw new Error(e.message); }
}

function registerHire(jobId, hiredData) {
  try {
    const sheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!sheet || !mSheet) throw new Error("シートへのアクセスに失敗しました。");

    const mCol = getMasterColumnMap(mSheet);
    const data = sheet.getDataRange().getValues();
    const mData = mSheet.getDataRange().getValues();
    
    let companyNamesText = "", rawInterviewDate = "", allCandidatesRaw = "", targetJobRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(jobId).trim()) {
        companyNamesText = String(data[i][3]).trim();
        allCandidatesRaw = String(data[i][5]);
        rawInterviewDate = data[i][6];
        targetJobRow = i + 1;
        break;
      }
    }
    if (!companyNamesText) throw new Error("案件が見つかりません。");
    if (!rawInterviewDate) throw new Error("面接日が設定されていません。\n先に「案件更新/削除」から面接日を登録してください。");

    let formattedDate = "日付不明";
    if (rawInterviewDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawInterviewDate, "JST", "yyyy/MM/dd");
    } else if (rawInterviewDate) {
      formattedDate = String(rawInterviewDate).replace(/[年月]/g, '/').replace(/日/g, '');
    }

    const companyNames = companyNamesText.split(/\r?\n/).filter(c => c.trim());
    const defaultCompany = companyNames.join('・') || "";

    const candDict = getCandidateDict();
    const allCandidateIds = allCandidatesRaw.split(/\r?\n/).map(line => line.split('-').slice(0, 2).join('-').trim()).filter(id => id !== "");
    
    const hiredIdMap = new Map();
    hiredData.forEach(item => {
       hiredIdMap.set(String(item.id).trim(), item.company);
    });
    
    allCandidateIds.forEach(candId => {
      const isHired = hiredIdMap.has(candId);
      const hiredComp = isHired ? hiredIdMap.get(candId) : defaultCompany;
      const resultText = isHired ? `（採用）` : "（不採用）";
      const newHistoryLine = `${formattedDate}：${hiredComp}${resultText}`;

      for (let j = 1; j < mData.length; j++) {
        if (String(mData[j][0]).trim() === candId) {
          const rowIdx = j + 1;
          if (isHired) {
            if (mCol['ステータス']) mSheet.getRange(rowIdx, mCol['ステータス']).setValue('採用');
            if (mCol['採用事業者']) mSheet.getRange(rowIdx, mCol['採用事業者']).setValue(hiredComp);
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
    if (hiredData.length > 0) {
      if (companyNames.length <= 1) {
        hiredNamesText = hiredData.map(item => {
           const name = candDict[item.id] ? candDict[item.id].name : "";
           return name ? `${item.id}-${name}` : `${item.id}`;
        }).join('\n');
      } else {
        const grouped = {};
        hiredData.forEach(item => {
          if (!grouped[item.company]) grouped[item.company] = [];
          const name = candDict[item.id] ? candDict[item.id].name : "";
          grouped[item.company].push(name ? `${item.id}-${name}` : `${item.id}`);
        });
        
        let lines = [];
        for (const [comp, cands] of Object.entries(grouped)) {
          lines.push(`【${comp}】`);
          lines.push(...cands);
        }
        hiredNamesText = lines.join('\n');
      }
    }

    sheet.getRange(targetJobRow, 8).setValue(hiredNamesText);
    sheet.getRange(targetJobRow, 2).setValue(hiredData.length > 0 ? '入国準備' : '終了');

    if (hiredData.length > 0) return `${hiredData.length} 名の面接結果（ステータス：入国準備）、および対象候補者全員の「面接履歴」への追記が完了しました。`;
    return `「採用者なし」として案件を終了し、対象候補者全員の「面接履歴」への追記が完了しました。`;
  } catch(e) { throw new Error(e.message); }
}

// ---------------------------------------------------------
// 2. 面接結果の修正・更新 (改修後: HIRE_EDIT) 用
// ---------------------------------------------------------

function getJobCandidatesEdit(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");
    if (!details.interviewDate) throw new Error("面接日が設定されていません。\n先に「案件更新/削除」から面接日を登録してください。");
    
    // ▼ バックエンドの強固なブロック（まだ結果がない場合は弾く）
    if (!details.hireNames || details.hireNames.trim() === "") {
        throw new Error("この案件はまだ面接結果が登録されていません。\n「面接結果登録」メニューを使用してください。");
    }

    const companies = details.company ? details.company.split(/\r?\n/).filter(c => c.trim()) : [];
    const ids = details.candidates ? details.candidates.split(/\r?\n/).filter(id => id.trim()) : [];

    const mSheet = getMasterSheet('登録者マスタ');
    const mData = mSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(mSheet);

    const formattedDate = String(details.interviewDate).replace(/-/g, '/');

    const candDict = getCandidateDict(); 
    const candidates = ids.map(id => {
      const cleanId = id.split('-').slice(0, 2).join('-').trim();
      
      let pastResult = "不採用"; 
      let pastCompany = "";

      for (let j = 1; j < mData.length; j++) {
        if (String(mData[j][0]).trim() === cleanId) {
          if (mCol['面接履歴']) {
            const history = String(mData[j][mCol['面接履歴']-1] || "");
            const lines = history.split(/\r?\n/);
            for (let k = lines.length - 1; k >= 0; k--) {
              const line = lines[k];
              if (line.startsWith(formattedDate + "：")) {
                const match = line.match(/：(.*?)(?:（(.*?)）)?$/);
                if (match) {
                  pastCompany = match[1].trim();
                  const resultText = match[2] ? match[2].trim() : "";
                  
                  if (resultText === "採用") pastResult = "採用";
                  else if (resultText === "不採用") pastResult = "不採用";
                  else if (resultText === "内定辞退") pastResult = "内定辞退（候補者都合）";
                  else if (resultText === "事業者都合取消") pastResult = "内定取消（事業者都合）";
                }
                break;
              }
            }
          }
          break;
        }
      }

      return { 
        id: cleanId, 
        display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId].name})` : id, 
        name: candDict[cleanId] ? candDict[cleanId].name : "",
        pastResult: pastResult,
        pastCompany: pastCompany
      };
    }).filter(c => c.id);

    return { candidates: candidates, companies: companies };
  } catch(e) { throw new Error(e.message); }
}

function updateHire(jobId, resultData) {
  try {
    const sheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!sheet || !mSheet) throw new Error("シートへのアクセスに失敗しました。");

    const mCol = getMasterColumnMap(mSheet);
    const data = sheet.getDataRange().getValues();
    const mData = mSheet.getDataRange().getValues();
    
    let companyNamesText = "", rawInterviewDate = "", targetJobRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(jobId).trim()) {
        companyNamesText = String(data[i][3]).trim();
        rawInterviewDate = data[i][6];
        targetJobRow = i + 1;
        break;
      }
    }
    if (!companyNamesText) throw new Error("案件が見つかりません。");
    if (!rawInterviewDate) throw new Error("面接日が設定されていません。");

    let formattedDate = "日付不明";
    if (rawInterviewDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawInterviewDate, "JST", "yyyy/MM/dd");
    } else if (rawInterviewDate) {
      formattedDate = String(rawInterviewDate).replace(/[年月]/g, '/').replace(/日/g, '');
    }

    const companyNames = companyNamesText.split(/\r?\n/).filter(c => c.trim());
    const defaultCompany = companyNames.join('・') || "";
    const candDict = getCandidateDict();
    
    let hiredList = [];
    
    resultData.forEach(item => {
      const candId = String(item.id).trim();
      const result = item.result;
      const comp = item.company || defaultCompany;
      
      let suffix = "";
      let newStatus = null;
      let writeHistory = false;
      let deleteHistory = false;

      if (result === '採用') {
          suffix = "（採用）";
          newStatus = "採用";
          writeHistory = true;
          hiredList.push(item);
      } else if (result === '不採用') {
          suffix = "（不採用）";
          newStatus = "未採用";
          writeHistory = true;
      } else if (result === '内定辞退（候補者都合）') {
          suffix = "（内定辞退）";
          newStatus = "辞退";
          writeHistory = true;
      } else if (result === '内定取消（事業者都合）') {
          suffix = "（事業者都合取消）";
          newStatus = "未採用";
          writeHistory = true;
      } else if (result === '面接結果の取消') {
          newStatus = "未採用";
          deleteHistory = true;
      }
      
      for (let j = 1; j < mData.length; j++) {
        if (String(mData[j][0]).trim() === candId) {
          const rowIdx = j + 1;
          
          if (newStatus && mCol['ステータス']) {
            mSheet.getRange(rowIdx, mCol['ステータス']).setValue(newStatus);
          }
          
          if (mCol['採用事業者']) {
              if (result === '採用') {
                  mSheet.getRange(rowIdx, mCol['採用事業者']).setValue(comp);
              } else {
                  mSheet.getRange(rowIdx, mCol['採用事業者']).setValue(""); 
              }
          }

          if (mCol['面接履歴']) {
            const historyCell = mSheet.getRange(rowIdx, mCol['面接履歴']);
            const currentHistory = String(historyCell.getValue() || "").trim();
            let lines = currentHistory.split(/\r?\n/).filter(l => l.trim() !== "");
            
            let existingIdx = lines.findIndex(l => l.startsWith(formattedDate + "："));
            
            if (deleteHistory) {
                if (existingIdx !== -1) lines.splice(existingIdx, 1);
            } else if (writeHistory) {
                let histComp = (result === '不採用') ? defaultCompany : comp;
                let newLine = `${formattedDate}：${histComp}${suffix}`;
                if (existingIdx !== -1) {
                    lines[existingIdx] = newLine; 
                } else {
                    lines.push(newLine); 
                }
            }
            historyCell.setValue(lines.join("\n"));
          }
          break;
        }
      }
    });
    
    let hiredNamesText = "採用者なし";
    if (hiredList.length > 0) {
        if (companyNames.length <= 1) {
            hiredNamesText = hiredList.map(item => {
                const name = candDict[item.id] ? candDict[item.id].name : "";
                return name ? `${item.id}-${name}` : `${item.id}`;
            }).join('\n');
        } else {
            const grouped = {};
            hiredList.forEach(item => {
                if (!grouped[item.company]) grouped[item.company] = [];
                const name = candDict[item.id] ? candDict[item.id].name : "";
                grouped[item.company].push(name ? `${item.id}-${name}` : `${item.id}`);
            });
            let lines = [];
            for (const [comp, cands] of Object.entries(grouped)) {
                lines.push(`【${comp}】`);
                lines.push(...cands);
            }
            hiredNamesText = lines.join('\n');
        }
    }
    
    sheet.getRange(targetJobRow, 8).setValue(hiredNamesText);
    sheet.getRange(targetJobRow, 2).setValue(hiredList.length > 0 ? '入国準備' : '終了');

    return `面接結果の更新が完了しました。\n（マスタのステータスと履歴が自動更新されました）`;
  } catch(e) { throw new Error(e.message); }
}