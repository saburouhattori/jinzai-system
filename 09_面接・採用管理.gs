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
    if (details.resultDetails && details.resultDetails.trim() !== "") {
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

    return { candidates: candidates, companies: companies, offerDate: details.offerDate };
  } catch(e) { throw new Error(e.message); }
}

function registerHire(jobId, hiredData, offerDateStr) {
  try {
    const sheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!sheet || !mSheet) throw new Error("シートへのアクセスに失敗しました。");

    const mCol = getMasterColumnMap(mSheet);
    const colMap = getMasterColumnMap(sheet);
    const data = sheet.getDataRange().getValues();
    const mData = mSheet.getDataRange().getValues();
    if (!colMap['案件ID']) throw new Error("案件管理シートに案件IDが見つかりません。");
    
    let companyNamesText = "", rawInterviewDate = "", allCandidatesRaw = "", targetJobRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colMap['案件ID'] - 1]).trim() === String(jobId).trim()) {
        companyNamesText = String(data[i][colMap['事業者名'] - 1]).trim();
        allCandidatesRaw = String(data[i][colMap['候補者名'] - 1]);
        rawInterviewDate = data[i][colMap['面接日'] - 1];
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

    let offerDateObj = "";
    if (offerDateStr) {
      offerDateObj = new Date(offerDateStr.replace(/-/g, '/'));
    }

    const companyNames = companyNamesText.split(/\r?\n/).filter(c => c.trim());
    const defaultCompany = companyNames.join('・') || "";

    const candDict = getCandidateDict();
    const allCandidateIds = allCandidatesRaw.split(/\r?\n/).map(line => line.split('-').slice(0, 2).join('-').trim()).filter(id => id !== "");
    
    const hiredIdMap = new Map();
    hiredData.forEach(item => {
       hiredIdMap.set(String(item.id).trim(), item.company);
    });
    
    let resultDataArray = [];

    allCandidateIds.forEach(candId => {
      const isHired = hiredIdMap.has(candId);
      const hiredComp = isHired ? hiredIdMap.get(candId) : defaultCompany;
      const resultText = isHired ? `（採用）` : "（不採用）";
      const newHistoryLine = `${formattedDate}：${hiredComp}${resultText}`;

      resultDataArray.push({ id: candId, company: hiredComp, result: isHired ? '採用' : '不採用' });

      for (let j = 1; j < mData.length; j++) {
        if (String(mData[j][0]).trim() === candId) {
          const rowIdx = j + 1;
          if (isHired) {
            if (mCol['ステータス']) mSheet.getRange(rowIdx, mCol['ステータス']).setValue('採用');
            if (mCol['採用事業者']) mSheet.getRange(rowIdx, mCol['採用事業者']).setValue(hiredComp);
            if (mCol['内定日']) {
              if (offerDateObj) {
                mSheet.getRange(rowIdx, mCol['内定日']).setValue(offerDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
              } else {
                mSheet.getRange(rowIdx, mCol['内定日']).setValue('');
              }
            }
          } else {
            if (mCol['内定日']) mSheet.getRange(rowIdx, mCol['内定日']).setValue('');
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

    // ▼ 採用者を上に、それ以外を下に並び替える
    resultDataArray.sort((a, b) => {
      if (a.result === '採用' && b.result !== '採用') return -1;
      if (a.result !== '採用' && b.result === '採用') return 1;
      return 0;
    });

    // ▼ 面接結果詳細テキストの構築と書式設定
    let resultDetailsText = "";
    if (companyNames.length <= 1) {
        resultDetailsText = resultDataArray.map(item => {
            const name = candDict[item.id] ? candDict[item.id].name : "";
            const prefix = name ? `${item.id}-${name}` : `${item.id}`;
            return `${prefix}（${item.result}）`;
        }).join('\n');
    } else {
        const grouped = {};
        resultDataArray.forEach(item => {
            if (!grouped[item.company]) grouped[item.company] = [];
            const name = candDict[item.id] ? candDict[item.id].name : "";
            const prefix = name ? `${item.id}-${name}` : `${item.id}`;
            grouped[item.company].push(`${prefix}（${item.result}）`);
        });
        let lines = [];
        for (const [comp, cands] of Object.entries(grouped)) {
            lines.push(`【${comp}】`);
            lines.push(...cands);
        }
        resultDetailsText = lines.join('\n');
    }

    if (colMap['面接結果詳細']) {
      const range = sheet.getRange(targetJobRow, colMap['面接結果詳細']);
      const richText = SpreadsheetApp.newRichTextValue().setText(resultDetailsText);
      const styleBlack = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setBold(false).build();
      const styleBlue = SpreadsheetApp.newTextStyle().setForegroundColor('#1a73e8').setBold(true).build();
      
      if (resultDetailsText.length > 0) {
        richText.setTextStyle(0, resultDetailsText.length, styleBlack);
      }

      const lines = resultDetailsText.split('\n');
      let currentPos = 0;
      lines.forEach(line => {
        if (line.includes('（採用）')) {
          richText.setTextStyle(currentPos, currentPos + line.length, styleBlue);
        }
        currentPos += line.length + 1;
      });
      range.setRichTextValue(richText.build());
    }

    if (colMap['ステータス']) sheet.getRange(targetJobRow, colMap['ステータス']).setValue(hiredData.length > 0 ? '入国準備' : '終了');
    if (colMap['内定日']) {
      if (offerDateObj && hiredData.length > 0) {
        sheet.getRange(targetJobRow, colMap['内定日']).setValue(offerDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
      } else {
        sheet.getRange(targetJobRow, colMap['内定日']).setValue('');
      }
    }

    if (hiredData.length > 0) return `${hiredData.length} 名の面接結果（ステータス：入国準備）、およびマスタの内定日・面接履歴の更新が完了しました。`;
    return `「採用者なし」として案件を終了し、対象候補者全員のマスタ更新が完了しました。`;
  } catch(e) { throw new Error(e.message); }
}

// ---------------------------------------------------------
// 2. 面接結果の修正・更新 (改修後: HIRE_EDIT) 用
// ---------------------------------------------------------

function getJobCandidatesEdit(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");
    
    if (!details.resultDetails || details.resultDetails.trim() === "") {
        throw new Error("この案件はまだ面接結果が登録されていません。\n「面接結果登録」メニューを使用してください。");
    }

    const companies = details.company ? details.company.split(/\r?\n/).filter(c => c.trim()) : [];
    const defaultCompany = companies.length > 0 ? companies[0] : "";
    const ids = details.candidates ? details.candidates.split(/\r?\n/).filter(id => id.trim()) : [];
    const candDict = getCandidateDict(); 

    // 面接結果詳細から現在のステータスを抽出
    const detailLines = details.resultDetails.split(/\r?\n/).filter(l => l.trim() !== "");
    const detailsMap = new Map();
    let currentCompDetail = defaultCompany;
    const knownStatuses = ['採用', '不採用', '内定辞退（候補者都合）', '内定取消（事業者都合）', '面接結果の取消'];

    for (const line of detailLines) {
        if (line.startsWith('【') && line.endsWith('】')) {
            currentCompDetail = line.slice(1, -1).trim();
        } else {
            const idMatch = line.match(/^(SD-\d+)/);
            if (idMatch) {
                const cid = idMatch[1];
                let parsedResult = '不採用';
                for (const status of knownStatuses) {
                    if (line.endsWith(`（${status}）`)) {
                        parsedResult = status;
                        break;
                    }
                }
                detailsMap.set(cid, { result: parsedResult, company: currentCompDetail });
            }
        }
    }

    const candidates = ids.map(id => {
      const cleanId = id.split('-').slice(0, 2).join('-').trim();
      let pastResult = '不採用';
      let pastCompany = defaultCompany;

      if (detailsMap.has(cleanId)) {
          pastResult = detailsMap.get(cleanId).result;
          pastCompany = detailsMap.get(cleanId).company;
      }

      return { 
        id: cleanId, 
        display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId].name})` : id, 
        name: candDict[cleanId] ? candDict[cleanId].name : "",
        pastResult: pastResult,
        pastCompany: pastCompany
      };
    }).filter(c => c.id);

    return { candidates: candidates, companies: companies, offerDate: details.offerDate };
  } catch(e) { throw new Error(e.message); }
}

function updateHire(jobId, resultData, offerDateStr) {
  try {
    const sheet = getMasterSheet('案件管理');
    const mSheet = getMasterSheet('登録者マスタ');
    if (!sheet || !mSheet) throw new Error("シートへのアクセスに失敗しました。");

    const mCol = getMasterColumnMap(mSheet);
    const colMap = getMasterColumnMap(sheet);
    const data = sheet.getDataRange().getValues();
    const mData = mSheet.getDataRange().getValues();
    if (!colMap['案件ID']) throw new Error("案件管理シートに案件IDが見つかりません。");
    
    let companyNamesText = "", rawInterviewDate = "", targetJobRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colMap['案件ID'] - 1]).trim() === String(jobId).trim()) {
        companyNamesText = String(data[i][colMap['事業者名'] - 1]).trim();
        rawInterviewDate = data[i][colMap['面接日'] - 1];
        targetJobRow = i + 1;
        break;
      }
    }
    if (!companyNamesText) throw new Error("案件が見つかりません。");

    let formattedDate = "日付不明";
    if (rawInterviewDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawInterviewDate, "JST", "yyyy/MM/dd");
    } else if (rawInterviewDate) {
      formattedDate = String(rawInterviewDate).replace(/[年月]/g, '/').replace(/日/g, '');
    }

    let offerDateObj = "";
    if (offerDateStr) {
      offerDateObj = new Date(offerDateStr.replace(/-/g, '/'));
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

          if (mCol['内定日']) {
              if (result === '採用') {
                  if (offerDateObj) {
                      mSheet.getRange(rowIdx, mCol['内定日']).setValue(offerDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
                  } else {
                      mSheet.getRange(rowIdx, mCol['内定日']).setValue('');
                  }
              } else {
                  mSheet.getRange(rowIdx, mCol['内定日']).setValue('');
              }
          }

          if (mCol['面接履歴']) {
            const historyCell = mSheet.getRange(rowIdx, mCol['面接履歴']);
            const currentHistory = String(historyCell.getValue() || "").trim();
            let lines = currentHistory.split(/\r?\n/).filter(l => l.trim() !== "");
            
            let existingIdx = lines.findIndex(l => {
               if (formattedDate === "日付不明") return false;
               if (l.startsWith(formattedDate + "：") || l.startsWith(formattedDate.replace('/0', '/').replace(/\/0(\d)$/, '/$1') + "：")) {
                  const match = l.match(/：(.*?)(?:（(.*?)）)?$/);
                  if (match && companyNames.includes(match[1].trim())) {
                     return true;
                  }
               }
               return false;
            });
            
            if (deleteHistory) {
                if (existingIdx !== -1) lines.splice(existingIdx, 1);
            } else if (writeHistory && formattedDate !== "日付不明") {
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

    // ▼ 採用者を上に、それ以外を下に並び替える
    resultData.sort((a, b) => {
      if (a.result === '採用' && b.result !== '採用') return -1;
      if (a.result !== '採用' && b.result === '採用') return 1;
      return 0;
    });

    // ▼ 面接結果詳細列用のテキスト作成と色付け
    let resultDetailsText = "";
    if (companyNames.length <= 1) {
        resultDetailsText = resultData.map(item => {
            const name = candDict[item.id] ? candDict[item.id].name : "";
            const prefix = name ? `${item.id}-${name}` : `${item.id}`;
            return `${prefix}（${item.result}）`;
        }).join('\n');
    } else {
        const grouped = {};
        resultData.forEach(item => {
            const comp = item.company || defaultCompany;
            if (!grouped[comp]) grouped[comp] = [];
            const name = candDict[item.id] ? candDict[item.id].name : "";
            const prefix = name ? `${item.id}-${name}` : `${item.id}`;
            grouped[comp].push(`${prefix}（${item.result}）`);
        });
        let lines = [];
        for (const [comp, cands] of Object.entries(grouped)) {
            lines.push(`【${comp}】`);
            lines.push(...cands);
        }
        resultDetailsText = lines.join('\n');
    }
    
    if (colMap['面接結果詳細']) {
      const range = sheet.getRange(targetJobRow, colMap['面接結果詳細']);
      const richText = SpreadsheetApp.newRichTextValue().setText(resultDetailsText);
      const styleBlack = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setBold(false).build();
      const styleBlue = SpreadsheetApp.newTextStyle().setForegroundColor('#1a73e8').setBold(true).build();

      if (resultDetailsText.length > 0) {
        richText.setTextStyle(0, resultDetailsText.length, styleBlack);
      }

      const lines = resultDetailsText.split('\n');
      let currentPos = 0;
      lines.forEach(line => {
        if (line.includes('（採用）')) {
          richText.setTextStyle(currentPos, currentPos + line.length, styleBlue);
        }
        currentPos += line.length + 1;
      });
      range.setRichTextValue(richText.build());
    }

    if (colMap['ステータス']) sheet.getRange(targetJobRow, colMap['ステータス']).setValue(hiredList.length > 0 ? '入国準備' : '終了');
    if (colMap['内定日']) {
      if (offerDateObj && hiredList.length > 0) {
        sheet.getRange(targetJobRow, colMap['内定日']).setValue(offerDateObj).setNumberFormat('yyyy"年"m"月"d"日"');
      } else {
        sheet.getRange(targetJobRow, colMap['内定日']).setValue('');
      }
    }

    return `面接結果の更新が完了しました。\n（マスタのステータスと履歴・内定日が自動更新されました）`;
  } catch(e) { throw new Error(e.message); }
}

// ---------------------------------------------------------
// 案件削除時のマスタ履歴・ステータスロールバック処理
// ---------------------------------------------------------
function rollbackMasterOnJobDelete(rawInterviewDate, companiesText, resultDetails, candidatesText) {
  try {
    const mSheet = getMasterSheet('登録者マスタ');
    if (!mSheet) return;
    const mCol = getMasterColumnMap(mSheet);
    const mData = mSheet.getDataRange().getValues();

    let formattedDate = "日付不明";
    if (rawInterviewDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawInterviewDate, "JST", "yyyy/MM/dd");
    } else if (rawInterviewDate) {
      formattedDate = String(rawInterviewDate).replace(/[年月]/g, '/').replace(/日/g, '');
    }

    const companies = String(companiesText).split(/\r?\n/).filter(c => c.trim());
    const defaultCompany = companies.length > 0 ? companies[0] : "";

    const targetCands = new Map();

    // 削除される案件に含まれていた候補者と結果をリストアップ
    if (resultDetails && resultDetails.trim() !== "") {
      const detailLines = String(resultDetails).split(/\r?\n/).filter(l => l.trim() !== "");
      let currentCompDetail = defaultCompany;
      const knownStatuses = ['採用', '不採用', '内定辞退（候補者都合）', '内定取消（事業者都合）', '面接結果の取消'];

      for (const line of detailLines) {
        if (line.startsWith('【') && line.endsWith('】')) {
          currentCompDetail = line.slice(1, -1).trim();
        } else {
          const idMatch = line.match(/^(SD-\d+)/);
          if (idMatch) {
            const cid = idMatch[1];
            let parsedResult = '不採用';
            for (const status of knownStatuses) {
                if (line.endsWith(`（${status}）`)) {
                    parsedResult = status;
                    break;
                }
            }
            targetCands.set(cid, { result: parsedResult, company: currentCompDetail });
          }
        }
      }
    } else if (candidatesText && candidatesText.trim() !== "") {
      const ids = String(candidatesText).split(/\r?\n/).filter(id => id.trim());
      ids.forEach(line => {
         const match = line.match(/^(SD-\d+)/);
         if (match) targetCands.set(match[1], { result: '未登録', company: defaultCompany });
      });
    }

    if (targetCands.size === 0) return;

    const idCol = mCol['登録者ID'] ? mCol['登録者ID'] - 1 : -1;
    const statusCol = mCol['ステータス'] ? mCol['ステータス'] - 1 : -1;
    const hiredCompCol = mCol['採用事業者'] ? mCol['採用事業者'] - 1 : -1;
    const offerDateCol = mCol['内定日'] ? mCol['内定日'] - 1 : -1;
    const historyCol = mCol['面接履歴'] ? mCol['面接履歴'] - 1 : -1;

    if (idCol === -1) return;

    for (let j = 1; j < mData.length; j++) {
      const rowIdx = j + 1;
      const cid = String(mData[j][idCol]).trim();
      if (!cid) continue;

      if (targetCands.has(cid)) {
        const candInfo = targetCands.get(cid);

        // ① 当該案件で「採用」になっていた場合のみステータス等をクリア
        if (candInfo.result === '採用') {
          if (statusCol !== -1) mSheet.getRange(rowIdx, statusCol + 1).setValue('未採用');
          if (hiredCompCol !== -1) mSheet.getRange(rowIdx, hiredCompCol + 1).clearContent();
          if (offerDateCol !== -1) mSheet.getRange(rowIdx, offerDateCol + 1).clearContent();
        }

        // ② 面接履歴の中から、対象の案件と同じ日付・企業名を持つ行を削除（ロールバック）
        if (historyCol !== -1 && mData[j][historyCol]) {
          const currentHistory = String(mData[j][historyCol]);
          const lines = currentHistory.split(/\r?\n/).filter(l => l.trim() !== "");
          let newLines = [];
          let isHistoryChanged = false;

          lines.forEach(l => {
            let shouldDelete = false;
            if (formattedDate !== "日付不明") {
              if (l.startsWith(formattedDate + "：") || l.startsWith(formattedDate.replace('/0', '/').replace(/\/0(\d)$/, '/$1') + "：")) {
                const match = l.match(/：(.*?)(?:（(.*?)）)?$/);
                if (match && companies.includes(match[1].trim())) {
                  shouldDelete = true;
                }
              }
            }

            if (shouldDelete) {
              isHistoryChanged = true;
            } else {
              newLines.push(l);
            }
          });

          if (isHistoryChanged) {
            mSheet.getRange(rowIdx, historyCol + 1).setValue(newLines.join('\n'));
          }
        }
      }
    }
  } catch(e) {
    console.error("案件削除に伴うマスタのロールバックに失敗しました: " + e.message);
  }
}