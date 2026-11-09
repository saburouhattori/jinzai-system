// =========================================
// 面接結果の新規登録処理 (HIRE)
// =========================================

function getJobCandidates(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details) throw new Error("該当する案件が見つかりません。");
    if (!details.interviewDate) throw new Error("面接日が設定されていません。\n先に「案件更新/削除」から面接日を登録してください。");
    
    // ▼ バックエンドの強固なブロック（既に結果がある場合は弾く）
    const isRegistered = (details.resultDetails && details.resultDetails.trim() !== "") || details.status === "終了" || details.status === "入国準備";
    if (isRegistered) {
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
          const currentStatus = mCol['ステータス'] ? String(mData[j][mCol['ステータス'] - 1] || '').trim() : '';

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
            // 他案件で既に「採用」でない場合のみ内定日をクリア
            if (currentStatus !== '採用' && mCol['内定日']) {
              mSheet.getRange(rowIdx, mCol['内定日']).setValue('');
            }
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

    SpreadsheetApp.flush();

    // ▼ 内定（採用）となった候補者の基本情報を外部シート「COE info」へ自動転記
    if (hiredData.length > 0) {
      const hiredIds = hiredData.map(item => String(item.id).trim().toUpperCase());
      syncCandidatesToCoeInfo(hiredIds);
      return `${hiredData.length} 名の面接結果（ステータス：入国準備）、マスタ更新、および外部シート「COE info」への転記が完了しました。`;
    }
    return `「採用者なし」として案件を終了し、対象候補者全員のマスタ更新が完了しました。`;
  } catch(e) { throw new Error(e.message); }
}