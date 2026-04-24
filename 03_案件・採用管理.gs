// =========================================
// 案件管理の操作（登録・更新・削除・採用）
// =========================================

/**
 * 補助：指定したセル内の複数のURLを「実際のファイル名」のリンクに変換する（疑似スマートチップ）
 */
function convertToSmartChips(sheet, row, col, urlText) {
  if (!urlText) {
    sheet.getRange(row, col).clearContent();
    return;
  }
  const urls = String(urlText).split(/\r?\n/).map(u => u.trim()).filter(u => u);
  if (urls.length === 0) {
    sheet.getRange(row, col).clearContent();
    return;
  }

  const range = sheet.getRange(row, col);
  const richTextValue = SpreadsheetApp.newRichTextValue();
  let fullText = "";
  let linkData = [];
  let currentPos = 0;

  urls.forEach((url, i) => {
    let fileName = url;
    try {
      let fileId = "";
      const idMatch = url.match(/\/d\/([-\w]{25,})/);
      if (idMatch) {
        fileId = idMatch[1];
      } else {
        const queryMatch = url.match(/id=([-\w]{25,})/);
        if (queryMatch) fileId = queryMatch[1];
      }

      if (fileId) {
        fileName = DriveApp.getFileById(fileId).getName();
      }
    } catch(ex) {
      fileName = "関連ファイル " + (i + 1);
    }
    
    const textPart = "📄 " + fileName;
    fullText += (i > 0 ? "\n" : "") + textPart;
    
    linkData.push({
      url: url,
      start: currentPos + (i > 0 ? 1 : 0),
      end: currentPos + (i > 0 ? 1 : 0) + textPart.length
    });
    currentPos = currentPos + (i > 0 ? 1 : 0) + textPart.length;
  });

  richTextValue.setText(fullText);
  linkData.forEach(ld => {
    richTextValue.setLinkUrl(ld.start, ld.end, ld.url);
  });
  range.setRichTextValue(richTextValue.build());
}

/**
 * 案件登録（新規事業者の自動マスタ登録付き）
 */
function addJob(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('案件管理');
    if (!sheet) throw new Error("「案件管理」シートが見つかりません。");

    const companyName = String(formData.company || "").trim();
    if (companyName) {
      const compSheet = ss.getSheetByName('事業者マスタ');
      if (compSheet) {
        const compData = compSheet.getDataRange().getValues();
        const exists = compData.some(row => String(row[1]).trim() === companyName);
        if (!exists) {
          let lastIdNum = 0;
          for (let i = 1; i < compData.length; i++) {
            let idVal = String(compData[i][0]);
            let match = idVal.match(/\d+/);
            if (match) {
              let num = parseInt(match[0], 10);
              if (num > lastIdNum) lastIdNum = num;
            }
          }
          const nextCompId = "CO-" + (lastIdNum + 1).toString().padStart(4, '0');
          compSheet.appendRow([nextCompId, companyName, "", "", "", "案件登録により自動追加"]);
        }
      }
    }

    // --- 書き込み行の特定ロジック ---
    const dataRange = sheet.getDataRange();
    const aVals = dataRange.getValues().map(r => r[0]); 
    let lastIdNum = 0;
    let targetRow = -1;

    for (let i = 1; i < aVals.length; i++) { 
      let val = String(aVals[i]).trim();
      let match = val.match(/\d+/);
      if (val.startsWith("JOB-") && match) {
        let num = parseInt(match[0], 10);
        if (num > lastIdNum) lastIdNum = num;
      }
      if (val === "" && targetRow === -1) {
        targetRow = i + 1;
      }
    }

    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
      sheet.insertRowAfter(sheet.getLastRow());
    } else {
      sheet.insertRowBefore(targetRow);
    }

    const nextId = "JOB-" + (lastIdNum + 1).toString().padStart(4, '0');
    
    // --- 日付処理 ---
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); 
    
    let interviewDate = '';
    if (formData.interviewDate) {
      const parts = formData.interviewDate.split('-'); 
      if (parts.length === 3) {
        interviewDate = new Date(parts[0], parts[1] - 1, parts[2]);
      }
    }
    
    const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
    const fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];
    const fileUrlsText = fileUrlsArr.join('\n');

    const rowData = [
      nextId,                           
      '未着手',                          
      today,                            
      companyName,                      
      formData.skill || '',             
      candidatesArr.join('\n'),         
      interviewDate,                    
      '',                               
      '',                               
      formData.memo || ''               
    ];

    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    
    try {
      if (fileUrlsText) {
        convertToSmartChips(sheet, targetRow, 9, fileUrlsText);
      }
      sheet.getRange(targetRow, 3).setNumberFormat('yyyy"年"m"月"d"日"');
      sheet.getRange(targetRow, 7).setNumberFormat('yyyy"年"m"月"d"日"');
    } catch(ex) {
      console.warn("装飾処理でエラー: " + ex.message);
    }

    return `案件登録が完了しました: ${nextId}`;

  } catch(e) {
    throw new Error("登録に失敗しました: " + e.message);
  }
}

/**
 * 案件詳細の取得
 */
function getJobDetails(jobId) {
  try {
    const sheet = getMasterSheet('案件管理');
    if (!sheet) return null;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const data = sheet.getRange(1, 1, lastRow, 10).getValues();
    const searchId = String(jobId).trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === searchId) {
        let rawUrls = "";
        try {
          const richText = sheet.getRange(i + 1, 9).getRichTextValue();
          if (richText) {
            const runs = richText.getRuns();
            const urlArray = [];
            runs.forEach(run => {
              const url = run.getLinkUrl();
              if (url) urlArray.push(url);
            });
            rawUrls = urlArray.join('\n');
          }
        } catch(e) {
          console.error("URL抽出エラー: " + e);
        }
        
        if (!rawUrls) rawUrls = String(data[i][8] || ""); 

        const toIsoDate = (val) => {
          if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy-MM-dd");
          if (typeof val === 'string' && val) {
            return val.replace(/[年月]/g, '-').replace(/日/g, '').replace(/\//g, '-');
          }
          return '';
        };

        return {
          row: i + 1,
          id: data[i][0],
          status: data[i][1],
          date: toIsoDate(data[i][2]),
          company: data[i][3],
          skill: data[i][4],
          candidates: String(data[i][5] || ""),
          interviewDate: toIsoDate(data[i][6]),
          hireNames: data[i][7],
          relatedFile: rawUrls, 
          memo: data[i][9]
        };
      }
    }
    return null;
  } catch(e) {
    throw new Error(e.message);
  }
}

/**
 * 案件情報の更新
 */
function updateJob(formData) {
  try {
    const sheet = getMasterSheet('案件管理');
    const row = Number(formData.row);
    if (!row || row < 2) throw new Error("無効な行番号です。");

    const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
    const fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];
    const fileUrlsText = fileUrlsArr.join('\n');
    
    sheet.getRange(row, 2).setValue(formData.status || '未着手');
    sheet.getRange(row, 4).setValue(formData.company || '');
    sheet.getRange(row, 5).setValue(formData.skill || '');
    sheet.getRange(row, 6).setValue(candidatesArr.join('\n'));
    
    let interviewDate = '';
    if (formData.interviewDate) {
      const parts = formData.interviewDate.split('-');
      if (parts.length === 3) {
        interviewDate = new Date(parts[0], parts[1] - 1, parts[2]);
      }
    }
    
    const cell = sheet.getRange(row, 7);
    cell.setValue(interviewDate).setNumberFormat('yyyy"年"m"月"d"日"');
    
    sheet.getRange(row, 10).setValue(formData.memo || '');
    
    try {
      convertToSmartChips(sheet, row, 9, fileUrlsText);
    } catch(ex) {
      console.warn("スマートチップ変換エラー: " + ex.message);
    }
    
    return "案件情報を更新しました。";
  } catch(e) {
    throw new Error(e.message);
  }
}

/**
 * 案件の削除
 */
function deleteJobRow(jobId) {
  try {
    const sheet = getMasterSheet('案件管理');
    if (!sheet) throw new Error("シートが見つかりません。");
    
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(jobId).trim()) {
        sheet.deleteRow(i + 1);
        return "案件を削除しました。";
      }
    }
    throw new Error("対象の案件が見つかりませんでした。");
  } catch(e) {
    throw new Error(e.message);
  }
}

/**
 * 案件に紐づく候補者情報のリスト取得
 */
function getJobCandidates(jobId) {
  try {
    const details = getJobDetails(jobId);
    if (!details || !details.candidates) return [];

    const candDict = getCandidateDict(); 
    const ids = details.candidates.split(/\r?\n/).filter(id => id.trim());

    return ids.map(id => {
      const cleanId = id.split('-').slice(0, 2).join('-').trim();
      return { 
        id: cleanId, 
        display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId]})` : id 
      };
    }).filter(c => c.id);
  } catch(e) {
    throw new Error(e.message);
  }
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
    
    let companyName = "";
    let rawInterviewDate = "";
    let allCandidatesRaw = "";
    let targetJobRow = -1;

    // 1. 案件管理シートから対象の案件情報を取得
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(jobId).trim()) {
        companyName = String(data[i][3]).trim();
        allCandidatesRaw = String(data[i][5]); // 候補者名（F列）
        rawInterviewDate = data[i][6];         // 面接日（G列）
        targetJobRow = i + 1;
        break;
      }
    }
    if (!companyName) throw new Error("案件が見つかりません。");

    // 面接日のフォーマット処理（yyyy/MM/dd形式へ）
    let formattedDate = "日付不明";
    if (rawInterviewDate instanceof Date) {
      formattedDate = Utilities.formatDate(rawInterviewDate, "JST", "yyyy/MM/dd");
    } else if (rawInterviewDate) {
      formattedDate = String(rawInterviewDate).replace(/[年月]/g, '/').replace(/日/g, '');
    }

    const candDict = getCandidateDict();
    
    // --- ▽ ここから面接履歴の追記ロジック ▽ ---
    
    // 案件に紐づく「全候補者」のIDリストを抽出
    const allCandidateIds = allCandidatesRaw.split(/\r?\n/)
                                            .map(line => line.split('-').slice(0, 2).join('-').trim())
                                            .filter(id => id !== "");
    
    // 採用者リスト（hiredIds）をSet化して判定を高速化
    const hiredIdSet = new Set(hiredIds.map(id => String(id).trim()));

    // 全候補者に対してループ処理
    allCandidateIds.forEach(candId => {
      // 採用か不採用かを判定
      const isHired = hiredIdSet.has(candId);
      const resultText = isHired ? "（採用）" : "（不採用）";
      
      // 追記する履歴文字列の生成
      const newHistoryLine = `${formattedDate}：${companyName}${resultText}`;

      // 登録者マスタを検索して履歴を追記
      for (let j = 1; j < mData.length; j++) {
        if (String(mData[j][0]).trim() === candId) {
          const rowIdx = j + 1;
          
          // 採用者の場合はステータスと採用事業者も更新
          if (isHired) {
            if (mCol['ステータス']) mSheet.getRange(rowIdx, mCol['ステータス']).setValue('採用');
            if (mCol['採用事業者']) mSheet.getRange(rowIdx, mCol['採用事業者']).setValue(companyName);
          }
          
          // 面接履歴の更新（改行して追記）
          if (mCol['面接履歴']) {
            const historyCell = mSheet.getRange(rowIdx, mCol['面接履歴']);
            const currentHistory = String(historyCell.getValue() || "").trim();
            
            // すでに履歴が入っている場合は改行を挟んで追加、空の場合はそのままセット
            if (currentHistory) {
              historyCell.setValue(currentHistory + "\n" + newHistoryLine);
            } else {
              historyCell.setValue(newHistoryLine);
            }
          }
          break; // この候補者の処理を終えて次の候補者へ
        }
      }
    });
    // --- △ ここまで △ ---

    // 案件管理シート側の更新（採用者名の書き込みとステータス変更）
    let hiredNamesText = "採用者なし";
    if (hiredIds.length > 0) {
      const hiredNames = hiredIds.map(id => {
        const name = candDict[id] || "";
        return name ? `${id}-${name}` : id;
      });
      hiredNamesText = hiredNames.join('\n');
    }

    sheet.getRange(targetJobRow, 8).setValue(hiredNamesText);
    sheet.getRange(targetJobRow, 2).setValue('終了');

    // ★修正：完了メッセージを分岐
    if (hiredIds.length > 0) {
      return `${hiredIds.length} 名の面接結果、および対象候補者全員の「面接履歴」への追記が完了しました。`;
    } else {
      return `「採用者なし」として案件を終了し、対象候補者全員の「面接履歴」への追記が完了しました。`;
    }
  } catch(e) {
    throw new Error(e.message);
  }
}