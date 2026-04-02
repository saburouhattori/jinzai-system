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

  const lastRow = sheet.getLastRow();
  let lastIdNum = 0;
  if (lastRow > 1) {
    const idVals = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (let i = 1; i < idVals.length; i++) {
      let idVal = String(idVals[i][0]).trim();
      let match = idVal.match(/\d+/);
      if (match) {
        let num = parseInt(match[0], 10);
        if (num > lastIdNum) lastIdNum = num;
      }
    }
  }

  const targetRow = lastRow + 1;
  const nextId = "JOB-" + (lastIdNum + 1).toString().padStart(4, '0');
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  
  const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
  const fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];
  const fileUrlsText = fileUrlsArr.join('\n');

  const rowData = [
    nextId,                           
    '未着手',                          
    todayStr,                         
    companyName,                      
    formData.skill || '',             
    candidatesArr.join('\n'),         
    formData.interviewDate || '',     
    '',                               
    '',                               
    '',                               
    formData.memo || ''               
  ];

  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  
  if (fileUrlsText) {
    convertToSmartChips(sheet, targetRow, 10, fileUrlsText);
  }
  
  sheet.getRange(targetRow, 3).setNumberFormat('yyyy/MM/dd');

  return `案件登録が完了しました: ${nextId}`;
}

/**
 * 案件詳細の取得
 */
function getJobDetails(jobId) {
  const sheet = getMasterSheet('案件管理');
  if (!sheet) return null;
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(1, 1, lastRow, 11).getValues();
  const searchId = String(jobId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === searchId) {
      let rawUrls = "";
      try {
        const richText = sheet.getRange(i + 1, 10).getRichTextValue();
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
      
      if (!rawUrls) rawUrls = String(data[i][9] || ""); 

      let ivDate = data[i][6];
      if (ivDate instanceof Date) ivDate = Utilities.formatDate(ivDate, "JST", "yyyy-MM-dd");
      let rDate = data[i][2];
      if (rDate instanceof Date) rDate = Utilities.formatDate(rDate, "JST", "yyyy-MM-dd");

      return {
        row: i + 1,
        id: data[i][0],
        status: data[i][1],
        date: rDate,
        company: data[i][3],
        skill: data[i][4],
        candidates: String(data[i][5] || ""),
        interviewDate: ivDate,
        hireNames: data[i][7],
        relatedFile: rawUrls, 
        memo: data[i][10]
      };
    }
  }
  return null;
}

/**
 * 案件情報の更新
 */
function updateJob(formData) {
  const sheet = getMasterSheet('案件管理');
  const row = Number(formData.row);
  if (!row || row < 2) return "エラー：無効な行番号です。";

  const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
  const fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];
  const fileUrlsText = fileUrlsArr.join('\n');
  
  sheet.getRange(row, 2).setValue(formData.status || '未着手');
  sheet.getRange(row, 4).setValue(formData.company || '');
  sheet.getRange(row, 5).setValue(formData.skill || '');
  sheet.getRange(row, 6).setValue(candidatesArr.join('\n'));
  sheet.getRange(row, 7).setValue(formData.interviewDate || '');
  sheet.getRange(row, 11).setValue(formData.memo || '');
  
  convertToSmartChips(sheet, row, 10, fileUrlsText);
  
  return "案件情報を更新しました。";
}

/**
 * 案件の削除
 */
function deleteJobRow(jobId) {
  const sheet = getMasterSheet('案件管理');
  if (!sheet) return "エラー：シートが見つかりません。";
  
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === String(jobId).trim()) {
      sheet.deleteRow(i + 1);
      return "案件を削除しました。";
    }
  }
  return "エラー：対象の案件が見つかりませんでした。";
}

/**
 * 案件に紐づく候補者情報のリスト取得
 */
function getJobCandidates(jobId) {
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
}

/**
 * 採用登録
 */
function registerHire(jobId, hiredIds) {
  const sheet = getMasterSheet('案件管理');
  const mSheet = getMasterSheet('登録者マスタ');
  if (!sheet || !mSheet) return "エラー：シートへのアクセスに失敗しました。";

  const mCol = getMasterColumnMap(mSheet);
  const data = sheet.getDataRange().getValues();
  const mData = mSheet.getDataRange().getValues();
  
  let companyName = "";
  let targetJobRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(jobId).trim()) {
      companyName = data[i][3];
      targetJobRow = i + 1;
      break;
    }
  }
  if (!companyName) return "エラー：案件が見つかりません。";

  const hiredNames = [];
  const candDict = getCandidateDict();

  hiredIds.forEach(id => {
    const name = candDict[id] || "";
    // ★ F列と同様に「ID-名前」の形式に変更
    const displayVal = name ? `${id}-${name}` : id;
    hiredNames.push(displayVal);
    
    for (let j = 1; j < mData.length; j++) {
      if (String(mData[j][0]).trim() === String(id).trim()) {
        if (mCol['ステータス']) mSheet.getRange(j + 1, mCol['ステータス']).setValue('採用');
        if (mCol['採用事業者']) mSheet.getRange(j + 1, mCol['採用事業者']).setValue(companyName);
        break;
      }
    }
  });

  // ★ F列と同様に改行(\n)区切りで書き込むよう変更
  sheet.getRange(targetJobRow, 8).setValue(hiredNames.join('\n'));
  sheet.getRange(targetJobRow, 2).setValue('終了');

  return `${hiredIds.length} 名の採用登録を完了しました。案件ステータスを「終了」にしました。`;
}