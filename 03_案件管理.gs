// =========================================
// 案件データの操作（登録・更新・削除・詳細取得）
// =========================================

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
      if (val === "" && targetRow === -1) targetRow = i + 1;
    }

    if (targetRow === -1) {
      targetRow = sheet.getLastRow() + 1;
      sheet.insertRowAfter(sheet.getLastRow());
    } else {
      sheet.insertRowBefore(targetRow);
    }

    const nextId = "JOB-" + (lastIdNum + 1).toString().padStart(4, '0');
    
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    let interviewDate = '';
    if (formData.interviewDate) {
      const parts = formData.interviewDate.split('-');
      if (parts.length === 3) interviewDate = new Date(parts[0], parts[1] - 1, parts[2]);
    }
    
    const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
    let fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];

    // 08_ドライブ連携.gs の共通処理を呼び出し
    fileUrlsArr = handleDriveUploads(nextId, companyName, fileUrlsArr, formData.uploadFiles);
    const fileUrlsText = fileUrlsArr.join('\n');

    const rowData = [
      nextId,                           
      formData.status || '未着手',                      
      today,                            
      companyName,                      
      formData.skill || '',             
      candidatesArr.join('\n'),         
      interviewDate,                    
      '',                               
      fileUrlsText,   
      formData.memo || ''               
    ];
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    
    try {
      if (fileUrlsText) convertToSmartChips(sheet, targetRow, 9, fileUrlsText);
      sheet.getRange(targetRow, 3).setNumberFormat('yyyy"年"m"月"d"日"');
      sheet.getRange(targetRow, 7).setNumberFormat('yyyy"年"m"月"d"日"');
    } catch(ex) { console.warn("装飾処理でエラー: " + ex.message); }

    return `案件登録が完了しました: ${nextId}`;
  } catch(e) { throw new Error("登録に失敗しました: " + e.message); }
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
            const urlArray = [];
            richText.getRuns().forEach(run => {
              const url = run.getLinkUrl();
              if (url) urlArray.push(url);
            });
            rawUrls = urlArray.join('\n');
          }
        } catch(e) {}
        
        if (!rawUrls) rawUrls = String(data[i][8] || "");
        const toIsoDate = (val) => {
          if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy-MM-dd");
          if (typeof val === 'string' && val) return val.replace(/[年月]/g, '-').replace(/日/g, '').replace(/\//g, '-');
          return '';
        };
        return {
          row: i + 1, id: data[i][0], status: data[i][1], date: toIsoDate(data[i][2]),
          company: data[i][3], skill: data[i][4], candidates: String(data[i][5] || ""),
          interviewDate: toIsoDate(data[i][6]), hireNames: data[i][7],
          relatedFile: rawUrls, memo: data[i][9]
        };
      }
    }
    return null;
  } catch(e) { throw new Error(e.message); }
}

/**
 * 案件情報の更新
 */
function updateJob(formData) {
  try {
    const sheet = getMasterSheet('案件管理');
    const row = Number(formData.row);
    if (!row || row < 2) throw new Error("無効な行番号です。");

    const companyName = String(formData.company || "").trim();
    const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
    let fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];

    // 08_ドライブ連携.gs の共通処理を呼び出し
    fileUrlsArr = handleDriveUploads(formData.id, companyName, fileUrlsArr, formData.uploadFiles);
    const fileUrlsText = fileUrlsArr.join('\n');
    
    sheet.getRange(row, 2).setValue(formData.status || '未着手');
    sheet.getRange(row, 4).setValue(companyName);
    sheet.getRange(row, 5).setValue(formData.skill || '');
    sheet.getRange(row, 6).setValue(candidatesArr.join('\n'));
    
    let interviewDate = '';
    if (formData.interviewDate) {
      const parts = formData.interviewDate.split('-');
      if (parts.length === 3) interviewDate = new Date(parts[0], parts[1] - 1, parts[2]);
    }
    
    sheet.getRange(row, 7).setValue(interviewDate).setNumberFormat('yyyy"年"m"月"d"日"');
    sheet.getRange(row, 10).setValue(formData.memo || '');
    
    try { convertToSmartChips(sheet, row, 9, fileUrlsText); } catch(ex) {}
    
    return "案件情報を更新しました。";
  } catch(e) { throw new Error(e.message); }
}

/**
 * 案件の削除
 */
function deleteJobRow(jobId) {
  try {
    const sheet = getMasterSheet('案件管理');
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(jobId).trim()) {
        sheet.deleteRow(i + 1);
        return "案件を削除しました。";
      }
    }
    throw new Error("対象の案件が見つかりませんでした。");
  } catch(e) { throw new Error(e.message); }
}