// =========================================
// 案件データの操作（登録・更新・削除・詳細取得）
// =========================================

function addJob(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('案件管理');
    if (!sheet) throw new Error("「案件管理」シートが見つかりません。");

    const colMap = getMasterColumnMap(sheet);
    if (!colMap['案件ID']) throw new Error("「案件管理」に案件IDの列が見つかりません。");

    const companiesArr = Array.isArray(formData.companies) ? formData.companies.map(c => String(c).trim()).filter(c => c) : [];

    const dataRange = sheet.getDataRange();
    const aVals = dataRange.getValues().map(r => r[colMap['案件ID'] - 1]); 
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

    const mainCompany = companiesArr.length > 0 ? companiesArr[0] : "";
    fileUrlsArr = handleDriveUploads(nextId, mainCompany, fileUrlsArr, formData.uploadFiles);
    const fileUrlsText = fileUrlsArr.join('\n');

    const safeMaxCol = Math.max(sheet.getLastColumn(), ...Object.values(colMap));
    const rowValues = new Array(safeMaxCol).fill("");

    const mapping = {
      '案件ID': nextId,
      'ステータス': formData.status || '未着手',
      '案件登録日': today,
      '事業者名': companiesArr.join('\n'),
      '技能分野': formData.skill || '',
      '候補者名': candidatesArr.join('\n'),
      '面接日': interviewDate,
      '内定日': '', 
      '面接結果詳細': '', 
      '関連フォルダ・ファイル': fileUrlsText,
      '備考・メモ': formData.memo || ''
    };

    for (let header in mapping) {
      if (colMap[header]) rowValues[colMap[header] - 1] = mapping[header];
    }

    sheet.getRange(targetRow, 1, 1, safeMaxCol).setValues([rowValues]);
    
    try {
      if (fileUrlsText && colMap['関連フォルダ・ファイル']) {
        convertToSmartChips(sheet, targetRow, colMap['関連フォルダ・ファイル'], fileUrlsText);
      }
      if (colMap['案件登録日']) sheet.getRange(targetRow, colMap['案件登録日']).setNumberFormat('yyyy"年"m"月"d"日"');
      if (colMap['面接日']) sheet.getRange(targetRow, colMap['面接日']).setNumberFormat('yyyy"年"m"月"d"日"');
      if (colMap['内定日']) sheet.getRange(targetRow, colMap['内定日']).setNumberFormat('yyyy"年"m"月"d"日"');
    } catch(ex) {}

    let resultMsg = `案件登録が完了しました: ${nextId}`;
    if (formData.uploadFiles && formData.uploadFiles.length > 0) {
      resultMsg += "\n（専用フォルダを作成・特定し、ファイルを保存しました）";
    }
    return resultMsg;

  } catch(e) { throw new Error("登録に失敗しました: " + e.message); }
}

function getJobDetails(jobId) {
  try {
    const sheet = getMasterSheet('案件管理');
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const data = sheet.getDataRange().getValues();
    const colMap = getMasterColumnMap(sheet);
    if (!colMap['案件ID']) return null;

    const searchId = String(jobId).replace(/\s/g, '').toUpperCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colMap['案件ID'] - 1]).replace(/\s/g, '').toUpperCase() === searchId) {
        let rawUrls = "";
        try {
          if (colMap['関連フォルダ・ファイル']) {
            const richText = sheet.getRange(i + 1, colMap['関連フォルダ・ファイル']).getRichTextValue();
            if (richText) {
              const urlArray = [];
              richText.getRuns().forEach(run => {
                const url = run.getLinkUrl();
                if (url) urlArray.push(url);
              });
              rawUrls = urlArray.join('\n');
            }
          }
        } catch(e) {}
        
        if (!rawUrls && colMap['関連フォルダ・ファイル']) {
           rawUrls = String(data[i][colMap['関連フォルダ・ファイル'] - 1] || "");
        }

        const toIsoDate = (val) => {
          if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy-MM-dd");
          if (typeof val === 'string' && val) return val.replace(/[年月]/g, '-').replace(/日/g, '').replace(/\//g, '-');
          return '';
        };

        const getVal = (header) => colMap[header] ? data[i][colMap[header] - 1] : "";

        return {
          row: i + 1, 
          id: getVal('案件ID'), 
          status: getVal('ステータス'), 
          date: toIsoDate(getVal('案件登録日')),
          company: getVal('事業者名'), 
          skill: getVal('技能分野'), 
          candidates: String(getVal('候補者名') || ""),
          interviewDate: toIsoDate(getVal('面接日')), 
          offerDate: toIsoDate(getVal('内定日')), 
          resultDetails: String(getVal('面接結果詳細') || ""), 
          relatedFile: rawUrls, 
          memo: getVal('備考・メモ')
        };
      }
    }
    return null;
  } catch(e) { throw new Error(e.message); }
}

function updateJob(formData) {
  try {
    const sheet = getMasterSheet('案件管理');
    const colMap = getMasterColumnMap(sheet);
    const row = Number(formData.row);
    if (!row || row < 2) throw new Error("無効な行番号です。");

    const companiesArr = Array.isArray(formData.companies) ? formData.companies.map(c => String(c).trim()).filter(c => c) : [];
    const candidatesArr = Array.isArray(formData.candidates) ? formData.candidates : [];
    let fileUrlsArr = Array.isArray(formData.relatedFiles) ? formData.relatedFiles : [];
    
    const mainCompany = companiesArr.length > 0 ? companiesArr[0] : "";
    fileUrlsArr = handleDriveUploads(formData.id, mainCompany, fileUrlsArr, formData.uploadFiles);
    const fileUrlsText = fileUrlsArr.join('\n');
    
    let interviewDate = '';
    if (formData.interviewDate) {
      const parts = formData.interviewDate.split('-');
      if (parts.length === 3) interviewDate = new Date(parts[0], parts[1] - 1, parts[2]);
    }

    const safeMaxCol = Math.max(sheet.getLastColumn(), ...Object.values(colMap));
    const currentRowRange = sheet.getRange(row, 1, 1, safeMaxCol);
    const currentRowData = currentRowRange.getValues()[0];

    const mapping = {
      'ステータス': formData.status || '未着手',
      '事業者名': companiesArr.join('\n'),
      '技能分野': formData.skill || '',
      '候補者名': candidatesArr.join('\n'),
      '面接日': interviewDate,
      // '内定日' および '面接結果詳細' はここでは更新対象から除外（既存の値を維持）
      '関連フォルダ・ファイル': fileUrlsText,
      '備考・メモ': formData.memo || ''
    };

    for (let header in mapping) {
      if (colMap[header]) currentRowData[colMap[header] - 1] = mapping[header];
    }
    
    sheet.getRange(row, 1, 1, safeMaxCol).setValues([currentRowData]);

    try {
      if (colMap['面接日']) sheet.getRange(row, colMap['面接日']).setNumberFormat('yyyy"年"m"月"d"日"');
      if (colMap['関連フォルダ・ファイル']) convertToSmartChips(sheet, row, colMap['関連フォルダ・ファイル'], fileUrlsText);
    } catch(ex) {}
    
    let resultMsg = "案件情報を更新しました。";
    if (formData.uploadFiles && formData.uploadFiles.length > 0) {
      resultMsg += "\n（専用フォルダを特定・作成し、ファイルを保存しました）";
    }
    return resultMsg;

  } catch(e) { throw new Error(e.message); }
}

function deleteJobRow(jobId) {
  try {
    const sheet = getMasterSheet('案件管理');
    const data = sheet.getDataRange().getValues();
    const colMap = getMasterColumnMap(sheet);
    if (!colMap['案件ID']) throw new Error("「案件管理」に案件IDの列が見つかりません。");

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][colMap['案件ID'] - 1]).trim() === String(jobId).trim()) {
        
        // ▼ 案件行を削除する前に、マスタの面接履歴・ステータスをロールバックする
        const rawInterviewDate = colMap['面接日'] ? data[i][colMap['面接日'] - 1] : "";
        const companiesText = colMap['事業者名'] ? String(data[i][colMap['事業者名'] - 1]) : "";
        const resultDetails = colMap['面接結果詳細'] ? String(data[i][colMap['面接結果詳細'] - 1]) : "";
        const candidatesText = colMap['候補者名'] ? String(data[i][colMap['候補者名'] - 1]) : "";
        
        rollbackMasterOnJobDelete(rawInterviewDate, companiesText, resultDetails, candidatesText);

        sheet.deleteRow(i + 1);
        return "案件を削除し、関連するマスタの面接履歴とステータスをロールバック（取消）しました。";
      }
    }
    throw new Error("対象の案件が見つかりませんでした。");
  } catch(e) { throw new Error(e.message); }
}

// ---------------------------------------------------------
// マスタ連動：候補者の名前変更時に案件管理の表記を自動置換する
// ---------------------------------------------------------
function updateJobCandidateName(adminId, newName) {
  try {
    const sheet = getMasterSheet('案件管理');
    if (!sheet) return;
    const colMap = getMasterColumnMap(sheet);
    const data = sheet.getDataRange().getValues();
    
    const candsCol = colMap['候補者名'] ? colMap['候補者名'] - 1 : -1;
    const detailsCol = colMap['面接結果詳細'] ? colMap['面接結果詳細'] - 1 : -1;

    if (candsCol === -1 && detailsCol === -1) return;

    // SD-XXXX-(旧名前) または SD-XXXX を SD-XXXX-(新名前) に置換する正規表現（カッコ等の直前まで）
    const replaceRegex = new RegExp(`(${adminId})(?:-[^\\n\\r（]*)?`, 'g');

    for (let i = 1; i < data.length; i++) {
      let newCands = "", newDetails = "";

      if (candsCol !== -1 && data[i][candsCol]) {
        const oldVal = String(data[i][candsCol]);
        newCands = oldVal.replace(replaceRegex, `$1-${newName}`);
        if (oldVal !== newCands) {
          sheet.getRange(i + 1, candsCol + 1).setValue(newCands);
        }
      }

      if (detailsCol !== -1 && data[i][detailsCol]) {
        const oldVal = String(data[i][detailsCol]);
        newDetails = oldVal.replace(replaceRegex, `$1-${newName}`);
        if (oldVal !== newDetails) {
          const range = sheet.getRange(i + 1, detailsCol + 1);
          const lines = newDetails.split('\n');
          const richText = SpreadsheetApp.newRichTextValue().setText(newDetails);
          const styleBlack = SpreadsheetApp.newTextStyle().setForegroundColor('#000000').setBold(false).build();
          const styleBlue = SpreadsheetApp.newTextStyle().setForegroundColor('#1a73e8').setBold(true).build();
          
          if (newDetails.length > 0) richText.setTextStyle(0, newDetails.length, styleBlack);

          let currentPos = 0;
          lines.forEach(line => {
            if (line.includes('（採用）')) {
              richText.setTextStyle(currentPos, currentPos + line.length, styleBlue);
            }
            currentPos += line.length + 1;
          });
          range.setRichTextValue(richText.build());
        }
      }
    }
  } catch(e) {
    console.error("候補者名の案件管理への連動更新に失敗しました: " + e.message);
  }
}