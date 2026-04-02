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
      let idMatch = url.match(/\/d\/([-\w]{25,})/);
      let fileId = idMatch ? idMatch[1] : (url.match(/id=([-\w]{25,})/) ? url.match(/id=([-\w]{25,})/)[1] : null);
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

  // --- ★追加機能：新規事業者の自動マスタ登録 ---
  const companyName = String(formData.company || "").trim();
  if (companyName) {
    const compSheet = ss.getSheetByName('事業者マスタ');
    if (compSheet) {
      const compData = compSheet.getDataRange().getValues();
      // 2列目（インデックス1）が事業者名と想定
      const exists = compData.some(row => String(row[1]).trim() === companyName);
      if (!exists) {
        // マスタに存在しない場合、新しく登録
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
        // 事業者ID, 名称, ヨミ, 所在地, HP, 備考 の順で空行を作成
        compSheet.appendRow([nextCompId, companyName, "", "", "", "案件登録により自動追加"]);
      }
    }
  }
  // ------------------------------------------

  const idVals = sheet.getRange("A1:A" + sheet.getMaxRows()).getValues();
  let lastDataRow = 1;
  let lastIdNum = 0;

  for (let i = 1; i < idVals.length; i++) {
    let idVal = String(idVals[i][0]).trim();
    if (idVal !== "") {
      lastDataRow = i + 1;
      let match = idVal.match(/\d+/);
      if (match) {
        let num = parseInt(match[0], 10);
        if (num > lastIdNum) lastIdNum = num;
      }
    }
  }

  sheet.insertRowAfter(lastDataRow);
  const targetRow = lastDataRow + 1;

  const nextId = "JOB-" + (lastIdNum + 1).toString().padStart(4, '0');
  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';

  // 11列分（K列：備考・メモまで）の配列を正確に構築します
  const rowData = [
    nextId,                           // 1: A列 案件ID
    '未着手',                          // 2: B列 ステータス
    todayStr,                         // 3: C列 案件登録日
    companyName,                      // 4: D列 事業者名
    formData.skill,                   // 5: E列 技能分野
    formData.candidates.join('\n'),   // 6: F列 候補者名
    formData.interviewDate || '',     // 7: G列 面接日
    '',                               // 8: H列 採用者氏名
    '',                               // 9: I列 求人票
    '',                               // 10: J列 関連ファイル（この後スマートチップで書き込み）
    formData.memo || ''               // 11: K列 備考・メモ
  ];

  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  // 第3引数を 9 から 10（J列）へ変更
  convertToSmartChips(sheet, targetRow, 10, fileUrls);
  sheet.getRange(targetRow, 3).setNumberFormat('yyyy/MM/dd');

  return `案件登録が完了しました: ${nextId}`;
}

function getJobDetails(jobId) {
  const sheet = getMasterSheet('案件管理');
  const data = sheet.getDataRange().getValues();
  const searchId = String(jobId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === searchId) {
      let rawUrls = "";
      try {
        // J列（10列目）から取得するように変更
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
      } catch(e) {}
      // data[i][9] が J列（関連ファイル）
      if (!rawUrls) rawUrls = String(data[i][9] || ""); 

      let ivDate = data[i][6];
      if (ivDate instanceof Date) ivDate = Utilities.formatDate(ivDate, "JST", "yyyy-MM-dd");
      let rDate = data[i][2];
      if (rDate instanceof Date) rDate = Utilities.formatDate(rDate, "JST", "yyyy-MM-dd");

      return {
        row: i + 1, id: data[i][0], status: data[i][1], date: rDate, company: data[i][3],
        skill: data[i][4], candidates: data[i][5], interviewDate: ivDate,
        hireNames: data[i][7], relatedFile: rawUrls, 
        memo: data[i][10] // K列（インデックス10）に変更
      };
    }
  }
  return null;
}

function updateJob(formData) {
  const sheet = getMasterSheet('案件管理');
  const row = Number(formData.row);

  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';
  
  sheet.getRange(row, 2).setValue(formData.status);
  sheet.getRange(row, 4).setValue(formData.company);
  sheet.getRange(row, 5).setValue(formData.skill);
  sheet.getRange(row, 6).setValue(formData.candidates.join('\n'));
  sheet.getRange(row, 7).setValue(formData.interviewDate);
  // メモの書き込み先を 11（K列）に変更
  sheet.getRange(row, 11).setValue(formData.memo);
  // 関連ファイルの書き込み先を 10（J列）に変更
  convertToSmartChips(sheet, row, 10, fileUrls);
  
  return "案件情報を更新しました。";
}

function deleteJobRow(jobId) {
  const sheet = getMasterSheet('案件管理');
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === jobId) {
      sheet.deleteRow(i + 1);
      return "案件を削除しました。";
    }
  }
  return "エラー：案件が見つかりませんでした。";
}

function getJobCandidates(jobId) {
  const details = getJobDetails(jobId);
  if (!details || !details.candidates) return [];

  const candDict = getCandidateDict();
  const ids = details.candidates.split(/\r?\n/);

  return ids.map(id => {
    const cleanId = id.split('-').slice(0,2).join('-').trim();
    return { id: cleanId, display: candDict[cleanId] ? `${cleanId} (${candDict[cleanId]})` : cleanId };
  }).filter(c => c.id);
}

function registerHire(jobId, hiredIds) {
  const sheet = getMasterSheet('案件管理');
  const mSheet = getMasterSheet('登録者マスタ');
  const mCol = getMasterColumnMap(mSheet);

  const data = sheet.getDataRange().getValues();
  const mData = mSheet.getDataRange().getValues();
  
  let companyName = "";
  let targetJobRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === jobId) {
      companyName = data[i][3];
      targetJobRow = i + 1;
      break;
    }
  }
  if (!companyName) return "エラー：案件が見つかりません。";

  const hiredNames = [];
  const candDict = getCandidateDict();

  hiredIds.forEach(id => {
    const name = candDict[id] || id;
    hiredNames.push(name);
    for (let j = 1; j < mData.length; j++) {
      if (String(mData[j][0]) === id) {
        if (mCol['ステータス']) mSheet.getRange(j + 1, mCol['ステータス']).setValue('採用');
        if (mCol['採用事業者']) mSheet.getRange(j + 1, mCol['採用事業者']).setValue(companyName);
        break;
      }
    }
  });

  sheet.getRange(targetJobRow, 8).setValue(hiredNames.join(', '));
  sheet.getRange(targetJobRow, 2).setValue('終了');

  return `${hiredIds.length} 名の採用登録を完了しました。案件ステータスを「終了」にしました。`;
}