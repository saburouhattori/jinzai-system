// =========================================
// 案件管理の操作（登録・更新・削除・採用）
// =========================================

/**
 * 補助：指定したセル内の複数のURLを、ドライブの「ファイル名」でリンク化する
 */
function convertToSmartChips(range, urlText) {
  if (!urlText) {
    range.clearContent();
    return;
  }
  const urls = String(urlText).split(/\r?\n/).map(u => u.trim()).filter(u => u);
  if (urls.length === 0) {
    range.clearContent();
    return;
  }
  
  const richTextValue = SpreadsheetApp.newRichTextValue();
  let fullText = "";
  let currentPos = 0;
  let linkData = [];
  
  urls.forEach((url, i) => {
    let fileName = url;
    try {
      // URLからドライブのファイルIDを抽出して実際のファイル名を取得する
      let idMatch = url.match(/[-\w]{25,}/);
      if (idMatch) {
        fileName = DriveApp.getFileById(idMatch[0]).getName();
      }
    } catch(e) {
      // 権限がない場合や取得に失敗した場合は代替テキスト
      fileName = "関連ファイル " + (i + 1);
    }
    
    const textPart = fileName;
    const line = (i > 0 ? "\n" : "") + textPart;
    fullText += line;
    
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

function addJob(formData) {
  const sheet = getMasterSheet('案件管理');
  const lastRow = sheet.getLastRow();
  let nextNumber = 1;
  if (lastRow >= 2) {
    const lastId = String(sheet.getRange(lastRow, 1).getValue());
    const match = lastId.match(/\d+/);
    if (match) nextNumber = parseInt(match[0], 10) + 1;
  }
  const nextId = "JOB-" + nextNumber.toString().padStart(4, '0');

  const todayStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");

  const row = [
    nextId,
    '未着手',
    todayStr,
    formData.company,
    formData.skill,
    formData.candidates.join('\n'),
    formData.interviewDate || '',
    '', // 採用者氏名
    '', // 関連ファイル（一旦空で追加）
    formData.memo || ''
  ];
  
  sheet.appendRow(row);
  const newRow = sheet.getLastRow();
  
  // スマートチップ（ファイル名リンク）への変換処理
  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';
  convertToSmartChips(sheet.getRange(newRow, 9), fileUrls);
  
  sheet.getRange(newRow, 3).setNumberFormat('yyyy/MM/dd');

  return `案件登録が完了しました: ${nextId}`;
}

function getJobDetails(jobId) {
  const sheet = getMasterSheet('案件管理');
  const data = sheet.getDataRange().getValues();
  const searchId = String(jobId).trim().toUpperCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === searchId) {
      
      // セルのリンクからURLを復元する
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
      } catch(e) {}
      
      if (!rawUrls) {
        rawUrls = String(data[i][8] || "");
      }

      // 日付データがエラーを起こさないように安全に文字列化する
      let ivDate = data[i][6];
      if (ivDate instanceof Date) {
        ivDate = Utilities.formatDate(ivDate, "JST", "yyyy-MM-dd");
      }

      let rDate = data[i][2];
      if (rDate instanceof Date) {
        rDate = Utilities.formatDate(rDate, "JST", "yyyy-MM-dd");
      }

      return {
        row: i + 1,
        id: data[i][0],
        status: data[i][1],
        date: rDate,
        company: data[i][3],
        skill: data[i][4],
        candidates: data[i][5],
        interviewDate: ivDate,
        hireNames: data[i][7],
        relatedFile: rawUrls,
        memo: data[i][9]
      };
    }
  }
  return null;
}

function updateJob(formData) {
  const sheet = getMasterSheet('案件管理');
  const row = Number(formData.row);
  
  sheet.getRange(row, 2).setValue(formData.status);
  sheet.getRange(row, 4).setValue(formData.company);
  sheet.getRange(row, 5).setValue(formData.skill);
  sheet.getRange(row, 6).setValue(formData.candidates.join('\n'));
  sheet.getRange(row, 7).setValue(formData.interviewDate);
  sheet.getRange(row, 10).setValue(formData.memo);
  
  // スマートチップ（ファイル名リンク）への変換処理
  const fileUrls = Array.isArray(formData.relatedFiles) ? formData.relatedFiles.join('\n') : '';
  convertToSmartChips(sheet.getRange(row, 9), fileUrls);
  
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
    
    // 登録者マスタの更新
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