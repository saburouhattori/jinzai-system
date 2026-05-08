// =========================================
// ドライブ連携・ファイル操作関連
// =========================================

/**
 * 関連ファイル（ファイル・フォルダ）の検索
 */
function searchDriveFiles(fileNameQuery) {
  try {
    const results = [];
    let query = 'trashed = false';
    if (fileNameQuery) query += ' and title contains "' + fileNameQuery + '"';
    
    // フォルダの検索
    let count = 0;
    try {
      const folderIter = DriveApp.searchFolders(query);
      while (folderIter.hasNext() && count < 10) {
        const folder = folderIter.next();
        results.push({ name: "📁 " + folder.getName(), url: folder.getUrl(), type: "folder" });
        count++;
      }
    } catch(e) {}
    
    // ファイルの検索
    count = 0;
    try {
      const fileIter = DriveApp.searchFiles(query);
      while (fileIter.hasNext() && count < 15) {
        const file = fileIter.next();
        results.push({ name: "📄 " + file.getName(), url: file.getUrl(), type: file.getMimeType() });
        count++;
      }
    } catch(e) {}
    
    return results;
  } catch (e) {
    return [];
  }
}

/**
 * フォルダ自動作成およびファイルアップロード処理の共通化
 */
function handleDriveUploads(jobId, companyName, existingUrls, uploadFiles) {
  if (!uploadFiles || uploadFiles.length === 0) return existingUrls;

  const parentFolderId = '1UuwRUPmldGgBR6dVUjldMI0vwZOjab0t';
  let targetFolder = null;

  // 1. 既存URL内にフォルダ（/folders/）があれば特定
  const existingFolderUrl = existingUrls.find(u => u.includes('/folders/'));
  if (existingFolderUrl) {
    const folderIdMatch = existingFolderUrl.match(/\/folders\/([-\w]{25,})/);
    if (folderIdMatch) {
      try { targetFolder = DriveApp.getFolderById(folderIdMatch[1]); } catch(e) {}
    }
  }

  // 2. フォルダがなければ新規作成
  if (!targetFolder) {
    try {
      const parentFolder = DriveApp.getFolderById(parentFolderId);
      const safeJobId = jobId || 'JOB-UNKNOWN';
      const safeComp = companyName || '名称未設定';
      targetFolder = parentFolder.createFolder(`${safeJobId}_${safeComp}`);
      existingUrls.push(targetFolder.getUrl());
    } catch (e) {
      console.warn("親フォルダの取得または新規作成エラー: " + e.message);
      return existingUrls;
    }
  }

  // 3. ファイルのアップロード・保存
  try {
    uploadFiles.forEach(f => {
      const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType || 'application/octet-stream', f.name);
      targetFolder.createFile(blob);
    });
  } catch (e) {
    console.warn("ファイル保存エラー: " + e.message);
  }

  return existingUrls;
}

/**
 * 補助：指定したセル内の複数のURLを「実際のファイル/フォルダ名」のリンクに変換する（疑似スマートチップ）
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
    let itemName = url;
    let icon = "📄"; // デフォルトはファイルアイコン
    
    try {
      let itemId = "";
      let isFolder = false;
      const folderMatch = url.match(/\/folders\/([-\w]{25,})/);
      const fileMatch = url.match(/\/d\/([-\w]{25,})/);
      const queryMatch = url.match(/id=([-\w]{25,})/);

      if (folderMatch) { itemId = folderMatch[1]; isFolder = true; }
      else if (fileMatch) { itemId = fileMatch[1]; isFolder = false; }
      else if (queryMatch) { itemId = queryMatch[1]; isFolder = false; }

      if (itemId) {
        if (isFolder) {
          itemName = DriveApp.getFolderById(itemId).getName();
          icon = "📁";
        } else {
          try {
            itemName = DriveApp.getFileById(itemId).getName();
            icon = "📄";
          } catch (fileEx) {
            itemName = DriveApp.getFolderById(itemId).getName();
            icon = "📁";
          }
        }
      }
    } catch(ex) {
      itemName = "関連リンク " + (i + 1);
      icon = "🔗";
    }
    
    const textPart = icon + " " + itemName;
    fullText += (i > 0 ? "\n" : "") + textPart;
    linkData.push({ url: url, start: currentPos + (i > 0 ? 1 : 0), end: currentPos + (i > 0 ? 1 : 0) + textPart.length });
    currentPos = currentPos + (i > 0 ? 1 : 0) + textPart.length;
  });

  richTextValue.setText(fullText);
  linkData.forEach(ld => richTextValue.setLinkUrl(ld.start, ld.end, ld.url));
  range.setRichTextValue(richTextValue.build());
}