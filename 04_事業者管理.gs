// =========================================
// 事業者マスタの操作
// =========================================

/**
 * 事業者の重複チェック
 */
function checkCompanyDuplicate(name) {
  const sheet = getMasterSheet('事業者マスタ');
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const searchName = String(name || "").normalize('NFC').trim();
  return data.some(row => String(row[1] || "").normalize('NFC').trim() === searchName);
}

/**
 * 事業者の新規追加
 */
function addCompany(formData) {
  const sheet = getMasterSheet('事業者マスタ');
  const data = sheet.getDataRange().getValues();
  
  const cleanName = String(formData.name || "").normalize('NFC').trim();
  const cleanYomi = String(formData.yomi || "").normalize('NFC').trim();
  const cleanAddress = String(formData.address || "").normalize('NFC').trim();
  const cleanUrl = String(formData.url || "").trim();
  const cleanNote = String(formData.note || "").normalize('NFC');

  // 重複チェック
  if (checkCompanyDuplicate(cleanName)) {
    return "エラー：この事業者は既に登録されています。";
  }

  // IDの発行
  let lastIdNum = 0;
  for (let i = 1; i < data.length; i++) {
    let idVal = String(data[i][0]);
    let match = idVal.match(/\d+/);
    if (match) {
      let num = parseInt(match[0], 10);
      if (num > lastIdNum) lastIdNum = num;
    }
  }
  const nextId = "CO-" + (lastIdNum + 1).toString().padStart(4, '0');

  // 新しい行を追加
  const newRow = [
    nextId,
    cleanName,
    cleanYomi,
    cleanAddress,
    cleanUrl,
    cleanNote
  ];
  
  sheet.appendRow(newRow);
  return `事業者登録が完了しました: ${nextId}`;
}