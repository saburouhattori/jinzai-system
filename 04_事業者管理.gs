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
  const searchName = String(name).trim();
  return data.some(row => String(row[1]).trim() === searchName);
}

/**
 * 事業者の新規追加
 */
function addCompany(formData) {
  const sheet = getMasterSheet('事業者マスタ');
  const data = sheet.getDataRange().getValues();
  
  // 重複チェック
  if (checkCompanyDuplicate(formData.name)) {
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
    formData.name,
    formData.yomi || "",
    formData.address || "",
    formData.url || "",
    formData.note || ""
  ];
  
  sheet.appendRow(newRow);
  return `事業者登録が完了しました: ${nextId}`;
}