// =========================================
// 事業者データの操作
// =========================================

function addCompany(formData) {
  const sheet = getMasterSheet('事業者マスタ');
  if (!sheet) return "エラー：事業者マスタが見つかりません。";
  const lastRow = sheet.getLastRow();
  let nextId = "CO-0001";
  if (lastRow >= 2) {
    const lastId = String(sheet.getRange(lastRow, 1).getValue());
    const nextNum = (parseInt(lastId.match(/\d+/)[0], 10) + 1);
    nextId = "CO-" + nextNum.toString().padStart(4, '0');
  }
  sheet.appendRow([nextId, formData.name, formData.yomi, formData.address, formData.url, formData.note]);
  return `事業者登録完了: ${nextId}`;
}

function getCompanyList() {
  const sheet = getMasterSheet('事業者マスタ');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => row[1]).filter(n => n);
}

function checkCompanyDuplicate(name) {
  const list = getCompanyList();
  return list.indexOf(name.trim()) !== -1;
}