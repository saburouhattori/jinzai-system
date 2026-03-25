/**
 * 事業者名の重複チェック用
 */
function checkDuplicateCompanyName(name) {
  const sheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('事業者マスタ');
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  const targetName = name.trim().toLowerCase();
  
  const isDuplicate = data.some((row, index) => 
    index > 0 && row[1].toString().trim().toLowerCase() === targetName
  );
  
  return isDuplicate;
}

/**
 * 事業者マスタへの新規登録（外部マスタ対応）
 */
function addCompanyRow(formData) {
  const sheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('事業者マスタ');
  if (!sheet) return 'エラー：外部の「事業者マスタ」が見つかりません。';

  const companyName = formData.name.trim();
  if (!companyName) return 'エラー：事業者名称を入力してください。';

  const lastRow = sheet.getLastRow();
  let nextNumber = 1;
  if (lastRow >= 2) {
    const lastValue = sheet.getRange(lastRow, 1).getValue().toString();
    const lastNumMatch = lastValue.match(/\d+/);
    if (lastNumMatch) nextNumber = parseInt(lastNumMatch[0], 10) + 1;
  }
  const nextId = "CO-" + nextNumber.toString().padStart(4, '0');
  
  const rowData = [
    nextId,
    companyName,
    formData.yomi || "",
    formData.address || "",
    formData.url || "",
    formData.note || ""
  ];
  
  try {
    sheet.appendRow(rowData);
    return `登録完了: ${nextId}\n${companyName}`;
  } catch (e) {
    return "登録エラー: " + e.toString();
  }
}