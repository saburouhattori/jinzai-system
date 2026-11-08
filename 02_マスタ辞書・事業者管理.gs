// =========================================
// 事業者マスタの操作・各種マスタ辞書取得
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

// =========================================
// フロントエンド用 各種マスタ・辞書データ取得
// =========================================

function getAgentList() {
  const sheet = getMasterSheet('送り出し機関マスタ');
  return sheet ?
    [...new Set(sheet.getDataRange().getValues().slice(1).map(row => String(row[1] || '').normalize('NFC').trim()).filter(n => n))].sort() : [];
}

function getSchoolList() {
  const sheet = getMasterSheet('日本語学校マスタ');
  return sheet ?
    [...new Set(sheet.getDataRange().getValues().slice(1).map(row => String(row[1] || '').normalize('NFC').trim()).filter(n => n))].sort() : [];
}

function getCompanyList() {
  const sheet = getMasterSheet('事業者マスタ');
  return sheet ?
    [...new Set(sheet.getDataRange().getValues().slice(1).map(row => String(row[1] || '').normalize('NFC').trim()).filter(n => n))].sort() : [];
}

function getCandidateDict() {
  const sheet = getMasterSheet('登録者マスタ');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach((h, i) => {
    col[String(h).replace(/\n/g, '').replace(/\s/g, '').trim()] = i;
  });
  
  const idIdx = col['登録者ID'];
  const nameIdx = col['名前'];
  const statusIdx = col['ステータス'];
  const birthIdx = col['生年月日'];

  const formatBirthIso = (val) => {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy-MM-dd");
    const s = String(val).trim().replace(/[年月\/]/g, '-').replace(/日/g, '');
    const parts = s.split('-');
    if (parts.length === 3) {
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }
    return s;
  };

  const dict = {};
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idIdx]) {
      const id = String(row[idIdx]).replace(/\s/g, '').toUpperCase();
      dict[id] = {
        name: nameIdx !== undefined ? String(row[nameIdx]).trim() : "",
        status: statusIdx !== undefined ? String(row[statusIdx]).trim() : "",
        birthday: birthIdx !== undefined ? formatBirthIso(row[birthIdx]) : ""
      };
    }
  }
  return dict;
}

function getJobDict() {
  const sheet = getMasterSheet('案件管理');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[String(h).replace(/\s/g, '')] = i);
  
  const idIdx = colMap['案件ID'];
  const compIdx = colMap['事業者名'];
  const statusIdx = colMap['ステータス'];
  const hireIdx = colMap['面接結果詳細']; 

  const dict = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idIdx]) {
      const id = String(row[idIdx]).replace(/\s/g, '').toUpperCase();
      const status = statusIdx !== undefined ? String(row[statusIdx] || "").trim() : "";
      const resultDetails = hireIdx !== undefined ? String(row[hireIdx] || "").trim() : "";
      const isRegistered = resultDetails !== "" || status === "終了" || status === "入国準備";
      dict[id] = {
        text: `${row[compIdx] || ""} (${status})`,
        isRegistered: isRegistered
      };
    }
  }
  return dict;
}
