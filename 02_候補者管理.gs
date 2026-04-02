// =========================================
// 候補者データの操作（検索・登録・更新・削除）
// =========================================

function safeSearchByAdminId(id) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    const col = getMasterColumnMap(sheet);
    const searchId = String(id).trim().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      const sheetId = String(data[i][0] || "").trim().toUpperCase();
      if (sheetId === searchId) {
        const rowData = data[i];
        const res = { row: i + 1 };
        
        const getVal = (name) => {
          const cIdx = col[name.replace(/\s/g, '')];
          if (!cIdx) return "";
          const val = rowData[cIdx - 1];
          if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy-MM-dd");
          return String(val || "").trim();
        };

        // UI項目へのマッピング
        res.name = getVal('名前'); res.furigana = getVal('フリガナ');
        res.nickname = getVal('呼び名'); 
        res.birthday = getVal('生年月日'); res.gender = getVal('性別'); res.spouse = getVal('配偶者'); 
        res.height = getVal('身長'); res.weight = getVal('体重');
        res.address = getVal('現住所'); 
        res.birthplace = getVal('住所（出身地）'); res.email = getVal('メールアドレス'); res.school = getVal('所属日本語学校');
        res.eduSchool = getVal('学歴＞学校名'); res.eduDept = getVal('学歴＞学部・学科・専攻');
        res.eduStatus = getVal('学歴＞状況');
        res.eduStart = getVal('学歴＞入学年月'); res.eduEnd = getVal('学歴＞卒業/中退年月'); res.eduNote = getVal('学歴＞補足');
        res.expPeriod1 = getVal('職歴①＞期間'); res.expContent1 = getVal('職歴①＞内容');
        res.expPeriod2 = getVal('職歴②＞期間'); res.expContent2 = getVal('職歴②＞内容');
        res.expPeriod3 = getVal('職歴③＞期間'); res.expContent3 = getVal('職歴③＞内容');
        res.jlptLevel = getVal('特定技能要件＞JLPTレベル'); res.jlptDate = getVal('特定技能要件＞JLPT取得年月');
        res.jftLevel = getVal('特定技能要件＞JFTBasicレベル'); res.jftDate = getVal('特定技能要件＞JFT取得年月');
        res.kaigoSkill = getVal('特定技能要件＞介護技能評価試験'); res.kaigoSkillDate = getVal('特定技能要件＞介護技能取得年月');
        res.kaigoLang = getVal('特定技能要件＞介護日本語評価試験'); res.kaigoLangDate = getVal('特定技能要件＞介護日本語取得年月');
        res.otherJapanese = getVal('その他の日本語能力試験'); res.otherJapaneseDate = getVal('取得年月');
        res.comment = getVal('修正前コメント'); 
        res.relative = getVal('日本在住の親族について');
        // 追加情報
        res.agent = getVal('所属送り出し機関'); res.offerDate = getVal('内定日'); res.birthCity = getVal('出生地（都市名）');
        res.addressDetail = getVal('住所詳細'); res.passportNum = getVal('パスポート番号'); res.passportExp = getVal('パスポート有効期限');
        res.job = getVal('職業'); res.traineeExp = getVal('技能実習の経験の有無'); res.traineeCert = getVal('技能実習修了書の有無');
        res.crime = getVal('犯罪歴の有無'); res.applyCount = getVal('在留資格交付申請の回数'); res.rejectCount = getVal('不許可となった在留資格交付申請の回数');
        res.overseasExp = getVal('海外への出入国歴の有無'); res.travelCount = getVal('出入国の回数'); res.lastInDate = getVal('直近の入国日');
        res.lastOutDate = getVal('直近の出国日');
        res.relName2 = getVal('日本在住の親族情報親族の名前'); res.relRelation2 = getVal('日本在住の親族情報続柄'); res.relBirth2 = getVal('日本在住の親族情報親族の生年月日');
        res.relCountry2 = getVal('日本在住の親族情報親族の国籍・地域'); res.relLive2 = getVal('日本在住の親族情報親族との同居予定の有無');
        res.relWork2 = getVal('日本在住の親族情報親族の勤務先・通学先'); res.relCard2 = getVal('日本在住の親族情報親族の在留カード番号'); res.memo = getVal('備考・メモ');

        return res;
      }
    }
    return null;
  } catch(e) { throw new Error("検索エラー: " + e.message);
  }
}

function addNewRow(formData) {
  const monthFields = ['eduStart', 'eduEnd', 'jlptDate', 'jftDate', 'kaigoSkillDate', 'kaigoLangDate', 'otherJapaneseDate'];
  monthFields.forEach(f => { if (formData[f]) formData[f] = normalizeYearMonth(formData[f]); });
  if (formData.birthday) formData.birthday = new Date(formData.birthday.replace(/-/g, '/'));

  const masterSheet = getMasterSheet('登録者マスタ');
  const col = getMasterColumnMap(masterSheet);

  const lastRow = masterSheet.getLastRow();
  let nextNumber = 1;
  if (lastRow >= 2) {
    const lastId = String(masterSheet.getRange(lastRow, 1).getValue());
    const match = lastId.match(/\d+/);
    if (match) nextNumber = parseInt(match[0], 10) + 1;
  }
  const nextId = "SD-" + nextNumber.toString().padStart(4, '0');
  
  const safeMaxCol = Math.max(masterSheet.getLastColumn(), ...Object.values(col));
  const rowValues = new Array(safeMaxCol).fill("");
  
  const mapping = {
    '登録者ID': nextId, '名前': formData.name, 'フリガナ': formData.furigana, '呼び名': formData.nickname,
    '生年月日': formData.birthday, '性別': formData.gender, '配偶者': formData.spouse, 
    '身長': formData.height, '体重': formData.weight, '現住所': formData.address, 
    '住所（出身地）': formData.birthplace, 'メールアドレス': formData.email, '所属日本語学校': formData.school,
    '学歴＞学校名': formData.eduSchool, '学歴＞学部・学科・専攻': formData.eduDept, '学歴＞状況': formData.eduStatus,
    '学歴＞入学年月': formData.eduStart, '学歴＞卒業/中退年月': formData.eduEnd, '学歴＞補足': formData.eduNote,
    '職歴①＞期間': formData.expPeriod1, '職歴①＞内容': formData.expContent1,
    '職歴②＞期間': formData.expPeriod2, '職歴②＞内容': formData.expContent2,
    '職歴③＞期間': formData.expPeriod3, '職歴③＞内容': formData.expContent3,
    '特定技能要件＞JLPTレベル': formData.jlptLevel, '特定技能要件＞JLPT取得年月': formData.jlptDate,
    '特定技能要件＞JFTBasicレベル': formData.jftLevel, 
    '特定技能要件＞JFT取得年月': formData.jftDate,
    '特定技能要件＞介護技能評価試験': formData.kaigoSkill, '特定技能要件＞介護技能取得年月': formData.kaigoSkillDate,
    '特定技能要件＞介護日本語評価試験': formData.kaigoLang, '特定技能要件＞介護日本語取得年月': formData.kaigoLangDate,
    'その他の日本語能力試験': formData.otherJapanese, '取得年月': formData.otherJapaneseDate,
    '修正前コメント': formData.comment, 
    '日本在住の親族について': formData.relative, 'ステータス': '未採用'
  };
  for (let header in mapping) {
    const h = header.replace(/\s/g, '');
    if (col[h]) rowValues[col[h]-1] = mapping[header];
  }

  masterSheet.appendRow(rowValues);
  const newRow = masterSheet.getLastRow();

  if (col['生年月日']) masterSheet.getRange(newRow, col['生年月日']).setNumberFormat('yyyy"年"m"月"d"日"');
  
  if (formData.imageFile && col['顔写真']) {
    try {
      const dataUri = `data:${formData.imageFile.mimeType};base64,${formData.imageFile.contents}`;
      const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUri).build();
      masterSheet.getRange(newRow, col['顔写真']).setValue(cellImage);
      masterSheet.setRowHeight(newRow, 80);
    } catch (e) {}
  }
  
  updateAges(newRow);
  return `登録完了: ${nextId}`;
}

function updateRow(formData) {
  const monthFields = ['eduStart', 'eduEnd', 'jlptDate', 'jftDate', 'kaigoSkillDate', 'kaigoLangDate', 'otherJapaneseDate'];
  monthFields.forEach(f => { if (formData[f]) formData[f] = normalizeYearMonth(formData[f]); });
  if (formData.birthday) formData.birthday = new Date(formData.birthday.replace(/-/g, '/'));

  const masterSheet = getMasterSheet('登録者マスタ');
  const col = getMasterColumnMap(masterSheet);
  const row = Number(formData.row);
  if (!row) return "エラー：行が不明です。";
  
  const mapping = {
    '名前': formData.name, 'フリガナ': formData.furigana, '呼び名': formData.nickname, '生年月日': formData.birthday,
    '性別': formData.gender, '配偶者': formData.spouse, '身長': formData.height, '体重': formData.weight,
    '現住所': formData.address, '住所（出身地）': formData.birthplace, 'メールアドレス': formData.email,
    '所属日本語学校': formData.school, '学歴＞学校名': formData.eduSchool, '学歴＞学部・学科・専攻': formData.eduDept,
    '学歴＞状況': formData.eduStatus, '学歴＞入学年月': formData.eduStart, '学歴＞卒業/中退年月': formData.eduEnd,
    '学歴＞補足': formData.eduNote, '職歴①＞期間': formData.expPeriod1, '職歴①＞内容': formData.expContent1,
    '職歴②＞期間': formData.expPeriod2, '職歴②＞内容': formData.expContent2, '職歴③＞期間': formData.expPeriod3,
    '職歴③＞内容': formData.expContent3, '特定技能要件＞JLPTレベル': formData.jlptLevel,
    '特定技能要件＞JLPT取得年月': formData.jlptDate, '特定技能要件＞JFTBasicレベル': formData.jftLevel,
    '特定技能要件＞JFT取得年月': formData.jftDate, '特定技能要件＞介護技能評価試験': formData.kaigoSkill,
    '特定技能要件＞介護技能取得年月': formData.kaigoSkillDate, '特定技能要件＞介護日本語評価試験': formData.kaigoLang,
    '特定技能要件＞介護日本語取得年月': formData.kaigoLangDate, 'その他の日本語能力試験': formData.otherJapanese,
    '取得年月': formData.otherJapaneseDate, 
    '修正前コメント': formData.comment, 
    '日本在住の親族について': formData.relative
  };

  const photoIdx = col['顔写真'] ? col['顔写真'] - 1 : -1;
  const safeMaxCol = Math.max(masterSheet.getLastColumn(), ...Object.values(col));
  const currentRowRange = masterSheet.getRange(row, 1, 1, safeMaxCol);
  const currentRowData = currentRowRange.getValues()[0];

  for (let header in mapping) {
    const h = header.replace(/\s/g, '');
    if (col[h] && mapping[header] !== undefined) {
      currentRowData[col[h] - 1] = mapping[header];
    }
  }
  
  if (photoIdx !== -1 && photoIdx < safeMaxCol) {
    if (photoIdx > 0) {
      masterSheet.getRange(row, 1, 1, photoIdx).setValues([currentRowData.slice(0, photoIdx)]);
    }
    if (photoIdx < safeMaxCol - 1) {
      masterSheet.getRange(row, photoIdx + 2, 1, safeMaxCol - (photoIdx + 1)).setValues([currentRowData.slice(photoIdx + 1)]);
    }
  } else {
    currentRowRange.setValues([currentRowData]);
  }

  if (col['生年月日']) masterSheet.getRange(row, col['生年月日']).setNumberFormat('yyyy"年"m"月"d"日"');

  if (formData.imageFile && col['顔写真']) {
    try {
      const dataUri = `data:${formData.imageFile.mimeType};base64,${formData.imageFile.contents}`;
      const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUri).build();
      masterSheet.getRange(row, col['顔写真']).setValue(cellImage);
      masterSheet.setRowHeight(row, 80);
    } catch (e) {}
  }
  
  updateAges(row); 
  return `更新が完了しました。`;
}

function updateAddInfoRow(formData) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) return "エラー：登録者マスタが見つかりません。";
    const col = getMasterColumnMap(sheet);
    const row = Number(formData.row);
    const mapping = {
      agent: '所属送り出し機関', offerDate: '内定日', birthCity: '出生地（都市名）',
      addressDetail: '住所詳細', passportNum: 'パスポート番号', passportExp: 'パスポート有効期限',
      job: '職業', traineeExp: '技能実習の経験の有無', traineeCert: '技能実習修了書の有無',
      crime: '犯罪歴の有無', applyCount: '在留資格交付申請の回数', rejectCount: '不許可となった在留資格交付申請の回数',
      overseasExp: '海外への出入国歴の有無', travelCount: '出入国の回数', lastInDate: '直近の入国日', lastOutDate: '直近の出国日',
      relName2: '日本在住の親族情報親族の名前', relRelation2: '日本在住の親族情報続柄',
      relBirth2: '日本在住の親族情報親族の生年月日', relCountry2: '日本在住の親族情報親族の国籍・地域',
      relLive2: '日本在住の親族情報親族との同居予定の有無', relWork2: '日本在住の親族情報親族の勤務先・通学先',
      relCard2: '日本在住の親族情報親族の在留カード番号', memo: '備考・メモ'
    };

    const photoIdx = col['顔写真'] ? col['顔写真'] - 1 : -1;
    const safeMaxCol = Math.max(sheet.getLastColumn(), ...Object.values(col));
    const currentRowRange = sheet.getRange(row, 1, 1, safeMaxCol);
    const currentRowData = currentRowRange.getValues()[0];

    for (let key in mapping) {
      const h = mapping[key].replace(/\s/g, '');
      if (col[h] && formData[key] !== undefined) {
        currentRowData[col[h] - 1] = formData[key];
      }
    }
    
    if (photoIdx !== -1 && photoIdx < safeMaxCol) {
      if (photoIdx > 0) {
        sheet.getRange(row, 1, 1, photoIdx).setValues([currentRowData.slice(0, photoIdx)]);
      }
      if (photoIdx < safeMaxCol - 1) {
        sheet.getRange(row, photoIdx + 2, 1, safeMaxCol - (photoIdx + 1)).setValues([currentRowData.slice(photoIdx + 1)]);
      }
    } else {
      currentRowRange.setValues([currentRowData]);
    }

    return `追加情報の登録が完了しました。`;
  } catch (e) { return "エラー: " + e.message; }
}

function deleteCandidate(id) {
  if (!id) return "IDが指定されていません。";
  try {
    const masterSheet = getMasterSheet('登録者マスタ');
    const masterData = masterSheet.getDataRange().getValues();
    const searchId = String(id).trim().toUpperCase();
    for (let i = masterData.length - 1; i >= 1; i--) {
      if (String(masterData[i][0]).trim().toUpperCase() === searchId) {
        masterSheet.deleteRow(i + 1);
        return "登録者情報をマスタから削除しました。";
      }
    }
    return "エラー: 指定されたIDが見つかりませんでした。";
  } catch (e) { return "エラー: " + e.toString(); }
}

function updateAges(targetRow) {
  const sheet = getMasterSheet('登録者マスタ');
  const col = getMasterColumnMap(sheet);
  if (!col['生年月日'] || !col['満年齢']) return;
  const today = new Date();
  const birthday = sheet.getRange(targetRow, col['生年月日']).getValue();
  if (birthday instanceof Date) {
    let age = today.getFullYear() - birthday.getFullYear();
    const m = today.getMonth() - birthday.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < birthday.getDate())) age--;
    sheet.getRange(targetRow, col['満年齢']).setValue(age);
  }
}

function normalizeYearMonth(val) {
  if (!val) return "";
  let str = val.toString().trim().replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  let match = str.match(/(\d{4})[-\/年](\d{1,2})/);
  return match ? "'" + match[1] + "年" + parseInt(match[2], 10) + "月" : str;
}