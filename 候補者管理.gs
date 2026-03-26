/**
 * 検索機能：管理番号(SD-xxxx)からデータを取得（外部マスタ対応）
 */
function searchByAdminId(adminId) {
  const sheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('登録者マスタ');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const formatDate = (val) => (val instanceof Date) ? Utilities.formatDate(val, "JST", "yyyy-MM-dd") : val;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === adminId) { 
      return {
        row: i + 1, 
        name: data[i][1], furigana: data[i][3], nickname: data[i][4], birthday: formatDate(data[i][5]),
        age: data[i][6], gender: data[i][7], spouse: data[i][8], height: data[i][9], weight: data[i][10], 
        address: data[i][11], birthplace: data[i][12], email: data[i][13], school: data[i][14], eduSchool: data[i][15], 
        eduDept: data[i][16], eduStatus: data[i][17], 
        eduStart: data[i][18], eduEnd: data[i][19], eduNote: data[i][20],
        expPeriod1: data[i][21], expContent1: data[i][22], expPeriod2: data[i][23], expContent2: data[i][24], expPeriod3: data[i][25], expContent3: data[i][26],
        jlptLevel: data[i][27], jlptDate: data[i][28], jftLevel: data[i][29], jftDate: data[i][30], kaigoSkill: data[i][31], kaigoSkillDate: data[i][32],
        kaigoLang: data[i][33], kaigoLangDate: data[i][34], otherExam: data[i][35], otherExamDate: data[i][36],
        otherJapanese: data[i][37], otherJapaneseDate: data[i][38], comment: data[i][40], relative: data[i][41]
      };
    }
  }
  return null;
}

/**
 * 新規登録：テキストは外部マスタ、画像は自シート「候補者写真」へ
 */
function addNewRow(formData) {
  // 保存前に年月のフォーマットを「'YYYY年M月」に統一する
  const monthFields = ['eduStart', 'eduEnd', 'jlptDate', 'jftDate', 'kaigoSkillDate', 'kaigoLangDate', 'otherJapaneseDate'];
  monthFields.forEach(f => { if (formData[f]) formData[f] = normalizeYearMonth(formData[f]); });

  const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
  const masterSheet = masterSs.getSheetByName('登録者マスタ');
  const localSs = SpreadsheetApp.getActiveSpreadsheet();
  const photoSheet = localSs.getSheetByName('候補者写真');

  if (!masterSheet || !photoSheet) return 'エラー：シートが見つかりません。';

  // ID生成
  const lastRow = masterSheet.getLastRow();
  let nextNumber = 1;
  if (lastRow >= 2) {
    const lastValue = masterSheet.getRange(lastRow, 1).getValue().toString();
    const lastNumMatch = lastValue.match(/\d+/);
    if (lastNumMatch) nextNumber = parseInt(lastNumMatch[0], 10) + 1;
  }
  const nextId = "SD-" + nextNumber.toString().padStart(4, '0');

  // テキストデータの準備（42列）
  const rowData = new Array(42).fill(""); 
  rowData[0] = nextId; 
  rowData[1] = formData.name; 
  rowData[3] = formData.furigana;
  rowData[4] = formData.nickname;
  rowData[5] = formData.birthday; 
  rowData[7] = formData.gender;
  rowData[8] = formData.spouse; 
  rowData[9] = formData.height; 
  rowData[10] = formData.weight;
  rowData[11] = formData.address;
  rowData[12] = formData.birthplace; 
  rowData[13] = formData.email;
  rowData[14] = formData.school; 
  rowData[15] = formData.eduSchool; 
  rowData[16] = formData.eduDept;
  rowData[17] = formData.eduStatus;
  rowData[18] = formData.eduStart; 
  rowData[19] = formData.eduEnd;
  rowData[20] = formData.eduNote; 
  rowData[21] = formData.expPeriod1; 
  rowData[22] = formData.expContent1;
  rowData[23] = formData.expPeriod2;
  rowData[24] = formData.expContent2;
  rowData[25] = formData.expPeriod3; 
  rowData[26] = formData.expContent3;
  rowData[27] = formData.jlptLevel; 
  rowData[28] = formData.jlptDate;
  rowData[29] = formData.jftLevel;
  rowData[30] = formData.jftDate; 
  rowData[31] = formData.kaigoSkill;
  rowData[32] = formData.kaigoSkillDate; 
  rowData[33] = formData.kaigoLang; 
  rowData[34] = formData.kaigoLangDate;
  rowData[35] = formData.otherExam;
  rowData[36] = formData.otherExamDate; 
  rowData[37] = formData.otherJapanese;
  rowData[38] = formData.otherJapaneseDate; 
  rowData[40] = formData.comment; 
  rowData[41] = formData.relative;

  masterSheet.appendRow(rowData);

  // 画像データの直接書き込み
  if (formData.imageFile) {
    try {
      const dataUri = `data:${formData.imageFile.mimeType};base64,${formData.imageFile.contents}`;
      const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUri).build();
      
      photoSheet.appendRow([nextId, ""]); 
      const photoLastRow = photoSheet.getLastRow();
      photoSheet.getRange(photoLastRow, 2).setValue(cellImage);
      photoSheet.setRowHeight(photoLastRow, 80);
    } catch (e) {
      console.log("画像挿入エラー: " + e);
    }
  }
  
  updateAges(masterSheet.getLastRow());
  return `登録完了: ${nextId}`;
}

/**
 * データ更新：テキストはマスタ、画像は自シートを検索して上書き
 */
function updateRow(formData) {
  // 保存前に年月のフォーマットを「'YYYY年M月」に統一する
  const monthFields = ['eduStart', 'eduEnd', 'jlptDate', 'jftDate', 'kaigoSkillDate', 'kaigoLangDate', 'otherJapaneseDate'];
  monthFields.forEach(f => { if (formData[f]) formData[f] = normalizeYearMonth(formData[f]); });

  const masterSheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('登録者マスタ');
  const photoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('候補者写真');

  const row = Number(formData.row);
  if (!row) return "エラー：更新対象の行が特定できません。";
  const adminId = masterSheet.getRange(row, 1).getValue();

  const colMap = {
    name: 2, furigana: 4, nickname: 5, birthday: 6, gender: 8, spouse: 9, 
    height: 10, weight: 11, address: 12, birthplace: 13, email: 14, school: 15,
    eduSchool: 16, eduDept: 17, eduStatus: 18, eduStart: 19, eduEnd: 20, eduNote: 21,
    expPeriod1: 22, expContent1: 23, expPeriod2: 24, expContent2: 25, expPeriod3: 26, expContent3: 27,
    jlptLevel: 28, jlptDate: 29, jftLevel: 30, jftDate: 31, kaigoSkill: 32, kaigoSkillDate: 33,
    kaigoLang: 34, kaigoLangDate: 35, otherExam: 36, otherExamDate: 37,
    otherJapanese: 38, otherJapaneseDate: 39, comment: 41, 
    relative: 42
  };

  for (let key in colMap) {
    if (formData[key] !== undefined) {
      masterSheet.getRange(row, colMap[key]).setValue(formData[key]);
    }
  }
  
  if (formData.imageFile) {
    try {
      const dataUri = `data:${formData.imageFile.mimeType};base64,${formData.imageFile.contents}`;
      const cellImage = SpreadsheetApp.newCellImage().setSourceUrl(dataUri).build();
      
      const photoData = photoSheet.getDataRange().getValues();
      let found = false;
      for (let i = 0; i < photoData.length; i++) {
        if (photoData[i][0] === adminId) {
          photoSheet.getRange(i + 1, 2).setValue(cellImage);
          found = true;
          break;
        }
      }
      if (!found) {
        photoSheet.appendRow([adminId, ""]);
        photoSheet.getRange(photoSheet.getLastRow(), 2).setValue(cellImage);
        photoSheet.setRowHeight(photoSheet.getLastRow(), 80);
      }
    } catch (e) {
      console.log("画像更新エラー: " + e);
    }
  }
  
  updateAges(row); 
  return `更新が完了しました。`;
}

/**
 * 年齢計算
 */
function updateAges(targetRow) {
  const sheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('登録者マスタ');
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const today = new Date();

  if (targetRow) {
    const birthday = sheet.getRange(targetRow, 6).getValue();
    if (birthday instanceof Date) {
      let age = today.getFullYear() - birthday.getFullYear();
      const m = today.getMonth() - birthday.getMonth();
      if (m < 0 || (m === 0 && today.getDate() < birthday.getDate())) age--;
      sheet.getRange(targetRow, 7).setValue(age);
    }
  } else {
    const data = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    const ages = data.map(row => {
      const birthday = row[0];
      if (birthday instanceof Date) {
        let age = today.getFullYear() - birthday.getFullYear();
        const m = today.getMonth() - birthday.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < birthday.getDate())) age--;
        return [age];
      }
      return [""];
    });
    sheet.getRange(2, 7, ages.length, 1).setValues(ages);
  }
}

/**
 * 日付（年月）の表記ゆれを「'YYYY年M月」に強制統一し、日付変換をブロックする関数
 */
function normalizeYearMonth(val) {
  if (!val) return "";
  let str = val.toString().trim();
  
  // 1. 全角数字を半角に変換（例：「３」→「3」）
  str = str.replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  
  // 2. 年月の数字を抽出して「'YYYY年M月」に組み直す（一桁の月は0を消す）
  let match = str.match(/(\d{4})[-\/年](\d{1,2})/);
  if (match) {
    // ★ スプレッドシートの自動日付変換を防ぐために、先頭に「'（シングルクォーテーション）」を付与する
    return "'" + match[1] + "年" + parseInt(match[2], 10) + "月";
  }
  
  return str; // 変換できない文字列の場合はそのまま返す
}