// =========================================
// その他の便利ツール・マスタ連携
// =========================================

function getAgentList() {
  const sheet = getMasterSheet('送り出し機関マスタ');
  return sheet ? [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getSchoolList() {
  const sheet = getMasterSheet('日本語学校マスタ');
  return sheet ? [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getCompanyList() {
  const sheet = getMasterSheet('事業者マスタ');
  return sheet ? [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getCandidateDict() {
  const sheet = getMasterSheet('登録者マスタ');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const dict = {};
  data.forEach(row => { if (row[0]) dict[String(row[0]).trim()] = String(row[1]); });
  return dict;
}

function getJobDict() {
  const sheet = getMasterSheet('案件管理');
  if (!sheet) return {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const dict = {};
  data.forEach(row => { if (row[0]) dict[String(row[0]).trim()] = `${row[3]} (${row[1]})`; });
  return dict;
}

function searchDriveFiles(fileNameQuery) {
  try {
    const files = [];
    let query = 'trashed = false';
    if (fileNameQuery) {
      query += ' and title contains "' + fileNameQuery + '"';
    }
    
    const iter = DriveApp.searchFiles(query);
    let count = 0;
    while (iter.hasNext() && count < 15) {
      const file = iter.next();
      files.push({
        name: file.getName(),
        url: file.getUrl(),
        type: file.getMimeType()
      });
      count++;
    }
    return files;
  } catch (e) { return []; }
}

function generateSimpleList(candIds) {
  try {
    const masterSheet = getMasterSheet('登録者マスタ');
    const masterData = masterSheet.getDataRange().getValues();
    const col = getMasterColumnMap(masterSheet);
    const listSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('簡易リスト');
    
    const lastRowList = listSheet.getLastRow();
    if (lastRowList >= 2) {
      listSheet.getRange(2, 2, lastRowList, 11).clearContent();
    }

    const result = [];
    const formulas = [];
    candIds.forEach(id => {
      let rowData = null;
      const sid = String(id).trim().toUpperCase();
      for (let i = 1; i < masterData.length; i++) { 
        if (String(masterData[i][0]).trim().toUpperCase() === sid) { 
          rowData = masterData[i]; 
          break; 
        } 
      }
      if (rowData) {
        const getVal = (name) => col[name.replace(/\s/g, '')] ? rowData[col[name.replace(/\s/g, '')]-1] : "";
        result.push([
          getVal('名前'), 
          getVal('フリガナ'), 
          getVal('満年齢'), 
          getVal('性別'), 
          getVal('学歴＞学校名'), 
          getVal('学歴＞状況'), 
          getVal('特定技能要件＞JLPTレベル') || "×", 
          getVal('特定技能要件＞JFTBasicレベル') || "×", 
          getVal('その他の日本語能力試験'), 
          id
        ]);
        formulas.push(['=IFERROR(VLOOKUP(L' + (result.length + 1) + ', \'登録者マスタ\'!$A:$C, 3, FALSE), "")']);
      }
    });
    if (result.length) {
      listSheet.getRange(2, 3, result.length, 10).setValues(result);
      listSheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
    }
    return `${result.length}名の簡易リストを作成しました。`;
  } catch(e) { return "エラー: " + e.message; }
}

// 採用者・未採用者一覧を「静的テキスト」で同期する処理
function syncListSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('登録者マスタ');
    if (!masterSheet) return "登録者マスタが見つかりません。";

    const mData = masterSheet.getDataRange().getValues();
    if (mData.length < 2) return "マスタにデータがありません。";

    // --- 採用者一覧の列指定 (Masterの列番号を1ベースで指定) ---
    // 1:採用事業者(45), 2:技能分野(46), 3:内定日(49), 4:ID(1), 5:名前(2), 6:フリガナ(4)...
    const adoptedCols = [
      45, 46, 49, 1, 2, 4, 5, 6, 7, 8, 9, 14, 28, 30, 15, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71
    ].map(c => c - 1);
    
    const unadoptedData = [];
    const adoptedData = [];

    const formatDate = (val) => {
      if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
      return val;
    };

    for (let i = 1; i < mData.length; i++) {
      const row = mData[i];
      const status = String(row[43]).trim(); // AR列: ステータス

      if (status === '未採用') {
        let quals = [];
        // JLPT (AB列: 28)
        let jlpt = String(row[27] || "").trim();
        if (jlpt && jlpt !== "-" && jlpt !== "×" && !jlpt.includes("予定") && !jlpt.includes("不合格")) quals.push(jlpt);
        
        // JFT (AD列: 30)
        let jft = String(row[29] || "").trim();
        if (jft && jft !== "-" && jft !== "×" && !jft.includes("予定") && !jft.includes("不合格")) quals.push(jft);
        
        // 介護 (AF, AH)
        if (String(row[31]).includes("予定")) quals.push("介護技能（受験予定）");
        else if (row[31]) quals.push("介護技能（合格）");
        if (String(row[33]).includes("予定")) quals.push("介護日本語（受験予定）");
        else if (row[33]) quals.push("介護日本語（合格）");

        // その他 (AJ列: 36)
        let other = String(row[35] || "").trim();
        if (other && !other.includes("不合格")) {
          if (other.includes("外食業")) quals.push("外食業技能（合格）");
          else if (other.includes("飲食料品製造")) quals.push("飲食料品製造技能（合格）");
          else quals.push(other + (other.includes("合格") ? "" : "（合格）"));
        }

        unadoptedData.push([
          formatDate(row[0]), formatDate(row[1]), formatDate(row[6]), formatDate(row[7]), 
          quals.join(", "), formatDate(row[42]), formatDate(row[39])
        ]);
      } else if (status === '採用' || status === '内定') {
        adoptedData.push(adoptedCols.map(idx => row[idx] !== undefined ? formatDate(row[idx]) : ""));
      }
    }

    const writeToSheet = (name, data) => {
      const target = ss.getSheetByName(name);
      if (!target) return;
      if (target.getLastRow() > 1) target.getRange(2, 1, target.getLastRow()-1, target.getLastColumn()).clearContent();
      if (data.length) target.getRange(2, 1, data.length, data[0].length).setValues(data);
    };

    writeToSheet('未採用者一覧', unadoptedData);
    writeToSheet('採用者一覧', adoptedData);

    return "リストの同期が完了しました。";
  } catch (e) { return "エラー: " + e.message; }
}