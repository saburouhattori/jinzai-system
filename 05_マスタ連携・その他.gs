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
    listSheet.getRange('B2:L51').clearContent();

    const result = [];
    const formulas = [];
    candIds.forEach(id => {
      let rowData = null;
      const sid = String(id).trim().toUpperCase();
      for (let i = 1; i < masterData.length; i++) { if (String(masterData[i][0]).trim().toUpperCase() === sid) { rowData = masterData[i]; break; } }
      if (rowData) {
        const getVal = (name) => col[name.replace(/\s/g, '')] ? rowData[col[name.replace(/\s/g, '')]-1] : "";
        result.push([getVal('名前'), getVal('フリガナ'), getVal('満年齢'), getVal('性別'), getVal('学歴＞学校名'), getVal('学歴＞状況'), 
                     getVal('特定技能要件＞JLPTレベル') || "×", getVal('特定技能要件＞JFTBasicレベル') || "×", getVal('その他の日本語能力試験'), id]);
        formulas.push(['=IFERROR(VLOOKUP(L' + (result.length + 1) + ', \'候補者写真\'!$A:$B, 2, FALSE), "")']);
      }
    });
    if (result.length) {
      listSheet.getRange(2, 3, result.length, 10).setValues(result);
      listSheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
    }
    return `${result.length}名の簡易リストを作成しました。`;
  } catch(e) { return "エラー: " + e.message; }
}