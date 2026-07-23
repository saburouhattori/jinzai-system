// =========================================
// その他の便利ツール・マスタ連携
// =========================================

function getAgentList() {
  const sheet = getMasterSheet('送り出し機関マスタ');
  return sheet ?
    [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getSchoolList() {
  const sheet = getMasterSheet('日本語学校マスタ');
  return sheet ?
    [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
}

function getCompanyList() {
  const sheet = getMasterSheet('事業者マスタ');
  return sheet ?
    [...new Set(sheet.getDataRange().getValues().slice(1).map(row => row[1]).filter(n => n))].sort() : [];
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
    if (result.length > 0) {
      listSheet.getRange(2, 3, result.length, 10).setValues(result);
      listSheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
    }
    return `${result.length}名の簡易リストを作成しました。`;
  } catch(e) {
    return "エラー: " + e.message;
  }
}

// ====== マスタ連携・リスト同期処理 ======

function syncListSheets() {
  updateCandidateLists(true);
  return 'リストの同期が完了しました。';
}

function normalize_(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/[\s \n\r]+/g, '').toLowerCase();
}

function buildRowByHeaders_(headers, dataMap) {
  return headers.map(h => {
    const key = normalize_(h);
    return dataMap[key] !== undefined ? dataMap[key] : '';
  });
}

function updateCandidateLists(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('登録者マスタ');
  const hiredSheet = ss.getSheetByName('採用者一覧');
  const unhiredSheet = ss.getSheetByName('未採用者一覧');
  const jobSheet = ss.getSheetByName('案件管理');
  if (!masterSheet || !hiredSheet || !unhiredSheet || !jobSheet) throw new Error('必要なシートが見つかりません。');

  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData.shift();
  const hiredHeaders = hiredSheet.getRange(1, 1, 1, hiredSheet.getLastColumn()).getValues()[0];
  const unhiredHeaders = unhiredSheet.getRange(1, 1, 1, unhiredSheet.getLastColumn()).getValues()[0];
  
  const jobData = jobSheet.getDataRange().getValues();
  const jobHeaders = jobData.shift();
  const jobList = jobData.map(row => {
    const jobMap = {};
    for (let i = 0; i < jobHeaders.length; i++) jobMap[normalize_(jobHeaders[i])] = row[i];
    return jobMap;
  });
  
  if (masterHeaders.findIndex(h => normalize_(h) === normalize_('登録者ID')) === -1) {
    throw new Error('登録者マスタに「登録者ID」列が見つかりません。');
  }

  const hiredCandidatesMap = new Map();

  for (const job of jobList) {
    const jobId = job[normalize_('案件ID')] || '';
    const skillField = job[normalize_('技能分野')] || '';
    const hiredText = String(job[normalize_('採用者名')] || '');

    if (!hiredText || hiredText.includes("採用者なし")) continue;

    const compText = String(job[normalize_('事業者名')] || '');
    const defaultCompany = compText.split(/\r?\n/)[0].trim();
    let currentCompany = defaultCompany;

    const hiredLines = hiredText.split(/\r?\n/).filter(line => line.trim() !== "");
    for (const line of hiredLines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentCompany = line.slice(1, -1).trim();
        continue;
      }

      const match = line.match(/^(SD-\d+)/);
      if (match) {
        hiredCandidatesMap.set(match[1], {
          jobId: jobId,
          skillField: skillField,
          company: currentCompany
        });
      }
    }
  }

  const hiredData = [];
  const unhiredData = [];

  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];
    const dataMap = {};
    for (let c = 0; c < masterHeaders.length; c++) {
      let val = row[c];
      if (val instanceof Date) val = Utilities.formatDate(val, "JST", "yyyy/MM/dd");
      dataMap[normalize_(masterHeaders[c])] = val;
    }

    const candidateId = dataMap[normalize_('登録者ID')];
    if (!candidateId) continue;

    const jlpt = dataMap[normalize_('特定技能要件＞JLPTレベル')];
    const jft = dataMap[normalize_('特定技能要件＞JFT Basicレベル')];
    const kaigoGinou = dataMap[normalize_('特定技能要件＞介護技能評価試験')];
    const kaigoNihongo = dataMap[normalize_('特定技能要件＞介護日本語評価試験')];
    
    let reqs = [];
    if (jlpt && jlpt !== "-" && jlpt !== "×" && !jlpt.includes("予定") && !jlpt.includes("不合格")) reqs.push(jlpt);
    if (jft && jft !== "-" && jft !== "×" && !jft.includes("予定") && !jft.includes("不合格")) reqs.push(jft);
    if (kaigoGinou && kaigoGinou !== "-" && kaigoGinou !== "×" && !kaigoGinou.includes("不合格")) {
       if (kaigoGinou.includes("予定")) reqs.push("介護技能（受験予定）");
       else reqs.push("介護技能（合格）");
    }
    if (kaigoNihongo && kaigoNihongo !== "-" && kaigoNihongo !== "×" && !kaigoNihongo.includes("不合格")) {
       if (kaigoNihongo.includes("予定")) reqs.push("介護日本語（受験予定）");
       else reqs.push("介護日本語（合格）");
    }
    dataMap[normalize_('特定技能要件')] = reqs.join(', ');

    dataMap[normalize_('JLPT')] = jlpt;
    dataMap[normalize_('JFT Basic')] = jft;
    dataMap[normalize_('在留資格交付申請の有無')] = dataMap[normalize_('在留資格交付申請の回数')];

    if (hiredCandidatesMap.has(candidateId)) {
      const hireInfo = hiredCandidatesMap.get(candidateId);
      dataMap[normalize_('案件ID')] = hireInfo.jobId;
      dataMap[normalize_('技能分野')] = hireInfo.skillField;
      dataMap[normalize_('採用事業者名')] = hireInfo.company;
      dataMap[normalize_('採用事業者')] = hireInfo.company; 

      hiredData.push(buildRowByHeaders_(hiredHeaders, dataMap));
    } else {
      const status = dataMap[normalize_('ステータス')];
      if (status === '未採用') {
        unhiredData.push(buildRowByHeaders_(unhiredHeaders, dataMap));
      }
    }
  }

  if (hiredData.length > 0) {
    const lastRow = hiredSheet.getLastRow();
    if (lastRow > 1) hiredSheet.getRange(2, 1, lastRow - 1, hiredHeaders.length).clearContent();
    hiredSheet.getRange(2, 1, hiredData.length, hiredHeaders.length).setValues(hiredData);
  } else {
    const lastRow = hiredSheet.getLastRow();
    if (lastRow > 1) hiredSheet.getRange(2, 1, lastRow - 1, hiredHeaders.length).clearContent();
  }

  if (unhiredData.length > 0) {
    const lastRow = unhiredSheet.getLastRow();
    if (lastRow > 1) unhiredSheet.getRange(2, 1, lastRow - 1, unhiredHeaders.length).clearContent();
    unhiredSheet.getRange(2, 1, unhiredData.length, unhiredHeaders.length).setValues(unhiredData);
  } else {
    const lastRow = unhiredSheet.getLastRow();
    if (lastRow > 1) unhiredSheet.getRange(2, 1, lastRow - 1, unhiredHeaders.length).clearContent();
  }

  if (!silent) {
    try { SpreadsheetApp.getUi().alert('リストの更新が完了しました。'); } catch(e) {}
  }
}