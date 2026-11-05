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
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach((h, i) => {
    col[String(h).replace(/\n/g, '').replace(/\s/g, '').trim()] = i;
  });
  
  const idIdx = col['登録者ID'];
  const nameIdx = col['名前'];
  const statusIdx = col['ステータス'];

  const dict = {};
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idIdx]) {
      const id = String(row[idIdx]).replace(/\s/g, '').toUpperCase();
      dict[id] = {
        name: nameIdx !== undefined ? String(row[nameIdx]) : "",
        status: statusIdx !== undefined ? String(row[statusIdx]).trim() : ""
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

  // 日付比較・整形用のヘルパー関数
  const toDateStr = (val) => {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
    const s = String(val).trim().replace(/[年月\-]/g, '/').replace(/日/g, '');
    const parts = s.split('/');
    if (parts.length === 3) {
      return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
    }
    return s;
  };

  const toDateObj = (val) => {
    if (!val) return '';
    if (val instanceof Date) return val;
    const str = toDateStr(val);
    const d = new Date(str);
    return isNaN(d.getTime()) ? val : d;
  };

  const hiredCandidatesMap = new Map();
  const declinedCandidatesSet = new Set();

  for (const job of jobList) {
    const jobId = job[normalize_('案件ID')] || '';
    const skillField = job[normalize_('技能分野')] || '';
    const detailsText = String(job[normalize_('面接結果詳細')] || '');
    const rawOfferDate = job[normalize_('内定日')];
    const offerDateStr = toDateStr(rawOfferDate);
    const offerDateObj = toDateObj(rawOfferDate);

    if (!detailsText) continue;

    const compText = String(job[normalize_('事業者名')] || '');
    const defaultCompany = compText.split(/\r?\n/)[0].trim();
    let currentCompany = defaultCompany;

    const dLines = detailsText.split(/\r?\n/).filter(line => line.trim() !== "");
    for (const line of dLines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentCompany = line.slice(1, -1).trim();
        continue;
      }

      const match = line.match(/^(SD-\d+)/);
      if (!match) continue;
      const cid = match[1].toUpperCase();

      if (line.includes('（採用）')) {
        hiredCandidatesMap.set(cid, {
          jobId: jobId,
          skillField: skillField,
          company: currentCompany,
          offerDateStr: offerDateStr,
          offerDateObj: offerDateObj
        });
      } else if (line.includes('（内定辞退')) {
        declinedCandidatesSet.add(cid);
      }
    }
  }

  const hiredData = [];
  const unhiredData = [];

  // 登録者マスタの「ステータス」「採用事業者」「内定日」列を案件管理（正本）のデータで同期するための準備
  const masterStatusColIdx = masterHeaders.findIndex(h => normalize_(h) === normalize_('ステータス'));
  const masterCompColIdx = masterHeaders.findIndex(h => normalize_(h) === normalize_('採用事業者'));
  const masterOfferColIdx = masterHeaders.findIndex(h => normalize_(h) === normalize_('内定日'));

  const masterStatusUpdates = [];
  const masterCompUpdates = [];
  const masterOfferUpdates = [];
  let isMasterStatusChanged = false;
  let isMasterCompChanged = false;
  let isMasterOfferChanged = false;

  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];
    const dataMap = {};
    for (let c = 0; c < masterHeaders.length; c++) {
      let val = row[c];
      if (val instanceof Date) val = Utilities.formatDate(val, "JST", "yyyy/MM/dd");
      dataMap[normalize_(masterHeaders[c])] = val;
    }

    const candidateId = String(dataMap[normalize_('登録者ID')] || '').trim().toUpperCase();
    const currentMasterStatus = masterStatusColIdx !== -1 ? String(row[masterStatusColIdx] || '').trim() : '';
    const currentMasterComp = masterCompColIdx !== -1 ? String(row[masterCompColIdx] || '').trim() : '';
    const currentMasterOfferRaw = masterOfferColIdx !== -1 ? row[masterOfferColIdx] : '';
    const currentMasterOfferStr = toDateStr(currentMasterOfferRaw);

    if (!candidateId) {
      if (masterStatusColIdx !== -1) masterStatusUpdates.push([currentMasterStatus]);
      if (masterCompColIdx !== -1) masterCompUpdates.push([currentMasterComp]);
      if (masterOfferColIdx !== -1) masterOfferUpdates.push([currentMasterOfferRaw]);
      continue;
    }

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
      dataMap[normalize_('ステータス')] = '採用';
      dataMap[normalize_('内定日')] = hireInfo.offerDateStr;

      if (masterStatusColIdx !== -1) {
        if (currentMasterStatus !== '採用') {
          masterStatusUpdates.push(['採用']);
          isMasterStatusChanged = true;
        } else {
          masterStatusUpdates.push([currentMasterStatus]);
        }
      }
      if (masterCompColIdx !== -1) {
        if (currentMasterComp !== hireInfo.company) {
          masterCompUpdates.push([hireInfo.company]);
          isMasterCompChanged = true;
        } else {
          masterCompUpdates.push([currentMasterComp]);
        }
      }
      if (masterOfferColIdx !== -1) {
        if (currentMasterOfferStr !== hireInfo.offerDateStr) {
          masterOfferUpdates.push([hireInfo.offerDateObj]);
          isMasterOfferChanged = true;
        } else {
          masterOfferUpdates.push([currentMasterOfferRaw]);
        }
      }

      hiredData.push(buildRowByHeaders_(hiredHeaders, dataMap));
    } else {
      // 採用されていない候補者のステータス判定（「辞退」「不採用」はそのまま維持、空欄や旧「採用」は「未採用」へ）
      let targetStatus = currentMasterStatus;
      if (currentMasterStatus === '辞退' || declinedCandidatesSet.has(candidateId)) {
        targetStatus = '辞退';
      } else if (currentMasterStatus === '不採用') {
        targetStatus = '不採用';
      } else {
        targetStatus = '未採用';
      }
      dataMap[normalize_('ステータス')] = targetStatus;

      if (masterStatusColIdx !== -1) {
        if (currentMasterStatus !== targetStatus) {
          masterStatusUpdates.push([targetStatus]);
          isMasterStatusChanged = true;
        } else {
          masterStatusUpdates.push([currentMasterStatus]);
        }
      }
      if (masterCompColIdx !== -1) {
        if (currentMasterComp !== '') {
          masterCompUpdates.push(['']);
          isMasterCompChanged = true;
        } else {
          masterCompUpdates.push(['']);
        }
      }
      if (masterOfferColIdx !== -1) {
        if (currentMasterOfferStr !== '') {
          masterOfferUpdates.push(['']);
          isMasterOfferChanged = true;
        } else {
          masterOfferUpdates.push([currentMasterOfferRaw]);
        }
      }

      if (targetStatus === '未採用') {
        unhiredData.push(buildRowByHeaders_(unhiredHeaders, dataMap));
      }
    }
  }

  // 登録者マスタの「ステータス」「採用事業者」「内定日」列に差分があった場合、一括更新
  if (masterStatusColIdx !== -1 && isMasterStatusChanged && masterStatusUpdates.length > 0) {
    masterSheet.getRange(2, masterStatusColIdx + 1, masterStatusUpdates.length, 1).setValues(masterStatusUpdates);
  }
  if (masterCompColIdx !== -1 && isMasterCompChanged && masterCompUpdates.length > 0) {
    masterSheet.getRange(2, masterCompColIdx + 1, masterCompUpdates.length, 1).setValues(masterCompUpdates);
  }
  if (masterOfferColIdx !== -1 && isMasterOfferChanged && masterOfferUpdates.length > 0) {
    const offerRange = masterSheet.getRange(2, masterOfferColIdx + 1, masterOfferUpdates.length, 1);
    offerRange.setValues(masterOfferUpdates);
    offerRange.setNumberFormat('yyyy"年"m"月"d"日"');
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

// =========================================
// 「案件管理」を正本とした登録者マスタの一括再構築（面接履歴・ステータス・名前表記統一・メモ分離）
// =========================================
function rebuildMasterFromJobSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('登録者マスタ');
  const jobSheet = ss.getSheetByName('案件管理');
  if (!masterSheet || !jobSheet) throw new Error('「登録者マスタ」または「案件管理」シートが見つかりません。');

  // 1. 候補者名の表記統一・(2)除去、および名前欄の改行メモの「備考・メモ」列への移動
  const nameFixMap = {
    'SD-0142': 'THANDAR LIN',
    'SD-0966': 'HNAUNG THARAPHI OO',
    'SD-1212': 'EI SHWE SIN',
    'SD-1248': 'SAUNG HNIN PHYU',
    'SD-1251': 'AYE SU MOE',
    'SD-1269': 'NYEIN NYEIN AYE',
    'SD-1270': 'HNIN EI PHYU',
    'SD-1281': 'AYE AYE NYEIN',
    'SD-1285': 'MAY ZIN PHYO'
  };

  const mCol = getMasterColumnMap(masterSheet);
  const mData = masterSheet.getDataRange().getValues();
  const idIdx = mCol['登録者ID'] - 1;
  const nameIdx = mCol['名前'] - 1;
  const memoIdx = mCol['備考・メモ'] ? mCol['備考・メモ'] - 1 : -1;

  let memoMovedCount = 0;

  for (let i = 1; i < mData.length; i++) {
    const cid = String(mData[i][idIdx] || '').trim().toUpperCase();
    if (!cid) continue;

    let rawName = String(mData[i][nameIdx] || '');

    // ① 名前セルに改行が含まれている場合：1行目を名前、2行目以降を「備考・メモ」へ移動
    if (rawName.includes('\n') || rawName.includes('\r')) {
      const parts = rawName.split(/\r?\n/).map(s => s.trim()).filter(s => s !== '');
      if (parts.length >= 2) {
        const cleanName = parts[0];
        const extractedNote = parts.slice(1).join('\n');

        masterSheet.getRange(i + 1, nameIdx + 1).setValue(cleanName);
        mData[i][nameIdx] = cleanName;
        rawName = cleanName;

        if (memoIdx !== -1) {
          const currentMemo = String(mData[i][memoIdx] || '').trim();
          let newMemo = currentMemo;
          if (!currentMemo) {
            newMemo = extractedNote;
          } else if (!currentMemo.includes(extractedNote)) {
            newMemo = currentMemo + '\n' + extractedNote;
          }
          if (newMemo !== currentMemo) {
            masterSheet.getRange(i + 1, memoIdx + 1).setValue(newMemo);
            mData[i][memoIdx] = newMemo;
          }
        }
        updateJobCandidateName(cid, cleanName);
        memoMovedCount++;
      }
    }

    // ② 個別の名前表記統一・(2)除去
    if (nameFixMap[cid]) {
      const targetName = nameFixMap[cid];
      if (String(mData[i][nameIdx] || '').trim() !== targetName) {
        masterSheet.getRange(i + 1, nameIdx + 1).setValue(targetName);
        mData[i][nameIdx] = targetName;
      }
      updateJobCandidateName(cid, targetName);
    }
  }

  // 2. 案件管理データの読み込みと各候補者の面接履歴・採用結果の集計
  const jobData = jobSheet.getDataRange().getValues();
  const jCol = getMasterColumnMap(jobSheet);
  const jIdIdx = jCol['案件ID'] - 1;
  const jCompIdx = jCol['事業者名'] - 1;
  const jDateIdx = jCol['面接日'] - 1;
  const jDetailsIdx = jCol['面接結果詳細'] - 1;

  const toSlashDate = (val) => {
    if (!val) return '日付不明';
    if (val instanceof Date) return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
    const s = String(val).trim().replace(/[年月\-]/g, '/').replace(/日/g, '');
    const parts = s.split('/');
    if (parts.length === 3) {
      return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
    }
    return s || '日付不明';
  };

  const knownStatuses = ['採用', '不採用', '内定辞退（候補者都合）', '内定取消（事業者都合）', '面接結果の取消'];
  const historyMap = new Map(); // cid -> Array of { sortDate, jobId, line }

  for (let i = 1; i < jobData.length; i++) {
    const row = jobData[i];
    const jobId = String(row[jIdIdx] || '').trim();
    const detailsText = String(row[jDetailsIdx] || '').trim();
    if (!jobId || !detailsText) continue;

    const compText = String(row[jCompIdx] || '').trim();
    const companies = compText.split(/\r?\n/).map(c => c.trim()).filter(c => c);
    const defaultJoinedCompany = companies.join('・');
    let currentCompany = companies.length > 0 ? companies[0] : '';

    const formattedDate = toSlashDate(row[jDateIdx]);
    const sortDateKey = (formattedDate === '日付不明') ? '0000/00/00' : formattedDate;

    const lines = detailsText.split(/\r?\n/).map(l => l.trim()).filter(l => l);
    for (const line of lines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentCompany = line.slice(1, -1).trim();
        continue;
      }

      const idMatch = line.match(/^(SD-\d+)/i);
      if (!idMatch) continue;
      const cid = idMatch[1].toUpperCase();

      let parsedResult = null;
      for (const st of knownStatuses) {
        if (line.endsWith(`（${st}）`)) {
          parsedResult = st;
          break;
        }
      }
      if (!parsedResult || parsedResult === '面接結果の取消') continue;

      let histComp = currentCompany;
      let suffix = '';
      if (parsedResult === '採用') {
        histComp = currentCompany;
        suffix = '（採用）';
      } else if (parsedResult === '不採用') {
        histComp = defaultJoinedCompany;
        suffix = '（不採用）';
      } else if (parsedResult === '内定辞退（候補者都合）') {
        histComp = currentCompany;
        suffix = '（内定辞退）';
      } else if (parsedResult === '内定取消（事業者都合）') {
        histComp = currentCompany;
        suffix = '（事業者都合取消）';
      }

      const histLine = `${formattedDate}：${histComp}${suffix}`;
      if (!historyMap.has(cid)) historyMap.set(cid, []);
      historyMap.get(cid).push({
        sortDate: sortDateKey,
        jobId: jobId,
        line: histLine
      });
    }
  }

  // 3. 登録者マスタの「面接履歴」列を一括更新
  const historyColIdx = mCol['面接履歴'] - 1;
  const historyUpdates = [];
  let historyUpdatedCount = 0;

  for (let i = 1; i < mData.length; i++) {
    const cid = String(mData[i][idIdx] || '').trim().toUpperCase();
    const oldHist = String(mData[i][historyColIdx] || '').trim();

    let newHist = '';
    if (cid && historyMap.has(cid)) {
      const items = historyMap.get(cid);
      items.sort((a, b) => {
        if (a.sortDate !== b.sortDate) return a.sortDate.localeCompare(b.sortDate);
        return a.jobId.localeCompare(b.jobId);
      });
      // 完全一致の重複行を除外しつつ結合
      const uniqueLines = [...new Set(items.map(item => item.line))];
      newHist = uniqueLines.join('\n');
    }

    if (oldHist !== newHist) {
      historyUpdatedCount++;
    }
    historyUpdates.push([newHist]);
  }

  if (historyUpdates.length > 0) {
    masterSheet.getRange(2, historyColIdx + 1, historyUpdates.length, 1).setValues(historyUpdates);
  }

  // 4. ステータス・採用事業者・内定日の一括同期および採用者一覧・未採用者一覧の更新
  updateCandidateLists(true);

  SpreadsheetApp.getUi().alert(
    `案件管理からの一括反映が完了しました。\n` +
    `・名前欄からの日本語メモ分離・備考移動：${memoMovedCount} 名\n` +
    `・面接履歴の再構築・更新：${historyUpdatedCount} 名\n` +
    `・ステータス／採用事業者／内定日／各種リストの同期完了`
  );
}