// =========================================
// シート間リスト同期連携（採用者一覧・未採用者一覧・登録者マスタ同期）
// =========================================

function syncListSheets() {
  updateCandidateLists(true);
  return 'リストの同期が完了しました。';
}

function normalize_(str) {
  if (str === null || str === undefined) return '';
  return String(str).normalize('NFC').replace(/[\s \n\r]+/g, '').toLowerCase();
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
    const defaultCompany = compText.split(/\r?\n/)[0].normalize('NFC').trim();
    let currentCompany = defaultCompany;

    const dLines = detailsText.split(/\r?\n/).filter(line => line.trim() !== "");
    for (const line of dLines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentCompany = line.slice(1, -1).normalize('NFC').trim();
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
    const currentMasterComp = masterCompColIdx !== -1 ? String(row[masterCompColIdx] || '').normalize('NFC').trim() : '';
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