// =========================================================
// 外部連携（Funtoco「COE info」シートへの同期・転記）
// =========================================================

const EXTERNAL_SHEET_COE_INFO = "COE info";

/**
 * COE info シートの標準30列ヘッダー定義（A列：内定日、B列：案件ID、C列：採用事業者名、D列：登録者ID）
 */
function getCoeInfoHeaders_() {
  return [
    "内定日\nOffer Date",                                                                                           // 1 (A)
    "案件ID\nJob ID",                                                                                               // 2 (B)
    "採用事業者名\nCompany",                                                                                        // 3 (C)
    "登録者ID\nCandidate ID",                                                                                       // 4 (D)
    "Name\n名前",                                                                                                   // 5 (E)
    "Nick Name\n呼び名",                                                                                            // 6 (F)
    "Sex\n性別",                                                                                                    // 7 (G)
    "Assignment\n所属送り出し機関",                                                                                 // 8 (H)
    "Date of birth\n生年月日",                                                                                      // 9 (I)
    "Place of birth\n出生地（都市名）",                                                                             // 10 (J)
    "Passport Number\nパスポート番号",                                                                              // 11 (K)
    "Passport Expire Date\nパスポート有効期限",                                                                     // 12 (L)
    "Certification of JFT or JLPT\n日本語試験",                                                                     // 13 (M)
    "Certification of Nurse Skill\n介護技能評価試験",                                                               // 14 (N)
    "Certification of Nurse Japanese\n介護日本語評価試験",                                                          // 15 (O)
    "Married or Single\n配偶者",                                                                                    // 16 (P)
    "Address of Myanmar\n住所詳細",                                                                                 // 17 (Q)
    "Occupation\n職業",                                                                                             // 18 (R)
    "CV Form Detail / Final educational background\n最終学歴・状況",                                                // 19 (S)
    "CV Form Detail / Final educational school name・Year/Month・Field\n最終学歴学校名・年月・分野",                 // 20 (T)
    "Have you applied for a status of residence in the past? (If yes, how many times?)\n過去の交付申請歴・回数",    // 21 (U)
    "Number of times of non-delivery\n不許可回数",                                                                  // 22 (V)
    "Past entry into / departure from Japan\n過去の出入国歴",                                                       // 23 (W)
    "If yes: How many times and The latest entry from - to -\n出入国回数・直近期間",                                // 24 (X)
    "Family in Japan\n日本在住の親族の有無",                                                                        // 25 (Y)
    "If yes: Relationship, Name, Date of birth, Nationality, Place of employment/school, Residence card number\n在日親族詳細", // 26 (Z)
    "Criminal record (in Japan / overseas)\n犯罪歴の有無",                                                          // 27 (AA)
    "Experience of Ginojissyu\n技能実習経験の有無",                                                                 // 28 (AB)
    "If yes: Certification of Ginojissyu\n技能実習修了書の有無",                                                    // 29 (AC)
    "Note\n備考"                                                                                                    // 30 (AD)
  ];
}

/**
 * ヘッダー文字列から標準フィールドキーを判定するヘルパー関数
 * ※ Z列（"...Relationship, Name, Date of birth...") が E列（Name）や I列（Date of birth）に誤判定されないよう、
 *   長文・固有キーワードを先に判定する順序としています。
 */
function identifyCoeFieldKey_(headerStr) {
  const h = String(headerStr || "").normalize('NFC').replace(/\s+/g, ' ').trim();
  if (!h) return null;

  // 1. 複合文・固有キーワードの優先判定（Z列などが Name / Date of birth に誤爆するのを防止）
  if (h.includes('Residence card number') || h.includes('在日親族詳細')) return 'familyDetail';
  if (h.includes('Family in Japan') || h.includes('日本在住の親族の有無')) return 'hasFamily';
  if (h.includes('Final educational school name') || h.includes('最終学歴学校名')) return 'eduDetail';
  if (h.includes('Final educational background') || h.includes('最終学歴・状況')) return 'eduBgSummary';
  if (h.includes('applied for a status of residence') || h.includes('過去の交付申請歴')) return 'applyStr';
  if (h.includes('non-delivery') || h.includes('不許可回数')) return 'rejectStr';
  if (h.includes('The latest entry from') || h.includes('出入国回数・直近期間')) return 'travelDetail';
  if (h.includes('Past entry into') || h.includes('過去の出入国歴')) return 'overseasExp';
  if (h.includes('Criminal record') || h.includes('犯罪歴の有無')) return 'crime';
  if (h.includes('Certification of Gino') || h.includes('技能実習修了書の有無')) return 'traineeCert';
  if (h.includes('Experience of Gino') || h.includes('技能実習経験の有無')) return 'traineeExp';
  if (h.includes('Certification of JFT or JLPT') || h.includes('日本語試験')) return 'certLang';
  if (h.includes('Nurse Skill') || h.includes('介護技能評価試験')) return 'kaigoSkill';
  if (h.includes('Nurse Japanese') || h.includes('介護日本語評価試験')) return 'kaigoLang';
  if (h.includes('Passport Number') || h.includes('パスポート番号')) return 'passportNum';
  if (h.includes('Passport Exp') || h.includes('パスポート有効期限')) return 'passportExp';
  if (h.includes('Address of Myanmar') || h.includes('Adress of Myanmar') || h.includes('住所詳細')) return 'addressDetail';
  if (h.includes('Place of birth') || h.includes('出生地')) return 'birthCity';
  if (h.includes('Date of birth') || h.includes('生年月日')) return 'birthday';
  if (h.includes('Married or Single') || h.includes('配偶者')) return 'spouse';
  if (h.includes('Occupation') || h.includes('職業')) return 'job';
  if (h.includes('Assignment') || h.includes('送り出し機関')) return 'agent';
  if (h.includes('Sex') || h.includes('性別')) return 'sex';
  if (h.includes('Nick Name') || h.includes('呼び名')) return 'nickname';

  // 2. 管理キー・基本項目の判定
  if (h.includes('内定日') || h.includes('Offer Date')) return 'offerDate';
  if (h.includes('案件ID') || h.includes('Job ID')) return 'jobId';
  if (h.includes('採用事業者') || h.includes('Company')) return 'company';
  if (h.includes('登録者ID') || h.includes('Candidate ID')) return 'candId';
  if (h.includes('Name') || h.includes('名前')) return 'name';
  if (h.includes('Note') || h.includes('備考')) return 'note';
  return null;
}

/**
 * COE info シートの1行目ヘッダーを読み取り、各列とフィールドキーの対応マップを返す関数
 */
function getCoeSheetLayout_(coeSheet) {
  ensureCoeInfoSheetInitialized_(coeSheet);
  const lastCol = coeSheet.getLastColumn();
  const headers = coeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const colKeys = headers.map(h => identifyCoeFieldKey_(h));
  const candIdColIdx = colKeys.indexOf('candId'); // 0-based
  const jobIdColIdx = colKeys.indexOf('jobId');   // 0-based
  return {
    lastCol: lastCol,
    colKeys: colKeys,
    candIdColIdx: candIdColIdx !== -1 ? candIdColIdx : 3,
    jobIdColIdx: jobIdColIdx !== -1 ? jobIdColIdx : 1
  };
}

/**
 * COE info シートが空の場合にヘッダーと書式を初期構築するヘルパー関数
 */
function ensureCoeInfoSheetInitialized_(coeSheet) {
  const headers = getCoeInfoHeaders_();
  const lastCol = coeSheet.getLastColumn();
  let needsHeader = false;

  if (lastCol === 0 || coeSheet.getLastRow() === 0) {
    needsHeader = true;
  } else {
    const firstRow = coeSheet.getRange(1, 1, 1, Math.min(lastCol, 5)).getValues()[0];
    if (firstRow.every(v => String(v || "").trim() === "")) {
      needsHeader = true;
    }
  }

  if (needsHeader) {
    const headerRange = coeSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#d2e3fc')
               .setFontColor('#174ea6')
               .setFontWeight('bold')
               .setFontSize(10)
               .setVerticalAlignment('middle')
               .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    coeSheet.setRowHeight(1, 48);
    coeSheet.setFrozenRows(1);

    const widths = [
      100, 95,  180, 95,  180, 110, 80,  150, 105, 130,
      120, 120, 120, 95,  95,  110, 260, 160, 180, 260,
      130, 100, 100, 200, 100, 280, 100, 100, 100, 220
    ];
    widths.forEach((w, idx) => coeSheet.setColumnWidth(idx + 1, w));
  }
}

/**
 * 年月文字列（例: "2020年3月" や "2020-03"）を英語表記（例: "2020 MARCH"）に変換するヘルパー関数
 */
function formatYearMonthEnglish_(val) {
  if (!val) return "";
  const s = String(val).replace(/^'/, '').trim();
  const m = s.match(/(\d{4})[-\/年](\d{1,2})/);
  if (!m) return s;
  const year = m[1];
  const monthIdx = parseInt(m[2], 10) - 1;
  const monthNames = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"
  ];
  if (monthIdx >= 0 && monthIdx < 12) {
    return `${year} ${monthNames[monthIdx]}`;
  }
  return s;
}

/**
 * 登録者マスタの1行データと案件情報から COE info シート用の行配列を生成する関数
 * ※ 未入力項目に対する他列からの勝手な自動補完は一切行わず、空欄は空欄のまま出力します。
 */
function buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys) {
  const getVal = (colName) => {
    const clean = colName.replace(/\s/g, '');
    const cIdx = mCol[clean];
    if (!cIdx) return "";
    const val = mRow[cIdx - 1];
    if (val === null || val === undefined) return "";
    if (val instanceof Date) return Utilities.formatDate(val, sourceTZ, "yyyy/MM/dd");
    const s = String(val).trim();
    return s.startsWith("'") ? s.substring(1).trim() : s;
  };

  const formatSlashDate = (val) => {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, sourceTZ, "yyyy/MM/dd");
    const s = String(val).trim().replace(/[年月\-]/g, '/').replace(/日/g, '');
    if (s === "NaN" || s === "undefined" || s === "Invalid Date") return "";
    const parts = s.split('/');
    if (parts.length === 3) {
      return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
    }
    return s;
  };

  const candId = getVal('登録者ID').toUpperCase();
  const name = getVal('名前');
  const nickname = getVal('呼び名');

  // 性別 (Male / Female)
  const rawGender = getVal('性別');
  let sexStr = rawGender;
  if (rawGender === '男性') sexStr = 'Male';
  else if (rawGender === '女性') sexStr = 'Female';

  const agent = getVal('所属送り出し機関');
  const birthday = formatSlashDate(getVal('生年月日'));

  // J列（出生地）と Q列（住所詳細）：空欄なら空欄のまま出力
  const birthCity = getVal('出生地（都市名）');
  const addressDetail = getVal('住所詳細');

  const passportNum = getVal('パスポート番号');
  const passportExp = formatSlashDate(getVal('パスポート有効期限'));

  // 日本語資格 (Certification of JFT or JLPT)
  const jlpt = getVal('特定技能要件＞JLPTレベル');
  const jft = getVal('特定技能要件＞JFTBasicレベル');
  let certLangParts = [];
  if (jlpt && jlpt !== '-' && jlpt !== '×') certLangParts.push(`JLPT ${jlpt}`);
  if (jft && jft !== '-' && jft !== '×') certLangParts.push(jft === 'A2' ? 'JFT A2' : `JFT ${jft}`);
  const certLangStr = certLangParts.join(' / ');

  // 介護技能・介護日本語 (YES / Scheduled)
  const formatExam = (val) => {
    if (!val || val === '-' || val === '×') return '';
    if (val === '合格') return 'YES';
    if (val.includes('予定')) return 'Scheduled (受験予定)';
    return val;
  };
  const kaigoSkillStr = formatExam(getVal('特定技能要件＞介護技能評価試験'));
  const kaigoLangStr = formatExam(getVal('特定技能要件＞介護日本語評価試験'));

  // 配偶者 (Single(未婚) / Married(既婚))
  const rawSpouse = getVal('配偶者');
  let spouseStr = rawSpouse;
  if (rawSpouse === '無') spouseStr = 'Single(未婚)';
  else if (rawSpouse === '有') spouseStr = 'Married(既婚)';

  // 職業 (Occupation)
  const rawJob = getVal('職業');
  let jobStr = rawJob;
  if (rawJob === '会社員') jobStr = 'company employee（会社員)';
  else if (rawJob === 'アルバイト') jobStr = 'part-time（アルバイト）';
  else if (rawJob === '無職') jobStr = 'unemployed(無職）';

  // 最終学歴・状況 (S列・T列)
  // ※ T列は COE用の「学歴＞学校名（英語）」「学歴＞学部・学科・専攻（英語）」が入力されている場合のみ出力し、未入力なら空欄とする
  const eduSchoolEn = getVal('学歴＞学校名（英語）');
  const eduDeptEn = getVal('学歴＞学部・学科・専攻（英語）');
  const eduStatus = getVal('学歴＞状況');
  const eduEndRaw = getVal('学歴＞卒業/中退年月');
  const eduEndEn = formatYearMonthEnglish_(eduEndRaw);

  let eduBgSummary = '';
  let eduDetailStr = '';

  if (eduSchoolEn || eduDeptEn) {
    const upperSchool = eduSchoolEn.toUpperCase();
    let schoolType = '';
    if (upperSchool.includes('UNIVERSITY')) schoolType = 'UNIVERSITY';
    else if (upperSchool.includes('HIGH')) schoolType = 'HIGH SCHOOL';
    else if (upperSchool.includes('COLLEGE')) schoolType = 'COLLEGE';

    let statusEn = eduStatus;
    if (eduStatus === '卒業') statusEn = 'PASSED (卒業)';
    else if (eduStatus === '中途退学') statusEn = 'NOT PASSED (中途退学)';
    else if (eduStatus === '在学中') statusEn = 'ENROLLED (在学中)';
    else if (eduStatus === '休学中') statusEn = 'ON LEAVE (休学中)';

    eduBgSummary = schoolType ? `${schoolType} ${statusEn}`.trim() : statusEn;

    const eduDetailParts = [eduSchoolEn, eduEndEn, eduDeptEn].filter(p => p !== '');
    eduDetailStr = eduDetailParts.join(', ');
  }

  // 交付申請回数・不許可回数 (U列・V列)
  const applyCountRaw = getVal('在留資格交付申請の回数');
  let applyStr = '';
  if (applyCountRaw !== '') {
    const num = Number(applyCountRaw);
    if (num === 0 || applyCountRaw === '無' || applyCountRaw === 'なし') {
      applyStr = 'No (無)';
    } else if (!isNaN(num) && num > 0) {
      applyStr = `YES (${num} TIME${num > 1 ? 'S' : ''})`;
    } else {
      applyStr = applyCountRaw;
    }
  }

  const rejectCountRaw = getVal('不許可となった在留資格交付申請の回数');
  let rejectStr = '';
  if (rejectCountRaw !== '') {
    const num = Number(rejectCountRaw);
    if (num === 0 || rejectCountRaw === '無' || rejectCountRaw === 'なし') {
      rejectStr = 'No (無)';
    } else {
      rejectStr = rejectCountRaw;
    }
  }

  // 有 / 無 変換ヘルパー
  const toYesNo = (val) => {
    if (val === '有') return 'YES (有)';
    if (val === '無') return 'No (無)';
    return val;
  };

  // 海外出入国歴・期間 (W列・X列)
  const overseasExp = getVal('海外への出入国歴の有無');
  const overseasExpStr = toYesNo(overseasExp);
  const travelCount = getVal('出入国の回数');
  const lastInDate = formatSlashDate(getVal('直近の入国日'));
  const lastOutDate = formatSlashDate(getVal('直近の出国日'));

  let travelDetailParts = [];
  if (travelCount && travelCount !== '0') travelDetailParts.push(`${travelCount} Time(s)`);
  if (lastInDate || lastOutDate) travelDetailParts.push(`${lastInDate} 〜 ${lastOutDate}`);
  const travelDetailStr = (overseasExp === '無') ? '' : travelDetailParts.join('\n');

  // 日本在住の親族情報 (Y列・Z列：COE用の親族情報列のみを参照し、未入力なら空欄)
  const relName = getVal('日本在住の親族情報親族の名前');
  const relRelation = getVal('日本在住の親族情報続柄');
  const relBirth = formatSlashDate(getVal('日本在住の親族情報親族の生年月日'));
  const relCountry = getVal('日本在住の親族情報親族の国籍・地域');
  const relLive = getVal('日本在住の親族情報親族との同居予定の有無');
  const relWork = getVal('日本在住の親族情報親族の勤務先・通学先');
  const relCard = getVal('日本在住の親族情報親族の在留カード番号');

  let hasFamilyStr = '';
  let familyDetailStr = '';
  const hasDetailedRel = [relName, relRelation, relBirth, relCountry, relLive, relWork, relCard].some(v => v !== '');

  if (hasDetailedRel) {
    hasFamilyStr = 'YES (有)';
    const fParts = [];
    if (relRelation) fParts.push(`続柄: ${relRelation}`);
    if (relName) fParts.push(`氏名: ${relName}`);
    if (relBirth) fParts.push(`生年月日: ${relBirth}`);
    if (relCountry) fParts.push(`国籍: ${relCountry}`);
    if (relLive) fParts.push(`同居予定: ${relLive}`);
    if (relWork) fParts.push(`勤務先/通学先: ${relWork}`);
    if (relCard) fParts.push(`在留カード: ${relCard}`);
    familyDetailStr = fParts.join(', ');
  }

  // 犯罪歴・技能実習経験・修了書 (AA〜AC列)
  const crimeStr = toYesNo(getVal('犯罪歴の有無'));
  const traineeExpStr = toYesNo(getVal('技能実習の経験の有無'));
  const traineeCertStr = toYesNo(getVal('技能実習修了書の有無'));

  // 備考 (AD列：その他の評価試験・その他の日本語能力試験・備考メモを結合)
  const otherExam = getVal('特定技能要件＞その他の評価試験');
  const otherJapanese = getVal('その他の日本語能力試験');
  const memo = getVal('備考・メモ');
  const noteParts = [];
  if (otherExam && otherExam !== '-' && otherExam !== '×') noteParts.push(otherExam);
  if (otherJapanese && otherJapanese !== '-' && otherJapanese !== '×') noteParts.push(otherJapanese);
  if (memo) noteParts.push(memo);
  const noteStr = noteParts.join(' / ');

  const jobId = jobInfo ? (jobInfo.jobId || '') : '';
  const company = jobInfo ? (jobInfo.company || getVal('採用事業者')) : getVal('採用事業者');
  const offerDate = formatSlashDate((jobInfo && jobInfo.offerDate) ? jobInfo.offerDate : getVal('内定日'));

  const fieldMap = {
    offerDate: offerDate,
    jobId: jobId,
    company: company,
    candId: candId,
    name: name,
    nickname: nickname,
    sex: sexStr,
    agent: agent,
    birthday: birthday,
    birthCity: birthCity,
    passportNum: passportNum,
    passportExp: passportExp,
    certLang: certLangStr,
    kaigoSkill: kaigoSkillStr,
    kaigoLang: kaigoLangStr,
    spouse: spouseStr,
    addressDetail: addressDetail,
    job: jobStr,
    eduBgSummary: eduBgSummary,
    eduDetail: eduDetailStr,
    applyStr: applyStr,
    rejectStr: rejectStr,
    overseasExp: overseasExpStr,
    travelDetail: travelDetailStr,
    hasFamily: hasFamilyStr,
    familyDetail: familyDetailStr,
    crime: crimeStr,
    traineeExp: traineeExpStr,
    traineeCert: traineeCertStr,
    note: noteStr
  };

  return colKeys.map(k => (k && fieldMap[k] !== undefined) ? fieldMap[k] : "");
}

/**
 * 案件管理シートから現在「採用」となっている全候補者の案件情報をマップで取得するヘルパー関数
 * @return {Map} key: SD-XXXX -> { jobId, company, offerDate }
 */
function getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ) {
  const jobData = jobSheet.getDataRange().getValues();
  const jCol = getMasterColumnMap(jobSheet);
  const hiredMap = new Map();

  if (jobData.length < 2 || !jCol['案件ID'] || !jCol['面接結果詳細']) return hiredMap;

  const jIdIdx = jCol['案件ID'] - 1;
  const jCompIdx = jCol['事業者名'] ? jCol['事業者名'] - 1 : -1;
  const jOfferIdx = jCol['内定日'] ? jCol['内定日'] - 1 : -1;
  const jDetailsIdx = jCol['面接結果詳細'] - 1;

  const formatSlashDate = (val) => {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, sourceTZ, "yyyy/MM/dd");
    const s = String(val).trim().replace(/[年月\-]/g, '/').replace(/日/g, '');
    const parts = s.split('/');
    if (parts.length === 3) {
      return `${parts[0]}/${parts[1].padStart(2, '0')}/${parts[2].padStart(2, '0')}`;
    }
    return s;
  };

  for (let i = 1; i < jobData.length; i++) {
    const row = jobData[i];
    const jobId = String(row[jIdIdx] || "").trim();
    const detailsText = String(row[jDetailsIdx] || "").trim();
    if (!jobId || !detailsText) continue;

    const compCell = jCompIdx !== -1 ? String(row[jCompIdx] || "").normalize('NFC') : "";
    const defaultComp = compCell.split(/\r?\n/)[0].trim();
    let currentComp = defaultComp;
    const offerDateStr = jOfferIdx !== -1 ? formatSlashDate(row[jOfferIdx]) : "";

    const lines = detailsText.split(/\r?\n/).map(l => l.trim()).filter(l => l !== "");
    for (const line of lines) {
      if (line.startsWith('【') && line.endsWith('】')) {
        currentComp = line.slice(1, -1).normalize('NFC').trim();
        continue;
      }
      if (line.includes('（採用）')) {
        const m = line.match(/^(SD-\d+)/i);
        if (m) {
          const cid = m[1].toUpperCase();
          hiredMap.set(cid, {
            jobId: jobId,
            company: currentComp,
            offerDate: offerDateStr
          });
        }
      }
    }
  }
  return hiredMap;
}

/**
 * 特定の候補者（複数可）を外部シート「COE info」へ即時同期（Upsert）する関数
 * ・スマートテーブル（Tables）構造および動的ヘッダー順に対応
 */
function syncCandidatesToCoeInfo(candidateIds, removeCandidateIds = []) {
  try {
    const targetIds = (Array.isArray(candidateIds) ? candidateIds : [candidateIds])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(id => id !== "");
    const removeIds = (Array.isArray(removeCandidateIds) ? removeCandidateIds : [removeCandidateIds])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(id => id !== "");

    if (targetIds.length === 0 && removeIds.length === 0) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const masterSheet = ss.getSheetByName('登録者マスタ');
    const jobSheet = ss.getSheetByName('案件管理');
    if (!masterSheet || !jobSheet) return;

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const coeSheet = targetSS.getSheetByName(EXTERNAL_SHEET_COE_INFO);
    if (!coeSheet) {
      console.warn(`外部シートに「${EXTERNAL_SHEET_COE_INFO}」が見つかりません。`);
      return;
    }

    const layout = getCoeSheetLayout_(coeSheet);
    const { lastCol, colKeys, candIdColIdx, jobIdColIdx } = layout;

    const hiredJobMap = getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ);
    const masterData = masterSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(masterSheet);
    const idIdx = mCol['登録者ID'] - 1;

    const masterRowMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      const cid = String(masterData[i][idIdx] || "").trim().toUpperCase();
      if (cid) masterRowMap.set(cid, masterData[i]);
    }

    // 1. 取消・辞退となった候補者の行があれば先に削除（降順）
    if (removeIds.length > 0 && coeSheet.getLastRow() >= 2) {
      const removeSet = new Set(removeIds);
      const coeData = coeSheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let r = 1; r < coeData.length; r++) {
        const existingCid = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
        if (existingCid && removeSet.has(existingCid) && !hiredJobMap.has(existingCid)) {
          rowsToDelete.push(r + 1);
        }
      }
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach(rowNum => coeSheet.deleteRow(rowNum));
      if (rowsToDelete.length > 0) SpreadsheetApp.flush();
    }

    if (targetIds.length === 0) return;

    // 2. 現在の COE info シートの既存行と最終データ行（スマートテーブルのフッター直前）を特定
    const existingRowMap = new Map(); // cid -> 1-based row number
    let lastValidRowIndex = 1;
    const coeData = coeSheet.getDataRange().getValues();

    for (let r = 1; r < coeData.length; r++) {
      const jId = String(coeData[r][jobIdColIdx] || "").trim();
      const cId = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
      if (jId || cId) {
        lastValidRowIndex = r + 1;
      }
      if (cId && !existingRowMap.has(cId)) {
        existingRowMap.set(cId, r + 1);
      }
    }

    const rowsToAppend = [];

    for (const cid of targetIds) {
      const mRow = masterRowMap.get(cid);
      if (!mRow) continue;

      const jobInfo = hiredJobMap.get(cid) || null;
      const rowArray = buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys);

      if (existingRowMap.has(cid)) {
        const targetRowNum = existingRowMap.get(cid);
        coeSheet.getRange(targetRowNum, 1, 1, lastCol).setValues([rowArray]);
      } else {
        rowsToAppend.push(rowArray);
      }
    }

    if (rowsToAppend.length > 0) {
      const insertRowIndex = lastValidRowIndex + 1;
      if (insertRowIndex <= coeSheet.getMaxRows()) {
        coeSheet.insertRowsBefore(insertRowIndex, rowsToAppend.length);
      } else {
        coeSheet.insertRowsAfter(coeSheet.getMaxRows(), rowsToAppend.length);
      }
      coeSheet.getRange(insertRowIndex, 1, rowsToAppend.length, lastCol).setValues(rowsToAppend);
    }

    SpreadsheetApp.flush();
  } catch (e) {
    console.error("COE info シートへの自動同期エラー: " + e.message);
  }
}

/**
 * 管理者メニュー用：全採用者のCOE情報を外部シート「COE info」へ一括同期する関数
 * ・スマートテーブル（Tables）および動的ヘッダー順に対応
 */
function syncAllToCoeInfoSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const masterSheet = ss.getSheetByName('登録者マスタ');
    const jobSheet = ss.getSheetByName('案件管理');
    if (!masterSheet || !jobSheet) throw new Error("「登録者マスタ」または「案件管理」シートが見つかりません。");

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const coeSheet = targetSS.getSheetByName(EXTERNAL_SHEET_COE_INFO);
    if (!coeSheet) throw new Error(`外部シートに「${EXTERNAL_SHEET_COE_INFO}」シートが見つかりません。`);

    const layout = getCoeSheetLayout_(coeSheet);
    const { lastCol, colKeys, candIdColIdx, jobIdColIdx } = layout;

    const hiredJobMap = getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ);
    const masterData = masterSheet.getDataRange().getValues();
    const mCol = getMasterColumnMap(masterSheet);
    const idIdx = mCol['登録者ID'] - 1;

    const masterRowMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      const cid = String(masterData[i][idIdx] || "").trim().toUpperCase();
      if (cid) masterRowMap.set(cid, masterData[i]);
    }

    // 案件ID順・登録者ID順でソートして同期対象リストを作成
    const sortedEntries = Array.from(hiredJobMap.entries()).sort((a, b) => {
      const jA = a[1].jobId || "";
      const jB = b[1].jobId || "";
      if (jA !== jB) return jA.localeCompare(jB);
      return a[0].localeCompare(b[0]);
    });

    // 1. 既存シートの行を取得し、採用取り消し・辞退となった行を先に削除（降順）
    let coeData = coeSheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let r = 1; r < coeData.length; r++) {
      const existingJid = String(coeData[r][jobIdColIdx] || "").trim();
      const existingCid = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
      if (!existingJid && !existingCid) continue; // フッターや空行はスキップ
      if (existingCid && !hiredJobMap.has(existingCid)) {
        rowsToDelete.push(r + 1);
      }
    }
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(rowNum => coeSheet.deleteRow(rowNum));
    if (rowsToDelete.length > 0) {
      SpreadsheetApp.flush();
      coeData = coeSheet.getDataRange().getValues();
    }

    // 2. 削除後の最新データで既存行マップと最終データ行（フッター直前）を特定
    const existingRowMap = new Map();
    let lastValidRowIndex = 1;
    for (let r = 1; r < coeData.length; r++) {
      const jId = String(coeData[r][jobIdColIdx] || "").trim();
      const cId = String(coeData[r][candIdColIdx] || "").trim().toUpperCase();
      if (jId || cId) {
        lastValidRowIndex = r + 1;
      }
      if (cId && !existingRowMap.has(cId)) {
        existingRowMap.set(cId, r + 1);
      }
    }

    let updateCount = 0;
    let appendCount = 0;
    const rowsToAppend = [];

    for (const [cid, jobInfo] of sortedEntries) {
      const mRow = masterRowMap.get(cid);
      if (!mRow) continue;
      const newRowArr = buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys);

      if (existingRowMap.has(cid)) {
        const rowNum = existingRowMap.get(cid);
        const oldRowArr = coeData[rowNum - 1];
        const isChanged = newRowArr.some((val, idx) => String(val || "").trim() !== String(oldRowArr[idx] || "").trim());
        if (isChanged) {
          coeSheet.getRange(rowNum, 1, 1, lastCol).setValues([newRowArr]);
          updateCount++;
        }
      } else {
        rowsToAppend.push(newRowArr);
        appendCount++;
      }
    }

    if (rowsToAppend.length > 0) {
      const insertRowIndex = lastValidRowIndex + 1;
      if (insertRowIndex <= coeSheet.getMaxRows()) {
        coeSheet.insertRowsBefore(insertRowIndex, rowsToAppend.length);
      } else {
        coeSheet.insertRowsAfter(coeSheet.getMaxRows(), rowsToAppend.length);
      }
      coeSheet.getRange(insertRowIndex, 1, rowsToAppend.length, lastCol).setValues(rowsToAppend);
    }

    SpreadsheetApp.flush();
    return `「COE info」シートへの同期が完了しました。\n新規追加: ${appendCount}件\n情報更新: ${updateCount}件\n削除: ${rowsToDelete.length}件`;
  } catch (e) {
    throw new Error("COE info 同期エラー: " + e.message);
  }
}