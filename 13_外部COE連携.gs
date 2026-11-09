// =========================================================
// 外部連携（Funtoco「COE info」シートへの同期・転記）
// =========================================================

const EXTERNAL_SHEET_COE_INFO = "COE info";

/**
 * COE info シートの標準33列ヘッダー定義（1行目：日本語、2行目：英語）
 */
function getCoeInfoHeaders_() {
  return [
    "内定日\nOffer Date",                                                               // 1 (A)
    "案件ID\nJob ID",                                                                   // 2 (B)
    "採用事業者名\nAccepting Organization",                                             // 3 (C)
    "技能分野\nJob Field",                                                              // 4 (D)
    "登録者ID\nCandidate ID",                                                           // 5 (E)
    "氏名\nName",                                                                       // 6 (F)
    "呼び名\nNickname",                                                                 // 7 (G)
    "性別\nSex",                                                                        // 8 (H)
    "所属送り出し機関\nSending Organization",                                           // 9 (I)
    "生年月日\nDate of birth",                                                          // 10 (J)
    "出生地（都市名）\nPlace of birth",                                                 // 11 (K)
    "旅券番号（パスポート番号）\nPassport Number",                                      // 12 (L)
    "旅券有効期限（パスポート有効期限）\nPassport Date of expiration",                  // 13 (M)
    "日本語能力試験（JLPT / JFT）\nCertification of JLPT or JFT-Basic",                 // 14 (N)
    "介護技能評価試験\nNursing Care Skills Exam",                                       // 15 (O)
    "介護日本語評価試験\nNursing Care Japanese Exam",                                   // 16 (P)
    "外食業技能測定試験\nFood Service Skills Exam",                                     // 17 (Q)
    "飲食料品製造業技能測定試験\nFood & Beverage Mfg. Skills Exam",                     // 18 (R)
    "配偶者の有無\nMarital status",                                                     // 19 (S)
    "本国における居住地（住所詳細）\nHome address in Myanmar",                          // 20 (T)
    "職業\nOccupation",                                                                 // 21 (U)
    "最終学歴・状況\nFinal educational background",                                     // 22 (V)
    "最終学歴（学校名・卒業年月・専攻）\nName of school / Date of graduation / Major",  // 23 (W)
    "過去の交付申請歴・回数\nPast COE application history & count",                     // 24 (X)
    "うち不交付となった回数\nNumber of times of non-issuance",                          // 25 (Y)
    "過去の出入国歴\nPast entry into / departure from Japan",                           // 26 (Z)
    "出入国回数・直近の出入国期間\nNumber of times & Latest entry/departure period",    // 27 (AA)
    "在日親族及び同居者の有無\nFamily in Japan or co-residents",                        // 28 (AB)
    "在日親族・同居者の詳細\nDetails of family in Japan",                               // 29 (AC)
    "犯罪歴の有無（日本・海外）\nCriminal record (in Japan / overseas)",                // 30 (AD)
    "技能実習経験の有無\nExperience of Ginojissyu",                                     // 31 (AE)
    "技能実習修了証の有無\nCertification of Ginojissyu",                                // 32 (AF)
    "備考（その他日本語試験等）\nNote"                                                  // 33 (AG)
  ];
}

/**
 * ヘッダー文字列から標準フィールドキーを判定するヘルパー関数
 * ※ AG列「備考（その他日本語試験等）」が N列「日本語試験」に誤判定されないよう、「備考 / Note」を最優先で判定します。
 */
function identifyCoeFieldKey_(headerStr) {
  const h = String(headerStr || "").normalize('NFC').replace(/\s+/g, ' ').trim();
  if (!h) return null;

  // 0. 備考列の最優先判定（「その他日本語試験等」という副題が「日本語試験」に誤爆するのを防止）
  if (h.includes('備考') || h.includes('Note')) return 'note';

  // 1. 複合文・固有キーワードの優先判定（在日親族詳細や学校名が Name 等に誤判定されるのを防止）
  if (h.includes('在日親族・同居者の詳細') || h.includes('在日親族詳細') || h.includes('Details of family in Japan') || h.includes('Residence card number')) return 'familyDetail';
  if (h.includes('在日親族及び同居者の有無') || h.includes('日本在住の親族の有無') || h.includes('Family in Japan')) return 'hasFamily';
  if (h.includes('学校名・卒業年月・専攻') || h.includes('最終学歴学校名') || h.includes('Name of school')) return 'eduDetail';
  if (h.includes('最終学歴・状況') || h.includes('Final educational background')) return 'eduBgSummary';
  if (h.includes('過去の交付申請歴') || h.includes('Past COE application') || h.includes('applied for a status of residence')) return 'applyStr';
  if (h.includes('不交付となった回数') || h.includes('不許可回数') || h.includes('non-issuance') || h.includes('non-delivery')) return 'rejectStr';
  if (h.includes('出入国回数・直近') || h.includes('Latest entry/departure period') || h.includes('The latest entry from')) return 'travelDetail';
  if (h.includes('過去の出入国歴') || h.includes('Past entry into')) return 'overseasExp';
  if (h.includes('犯罪歴の有無') || h.includes('Criminal record')) return 'crime';
  if (h.includes('技能実習修了') || h.includes('Certification of Gino')) return 'traineeCert';
  if (h.includes('技能実習経験') || h.includes('Experience of Gino')) return 'traineeExp';
  if (h.includes('日本語能力試験') || h.includes('JLPT or JFT') || h.includes('Certification of JFT or JLPT')) return 'certLang';
  if (h.includes('介護技能評価試験') || h.includes('Nursing Care Skills') || h.includes('Nurse Skill')) return 'kaigoSkill';
  if (h.includes('介護日本語評価試験') || h.includes('Nursing Care Japanese') || h.includes('Nurse Japanese')) return 'kaigoLang';
  if (h.includes('外食業技能測定試験') || h.includes('Food Service Skills')) return 'foodServiceExam';
  if (h.includes('飲食料品製造業技能測定試験') || h.includes('Food & Beverage Mfg')) return 'foodMfgExam';
  if (h.includes('旅券番号') || h.includes('パスポート番号') || h.includes('Passport Number')) return 'passportNum';
  if (h.includes('旅券有効期限') || h.includes('パスポート有効期限') || h.includes('Passport Date of expiration') || h.includes('Passport Exp')) return 'passportExp';
  if (h.includes('本国における居住地') || h.includes('住所詳細') || h.includes('Home address in Myanmar') || h.includes('Address of Myanmar')) return 'addressDetail';
  if (h.includes('出生地') || h.includes('Place of birth')) return 'birthCity';
  if (h.includes('生年月日') || h.includes('Date of birth')) return 'birthday';
  if (h.includes('配偶者') || h.includes('Marital status') || h.includes('Married or Single')) return 'spouse';
  if (h.includes('職業') || h.includes('Occupation')) return 'job';
  if (h.includes('送り出し機関') || h.includes('Sending Organization') || h.includes('Assignment')) return 'agent';
  if (h.includes('性別') || h.includes('Sex')) return 'sex';
  if (h.includes('呼び名') || h.includes('Nickname') || h.includes('Nick Name')) return 'nickname';

  // 2. 管理キー・基本項目の判定
  if (h.includes('内定日') || h.includes('Offer Date')) return 'offerDate';
  if (h.includes('案件ID') || h.includes('Job ID')) return 'jobId';
  if (h.includes('採用事業者') || h.includes('Accepting Organization') || h.includes('Company')) return 'company';
  if (h.includes('技能分野') || h.includes('Job Field')) return 'skillField';
  if (h.includes('登録者ID') || h.includes('Candidate ID')) return 'candId';
  if (h.includes('氏名') || h.includes('名前') || h.includes('Name')) return 'name';
  return null;
}

/**
 * COE info シートの構造を確認し、旧30列から新33列への安全な列挿入およびヘッダー更新を行う関数
 */
function ensureCoeInfoSheetInitialized_(coeSheet) {
  const headers = getCoeInfoHeaders_();
  const targetColCount = headers.length; // 33列
  let lastCol = coeSheet.getLastColumn();

  if (lastCol === 0 || coeSheet.getLastRow() === 0) {
    coeSheet.getRange(1, 1, 1, targetColCount).setValues([headers]);
    return;
  }

  let currentHeaders = coeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let currentKeys = currentHeaders.map(h => identifyCoeFieldKey_(h));

  // 1. 「技能分野 (skillField)」列がまだ存在しない場合、採用事業者名 (company) の右隣に1列挿入
  if (!currentKeys.includes('skillField')) {
    const compIdx = currentKeys.indexOf('company'); // 0-based
    const insertAfterCol = compIdx !== -1 ? compIdx + 1 : 3;
    coeSheet.insertColumnAfter(insertAfterCol);
    SpreadsheetApp.flush();
    lastCol = coeSheet.getLastColumn();
    currentHeaders = coeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    currentKeys = currentHeaders.map(h => identifyCoeFieldKey_(h));
  }

  // 2. 「外食業 (foodServiceExam)」「飲食料品製造業 (foodMfgExam)」列がまだ存在しない場合、介護日本語 (kaigoLang) の右隣に2列挿入
  if (!currentKeys.includes('foodServiceExam') && !currentKeys.includes('foodMfgExam')) {
    const kLangIdx = currentKeys.indexOf('kaigoLang'); // 0-based
    const insertAfterCol = kLangIdx !== -1 ? kLangIdx + 1 : 16;
    coeSheet.insertColumnsAfter(insertAfterCol, 2);
    SpreadsheetApp.flush();
    lastCol = coeSheet.getLastColumn();
    currentHeaders = coeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    currentKeys = currentHeaders.map(h => identifyCoeFieldKey_(h));
  }

  // 3. 33列のヘッダー文言に差分があれば1行目だけを一括更新
  if (lastCol >= targetColCount) {
    let headerDiff = false;
    for (let i = 0; i < targetColCount; i++) {
      if (String(currentHeaders[i] || "").trim() !== headers[i]) {
        headerDiff = true;
        break;
      }
    }
    if (headerDiff) {
      coeSheet.getRange(1, 1, 1, targetColCount).setValues([headers]);
      SpreadsheetApp.flush();
    }
  }
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
    candIdColIdx: candIdColIdx !== -1 ? candIdColIdx : 4,
    jobIdColIdx: jobIdColIdx !== -1 ? jobIdColIdx : 1
  };
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
 * ※ 技能分野の対象外である試験列には「-」（対象外）を出力し、未入力の列は空欄のまま出力します。
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

  const jobId = jobInfo ? (jobInfo.jobId || '') : '';
  const company = jobInfo ? (jobInfo.company || getVal('採用事業者')) : getVal('採用事業者');
  const skillField = jobInfo ? (jobInfo.skillField || '') : '';
  const offerDate = formatSlashDate((jobInfo && jobInfo.offerDate) ? jobInfo.offerDate : getVal('内定日'));

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
  const birthCity = getVal('出生地（都市名）');
  const addressDetail = getVal('住所詳細');
  const passportNum = getVal('パスポート番号');
  const passportExp = formatSlashDate(getVal('パスポート有効期限'));

  // 日本語資格 (Certification of JLPT or JFT-Basic：14列目 N列のみに出力)
  const jlpt = getVal('特定技能要件＞JLPTレベル');
  const jft = getVal('特定技能要件＞JFTBasicレベル');
  let certLangParts = [];
  if (jlpt && jlpt !== '-' && jlpt !== '×') certLangParts.push(`JLPT ${jlpt}`);
  if (jft && jft !== '-' && jft !== '×') certLangParts.push(jft === 'A2' ? 'JFT A2' : `JFT ${jft}`);
  const certLangStr = certLangParts.join(' / ');

  // 特定技能評価試験（介護技能・介護日本語・外食業・飲食料品製造業）
  const formatExam = (val) => {
    if (!val || val === '-' || val === '×') return '';
    if (val === '合格') return 'YES';
    if (val.includes('予定')) return 'Scheduled (受験予定)';
    return val;
  };

  const legacyOtherExam = getVal('特定技能要件＞その他の評価試験');
  let rawKaigoSkill = getVal('特定技能要件＞介護技能評価試験');
  let rawKaigoLang = getVal('特定技能要件＞介護日本語評価試験');
  let rawFoodService = getVal('特定技能要件＞外食業技能測定試験');
  if (!rawFoodService && legacyOtherExam && legacyOtherExam.includes('外食業')) rawFoodService = '合格';
  let rawFoodMfg = getVal('特定技能要件＞飲食料品製造業技能測定試験');
  if (!rawFoodMfg && legacyOtherExam && legacyOtherExam.includes('飲食料品')) rawFoodMfg = '合格';

  // 技能分野に応じて関係のない試験列には「-」（対象外）を出力する
  let kaigoSkillStr = formatExam(rawKaigoSkill);
  let kaigoLangStr = formatExam(rawKaigoLang);
  let foodServiceStr = formatExam(rawFoodService);
  let foodMfgStr = formatExam(rawFoodMfg);

  if (skillField === '介護') {
    foodServiceStr = '-';
    foodMfgStr = '-';
  } else if (skillField === '外食業') {
    kaigoSkillStr = '-';
    kaigoLangStr = '-';
    foodMfgStr = '-';
  } else if (skillField === '飲食料品製造業') {
    kaigoSkillStr = '-';
    kaigoLangStr = '-';
    foodServiceStr = '-';
  }

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

  // 最終学歴・状況（COE用の英語学校名・英語専攻が入力されている場合のみ出力）
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

  // 有 / 無 変換ヘルパー
  const toYesNo = (val) => {
    if (val === '有') return 'YES (有)';
    if (val === '無') return 'No (無)';
    return val;
  };

  // 過去の交付申請歴・回数 ＆ うち不交付回数
  const applyExpRaw = getVal('在留資格交付申請歴の有無');
  const applyCountRaw = getVal('在留資格交付申請の回数');
  const rejectCountRaw = getVal('不許可となった在留資格交付申請の回数');

  let applyStr = '';
  let rejectStr = '';

  if (applyExpRaw === '無') {
    applyStr = 'No (無)';
    rejectStr = 'No (無)';
  } else if (applyExpRaw === '有') {
    const num = Number(applyCountRaw);
    if (!isNaN(num) && num > 0) {
      applyStr = `YES (${num} TIME${num > 1 ? 'S' : ''})`;
    } else {
      applyStr = 'YES (有)';
    }
    if (rejectCountRaw !== '') {
      const rNum = Number(rejectCountRaw);
      rejectStr = (rNum === 0 || rejectCountRaw === '無' || rejectCountRaw === 'なし') ? 'No (無)' : rejectCountRaw;
    }
  } else if (applyCountRaw !== '') {
    const num = Number(applyCountRaw);
    if (num === 0 || applyCountRaw === '無' || applyCountRaw === 'なし') {
      applyStr = 'No (無)';
      rejectStr = 'No (無)';
    } else if (!isNaN(num) && num > 0) {
      applyStr = `YES (${num} TIME${num > 1 ? 'S' : ''})`;
      if (rejectCountRaw !== '') {
        const rNum = Number(rejectCountRaw);
        rejectStr = (rNum === 0 || rejectCountRaw === '無' || rejectCountRaw === 'なし') ? 'No (無)' : rejectCountRaw;
      }
    } else {
      applyStr = applyCountRaw;
    }
  }

  // 海外出入国歴・期間
  const overseasExp = getVal('海外への出入国歴の有無');
  const overseasExpStr = toYesNo(overseasExp);
  const travelCount = getVal('出入国の回数');
  const lastInDate = formatSlashDate(getVal('直近の入国日'));
  const lastOutDate = formatSlashDate(getVal('直近の出国日'));

  let travelDetailParts = [];
  if (travelCount && travelCount !== '0') travelDetailParts.push(`${travelCount} Time(s)`);
  if (lastInDate || lastOutDate) travelDetailParts.push(`${lastInDate} 〜 ${lastOutDate}`);
  const travelDetailStr = (overseasExp === '無') ? '' : travelDetailParts.join('\n');

  // 在日親族及び同居者情報
  const relExp = getVal('日本在住の親族情報＞有無');
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

  if (relExp === '無') {
    hasFamilyStr = 'No (無)';
    familyDetailStr = '';
  } else if (relExp === '有' || hasDetailedRel) {
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

  // 犯罪歴・技能実習経験・修了証
  const crimeStr = toYesNo(getVal('犯罪歴の有無'));
  const traineeExpStr = toYesNo(getVal('技能実習の経験の有無'));
  const traineeCertStr = toYesNo(getVal('技能実習修了書の有無'));

  // 備考（33列目 AG列：その他の日本語能力試験［NAT-TEST等］と備考・メモのみを出力し、JLPT/JFTは含めない）
  const otherJapanese = getVal('その他の日本語能力試験');
  const memo = getVal('備考・メモ');
  const noteParts = [];
  if (otherJapanese && otherJapanese !== '-' && otherJapanese !== '×') noteParts.push(otherJapanese);
  if (memo) noteParts.push(memo);
  const noteStr = noteParts.join(' / ');

  const fieldMap = {
    offerDate: offerDate,
    jobId: jobId,
    company: company,
    skillField: skillField,
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
    foodServiceExam: foodServiceStr,
    foodMfgExam: foodMfgStr,
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
 * @return {Map} key: SD-XXXX -> { jobId, company, skillField, offerDate }
 */
function getHiredJobInfoMapFromJobSheet_(jobSheet, sourceTZ) {
  const jobData = jobSheet.getDataRange().getValues();
  const jCol = getMasterColumnMap(jobSheet);
  const hiredMap = new Map();

  if (jobData.length < 2 || !jCol['案件ID'] || !jCol['面接結果詳細']) return hiredMap;

  const jIdIdx = jCol['案件ID'] - 1;
  const jCompIdx = jCol['事業者名'] ? jCol['事業者名'] - 1 : -1;
  const jSkillIdx = jCol['技能分野'] ? jCol['技能分野'] - 1 : -1;
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
    const skillFieldStr = jSkillIdx !== -1 ? String(row[jSkillIdx] || "").trim() : "";
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
            skillField: skillFieldStr,
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
 * @return {Array<number>} 更新または追加されたシート上の行番号の配列
 */
function syncCandidatesToCoeInfo(candidateIds, removeCandidateIds = []) {
  try {
    const targetIds = (Array.isArray(candidateIds) ? candidateIds : [candidateIds])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(id => id !== "");
    const removeIds = (Array.isArray(removeCandidateIds) ? removeCandidateIds : [removeCandidateIds])
      .map(id => String(id || "").trim().toUpperCase())
      .filter(id => id !== "");

    if (targetIds.length === 0 && removeIds.length === 0) return [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceTZ = ss.getSpreadsheetTimeZone();
    const masterSheet = ss.getSheetByName('登録者マスタ');
    const jobSheet = ss.getSheetByName('案件管理');
    if (!masterSheet || !jobSheet) return [];

    const targetSS = SpreadsheetApp.openById(EXTERNAL_SS_ID_FUNTOCO);
    const coeSheet = targetSS.getSheetByName(EXTERNAL_SHEET_COE_INFO);
    if (!coeSheet) {
      console.warn(`外部シートに「${EXTERNAL_SHEET_COE_INFO}」が見つかりません。`);
      return [];
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

    if (targetIds.length === 0) return [];

    // 2. 現在の COE info シートの既存行と最終データ行（スマートテーブルのフッター直前）を特定
    const existingRowMap = new Map(); // cid -> 1-based row number
    let lastValidRowIndex = 1;
    const lastSheetRow = coeSheet.getLastRow();

    if (lastSheetRow >= 2) {
      const keyColsData = coeSheet.getRange(2, 1, lastSheetRow - 1, Math.max(candIdColIdx, jobIdColIdx) + 1).getValues();
      for (let i = 0; i < keyColsData.length; i++) {
        const jId = String(keyColsData[i][jobIdColIdx] || "").trim();
        const cId = String(keyColsData[i][candIdColIdx] || "").trim().toUpperCase();
        if (jId || cId) {
          lastValidRowIndex = i + 2;
        }
        if (cId && !existingRowMap.has(cId)) {
          existingRowMap.set(cId, i + 2);
        }
      }
    }

    const rowsToAppend = [];
    const affectedRows = [];

    for (const cid of targetIds) {
      const mRow = masterRowMap.get(cid);
      if (!mRow) continue;

      const jobInfo = hiredJobMap.get(cid) || null;
      const rowArray = buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys);

      if (existingRowMap.has(cid)) {
        const targetRowNum = existingRowMap.get(cid);
        coeSheet.getRange(targetRowNum, 1, 1, lastCol).setValues([rowArray]);
        affectedRows.push(targetRowNum);
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
      for (let k = 0; k < rowsToAppend.length; k++) {
        affectedRows.push(insertRowIndex + k);
      }
    }

    SpreadsheetApp.flush();
    return affectedRows;
  } catch (e) {
    console.error("COE info シートへの自動同期エラー: " + e.message);
    return [];
  }
}

/**
 * 管理者メニュー用：全採用者のCOE情報を外部シート「COE info」へ一括同期する関数
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

    // 既存行の一括更新用バッファ
    const numExistingDataRows = lastValidRowIndex - 1;
    const updatedMatrix = numExistingDataRows > 0 ? coeData.slice(1, lastValidRowIndex).map(row => {
      const padded = row.slice(0, lastCol);
      while (padded.length < lastCol) padded.push("");
      return padded;
    }) : [];
    let isMatrixChanged = false;

    for (const [cid, jobInfo] of sortedEntries) {
      const mRow = masterRowMap.get(cid);
      if (!mRow) continue;
      const newRowArr = buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys);

      if (existingRowMap.has(cid)) {
        const rowNum = existingRowMap.get(cid);
        const matrixIdx = rowNum - 2;
        if (matrixIdx >= 0 && matrixIdx < numExistingDataRows) {
          const oldRowArr = updatedMatrix[matrixIdx];
          const isChanged = newRowArr.some((val, idx) => String(val || "").trim() !== String(oldRowArr[idx] || "").trim());
          if (isChanged) {
            updatedMatrix[matrixIdx] = newRowArr;
            isMatrixChanged = true;
            updateCount++;
          }
        }
      } else {
        rowsToAppend.push(newRowArr);
        appendCount++;
      }
    }

    if (numExistingDataRows > 0 && isMatrixChanged) {
      coeSheet.getRange(2, 1, numExistingDataRows, lastCol).setValues(updatedMatrix);
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