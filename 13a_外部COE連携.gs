// =========================================================
// 外部連携：COE info シートの列定義・レイアウト解析・データ変換 (13a_外部COEデータ変換.gs)
// =========================================================

const EXTERNAL_SHEET_COE_INFO = "COE info";

/**
 * COE info シートの標準34列ヘッダー定義（1行目：日本語、2行目：英語）
 */
function getCoeInfoHeaders_() {
  return [
    "状況\nStatus",                                                                     // 1 (A)
    "内定日\nOffer Date",                                                               // 2 (B)
    "案件ID\nJob ID",                                                                   // 3 (C)
    "採用事業者名\nAccepting Organization",                                             // 4 (D)
    "技能分野\nJob Field",                                                              // 5 (E)
    "登録者ID\nCandidate ID",                                                           // 6 (F)
    "氏名\nName",                                                                       // 7 (G)
    "呼び名\nNickname",                                                                 // 8 (H)
    "性別\nSex",                                                                        // 9 (I)
    "所属送り出し機関\nSending Organization",                                           // 10 (J)
    "生年月日\nDate of birth",                                                          // 11 (K)
    "出生地（都市名）\nPlace of birth",                                                 // 12 (L)
    "旅券番号（パスポート番号）\nPassport Number",                                      // 13 (M)
    "旅券有効期限（パスポート有効期限）\nPassport Date of expiration",                  // 14 (N)
    "日本語能力試験（JLPT / JFT）\nCertification of JLPT or JFT-Basic",                 // 15 (O)
    "介護技能評価試験\nNursing Care Skills Exam",                                       // 16 (P)
    "介護日本語評価試験\nNursing Care Japanese Exam",                                   // 17 (Q)
    "外食業技能測定試験\nFood Service Skills Exam",                                     // 18 (R)
    "飲食料品製造業技能測定試験\nFood & Beverage Mfg. Skills Exam",                     // 19 (S)
    "配偶者の有無\nMarital status",                                                     // 20 (T)
    "本国における居住地（住所詳細）\nHome address in Myanmar",                          // 21 (U)
    "職業\nOccupation",                                                                 // 22 (V)
    "最終学歴・状況\nFinal educational background",                                     // 23 (W)
    "最終学歴（学校名・卒業年月・専攻）\nName of school / Date of graduation / Major",  // 24 (X)
    "過去の交付申請歴・回数\nPast COE application history & count",                     // 25 (Y)
    "うち不交付となった回数\nNumber of times of non-issuance",                          // 26 (Z)
    "過去の出入国歴\nPast entry into / departure from Japan",                           // 27 (AA)
    "出入国回数・直近の出入国期間\nNumber of times & Latest entry/departure period",    // 28 (AB)
    "在日親族及び同居者の有無\nFamily in Japan or co-residents",                        // 29 (AC)
    "在日親族・同居者の詳細\nDetails of family in Japan",                               // 30 (AD)
    "犯罪歴の有無（日本・海外）\nCriminal record (in Japan / overseas)",                // 31 (AE)
    "技能実習経験の有無\nExperience of Ginojissyu",                                     // 32 (AF)
    "技能実習修了証の有無\nCertification of Ginojissyu",                                // 33 (AG)
    "備考（その他日本語試験等）\nNote"                                                  // 34 (AH)
  ];
}

/**
 * ヘッダー文字列から標準フィールドキーを判定するヘルパー関数
 * ※ AH列「備考（その他日本語試験等）」が O列「日本語試験」に誤判定されないよう、「備考 / Note」を最優先で判定します。
 */
function identifyCoeFieldKey_(headerStr) {
  const h = String(headerStr || "").normalize('NFC').replace(/\s+/g, ' ').trim();
  if (!h) return null;

  // 0. 備考列の最優先判定
  if (h.includes('備考') || h.includes('Note')) return 'note';

  // 1. 複合文・固有キーワードの優先判定（W列「最終学歴・状況」やT列「Marital status」とA列「状況 / Status」を明確に区別）
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

  // 2. A列「状況 / Status」（COE情報状況）の判定
  if (h.includes('状況') || h.includes('Status') || h.includes('ステータス')) return 'coeStatus';

  // 3. 管理キー・基本項目の判定
  if (h.includes('内定日') || h.includes('Offer Date')) return 'offerDate';
  if (h.includes('案件ID') || h.includes('Job ID')) return 'jobId';
  if (h.includes('採用事業者') || h.includes('Accepting Organization') || h.includes('Company')) return 'company';
  if (h.includes('技能分野') || h.includes('Job Field')) return 'skillField';
  if (h.includes('登録者ID') || h.includes('Candidate ID')) return 'candId';
  if (h.includes('氏名') || h.includes('名前') || h.includes('Name')) return 'name';
  return null;
}

/**
 * 状況列（A列等）に「入力済（緑）」「未入力（赤）」の条件付き書式を設定するヘルパー関数
 * ※ すでに設定済みの場合は再設定を行わずスキップします（スマートテーブルに負荷をかけません）。
 */
function ensureCoeStatusConditionalFormat_(coeSheet, statusColNum) {
  try {
    const rules = coeSheet.getConditionalFormatRules();
    const hasCompletedRule = rules.some(r => {
      const bc = r.getBooleanCondition();
      if (!bc) return false;
      const vals = bc.getCriteriaValues();
      return vals && vals.includes("入力済");
    });
    if (hasCompletedRule) return;

    const maxRows = Math.max(coeSheet.getMaxRows(), 1000);
    const targetRange = coeSheet.getRange(2, statusColNum, maxRows - 1, 1);

    // 既存の「状況」関連ルールがあれば除外して再構築
    const filteredRules = rules.filter(r => {
      const bc = r.getBooleanCondition();
      if (!bc) return true;
      const vals = bc.getCriteriaValues();
      if (vals && (vals.includes("入力済") || vals.includes("未入力") || vals.includes("入力済 (Completed)") || vals.includes("未入力 (Incomplete)"))) {
        return false;
      }
      return true;
    });

    const ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("入力済")
      .setBackground("#e6f4ea")
      .setFontColor("#137333")
      .setBold(true)
      .setRanges([targetRange])
      .build();

    const ruleIncomplete = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("未入力")
      .setBackground("#fce8e6")
      .setFontColor("#c5221f")
      .setBold(true)
      .setRanges([targetRange])
      .build();

    filteredRules.push(ruleCompleted, ruleIncomplete);
    coeSheet.setConditionalFormatRules(filteredRules);
  } catch (e) {
    console.warn("条件付き書式の設定をスキップしました: " + e.message);
  }
}

/**
 * COE info シートの1行目ヘッダーを読み取り、各列とフィールドキーの対応マップを返す関数
 * ※ A列に「状況\nStatus」列がまだ存在しない場合は最左列に挿入し、旧「COE情報状況」表記であれば「状況\nStatus」に更新します。
 */
function getCoeSheetLayout_(coeSheet) {
  let lastCol = coeSheet.getLastColumn();
  if (lastCol === 0 || coeSheet.getLastRow() === 0) {
    const defaultHeaders = getCoeInfoHeaders_();
    coeSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    lastCol = defaultHeaders.length;
  }

  let headers = coeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let colKeys = headers.map(h => identifyCoeFieldKey_(h));

  // A列に「状況 (coeStatus)」がまだない場合は挿入、旧ヘッダー名（「COE情報状況」等）の場合は「状況\nStatus」へ更新
  let statusColIdx = colKeys.indexOf('coeStatus');
  if (statusColIdx === -1) {
    coeSheet.insertColumnBefore(1);
    coeSheet.getRange(1, 1).setValue("状況\nStatus");
    coeSheet.setColumnWidth(1, 85);
    SpreadsheetApp.flush();
    lastCol = coeSheet.getLastColumn();
    headers = coeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    colKeys = headers.map(h => identifyCoeFieldKey_(h));
    statusColIdx = colKeys.indexOf('coeStatus');
  } else if (String(headers[statusColIdx] || "").trim() !== "状況\nStatus") {
    coeSheet.getRange(1, statusColIdx + 1).setValue("状況\nStatus");
    coeSheet.setColumnWidth(statusColIdx + 1, 85);
    headers[statusColIdx] = "状況\nStatus";
  }

  if (statusColIdx !== -1) {
    ensureCoeStatusConditionalFormat_(coeSheet, statusColIdx + 1);
  }

  const colLabels = headers.map(h => String(h || "").split(/\r?\n/)[0].trim());
  const candIdColIdx = colKeys.indexOf('candId'); // 0-based
  const jobIdColIdx = colKeys.indexOf('jobId');   // 0-based
  return {
    lastCol: lastCol,
    colKeys: colKeys,
    colLabels: colLabels,
    candIdColIdx: candIdColIdx !== -1 ? candIdColIdx : 5,
    jobIdColIdx: jobIdColIdx !== -1 ? jobIdColIdx : 2
  };
}

/**
 * 差分比較用の正規化ヘルパー関数
 * ※ スプレッドシート側で Date 型やゼロ埋めなし日付（例: 2026/11/6）に変換されていても同一値と判定させます。
 */
function normalizeCoeCompareVal_(val, tz) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz || "JST", "yyyy/MM/dd");
  }
  const s = String(val).normalize('NFC').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  const dateMatch = s.match(/^(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})日?$/);
  if (dateMatch) {
    return `${dateMatch[1]}/${dateMatch[2].padStart(2, '0')}/${dateMatch[3].padStart(2, '0')}`;
  }
  return s;
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
 * ※ 備考を除くすべての対象項目が埋まっている場合のみ、A列（coeStatus）を「入力済」とし、
 *   1つでも空欄がある場合は「未入力」とします。
 */
function buildCoeInfoRowByLayout_(mRow, mCol, jobInfo, sourceTZ, colKeys) {
  const getVal = (colName) => {
    const clean = colName.replace(/\s/g, '');
    const cIdx = mCol[clean];
    if (!cIdx) return "";
    const val = mRow[cIdx - 1];
    if (val === null || val === undefined) return "";
    if (val instanceof Date) return Utilities.formatDate(val, sourceTZ, "yyyy/MM/dd");
    const s = String(val).normalize('NFC').trim();
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

  // 日本語資格 (Certification of JLPT or JFT-Basic)
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

  // ① Y列（過去の交付申請歴・回数）＆ Z列（うち不交付となった回数）
  const applyExpRaw = getVal('在留資格交付申請歴の有無');
  const applyCountRaw = getVal('在留資格交付申請の回数');
  const rejectCountRaw = getVal('不許可となった在留資格交付申請の回数');

  let applyStr = '';
  let rejectStr = '';

  if (applyExpRaw === '無') {
    applyStr = 'No (無)';
    rejectStr = '-';
  } else if (applyExpRaw === '有') {
    const num = Number(applyCountRaw);
    if (!isNaN(num) && num > 0) {
      applyStr = `YES (${num} TIME${num > 1 ? 'S' : ''})`;
    } else {
      applyStr = 'YES (有)';
    }
    if (rejectCountRaw !== '') {
      const rNum = Number(rejectCountRaw);
      rejectStr = (rNum === 0 || rejectCountRaw === '無' || rejectCountRaw === 'なし') ? '0 (無)' : `${rejectCountRaw}`;
    }
  } else if (applyCountRaw !== '') {
    const num = Number(applyCountRaw);
    if (num === 0 || applyCountRaw === '無' || applyCountRaw === 'なし') {
      applyStr = 'No (無)';
      rejectStr = '-';
    } else if (!isNaN(num) && num > 0) {
      applyStr = `YES (${num} TIME${num > 1 ? 'S' : ''})`;
      if (rejectCountRaw !== '') {
        const rNum = Number(rejectCountRaw);
        rejectStr = (rNum === 0 || rejectCountRaw === '無' || rejectCountRaw === 'なし') ? '0 (無)' : `${rejectCountRaw}`;
      }
    } else {
      applyStr = applyCountRaw;
    }
  }

  // ② AA列（過去の出入国歴）＆ AB列（出入国回数・直近の出入国期間）
  const overseasExp = getVal('海外への出入国歴の有無');
  const overseasExpStr = toYesNo(overseasExp);
  const travelCount = getVal('出入国の回数');
  const lastInDate = formatSlashDate(getVal('直近の入国日'));
  const lastOutDate = formatSlashDate(getVal('直近の出国日'));

  let travelDetailStr = '';
  if (overseasExp === '無') {
    travelDetailStr = '-';
  } else {
    let travelDetailParts = [];
    if (travelCount && travelCount !== '0') travelDetailParts.push(`${travelCount} Time(s)`);
    if (lastInDate || lastOutDate) travelDetailParts.push(`${lastInDate} 〜 ${lastOutDate}`);
    travelDetailStr = travelDetailParts.join('\n');
  }

  // ③ AC列（在日親族及び同居者の有無）＆ AD列（在日親族・同居者の詳細）
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
    familyDetailStr = '-';
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

  // AE列：犯罪歴の有無
  const crimeStr = toYesNo(getVal('犯罪歴の有無'));

  // ④ AF列（技能実習経験の有無）＆ AG列（技能実習修了証の有無）
  const traineeExpRaw = getVal('技能実習の経験の有無');
  const traineeCertRaw = getVal('技能実習修了書の有無');
  const traineeExpStr = toYesNo(traineeExpRaw);
  let traineeCertStr = '';
  if (traineeExpRaw === '無') {
    traineeCertStr = '-';
  } else {
    traineeCertStr = toYesNo(traineeCertRaw);
  }

  // 備考（34列目 AH列：その他の日本語能力試験［NAT-TEST等］と備考・メモのみを出力し、JLPT/JFTは含めない）
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

  // ▼ A列「状況 (coeStatus)」の全項目完了判定
  // 備考 (note) を除くすべての必須項目が1つも欠けずに埋まっているかをチェック
  const requiredKeys = [
    'offerDate', 'jobId', 'company', 'skillField', 'candId',
    'name', 'nickname', 'sex', 'agent', 'birthday',
    'birthCity', 'passportNum', 'passportExp', 'certLang',
    'kaigoSkill', 'kaigoLang', 'foodServiceExam', 'foodMfgExam',
    'spouse', 'addressDetail', 'job', 'eduBgSummary', 'eduDetail',
    'applyStr', 'rejectStr', 'overseasExp', 'travelDetail',
    'hasFamily', 'familyDetail', 'crime', 'traineeExp', 'traineeCert'
  ];

  // さらに英語学歴の卒業年月（eduEndEn）が抜けていないかも厳密チェック
  const isAllFilled = requiredKeys.every(k => String(fieldMap[k] || "").trim() !== "") && (eduSchoolEn !== "" && eduEndEn !== "");
  fieldMap.coeStatus = isAllFilled ? "入力済" : "未入力";

  return colKeys.map(k => (k && fieldMap[k] !== undefined) ? fieldMap[k] : "");
}