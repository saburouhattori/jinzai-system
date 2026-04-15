// ====== マスタ連携・その他共通処理 ======

/**
 * 登録や更新後に各一覧シートを同期する統合関数
 * （02_候補者管理.gs などから呼び出されます）
 */
function syncListSheets() {
  // true を渡すことでポップアップアラートを非表示にしてバックグラウンド実行します
  updateCandidateLists(true);
  updateSimpleList();
}

/**
 * 全角半角スペース、改行などをすべて除去し、小文字化して比較用文字列を作る内部関数
 * （表記ゆれを完全に吸収します）
 */
function normalize_(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/[\s　\n\r]+/g, '').toLowerCase();
}

/**
 * ターゲットシートのヘッダー配列に従って、データマップから1行分の配列を生成する内部関数
 * （これによって列が入れ替わったり挿入されたりしても、自動で正しい位置にデータが入ります）
 */
function buildRowByHeaders_(headers, dataMap) {
  return headers.map(h => {
    const key = normalize_(h);
    return dataMap[key] !== undefined ? dataMap[key] : '';
  });
}

/**
 * 登録者マスタから「採用者一覧」「未採用者一覧」シートを更新する
 * @param {boolean} silent - true の場合、完了時のアラートを表示しない
 */
function updateCandidateLists(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('登録者マスタ');
  const hiredSheet = ss.getSheetByName('採用者一覧');
  const unhiredSheet = ss.getSheetByName('未採用者一覧');
  const jobSheet = ss.getSheetByName('案件管理');

  if (!masterSheet || !hiredSheet || !unhiredSheet || !jobSheet) {
    throw new Error('必要なシート（登録者マスタ、採用者一覧、未採用者一覧、案件管理）のいずれかが見つかりません。');
  }

  // マスタデータの取得
  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData.shift();

  // 出力先ヘッダーの取得（動的生成の要：シートの構成変更に自動追従します）
  const hiredHeaders = hiredSheet.getRange(1, 1, 1, hiredSheet.getLastColumn()).getValues()[0];
  const unhiredHeaders = unhiredSheet.getRange(1, 1, 1, unhiredSheet.getLastColumn()).getValues()[0];
  
  // 案件管理データの取得と事前マップ化（検索高速化のため）
  const jobData = jobSheet.getDataRange().getValues();
  const jobHeaders = jobData.shift();
  const jobList = jobData.map(row => {
    const jobMap = {};
    for (let i = 0; i < jobHeaders.length; i++) {
      jobMap[normalize_(jobHeaders[i])] = row[i];
    }
    return jobMap;
  });

  const hiredData = [];
  const unhiredData = [];

  // マスタの必須項目チェック
  if (masterHeaders.findIndex(h => normalize_(h) === normalize_('登録者ID')) === -1) {
    throw new Error('登録者マスタに「登録者ID」列が見つかりません。');
  }

  // マスタデータを1行ずつ処理
  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];
    
    // 1. 各行のデータを「ヘッダー名」をキーにした連想配列(dataMap)に変換
    const dataMap = {};
    for (let c = 0; c < masterHeaders.length; c++) {
      dataMap[normalize_(masterHeaders[c])] = row[c];
    }

    const candidateId = dataMap[normalize_('登録者ID')];
    const status = dataMap[normalize_('ステータス')];
    const companyName = dataMap[normalize_('採用事業者')];

    // 空行はスキップ
    if (!candidateId) continue;

    // 2. 特殊項目（結合や別名など）の計算・追加
    const jlpt = dataMap[normalize_('特定技能要件＞JLPTレベル')];
    const jft = dataMap[normalize_('特定技能要件＞JFT Basicレベル')];
    const kaigoGinou = dataMap[normalize_('特定技能要件＞介護技能評価試験')];
    const kaigoNihongo = dataMap[normalize_('特定技能要件＞介護日本語評価試験')];
    
    // ▽ 未採用者一覧用「特定技能要件」の結合文字列
    let reqs = [];
    if (jlpt) reqs.push(jlpt);
    if (jft) reqs.push(jft);
    if (kaigoGinou) reqs.push(`介護技能（${kaigoGinou}）`);
    if (kaigoNihongo) reqs.push(`介護日本語（${kaigoNihongo}）`);
    dataMap[normalize_('特定技能要件')] = reqs.join(', ');

    // ▽ ヘッダー名の揺れ（エイリアス）をマッピング
    dataMap[normalize_('JLPT')] = jlpt;
    dataMap[normalize_('JFT Basic')] = jft;
    dataMap[normalize_('採用事業者名')] = companyName; // マスタは「採用事業者」、一覧は「採用事業者名」
    dataMap[normalize_('在留資格交付申請の有無')] = dataMap[normalize_('在留資格交付申請の回数')];

    // 3. ステータスに応じた処理
    if (status === '採用') {
      let jobId = '';
      let skillField = '';
      
      // 案件管理から「案件ID」と「技能分野」を取得
      if (companyName) {
        // 事業者名が一致し、かつ候補者名か採用者名にこの登録者IDが含まれる案件を探す
        const matchedJob = jobList.find(job => {
          const jComp = job[normalize_('事業者名')];
          const jCands = String(job[normalize_('候補者名')] || '');
          const jHired = String(job[normalize_('採用者名')] || '');
          
          return jComp === companyName && (jCands.includes(candidateId) || jHired.includes(candidateId));
        });
        
        if (matchedJob) {
          jobId = matchedJob[normalize_('案件ID')] || '';
          skillField = matchedJob[normalize_('技能分野')] || '';
        }
      }
      
      // 取得した案件情報をdataMapに追加
      dataMap[normalize_('案件ID')] = jobId;
      dataMap[normalize_('技能分野')] = skillField;

      // 「採用者一覧」のヘッダー順に従って、dataMapから自動的に配列を組み立てる
      const outRow = buildRowByHeaders_(hiredHeaders, dataMap);
      hiredData.push(outRow);

    } else if (status === '未採用' || status === '辞退' || status === '保留') {
      // 「未採用者一覧」のヘッダー順に従って、自動的に配列を組み立てる
      const outRow = buildRowByHeaders_(unhiredHeaders, dataMap);
      unhiredData.push(outRow);
    }
  }

  // ----------------------------
  // シートへの書き込み処理
  // ----------------------------
  if (hiredData.length > 0) {
    const lastRow = hiredSheet.getLastRow();
    if (lastRow > 1) hiredSheet.getRange(2, 1, lastRow - 1, hiredHeaders.length).clearContent();
    hiredSheet.getRange(2, 1, hiredData.length, hiredHeaders.length).setValues(hiredData);
  }

  if (unhiredData.length > 0) {
    const lastRow = unhiredSheet.getLastRow();
    if (lastRow > 1) unhiredSheet.getRange(2, 1, lastRow - 1, unhiredHeaders.length).clearContent();
    unhiredSheet.getRange(2, 1, unhiredData.length, unhiredHeaders.length).setValues(unhiredData);
  }

  if (!silent) {
    try {
      SpreadsheetApp.getUi().alert('リストの更新が完了しました。');
    } catch(e) {
      // UIがない環境での実行時は無視する
    }
  }
}

/**
 * 簡易リストの更新処理
 */
function updateSimpleList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('登録者マスタ');
  const simpleSheet = ss.getSheetByName('簡易リスト');

  if (!masterSheet || !simpleSheet) {
    throw new Error('必要なシートが見つかりません。');
  }

  const masterData = masterSheet.getDataRange().getValues();
  const masterHeaders = masterData.shift();

  const simpleData = [];
  let no = 1;

  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];
    
    // こちらも同様にdataMap化
    const dataMap = {};
    for (let c = 0; c < masterHeaders.length; c++) {
      dataMap[normalize_(masterHeaders[c])] = row[c];
    }

    const candId = dataMap[normalize_('登録者ID')];
    if (!candId) continue;

    const jlpt = dataMap[normalize_('特定技能要件＞JLPTレベル')];
    const jft = dataMap[normalize_('特定技能要件＞JFT Basicレベル')];

    // ※簡易リストは「写真」などの特殊列（B列）を保持するため、これまで通りの列セットを行います
    simpleData.push([
      no++,                                      // A: No
      '',                                        // B: 写真（数式等のため空）
      dataMap[normalize_('名前')] || '',           // C: 名前
      dataMap[normalize_('フリガナ')] || '',       // D: フリガナ
      dataMap[normalize_('満年齢')] || '',         // E: 年齢
      dataMap[normalize_('性別')] || '',           // F: 性別
      dataMap[normalize_('学歴＞学校名')] || '',     // G: 学歴(学校名)
      dataMap[normalize_('学歴＞状況')] || '',       // H: 状況(卒業など)
      jlpt ? jlpt : '×',                          // I: JLPT
      jft ? jft : '×',                            // J: JFT
      dataMap[normalize_('その他の日本語能力試験')] || '', // K: その他の日本語
      candId                                     // L: 登録者ID
    ]);
  }

  if (simpleData.length > 0) {
    const lastRow = simpleSheet.getLastRow();
    if (lastRow > 1) {
      // B列（写真）はクリアしないよう分けてクリア
      simpleSheet.getRange(2, 1, lastRow - 1, 1).clearContent();
      simpleSheet.getRange(2, 3, lastRow - 1, 10).clearContent();
    }
    
    // 値のセット (B列はスキップする)
    const rangeA = simpleData.map(r => [r[0]]);
    const rangeC_L = simpleData.map(r => r.slice(2));
    
    simpleSheet.getRange(2, 1, simpleData.length, 1).setValues(rangeA);
    simpleSheet.getRange(2, 3, simpleData.length, 10).setValues(rangeC_L);
  }
}