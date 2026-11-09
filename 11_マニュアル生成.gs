// =========================================
// 操作マニュアルシートの自動生成・書式設定
// =========================================

function createOrUpdateManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '操作マニュアル';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    sheet.clearFormats();
  }

  // グリッド線を非表示にしてドキュメント風に見せる
  sheet.setHiddenGridlines(true);

  // 列幅の設定 (A: 余白, B: ステップ/項目, C: 操作メニュー/対象, D: 操作手順・詳細ルール)
  sheet.setColumnWidth(1, 25);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 580);

  const rows = [];
  const rowStyles = []; // { row, type }

  const addRow = (colB, colC, colD, type = 'NORMAL') => {
    rows.push(['', colB, colC, colD]);
    rowStyles.push({ row: rows.length, type: type });
  };

  // --- タイトル ---
  addRow('人材管理システム 操作マニュアル', '', '', 'TITLE');
  addRow('本マニュアルは、日常業務において人材管理システム（Googleスプレッドシート）を操作する手順とルールをまとめたものです。\nデータの登録・修正・出力は、すべてスプレッドシート上部のカスタムメニューから行います。', '', '', 'SUBTITLE');
  addRow('', '', '', 'SPACER');

  // --- 1. 重要ルール ---
  addRow('⚠️ 1. システム運用上の重要ルール（厳守事項）', '', '', 'SECTION');
  addRow('項目', 'ルール内容', '理由・注意点', 'TABLE_HEADER_WARN');
  addRow(
    '① 1行目（ヘッダー名）の変更禁止',
    '各シートの1行目の文字は絶対に変更しないでください。',
    'システムは1行目の項目名を読み取ってデータの読み書きやシート間の同期を行っています。文字を変更するとシステムが動かなくなります。',
    'WARN_ROW'
  );
  addRow(
    '② 数式列の直接編集禁止',
    '「満年齢」「顔写真」の列は直接文字を入力しないでください。',
    'これらの列は関数（数式）によって自動計算・自動表示されています。直接入力すると数式が壊れてしまいます。',
    'WARN_ROW'
  );
  addRow(
    '③ メニューからの操作徹底',
    'データの登録・更新・削除は必ず上部メニューのフォームから行ってください。',
    'シートへ直接手入力すると、関連する他シート（「案件管理」「採用者一覧」など）とのデータ連携が崩れる原因になります。',
    'WARN_ROW'
  );
  addRow('', '', '', 'SPACER');

  // --- 2. メニューの構成 ---
  addRow('📋 2. 上部メニューの構成', '', '', 'SECTION');
  addRow('メニュー名', '含まれる機能（番号順）', '主な用途', 'TABLE_HEADER');
  addRow(
    '1. 登録者マスタ管理',
    '1. 候補者の新規登録\n2. 登録者データの修正・更新\n3. 登録者データの削除',
    '候補者（登録者）の基本情報・学歴・職歴・資格・顔写真の登録、修正、および削除を行います。',
    'DATA_ROW'
  );
  addRow(
    '2. 案件・採用管理',
    '1. 案件の新規登録\n2. 簡易リスト作成\n3. 履歴書作成\n4. 案件情報の修正・削除\n5. 面接結果の登録\n6. 面接結果の修正・更新\n7. COE申請情報の登録・更新 (COE Info)\n8. 入国日・入職日の一括登録',
    '求人案件の作成から、企業提案用の書類作成、面接日の設定、合否結果・内定日の登録、内定後のCOE（在留資格）申請用情報や入国・入職日の登録までを時系列順に行います。',
    'DATA_ROW'
  );
  addRow(
    '3. 管理者用（マスタ・同期）',
    '1. 事業者の新規登録\n2. 採用・未採用リストの同期\n3. 支払い管理シートへ同期（Funtoco）',
    '新しい受入企業（事業者）のマスタ登録や、各シート間のデータ同期・外部シート連携を行います。',
    'DATA_ROW'
  );
  addRow('', '', '', 'SPACER');

  // --- 3. 日常業務の基本フロー ---
  addRow('🏢 3. 日常業務の基本フロー（Step 1 〜 Step 7）', '', '', 'SECTION');
  addRow('ステップ・業務タイミング', '操作するメニュー', '具体的な操作手順・入力ルール', 'TABLE_HEADER');

  addRow(
    'Step 1：候補者の新規登録\n（新規候補者の面談完了後）',
    '1. 登録者マスタ管理\n＞ 1. 候補者の新規登録',
    '1. メニューを開き、候補者の基本情報・学歴・職歴・資格・顔写真を入力して「新規登録」を押します。\n' +
    '2. 登録が完了すると「SD-XXXX」のIDが自動発行され、ステータスは「未採用」、登録日は当日日付が自動で入ります。\n\n' +
    '【入力ルール】\n' +
    '・名前：アルファベットは自動で「大文字」に変換されます。名前にカッコ () や数字（1, 2など）は絶対に含めないでください（同姓同名もIDで区別されます）。既存の候補者と同姓同名や同一生年月日の場合は自動で注意喚起が表示されます。\n' +
    '・学歴・試験年月・職歴期間：カレンダーから年月を選択します。職歴は開始年月と終了年月を選ぶと自動で「〜」で結合されます。',
    'DATA_ROW'
  );
  addRow(
    '💡 候補者データの修正・更新\n（登録済み情報の変更時）',
    '1. 登録者マスタ管理\n＞ 2. 登録者データの修正・更新',
    '1. 「管理番号で検索」にID（例：1 または SD-0001）を入力して「読み込む」を押します。\n' +
    '2. 内容を修正して「更新」を押します。\n' +
    '※名前を修正した場合、「案件管理」シート側の候補者名表記も自動的に新しい名前へ書き換わります。',
    'NOTE_ROW'
  );

  addRow(
    'Step 2：案件登録と提案書類の作成\n（求人依頼・候補者選定時）',
    '2. 案件・採用管理\n＞ 1. 案件の新規登録\n＞ 2. 簡易リスト作成\n＞ 3. 履歴書作成',
    '【① 案件の新規登録】\n' +
    '1. 事業者名を入力し、表示される候補リストから選択します（合同面接の場合は「+ 事業者名を追加」で複数追加可）。\n' +
    '   ※候補に出ない新規企業の場合は、先に「3. 管理者用」＞「1. 事業者の新規登録」で登録してください。\n' +
    '2. 技能分野を選び、エントリーする候補者IDを入力します（既に他案件で「採用」の候補者は登録不可）。\n' +
    '3. 求人票等のファイルをPCから選択して「登録」を押すと、Googleドライブに「案件ID_事業者名」フォルダが自動作成され、ファイルが保存されます。\n\n' +
    '【② 簡易リスト作成】\n' +
    '・候補者IDを入力して「簡易リストを作成」を押すと、「簡易リスト」シートに提案用の一覧が出力されます。\n\n' +
    '【③ 履歴書作成】\n' +
    '・ダイアログに候補者ID（例：0001 または SD-0001）を入力してOKを押すと、「履歴書」シートに反映されます。',
    'DATA_ROW'
  );

  addRow(
    'Step 3：面接日の設定・案件更新\n（面接日程の確定時・候補者変更時）',
    '2. 案件・採用管理\n＞ 4. 案件情報の修正・削除',
    '1. 「案件IDで検索」に対象の案件ID（例：225 または JOB-0225）を入力して「読み込む」を押します。\n' +
    '2. 確定した「面接日」を入力し、ステータスを「面接準備」などに変更して「更新を保存」を押します。\n' +
    '※面接日が設定されていない案件は、次のStep 4（面接結果の登録）に進むことができません。',
    'DATA_ROW'
  );

  addRow(
    'Step 4：面接結果・内定日の登録\n（面接実施・合否判明時）',
    '2. 案件・採用管理\n＞ 5. 面接結果の登録',
    '1. 「採用元の案件ID」に案件IDを入力して「候補者検索」を押します。\n' +
    '2. 【採用者がいる場合】：採用された候補者にチェックを入れ（合同面接の場合は右のプルダウンで採用企業を選択）、「内定日」を入力します。\n' +
    '3. 【全員不採用の場合】：下部の「今回の面接では『採用者なし』として結果を登録する」にチェックを入れます。\n' +
    '4. 「面接結果を登録」を押すと、案件管理の面接結果詳細が更新され、全候補者のマスタにステータス・採用事業者・内定日・面接履歴が自動反映されます。',
    'DATA_ROW'
  );
  addRow(
    '💡 面接結果の修正・更新\n（内定辞退や結果変更が生じた時）',
    '2. 案件・採用管理\n＞ 6. 面接結果の修正・更新',
    '・一度結果を登録した案件は「5. 面接結果の登録」からは操作できません。結果の変更や内定辞退が発生した場合はこちらを使用します。\n' +
    '・案件IDで検索し、各候補者のプルダウンから「採用 / 不採用 / 内定辞退（候補者都合） / 内定取消（事業者都合） / 面接結果の取消」を選択して更新します（マスタ側のステータスや履歴も自動で連動修正されます）。',
    'NOTE_ROW'
  );

  addRow(
    'Step 5：COE申請情報の登録・更新\n（在留資格申請の準備時・日英併記）',
    '2. 案件・採用管理\n＞ 7. COE申請情報の登録・更新 (COE Info)',
    '1. 「管理番号で検索 / Search by ID」に採用者のID（例：1 または SD-0001）を入力して「読み込む / Load」を押します。\n' +
    '2. 在留資格認定証明書（COE）交付申請等に必要な以下の情報を確認・入力し、「追加情報を保存 / Save Info」を押します（ミャンマー人スタッフ向けに日英併記されています）。\n' +
    '   ・基本・パスポート情報：名前、呼び名、生年月日、性別、配偶者、所属送り出し機関、出生地（都市名・自動で英語大文字化）、住所詳細、パスポート番号（半角英数字のみ）・有効期限\n' +
    '   ・学歴（最終学歴）：学校名、学部・学科・専攻、状況（卒業/中途退学など）、卒業/中退年月\n' +
    '   ・特定技能要件・日本語能力：JLPT、JFT Basic、介護技能/日本語評価試験、その他の評価試験、その他の日本語能力試験（面接時に「受験予定」だった候補者が合格した場合もここで変更可能）\n' +
    '   ・経歴・申請歴：職業、技能実習（Ginojissyu）経験・修了書の有無、犯罪歴の有無、交付申請回数・不許可回数、海外出入国歴・直近の出入国日\n' +
    '   ・日本在住の親族情報（氏名、続柄、生年月日、国籍、同居予定、勤務先・通学先、在留カード番号）',
    'DATA_ROW'
  );

  addRow(
    'Step 6：入国日・入職日の一括登録\n（渡航・配属スケジュールの確定時）',
    '2. 案件・採用管理\n＞ 8. 入国日・入職日の一括登録',
    '1. 「対象の案件ID」に案件IDを入力して「採用者検索」を押すと、その案件で「採用」となっている候補者のみが表示されます。\n' +
    '2. 【全員同じ日付の場合】：青枠内の「共通の入国日」「共通の入職日」に入力すると、下の全員の欄に一括で同じ日付がセットされます。\n' +
    '3. 【個別に日付が異なる場合】：一括セット後に個人の欄を直接書き換えるか、1名ずつ入力して調整します。\n' +
    '4. 「一括保存」を押すと、登録者マスタの入国日・入職日が一括更新され、採用者一覧へ自動同期されます。',
    'DATA_ROW'
  );

  addRow(
    'Step 7：新規事業者の登録・データ同期\n（新規取引開始時・定期同期時）',
    '3. 管理者用（マスタ・同期）\n＞ 1〜3の各メニュー',
    '【1. 事業者の新規登録】\n' +
    '・新しい受入企業をマスタに登録します。「事業者名称」にはスペース（空白）を含めずに入力してください。\n\n' +
    '【2. 採用・未採用リストの同期】\n' +
    '・「登録者マスタ」と「案件管理」の最新データをもとに、「採用者一覧」「未採用者一覧」を更新します。また、「案件管理」の内定日を正本としてマスタ側の内定日も一括で揃えます。\n\n' +
    '【3. 支払い管理シートへ同期（Funtoco）】\n' +
    '・確定した採用者情報や内定日・入国日・入職日を、Funtoco様の外部シート「支払い管理」へ同期します（辞退・取消となった候補者は自動削除されます）。',
    'DATA_ROW'
  );
  addRow('', '', '', 'SPACER');

  // --- 4. よくあるエラーと対処法 ---
  addRow('❓ 4. よくあるエラーメッセージと対処法', '', '', 'SECTION');
  addRow('表示されるメッセージ', '原因', '対処法', 'TABLE_HEADER');
  addRow(
    '「辞書データを読み込み中です。少々お待ちください。」',
    '画面を開いた直後で、システムが裏側でマスタデータを読み込んでいる途中です。',
    '画面を閉じずに2〜3秒待ってから、もう一度検索ボタンや登録ボタンを押してください。',
    'DATA_ROW'
  );
  addRow(
    '「未登録の事業者が含まれています」',
    '案件登録時に、事業者マスタに登録されていない企業名が入力されています。',
    '先に「3. 管理者用（マスタ・同期）」＞「1. 事業者の新規登録」から企業を登録するか、文字入力時に表示される候補リストから選択してください。',
    'DATA_ROW'
  );
  addRow(
    '「面接日が設定されていません」',
    '案件管理シートに「面接日」が入っていない状態で、面接結果を登録しようとしています。',
    '先に「2. 案件・採用管理」＞「4. 案件情報の修正・削除」を開き、その案件の「面接日」を入力・保存してから結果登録を行ってください。',
    'DATA_ROW'
  );
  addRow(
    '「名前にカッコや数字を含めることはできません」',
    '候補者の名前欄に () や 1, 2 などの数字が入力されています。',
    'カッコや数字を消してアルファベット（およびスペース）のみで登録してください。同姓同名の候補者は登録者ID（SD-XXXX）で自動的に区別されます。',
    'DATA_ROW'
  );

  // 一括書き込み
  const range = sheet.getRange(1, 1, rows.length, 4);
  range.setValues(rows);

  // 全体共通フォント・折り返し設定
  range.setFontFamily('Arial')
       .setFontSize(10)
       .setVerticalAlignment('top')
       .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // 行ごとのスタイル適用
  rowStyles.forEach(item => {
    const r = item.row;
    const type = item.type;
    const rowRange = sheet.getRange(r, 2, 1, 3); // B〜D列

    if (type === 'TITLE') {
      rowRange.merge();
      rowRange.setBackground('#1a73e8')
              .setFontColor('#ffffff')
              .setFontSize(16)
              .setFontWeight('bold')
              .setVerticalAlignment('middle');
      sheet.setRowHeight(r, 45);
    } else if (type === 'SUBTITLE') {
      rowRange.merge();
      rowRange.setBackground('#e8f0fe')
              .setFontColor('#1f1f1f')
              .setFontSize(10)
              .setVerticalAlignment('middle');
      sheet.setRowHeight(r, 42);
    } else if (type === 'SPACER') {
      sheet.setRowHeight(r, 18);
    } else if (type === 'SECTION') {
      rowRange.merge();
      rowRange.setBackground('#3c4043')
              .setFontColor('#ffffff')
              .setFontSize(12)
              .setFontWeight('bold')
              .setVerticalAlignment('middle');
      sheet.setRowHeight(r, 32);
    } else if (type === 'TABLE_HEADER') {
      rowRange.setBackground('#d2e3fc')
              .setFontColor('#174ea6')
              .setFontWeight('bold')
              .setVerticalAlignment('middle')
              .setBorder(true, true, true, true, true, true, '#aecbfa', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(r, 28);
    } else if (type === 'TABLE_HEADER_WARN') {
      rowRange.setBackground('#fad2cf')
              .setFontColor('#a50e0e')
              .setFontWeight('bold')
              .setVerticalAlignment('middle')
              .setBorder(true, true, true, true, true, true, '#f28b82', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(r, 28);
    } else if (type === 'WARN_ROW') {
      rowRange.setBackground('#fef7f6')
              .setBorder(true, true, true, true, true, true, '#f28b82', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(r, 2).setFontWeight('bold').setFontColor('#c5221f');
      sheet.getRange(r, 3).setFontWeight('bold');
    } else if (type === 'DATA_ROW') {
      rowRange.setBackground('#ffffff')
              .setBorder(true, true, true, true, true, true, '#dadce0', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(r, 2).setFontWeight('bold').setBackground('#f8f9fa');
      sheet.getRange(r, 3).setFontColor('#1a73e8').setFontWeight('bold');
    } else if (type === 'NOTE_ROW') {
      rowRange.setBackground('#e8f0fe')
              .setBorder(true, true, true, true, true, true, '#aecbfa', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(r, 2).setFontWeight('bold').setFontColor('#174ea6');
      sheet.getRange(r, 3).setFontColor('#174ea6').setFontWeight('bold');
    }
  });

  // シートをアクティブにして完了通知
  ss.setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('「操作マニュアル」シートの作成・書式設定が完了しました。');
}