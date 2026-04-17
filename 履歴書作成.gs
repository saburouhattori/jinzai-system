/**
 * 履歴書作成メイン処理
 */
function rirekisyo() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('履歴書作成', '作成したい候補者の「登録者ID番号」(例: 0001) を入力してください', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  let targetId = response.getResponseText().trim();
  if (!targetId) return;
  // IDが数字だけの場合に「SD-」を付ける処理
  if (!targetId.startsWith('SD-')) {
    targetId = 'SD-' + targetId;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 同じファイル内のマスタシートを取得する
  const mSheet = ss.getSheetByName('登録者マスタ');
  const rSheet = ss.getSheetByName('履歴書');

  if (!mSheet || !rSheet) {
    ui.alert('エラー：シートが見つかりません。');
    return;
  }

  // 特定セルのクリア（以前のデータ残り防止）
  rSheet.getRange('J33').clearContent();
  rSheet.getRange('M33').clearContent();

  const data = mSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert('エラー：登録者マスタにデータがありません。');
    return;
  }

  // 1行目のヘッダーを取得して、各項目名が何列目(配列の何番目)にあるかのマッピングを作成
  const headers = data[0];
  const col = {};
  headers.forEach((header, index) => {
    col[header.trim()] = index;
  });

  // 必須の列が存在するか簡易チェック
  if (col['登録者ID'] === undefined) {
    ui.alert('エラー：登録者マスタに「登録者ID」の列が見つかりません。ヘッダー名を確認してください。');
    return;
  }

  let targetData = null;
  // IDを検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][col['登録者ID']] === targetId) {
      targetData = data[i];
      break;
    }
  }

  if (!targetData) {
    ui.alert('エラー：ID ' + targetId + ' が見つかりません。');
    return;
  }

  // 値を取得するためのヘルパー関数（列が存在しない場合のundefined対策）
  const getVal = (colName) => {
    return col[colName] !== undefined ? targetData[col[colName]] : "";
  };

  // --- 書き出し処理（お客様の元のセル配置をそのまま使用） ---
  rSheet.getRange('B2').setValue(getVal('登録者ID'));
  rSheet.getRange('D4').setValue(getVal('フリガナ'));
  rSheet.getRange('D5').setValue(getVal('呼び名'));
  rSheet.getRange('D6').setValue(getVal('名前'));
  
  // 生年月日
  const birthday = getVal('生年月日');
  if (birthday instanceof Date) {
    rSheet.getRange('D9').setValue(Utilities.formatDate(birthday, "JST", "yyyy/MM/dd"));
  } else {
    rSheet.getRange('D9').setValue(birthday);
  }
  
  const age = getVal('満年齢');
  rSheet.getRange('G9').setValue(age ? age + '歳' : '');
  rSheet.getRange('D10').setValue(getVal('性別'));
  rSheet.getRange('K9').setValue(getVal('配偶者')); 
  
  const height = getVal('身長');
  rSheet.getRange('K10').setValue(height ? height + 'cm' : '');
  
  const weight = getVal('体重');
  rSheet.getRange('K11').setValue(weight ? weight + 'kg' : ''); 
  
  rSheet.getRange('E13').setValue(getVal('現住所'));
  rSheet.getRange('K13').setValue(getVal('メールアドレス'));
  rSheet.getRange('E14').setValue(getVal('住所（出身地）'));

  // 学歴
  rSheet.getRange('B17').setValue(getVal('学歴＞入学年月'));
  const schoolName = getVal('学歴＞学校名');
  rSheet.getRange('F17').setValue(schoolName ? schoolName + '　入学' : '');
  rSheet.getRange('B18').setValue(getVal('学歴＞卒業/中退年月'));
  rSheet.getRange('F18').setValue(getVal('学歴＞状況'));
  rSheet.getRange('F20').setValue(getVal('学歴＞補足'));
  
  // 職歴
  rSheet.getRange('C23').setValue(getVal('職歴①＞期間') + '　' + getVal('職歴①＞内容'));
  rSheet.getRange('C24').setValue(getVal('職歴②＞期間') + '　' + getVal('職歴②＞内容'));
  rSheet.getRange('B25').setValue(getVal('職歴③＞期間') + '　' + getVal('職歴③＞内容'));

  // 資格・試験
  const jlptLvl = getVal('特定技能要件＞JLPTレベル');
  const jlptDate = getVal('特定技能要件＞JLPT取得年月');
  if (!jlptLvl) {
    rSheet.getRange('E29').setValue('-');
  } else {
    rSheet.getRange('E29').setValue(jlptLvl + '合格（' + jlptDate + '）');
  }

  const jftLvl = getVal('特定技能要件＞JFT Basicレベル');
  const jftDate = getVal('特定技能要件＞JFT取得年月');
  if (!jftLvl) {
    rSheet.getRange('E30').setValue('-');
  } else {
    rSheet.getRange('E30').setValue(jftLvl + '合格（' + jftDate + '）');
  }

  const careSkill = getVal('特定技能要件＞介護技能評価試験');
  const careSkillDate = getVal('特定技能要件＞介護技能 取得年月');
  if (!careSkill) {
    rSheet.getRange('R28').setValue('-');
  } else {
    rSheet.getRange('R28').setValue(careSkill + '（' + careSkillDate + '）');
  }

  const careLang = getVal('特定技能要件＞介護日本語評価試験');
  const careLangDate = getVal('特定技能要件＞介護日本語 取得年月');
  if (!careLang) {
    rSheet.getRange('R29').setValue('-');
  } else {
    rSheet.getRange('R29').setValue(careLang + '（' + careLangDate + '）');
  }

  rSheet.getRange('J30').setValue(getVal('特定技能要件＞その他の評価試験'));

  const otherSkillDate = getVal('特定技能要件＞その他の評価試験の取得年月');
  if (otherSkillDate) {
    rSheet.getRange('R30').setValue('合格（' + otherSkillDate + '）');
  } else {
    rSheet.getRange('R30').clearContent();
  }

  const otherJlpt = getVal('その他の日本語能力試験');
  if (otherJlpt) {
    rSheet.getRange('B33').setValue(otherJlpt);
  } else {
    rSheet.getRange('B33').clearContent();
  }

  const otherJlptDate = getVal('取得年月');
  if (otherJlptDate) {
    rSheet.getRange('F33').setValue('合格（' + otherJlptDate + '）');
  } else {
    rSheet.getRange('F33').clearContent();
  }

  // コメント・備考
  rSheet.getRange('C36').setValue(getVal('コメント'));
  rSheet.getRange('C41').setValue(getVal('日本在住の親族について'));

  // 写真の取得
  // 登録者IDの列文字と、顔写真の列文字を動的に取得してINDEXとMATCHを使った数式を生成
  if (col['顔写真'] !== undefined && col['登録者ID'] !== undefined) {
    const idColLetter = columnToLetter_(col['登録者ID'] + 1);
    const photoColLetter = columnToLetter_(col['顔写真'] + 1);
    // 例: =IFERROR(INDEX('登録者マスタ'!C:C, MATCH(B2, '登録者マスタ'!A:A, 0)), "")
    const photoFormula = `=IFERROR(INDEX('登録者マスタ'!${photoColLetter}:${photoColLetter}, MATCH(B2, '登録者マスタ'!${idColLetter}:${idColLetter}, 0)), "")`;
    rSheet.getRange('Q3').setFormula(photoFormula);
  }

  ui.alert('ID: ' + targetId + ' の履歴書作成が完了しました。');
}

/**
 * 列番号（1始まり）をアルファベット（A, B, C...）に変換する補助関数
 */
function columnToLetter_(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}