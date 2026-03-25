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
  // マスタデータは外部から取得
  const mSheet = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('登録者マスタ');
  const rSheet = ss.getSheetByName('履歴書');

  if (!mSheet || !rSheet) {
    ui.alert('エラー：シートが見つかりません。');
    return;
  }

  // 特定セルのクリア（以前のデータ残り防止）
  rSheet.getRange('J33').clearContent();
  rSheet.getRange('M33').clearContent();

  const data = mSheet.getDataRange().getValues();
  let targetData = null;

  // IDを検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetId) {
      targetData = data[i];
      break;
    }
  }

  if (!targetData) {
    ui.alert('エラー：ID ' + targetId + ' が見つかりません。');
    return;
  }

  // --- 書き出し処理 ---
  rSheet.getRange('B2').setValue(targetData[0]);  // ID
  rSheet.getRange('D4').setValue(targetData[3]);  // フリガナ
  rSheet.getRange('D5').setValue(targetData[4]);  // 呼び名
  rSheet.getRange('D6').setValue(targetData[1]);  // 名前
  
  // 生年月日
  if (targetData[5] instanceof Date) {
    rSheet.getRange('D9').setValue(Utilities.formatDate(targetData[5], "JST", "yyyy/MM/dd"));
  } else {
    rSheet.getRange('D9').setValue(targetData[5]);
  }
  
  rSheet.getRange('G9').setValue(targetData[6] ? targetData[6] + '歳' : ''); 
  rSheet.getRange('D10').setValue(targetData[7]);
  rSheet.getRange('K9').setValue(targetData[8]);
  rSheet.getRange('K10').setValue(targetData[9] ? targetData[9] + 'cm' : ''); 
  rSheet.getRange('K11').setValue(targetData[10] ? targetData[10] + 'kg' : ''); 
  rSheet.getRange('E13').setValue(targetData[11]);
  rSheet.getRange('K13').setValue(targetData[13]);
  rSheet.getRange('E14').setValue(targetData[12]);

  // 学歴
  rSheet.getRange('B17').setValue(targetData[18]);
  rSheet.getRange('F17').setValue(targetData[15] ? targetData[15] + '　入学' : ''); 
  rSheet.getRange('B18').setValue(targetData[19]);
  rSheet.getRange('F18').setValue(targetData[17]);
  rSheet.getRange('F20').setValue(targetData[20]);
  
  // 職歴
  rSheet.getRange('B23').setValue(targetData[21]);
  rSheet.getRange('G23').setValue(targetData[22]);
  rSheet.getRange('B24').setValue(targetData[23]);
  rSheet.getRange('G24').setValue(targetData[24]);
  rSheet.getRange('B25').setValue(targetData[25]);
  rSheet.getRange('G25').setValue(targetData[26]);

  // 資格・試験
  if (!targetData[27]) {
    rSheet.getRange('E29').setValue('-');
  } else {
    rSheet.getRange('E29').setValue(targetData[27] + '合格（' + (targetData[28] || '') + '）');
  }

  if (!targetData[29]) {
    rSheet.getRange('E30').setValue('-');
  } else {
    rSheet.getRange('E30').setValue(targetData[29] + '合格（' + (targetData[30] || '') + '）');
  }

  if (!targetData[31]) {
    rSheet.getRange('R28').setValue('-');
  } else {
    rSheet.getRange('R28').setValue(targetData[31] + '（' + (targetData[32] || '') + '）');
  }

  if (!targetData[33]) {
    rSheet.getRange('R29').setValue('-');
  } else {
    rSheet.getRange('R29').setValue(targetData[33] + '（' + (targetData[34] || '') + '）');
  }

  rSheet.getRange('J30').setValue(targetData[35]);

  if (targetData[36]) {
    rSheet.getRange('R30').setValue('合格（' + targetData[36] + '）');
  } else {
    rSheet.getRange('R30').clearContent();
  }

  if (targetData[37]) {
    rSheet.getRange('B33').setValue(targetData[37]);
  } else {
    rSheet.getRange('B33').clearContent();
  }

  if (targetData[38]) {
    rSheet.getRange('F33').setValue('合格（' + targetData[38] + '）');
  } else {
    rSheet.getRange('F33').clearContent();
  }

  // コメント・備考
  rSheet.getRange('C36').setValue(targetData[39]);
  rSheet.getRange('C41').setValue(targetData[41]);

  // --- 【重要修正】「候補者写真」シートから直接VLOOKUPで写真をセット ---
  // B2に入力されたIDを元に、同じファイル内の「候補者写真」シートから画像を取得します
  const photoFormula = '=IFERROR(VLOOKUP(B2, \'候補者写真\'!A:B, 2, FALSE), "")';
  rSheet.getRange('Q3').setFormula(photoFormula);

  ui.alert('ID: ' + targetId + ' の履歴書作成が完了しました。');
}
