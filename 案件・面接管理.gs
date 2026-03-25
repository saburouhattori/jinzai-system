/**
 * 事業者マスタから名称リストを取得する（プルダウン用）
 */
function getVendorList() {
  const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
  const vendorSheet = masterSs.getSheetByName('事業者マスタ');
  const data = vendorSheet.getDataRange().getValues();
  const vendors = data.slice(1).map(row => row[1]).filter(name => name !== "");
  return [...new Set(vendors)].sort();
}

/**
 * 案件登録処理（複数候補者・改行・ID付与対応版）
 */
function addJobRow(formData) {
  if (!formData.company) return "事業者名を入力または選択してください。";
  if (!formData.candidateIds || formData.candidateIds.length === 0) return "候補者の登録者IDが入力されていません。";

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('案件管理');
    
    // 1. マスタ管理の「登録者マスタ」からIDを検索し、名前に変換する
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const candSheet = masterSs.getSheetByName('登録者マスタ');
    const candData = candSheet.getDataRange().getValues();
    
    // IDと名前(A列=0, B列=1)の対応辞書を作成
    const candMap = {};
    for (let i = 1; i < candData.length; i++) {
      if (candData[i][0]) {
        candMap[candData[i][0]] = candData[i][1];
      }
    }
    
    // 入力されたIDの配列を、「ID-名前」の形式に変換
    const candidateNames = formData.candidateIds.map(id => {
      // マスタに存在すれば「SD-0001-KYAW MIN OO」の形に、存在しなければ「ID(未登録)」を返す
      return candMap[id] ? id + "-" + candMap[id] : id + "(未登録)";
    });
    
    // 複数の候補者を改行コード（\n）で結合して1つの文字列にする
    const candidateNamesStr = candidateNames.join("\n");

    // 2. 事業者マスタの確認と自動登録
    const isDuplicate = checkDuplicateCompanyName(formData.company);
    if (!isDuplicate) {
      addCompanyRow({
        name: formData.company,
        note: "案件登録時に自動追加"
      });
    }

    // 3. 案件管理シートへ登録
    const lastRow = sheet.getLastRow();
    const nextId = "JOB-" + lastRow.toString().padStart(4, '0');
    
    // E列（インデックス4）に改行で結合した候補者名文字列を入力
    sheet.appendRow([nextId, "未着手", new Date(), formData.company, candidateNamesStr]);
    
    // 完了メッセージ（アラートが縦に長くなりすぎないよう、登録人数を表示する形に変更）
    let msg = "案件「" + nextId + "」を登録しました。\n登録候補者数: " + formData.candidateIds.length + "名";
    if (!isDuplicate) msg += "\n※新しい事業者をマスタに自動登録しました。";
    return msg;

  } catch (e) {
    return "エラーが発生しました: " + e.toString();
  }
}

/**
 * 一括面接登録処理
 */
function submitInterviews(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('面接管理');
  const rows = data.candidates.map(c => ["", data.date, data.company, c.id, c.name, ""]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  return "面談一括登録完了";
}
