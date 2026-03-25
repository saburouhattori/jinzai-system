// 【重要】外部マスタ管理スプレッドシートのID
const MASTER_SS_ID = '1cq4h6yI0on-bm_MlqUlUi6MMXBlFIXyTCZpzcTZlCMw';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu( '人材事業メニュー' )
    .addItem( '【新規】候補者登録' , 'showSidebarNew')
    .addItem( '【修正】データ更新' , 'showSidebarEdit')
    .addItem( '【コメント登録】' , 'showSidebarComment')
    .addSeparator()
    .addItem( '【削除】登録者削除' , 'showSidebarDelete')
    .addSeparator()
    .addItem( '【事業者】マスタ登録' , 'showSidebarCompany')
    .addSeparator()
    .addItem( '【作成】履歴書出力' , 'rirekisyo')
    .addItem( '【作成】簡易リスト出力' , 'showSidebarList') 
    .addSeparator()
    .addItem( '【新規】案件登録' , 'showSidebarJobNew')
    .addItem( '【登録】採用者登録' , 'showSidebarHire') 
    .addToUi();
}

/**
 * 統合版サイドバーを表示する共通関数
 */
function showMainSidebar(mode, title) {
  const html = HtmlService.createTemplateFromFile('MainSidebar');
  html.mode = mode;
  const output = html.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle(title)
    .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(output);
}

function showSidebarNew()     { showMainSidebar('NEW',  '【新規】候補者登録' ); }
function showSidebarEdit()    { showMainSidebar('EDIT',  '【修正】データ更新' ); }
function showSidebarComment() { showMainSidebar('COMMENT',  '【コメント登録】' ); }
function showSidebarCompany() { showMainSidebar('COMPANY',  '【事業者】マスタ登録' ); }
function showSidebarJobNew()  { showMainSidebar('JOB',  '【新規】案件登録' ); }
function showSidebarDelete()  { showMainSidebar('DELETE', '【削除】登録者削除'); }
function showSidebarHire()    { showMainSidebar('HIRE', '【登録】採用者登録'); }
function showSidebarList()    { showMainSidebar('LIST', '【作成】簡易リスト出力'); }

/**
 * 登録者の削除処理
 */
function deleteCandidate(id) {
  if (!id) return "IDが指定されていません。";
  try {
    let msg = "";
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const masterSheet = masterSs.getSheetByName('登録者マスタ');
    const masterData = masterSheet.getDataRange().getValues();
    let masterDeleted = false;
    for (let i = masterData.length - 1; i >= 1; i--) {
      if (masterData[i][0] === id) {
        masterSheet.deleteRow(i + 1);
        masterDeleted = true;
        break; 
      }
    }
    if (masterDeleted) msg += "・「登録者マスタ」から削除しました。\n";
    else msg += "・「登録者マスタ」に該当IDは見つかりませんでした。\n";

    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const photoSheet = activeSs.getSheetByName('候補者写真');
    if (photoSheet) {
      const photoData = photoSheet.getDataRange().getValues();
      let photoDeleted = false;
      for (let i = photoData.length - 1; i >= 1; i--) {
        if (photoData[i][0] === id) {
          photoSheet.deleteRow(i + 1);
          photoDeleted = true;
          break;
        }
      }
      if (photoDeleted) msg += "・「候補者写真」から削除しました。";
      else msg += "・「候補者写真」に該当IDは見つかりませんでした。";
    }
    return msg;
  } catch (e) {
    return "エラーが発生しました: " + e.toString();
  }
}

/**
 * 採用者一括登録処理
 */
function registerHire(jobId, candIds) {
  try {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const jobSheet = activeSs.getSheetByName('案件管理');
    const jobData = jobSheet.getDataRange().getValues();
    
    let jobRow = -1;
    let companyName = "";
    let existingHiredStr = "";
    let allCandidatesStr = "";
    let interviewDateRaw = "";
    
    for (let i = 1; i < jobData.length; i++) {
      if (jobData[i][0] === jobId) {
        jobRow = i + 1;
        companyName = jobData[i][3];
        allCandidatesStr = jobData[i][4] || "";
        existingHiredStr = jobData[i][5] || "";
        interviewDateRaw = jobData[i][6] || "";
        break;
      }
    }
    if (jobRow === -1) return "エラー: 指定された案件IDが存在しません。";

    let interviewDateStr = "日付不明";
    if (interviewDateRaw instanceof Date) {
      interviewDateStr = Utilities.formatDate(interviewDateRaw, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    } else if (interviewDateRaw) {
      interviewDateStr = String(interviewDateRaw);
    }

    const allCandIds = [];
    const candList = String(allCandidatesStr).split(/\r?\n|,/);
    candList.forEach(cand => {
      const str = cand.trim();
      if (str) {
        const match = str.match(/^([a-zA-Z]+-\d+)/);
        if (match) allCandIds.push(match[1]);
        else allCandIds.push(str.split('-')[0].trim());
      }
    });

    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const candSheet = masterSs.getSheetByName('登録者マスタ');
    const candData = candSheet.getDataRange().getValues();
    
    const headers = candData[0];
    
    let historyColIdx = headers.indexOf('面接履歴') !== -1 ? headers.indexOf('面接履歴') + 1 : 43; 
    let statusColIdx = headers.indexOf('ステータス') !== -1 ? headers.indexOf('ステータス') + 1 : 44; 
    let compColIdx = headers.indexOf('採用事業者') !== -1 ? headers.indexOf('採用事業者') + 1 : 
                     headers.indexOf('採用先') !== -1 ? headers.indexOf('採用先') + 1 : 45; 

    const neededCols = Math.max(historyColIdx, statusColIdx, compColIdx);
    if (candSheet.getMaxColumns() < neededCols) {
      candSheet.insertColumnsAfter(candSheet.getMaxColumns(), neededCols - candSheet.getMaxColumns());
    }

    const hiredNames = [];
    const candUpdates = []; 
    
    for (let i = 1; i < candData.length; i++) {
      const cId = candData[i][0];
      if (allCandIds.indexOf(cId) !== -1) {
        const isHired = (candIds.indexOf(cId) !== -1); 
        if (isHired) {
          hiredNames.push(cId + "-" + candData[i][1]); 
        }
        candUpdates.push({
          row: i + 1,
          isHired: isHired,
          currentHistory: candData[i][historyColIdx - 1] || "" 
        });
      }
    }

    if (hiredNames.length === 0) return "エラー: 指定された採用候補者IDがマスタに見つかりません。";

    let currentHiredArr = existingHiredStr ? String(existingHiredStr).split('\n') : [];
    hiredNames.forEach(name => {
      if (currentHiredArr.indexOf(name) === -1) {
        currentHiredArr.push(name);
      }
    });
    jobSheet.getRange(jobRow, 6).setValue(currentHiredArr.join('\n')); 
    jobSheet.getRange(jobRow, 2).setValue("入国準備"); 

    candUpdates.forEach(item => {
      try {
        let prefix = item.isHired ? "【採用】" : "【不採用】";
        let addText = prefix + interviewDateStr + "：" + companyName;
        let newHistory = String(item.currentHistory);
        
        if (newHistory.indexOf(addText) === -1) {
          newHistory = newHistory ? newHistory + "\n" + addText : addText;
          candSheet.getRange(item.row, historyColIdx).setValue(newHistory); 
        }

        if (item.isHired) {
          candSheet.getRange(item.row, statusColIdx).setValue("採用"); 
          candSheet.getRange(item.row, compColIdx).setValue(companyName);
        }
      } catch (innerErr) {
        console.error("行" + item.row + "の更新エラー: " + innerErr.message);
      }
    });

    return "採用登録が完了しました。\n案件: " + jobId + "\n採用人数: " + hiredNames.length + "名\n（非採用者の履歴も自動更新しました）";
    
  } catch (e) {
    return "処理中にエラーが発生しました: " + e.message;
  }
}

/**
 * 案件管理シートから指定案件IDの候補者リストを取得する
 */
function getJobCandidates(jobId) {
  if (!jobId) return [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const jobSheet = ss.getSheetByName('案件管理');
    const jobData = jobSheet.getDataRange().getValues();

    for (let i = 1; i < jobData.length; i++) {
      if (jobData[i][0] === jobId) {
        const candidatesStr = String(jobData[i][4]); 
        if (!candidatesStr) return [];
        
        const candList = candidatesStr.split(/\r?\n|,/);
        const result = [];
        
        candList.forEach(cand => {
          const str = cand.trim();
          if (str) {
            let extractedId = "";
            const match = str.match(/^([a-zA-Z]+-\d+)/);
            if (match) extractedId = match[1];
            else extractedId = str.split('-')[0].trim();
            
            result.push({ id: extractedId, display: str });
          }
        });
        return result;
      }
    }
    return null; 
  } catch (e) {
    return [];
  }
}

/**
 * 指定した登録者IDの配列から「簡易リスト」を作成する
 */
function generateSimpleList(candIds) {
  try {
    if (!candIds || candIds.length === 0) return "候補者IDが指定されていません。";

    // 1. マスタを開いてデータを取得
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const masterSheet = masterSs.getSheetByName('登録者マスタ');
    const masterData = masterSheet.getDataRange().getValues();
    const headers = masterData[0];

    // マスタの列番号を特定
    const colName = headers.indexOf('名前') !== -1 ? headers.indexOf('名前') : 1;
    const colFuri = headers.indexOf('フリガナ') !== -1 ? headers.indexOf('フリガナ') : 3;
    const colAge  = headers.indexOf('満年齢') !== -1 ? headers.indexOf('満年齢') : 6;
    const colSex  = headers.indexOf('性別') !== -1 ? headers.indexOf('性別') : 7;
    const colSch  = headers.indexOf('学歴＞学校名') !== -1 ? headers.indexOf('学歴＞学校名') : 15;
    const colStat = headers.indexOf('学歴＞状況') !== -1 ? headers.indexOf('学歴＞状況') : 17;
    const colJLPT = headers.indexOf('特定技能要件＞JLPTレベル') !== -1 ? headers.indexOf('特定技能要件＞JLPTレベル') : 27;
    const colJFT  = headers.indexOf('特定技能要件＞JFT Basicレベル') !== -1 ? headers.indexOf('特定技能要件＞JFT Basicレベル') : 29;
    const colOth  = headers.indexOf('その他の日本語能力試験') !== -1 ? headers.indexOf('その他の日本語能力試験') : 37;

    const listData = [];
    const formulas = []; 

    candIds.forEach((id, index) => {
      let found = false;
      for (let i = 1; i < masterData.length; i++) {
        if (String(masterData[i][0]) === id) {
          
          let jlpt = masterData[i][colJLPT] ? String(masterData[i][colJLPT]).trim() : "";
          if (jlpt === "") jlpt = "×";
          
          let jft = masterData[i][colJFT] ? String(masterData[i][colJFT]).trim() : "";
          if (jft === "") jft = "×";

          listData.push([
            masterData[i][colName] || "", 
            masterData[i][colFuri] || "", 
            masterData[i][colAge]  || "", 
            masterData[i][colSex]  || "", 
            masterData[i][colSch]  || "", 
            masterData[i][colStat] || "", 
            jlpt,                         
            jft,                          
            masterData[i][colOth]  || "", 
            id                            
          ]);
          
          formulas.push(['=IFERROR(VLOOKUP(L' + (listData.length + 1) + ', \'候補者写真\'!$A:$B, 2, FALSE), "")']);
          found = true;
          break;
        }
      }
      
      if (!found) {
        listData.push(["未登録", "", "", "", "", "", "×", "×", "", id]); 
        formulas.push(['']);
      }
    });

    if (listData.length === 0) return "該当するデータが見つかりませんでした。";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listSheet = ss.getSheetByName('簡易リスト');
    
    listSheet.getRange('B2:L51').clearContent();

    listSheet.getRange(2, 3, listData.length, 10).setValues(listData);
    listSheet.getRange(2, 2, formulas.length, 1).setFormulas(formulas);

    return listData.length + "名の簡易リストを作成しました。";

  } catch (e) {
    return "処理中にエラーが発生しました: " + e.message;
  }
}

/**
 * 日本語学校マスタから名称リストを取得する（プルダウン用）
 */
function getSchoolList() {
  try {
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const schoolSheet = masterSs.getSheetByName('日本語学校マスタ');
    if (!schoolSheet) return [];
    
    const data = schoolSheet.getDataRange().getValues();
    const schools = data.slice(1).map(row => row[1]).filter(name => name !== "");
    
    return [...new Set(schools)].sort();
  } catch(e) {
    return [];
  }
}

/**
 * 登録者マスタから「ID: 名前」の辞書オブジェクトを取得する
 */
function getCandidateDict() {
  try {
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const sheet = masterSs.getSheetByName('登録者マスタ');
    if (!sheet) return {};
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const dict = {};
    data.forEach(row => {
      if (row[0]) dict[String(row[0]).trim()] = String(row[1]);
    });
    return dict;
  } catch(e) {
    return {};
  }
}

/**
 * 時間主導型トリガー（毎日深夜）で実行される全候補者の年齢更新処理
 */
function updateAllAges() {
  try {
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const sheet = masterSs.getSheetByName('登録者マスタ');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const birthColIdx = headers.indexOf('生年月日');
    const ageColIdx = headers.indexOf('満年齢');

    if (birthColIdx === -1 || ageColIdx === -1) return;

    const today = new Date();
    const todayNum = today.getFullYear() * 10000 + (today.getMonth() + 1) * 100 + today.getDate();

    const ageUpdates = [];

    for (let i = 1; i < data.length; i++) {
      let birthVal = data[i][birthColIdx];
      let newAge = data[i][ageColIdx];

      if (birthVal) {
        let birthDate = null;
        if (birthVal instanceof Date) {
          birthDate = birthVal;
        } else if (typeof birthVal === 'string') {
          const match = birthVal.match(/(\d{4})[-\/\年](\d{1,2})[-\/\月](\d{1,2})/);
          if (match) {
            birthDate = new Date(match[1], match[2] - 1, match[3]);
          }
        }
        if (birthDate && !isNaN(birthDate.getTime())) {
          const birthNum = birthDate.getFullYear() * 10000 + (birthDate.getMonth() + 1) * 100 + birthDate.getDate();
          newAge = Math.floor((todayNum - birthNum) / 10000);
        }
      }
      ageUpdates.push([newAge]);
    }

    if (ageUpdates.length > 0) {
      sheet.getRange(2, ageColIdx + 1, ageUpdates.length, 1).setValues(ageUpdates);
    }
  } catch(e) {
    console.error("年齢自動更新エラー: " + e.message);
  }
}

// --- ここから追加・修正：データ更新用の「安全な」検索エンジン ---
/**
 * 登録者IDからデータを検索し、フォーム用のオブジェクトを返す（エラー回避強化版）
 */
function safeSearchByAdminId(id) {
  try {
    const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    const sheet = masterSs.getSheetByName('登録者マスタ');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // ★空白（半角・全角・ゼロ幅）を完全除去して大文字で比較（見つからないエラーを防止）
    const cleanId = (str) => String(str).replace(/[\s\uFEFF\xA0\u3000]/g, '').toUpperCase();
    const searchId = cleanId(id);

    for (let i = 1; i < data.length; i++) {
      if (cleanId(data[i][0]) === searchId) {
        const res = { row: i + 1 };
        
        const getIdx = (name) => headers.findIndex(h => String(h).replace(/\n/g, '').includes(name));
        
        // ★重要：全ての取得値を「文字列」に強制変換（Dateオブジェクトによるシリアライズエラーを完全防止）
        const safeGet = (name) => {
          const idx = getIdx(name);
          if (idx !== -1 && data[i][idx] != null && data[i][idx] !== '') {
            const val = data[i][idx];
            if (val instanceof Date) {
              return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            }
            return String(val).trim();
          }
          return '';
        };

        res.name = safeGet('名前');
        res.furigana = safeGet('フリガナ');
        res.nickname = safeGet('呼び名');

        const colBirth = getIdx('生年月日');
        if (colBirth !== -1 && data[i][colBirth] != null && data[i][colBirth] !== '') {
          let bDay = data[i][colBirth];
          if (bDay instanceof Date) {
            res.birthday = Utilities.formatDate(bDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          } else {
            const bStr = String(bDay).trim();
            const match = bStr.match(/(\d{4})[-\/\年](\d{1,2})[-\/\月](\d{1,2})/);
            if (match) {
              res.birthday = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
            } else {
              res.birthday = bStr;
            }
          }
        } else {
          res.birthday = '';
        }

        res.gender = safeGet('性別');
        res.spouse = safeGet('配偶者');
        res.height = safeGet('身長');
        res.weight = safeGet('体重');
        res.address = safeGet('現住所');
        res.birthplace = safeGet('住所（出身地）');
        res.email = safeGet('メールアドレス');
        res.school = safeGet('所属日本語学校');

        res.eduSchool = safeGet('学歴＞学校名');
        res.eduDept = safeGet('学歴＞学部・学科・専攻');
        res.eduStatus = safeGet('学歴＞状況');
        res.eduStart = safeGet('学歴＞入学年月');
        res.eduEnd = safeGet('学歴＞卒業/中退年月');
        res.eduNote = safeGet('学歴＞補足');

        res.expPeriod1 = safeGet('職歴①＞期間');
        res.expContent1 = safeGet('職歴①＞内容');
        res.expPeriod2 = safeGet('職歴②＞期間');
        res.expContent2 = safeGet('職歴②＞内容');
        res.expPeriod3 = safeGet('職歴③＞期間');
        res.expContent3 = safeGet('職歴③＞内容');

        res.jlptLevel = safeGet('特定技能要件＞JLPTレベル');
        res.jlptDate = safeGet('特定技能要件＞JLPT取得年月');
        res.jftLevel = safeGet('特定技能要件＞JFT Basicレベル');
        res.jftDate = safeGet('特定技能要件＞JFT取得年月');
        res.kaigoSkill = safeGet('特定技能要件＞介護技能評価試験');
        res.kaigoSkillDate = safeGet('特定技能要件＞介護技能 取得年月');
        res.kaigoLang = safeGet('特定技能要件＞介護日本語評価試験');
        res.kaigoLangDate = safeGet('特定技能要件＞介護日本語 取得年月');

        res.otherJapanese = safeGet('その他の日本語能力試験');
        const colOther = getIdx('その他の日本語能力試験');
        if (colOther !== -1 && colOther + 1 < headers.length && data[i][colOther + 1] != null && data[i][colOther + 1] !== '') {
            const dVal = data[i][colOther + 1];
            res.otherJapaneseDate = (dVal instanceof Date) ? Utilities.formatDate(dVal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(dVal).trim();
        } else {
            res.otherJapaneseDate = '';
        }

        res.comment = safeGet('コメント');
        
        const relName = safeGet('親族の名前');
        const relRel = safeGet('続柄');
        if (relName && relName !== "なし" && relName !== "無" && relName !== "日本に家族や親戚はいません。") {
          res.relative = relRel ? `${relRel}：${relName}` : relName;
        } else {
          res.relative = "日本に家族や親戚はいません。";
        }

        return res; 
      }
    }
    return null; 
  } catch(e) {
    throw new Error("検索中にエラーが発生しました: " + e.message);
  }
}
/**
 * HTMLファイルを分割して読み込むための共通関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// --- 追加ここまで ---