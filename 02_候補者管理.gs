// =========================================
// 候補者管理の操作（登録・更新・削除）
// =========================================

/**
 * 登録者マスタの見出しから列番号（1ベース）を取得
 */
function getMasterColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[String(h).replace(/\s/g, '')] = i + 1;
  });
  return map;
}

/**
 * 候補者（SD）の新規登録
 */
function addCandidate(formData) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) throw new Error("「登録者マスタ」シートが見つかりません。");

    // ★安全装置1：シートの列数が足りない場合は自動で拡張する
    const requiredCols = 71;
    if (sheet.getMaxColumns() < requiredCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
    }

    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    let targetRow = 1; 

    // A列（ID）が入っている本当の最終行を特定する
    for (let i = 1; i < data.length; i++) {
      let idStr = String(data[i][0]).trim();
      if (idStr.startsWith("SD-")) {
        let num = parseInt(idStr.replace("SD-", ""), 10);
        if (num > maxId) maxId = num;
        targetRow = i + 1; 
      }
    }
    
    const nextRow = targetRow + 1; 
    
    // ★安全装置2：シートの行数が足りない場合は自動で拡張する
    if (nextRow > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    }

    const nextId = "SD-" + (maxId + 1).toString().padStart(4, '0');
    
    // ★修正：fill('') だと ARRAYFORMULA を壊すため fill(null) に変更
    const rowData = new Array(sheet.getMaxColumns()).fill(null);
    
    // 全71項目の正確なマッピング（未定義の場合はnullを入れて数式を保護）
    rowData[0] = nextId; // A: 登録者ID
    rowData[1] = formData.name || null; // B: 名前
    // C(2): 顔写真 (新規は空)
    rowData[3] = formData.kana || null; // D: フリガナ
    rowData[4] = formData.nickname || null; // E: 呼び名
    rowData[5] = formData.birthday || null; // F: 生年月日
    // G(6): 満年齢 (数式のため null のまま保護)
    rowData[7] = formData.gender || null; // H: 性別
    rowData[8] = formData.marriage || null; // I: 配偶者
    rowData[9] = formData.height || null; // J: 身長
    rowData[10] = formData.weight || null; // K: 体重
    rowData[11] = formData.address || null; // L: 現住所
    rowData[12] = formData.origin || null; // M: 住所（出身地）
    rowData[13] = formData.email || null; // N: メールアドレス
    rowData[14] = formData.school || null; // O: 所属日本語学校
    rowData[15] = formData.eduName || null; // P: 学歴 学校名
    rowData[16] = formData.eduMajor || null; // Q: 学歴 学科
    rowData[17] = formData.eduStatus || null; // R: 学歴 状況
    rowData[18] = formData.eduStart || null; // S: 学歴 入学
    rowData[19] = formData.eduEnd || null; // T: 学歴 卒業
    rowData[20] = formData.eduMemo || null; // U: 学歴 補足
    rowData[21] = formData.work1Period || null; // V: 職歴1 期間
    rowData[22] = formData.work1Detail || null; // W: 職歴1 内容
    rowData[23] = formData.work2Period || null; // X: 職歴2 期間
    rowData[24] = formData.work2Detail || null; // Y: 職歴2 内容
    rowData[25] = formData.work3Period || null; // Z: 職歴3 期間
    rowData[26] = formData.work3Detail || null; // AA: 職歴3 内容
    rowData[27] = formData.jlpt || null; // AB: JLPT
    rowData[28] = formData.jlptDate || null; // AC: JLPT年月
    rowData[29] = formData.jft || null; // AD: JFT
    rowData[30] = formData.jftDate || null; // AE: JFT年月
    rowData[31] = formData.kaigoG || null; // AF: 介護技能
    rowData[32] = formData.kaigoGDate || null; // AG: 介護技能年月
    rowData[33] = formData.kaigoN || null; // AH: 介護日本語
    rowData[34] = formData.kaigoNDate || null; // AI: 介護日本語年月
    rowData[35] = formData.otherTest || null; // AJ: その他評価
    rowData[36] = formData.otherTestDate || null; // AK: その他年月
    rowData[37] = formData.otherSkill || null; // AL: その他能力試験
    rowData[38] = formData.otherSkillDate || null; // AM: その他取得年月
    rowData[39] = formData.memo || null; // AN: コメント
    // AO(40): 修正前コメント (新規は空)
    rowData[41] = formData.family || null; // AP: 親族について
    // AQ(42): 面接履歴 (新規は空)
    rowData[43] = '未採用'; // AR: ステータス
    // AS(44): 採用事業者 (新規は空)
    rowData[45] = formData.skillField || null; // AT: 技能分野
    rowData[46] = formData.agent || null; // AU: 所属送り出し機関
    rowData[47] = Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd"); // AV: 登録日
    // AW(48): 内定日
    rowData[49] = formData.birthPlace || null; // AX: 出生地
    rowData[50] = formData.addressDetail || null; // AY: 住所詳細
    rowData[51] = formData.passportNo || null; // AZ: パスポート番号
    rowData[52] = formData.passportExp || null; // BA: パスポート期限
    rowData[53] = formData.job || null; // BB: 職業
    rowData[54] = formData.expJissyu || null; // BC: 実習経験
    rowData[55] = formData.certJissyu || null; // BD: 修了書
    rowData[56] = formData.crime || null; // BE: 犯罪歴
    rowData[57] = formData.visaApplyCount || null; // BF: 申請回数
    rowData[58] = formData.visaRejectCount || null; // BG: 不許可回数
    rowData[59] = formData.overseasHist || null; // BH: 出入国歴
    rowData[60] = formData.overseasCount || null; // BI: 出入国回数
    rowData[61] = formData.lastEntry || null; // BJ: 入国日
    rowData[62] = formData.lastExit || null; // BK: 出国日
    rowData[63] = formData.relName || null; // BL: 親族名
    rowData[64] = formData.relType || null; // BM: 続柄
    rowData[65] = formData.relBirth || null; // BN: 親族生年月日
    rowData[66] = formData.relNat || null; // BO: 親族国籍
    rowData[67] = formData.relLive || null; // BP: 同居予定
    rowData[68] = formData.relWork || null; // BQ: 勤務先
    rowData[69] = formData.relCard || null; // BR: 在留カード
    rowData[70] = formData.generalMemo || null; // BS: 備考・メモ

    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    return `登録者 ${nextId} を新規登録しました。`;
  } catch (e) {
    // ★デバッグ強化：どこでエラーになったか詳細をフロントに返す
    throw new Error("バックエンド処理エラー: " + e.message + " (行: " + e.lineNumber + ")");
  }
}

/**
 * 登録者詳細の取得
 */
function getCandidateDetails(candId) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    const id = String(candId).trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === id) {
        const r = data[i];
        const toIso = (v) => (v instanceof Date) ? Utilities.formatDate(v, "JST", "yyyy-MM-dd") : v;
        
        return {
          row: i + 1, id: r[0], name: r[1], kana: r[3], nickname: r[4],
          birthday: toIso(r[5]), gender: r[7], marriage: r[8], height: r[9], weight: r[10],
          address: r[11], origin: r[12], email: r[13], school: r[14],
          eduName: r[15], eduMajor: r[16], eduStatus: r[17], eduStart: r[18], eduEnd: r[19], eduMemo: r[20],
          work1Period: r[21], work1Detail: r[22], work2Period: r[23], work2Detail: r[24], work3Period: r[25], work3Detail: r[26],
          jlpt: r[27], jlptDate: r[28], jft: r[29], jftDate: r[30],
          kaigoG: r[31], kaigoGDate: r[32], kaigoN: r[33], kaigoNDate: r[34],
          otherTest: r[35], otherTestDate: r[36], otherSkill: r[37], otherSkillDate: r[38],
          memo: r[39], family: r[41], interviewHist: r[42], status: r[43],
          hiredBy: r[44], skillField: r[45], agent: r[46], regDate: toIso(r[47]), hireDate: toIso(r[48]),
          birthPlace: r[49], addressDetail: r[50], passportNo: r[51], passportExp: toIso(r[52]),
          job: r[53], expJissyu: r[54], certJissyu: r[55], crime: r[56],
          visaApplyCount: r[57], visaRejectCount: r[58], overseasHist: r[59], overseasCount: r[60],
          lastEntry: toIso(r[61]), lastExit: toIso(r[62]), relName: r[63], relType: r[64],
          relBirth: toIso(r[65]), relNat: r[66], relLive: r[67], relWork: r[68], relCard: r[69],
          generalMemo: r[70]
        };
      }
    }
    return null;
  } catch(e) { throw new Error(e.message); }
}

/**
 * 候補者の更新（高速版）
 */
function updateCandidate(formData) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    const row = Number(formData.row);
    if (!row) throw new Error("行特定不可");

    const requiredCols = 71;
    if (sheet.getMaxColumns() < requiredCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
    }

    const currentValues = sheet.getRange(row, 1, 1, sheet.getMaxColumns()).getValues()[0];
    const photo = currentValues[2];
    const oldComment = currentValues[40];

    const rowData = new Array(sheet.getMaxColumns()).fill(null);
    rowData[0] = formData.id;
    rowData[1] = formData.name;
    rowData[2] = photo;
    rowData[3] = formData.kana;
    rowData[4] = formData.nickname;
    rowData[5] = formData.birthday;
    // 6: 満年齢(ARRAYFORMULAで自動計算のため保護)
    rowData[7] = formData.gender;
    rowData[8] = formData.marriage;
    rowData[9] = formData.height;
    rowData[10] = formData.weight;
    rowData[11] = formData.address;
    rowData[12] = formData.origin;
    rowData[13] = formData.email;
    rowData[14] = formData.school;
    rowData[15] = formData.eduName;
    rowData[16] = formData.eduMajor;
    rowData[17] = formData.eduStatus;
    rowData[18] = formData.eduStart;
    rowData[19] = formData.eduEnd;
    rowData[20] = formData.eduMemo;
    rowData[21] = formData.work1Period;
    rowData[22] = formData.work1Detail;
    rowData[23] = formData.work2Period;
    rowData[24] = formData.work2Detail;
    rowData[25] = formData.work3Period;
    rowData[26] = formData.work3Detail;
    rowData[27] = formData.jlpt;
    rowData[28] = formData.jlptDate;
    rowData[29] = formData.jft;
    rowData[30] = formData.jftDate;
    rowData[31] = formData.kaigoG;
    rowData[32] = formData.kaigoGDate;
    rowData[33] = formData.kaigoN;
    rowData[34] = formData.kaigoNDate;
    rowData[35] = formData.otherTest;
    rowData[36] = formData.otherTestDate;
    rowData[37] = formData.otherSkill;
    rowData[38] = formData.otherSkillDate;
    rowData[39] = formData.memo;
    rowData[40] = oldComment;
    rowData[41] = formData.family;
    rowData[42] = formData.interviewHist;
    rowData[43] = formData.status;
    rowData[44] = formData.hiredBy;
    rowData[45] = formData.skillField;
    rowData[46] = formData.agent;
    rowData[47] = formData.regDate;
    rowData[48] = formData.hireDate;
    rowData[49] = formData.birthPlace;
    rowData[50] = formData.addressDetail;
    rowData[51] = formData.passportNo;
    rowData[52] = formData.passportExp;
    rowData[53] = formData.job;
    rowData[54] = formData.expJissyu;
    rowData[55] = formData.certJissyu;
    rowData[56] = formData.crime;
    rowData[57] = formData.visaApplyCount;
    rowData[58] = formData.visaRejectCount;
    rowData[59] = formData.overseasHist;
    rowData[60] = formData.overseasCount;
    rowData[61] = formData.lastEntry;
    rowData[62] = formData.lastExit;
    rowData[63] = formData.relName;
    rowData[64] = formData.relType;
    rowData[65] = formData.relBirth;
    rowData[66] = formData.relNat;
    rowData[67] = formData.relLive;
    rowData[68] = formData.relWork;
    rowData[69] = formData.relCard;
    rowData[70] = formData.generalMemo;

    for (let i = 71; i < currentValues.length; i++) {
      rowData[i] = currentValues[i];
    }

    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    return "候補者情報を更新しました。";
  } catch(e) { throw new Error(e.message); }
}

/**
 * 候補者の削除
 */
function deleteCandidateRow(candId) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) throw new Error("「登録者マスタ」シートが見つかりません。");
    
    const data = sheet.getDataRange().getValues();
    const id = String(candId).trim();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === id) {
        sheet.deleteRow(i + 1);
        return `登録者 ${id} を削除しました。`;
      }
    }
    throw new Error("対象の登録者が見つかりませんでした。");
  } catch (e) {
    throw new Error("削除エラー: " + e.message);
  }
}