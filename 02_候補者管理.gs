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

    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      let idStr = String(data[i][0]);
      if (idStr.startsWith("SD-")) {
        let num = parseInt(idStr.replace("SD-", ""), 10);
        if (num > maxId) maxId = num;
      }
    }
    const nextId = "SD-" + (maxId + 1).toString().padStart(4, '0');
    
    // 現在の最新の列数（71列）に合わせた空配列を作成
    const rowData = new Array(71).fill('');
    rowData[0] = nextId;
    rowData[1] = formData.name;
    rowData[3] = formData.kana;
    rowData[4] = formData.nickname;
    rowData[5] = formData.birthday;
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
    rowData[41] = formData.family;
    rowData[43] = '未採用'; 

    sheet.appendRow(rowData);
    return `登録者 ${nextId} を新規登録しました。`;
  } catch (e) {
    throw new Error("登録エラー: " + e.message);
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

    // 既存の顔写真(C列)と修正前コメント(AO列)を保持
    const currentValues = sheet.getRange(row, 1, 1, 71).getValues()[0];
    const photo = currentValues[2];
    const oldComment = currentValues[40];

    const rowData = [
      formData.id, formData.name, photo, formData.kana, formData.nickname,
      formData.birthday, '', formData.gender, formData.marriage,
      formData.height, formData.weight, formData.address, formData.origin,
      formData.email, formData.school, formData.eduName, formData.eduMajor,
      formData.eduStatus, formData.eduStart, formData.eduEnd, formData.eduMemo,
      formData.work1Period, formData.work1Detail, formData.work2Period, formData.work2Detail,
      formData.work3Period, formData.work3Detail,
      formData.jlpt, formData.jlptDate, formData.jft, formData.jftDate,
      formData.kaigoG, formData.kaigoGDate, formData.kaigoN, formData.kaigoNDate,
      formData.otherTest, formData.otherTestDate, formData.otherSkill, formData.otherSkillDate,
      formData.memo, oldComment, formData.family, formData.interviewHist,
      formData.status, formData.hiredBy, formData.skillField, formData.agent,
      formData.regDate, formData.hireDate, formData.birthPlace, formData.addressDetail,
      formData.passportNo, formData.passportExp, formData.job, formData.expJissyu,
      formData.certJissyu, formData.crime, formData.visaApplyCount, formData.visaRejectCount,
      formData.overseasHist, formData.overseasCount, formData.lastEntry, formData.lastExit,
      formData.relName, formData.relType, formData.relBirth, formData.relNat,
      formData.relLive, formData.relWork, formData.relCard, formData.generalMemo
    ];

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