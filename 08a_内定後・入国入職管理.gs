// =========================================
// COE申請情報の登録・更新 (ADDINFO)
// =========================================

function updateAddInfoRow(formData) {
  try {
    const sheet = getMasterSheet('登録者マスタ');
    if (!sheet) return "エラー：登録者マスタが見つかりません。 / Error: Master sheet not found.";
    const col = getMasterColumnMap(sheet);
    const row = Number(formData.row);
    if (!row) return "エラー：行が不明です。 / Error: Invalid row.";

    if (formData.eduEnd) formData.eduEnd = normalizeYearMonth(formData.eduEnd);
    if (formData.birthday) formData.birthday = new Date(formData.birthday.replace(/-/g, '/'));

    const safeMaxCol = Math.max(sheet.getLastColumn(), ...Object.values(col));
    const currentRowRange = sheet.getRange(row, 1, 1, safeMaxCol);
    const currentRowData = currentRowRange.getValues()[0];

    const oldName = col['名前'] ? String(currentRowData[col['名前'] - 1]).trim() : "";
    const newName = formData.name !== undefined ? String(formData.name).trim() : oldName;
    const adminId = col['登録者ID'] ? String(currentRowData[col['登録者ID'] - 1]).trim() : "";

    const mapping = {
      name: '名前', nickname: '呼び名', birthday: '生年月日', gender: '性別', spouse: '配偶者',
      eduSchool: '学歴＞学校名', eduDept: '学歴＞学部・学科・専攻',
      eduSchoolEn: '学歴＞学校名（英語）', eduDeptEn: '学歴＞学部・学科・専攻（英語）',
      eduStatus: '学歴＞状況', eduEnd: '学歴＞卒業/中退年月',
      jlptLevel: '特定技能要件＞JLPTレベル', jftLevel: '特定技能要件＞JFTBasicレベル',
      kaigoSkill: '特定技能要件＞介護技能評価試験', kaigoLang: '特定技能要件＞介護日本語評価試験',
      otherExam: '特定技能要件＞その他の評価試験', otherJapanese: 'その他の日本語能力試験',
      agent: '所属送り出し機関', birthCity: '出生地（都市名）',
      addressDetail: '住所詳細', passportNum: 'パスポート番号', passportExp: 'パスポート有効期限',
      job: '職業', traineeExp: '技能実習の経験の有無', traineeCert: '技能実習修了書の有無',
      crime: '犯罪歴の有無', applyCount: '在留資格交付申請の回数', rejectCount: '不許可となった在留資格交付申請の回数',
      overseasExp: '海外への出入国歴の有無', travelCount: '出入国の回数', lastInDate: '直近の入国日', lastOutDate: '直近の出国日',
      relName2: '日本在住の親族情報親族の名前', relRelation2: '日本在住の親族情報続柄',
      relBirth2: '日本在住の親族情報親族の生年月日', relCountry2: '日本在住の親族情報親族の国籍・地域',
      relLive2: '日本在住の親族情報親族との同居予定の有無', relWork2: '日本在住の親族情報親族の勤務先・通学先',
      relCard2: '日本在住の親族情報親族の在留カード番号', memo: '備考・メモ'
    };

    for (let key in mapping) {
      const h = mapping[key].replace(/\s/g, '');
      if (col[h] && formData[key] !== undefined) {
        currentRowData[col[h] - 1] = formData[key];
      }
    }
    
    const skipIndices = [];
    if (col['顔写真']) skipIndices.push(col['顔写真'] - 1);
    if (col['満年齢']) skipIndices.push(col['満年齢'] - 1);
    
    skipIndices.sort((a, b) => a - b);
    
    let startCol = 0;
    skipIndices.forEach(skipIdx => {
      if (skipIdx > startCol) {
        const length = skipIdx - startCol;
        sheet.getRange(row, startCol + 1, 1, length).setValues([currentRowData.slice(startCol, skipIdx)]);
      }
      startCol = skipIdx + 1;
    });
    if (startCol < safeMaxCol) {
      const length = safeMaxCol - startCol;
      sheet.getRange(row, startCol + 1, 1, length).setValues([currentRowData.slice(startCol)]);
    }

    if (col['生年月日'] && formData.birthday) {
      sheet.getRange(row, col['生年月日']).setNumberFormat('yyyy"年"m"月"d"日"');
    }

    if (adminId && oldName && newName && oldName !== newName) {
      updateJobCandidateName(adminId, newName);
    }

    SpreadsheetApp.flush();

    // ▼ 外部シート「COE info」へ該当候補者の最新情報を即時転記・更新
    if (adminId) {
      syncCandidatesToCoeInfo([adminId]);
    }

    return `COE申請情報の登録・更新（および外部「COE info」シートへの転記）が完了しました。\nCOE application information saved and synced successfully.`;
  } catch (e) { return "エラー / Error: " + e.message; }
}