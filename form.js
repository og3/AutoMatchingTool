/** 自動マッチングツール：フォーム → 「人材DB」書き込み（GAS）
 *  - シート名:
 *    DB_SHEET_NAME = '人材DB'
 *    DIC_SHEET_NAME = 'DICT'  // ← 固定
 *
 *  - 「辞書」category は rolls / industries / skills を想定（roles→rolls の表記ゆれ吸収）
 *  - サジェストは canonical のみ
 *  - 送信時は aliases/表記ゆれを canonical に正規化して保存
 */

const DB_SHEET_NAME  = '人材DB';
const DIC_SHEET_NAME = 'DICT';

/** 画面提供（Thanks.html は使わず常に Index を返す。表示切替は Index 側JS） */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.baseUrl = ScriptApp.getService().getUrl(); // 使わなければ無視される
  return t.evaluate()
           .setTitle('人材登録フォーム')
           .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** submitForm：成功時は {status:'ok'} を返す */
function submitForm(form) {
  const payload = {
    name: String(form.name || '').trim(),                            // (1) 氏名
    dob: String(form.dob || '').trim(),                              // (3) 生年月日 (YYYY-MM-DD)
    email: String(form.email || '').trim(),                          // (4) メール
    tel: String(form.tel || '').trim(),                              // (5) TEL
    workstyle: String(form.workstyle || '').trim(),                  // (6) ワークスタイル
    employment: String(form.employment || '').trim(),                // (7) 雇用形態
    timeRange: String(form.timeRange || '').trim(),                  // (8) 希望時間帯 (hh:mm~hh:mm)
    monthlyHours: form.monthlyHours !== undefined && form.monthlyHours !== null && String(form.monthlyHours).trim() !== '' ? Number(form.monthlyHours) : '',
    monthlyUnit:  form.monthlyUnit  !== undefined && form.monthlyUnit  !== null && String(form.monthlyUnit).trim()  !== '' ? Number(form.monthlyUnit)  : '',
    acceptLower: !!form.acceptLower,                                 // (11) 許容フラグ
    desiredFree: String(form.desiredFree || '').trim(),              // (12) 希望職種（自由）
    expRoles: String(form.expRoles || '').trim(),                    // (13) 経験職種 CSV
    expIndustries: String(form.expIndustries || '').trim(),          // (14) 経験業界 CSV
    skills: String(form.skills || '').trim(),                        // (15) スキル CSV
    certsFree: String(form.certsFree || '').trim(),                  // (16) 保有資格（自由）
  };

  // 必須チェック（「自由」以外）
  const required = {
    name: '氏名',
    dob: '生年月日',
    email: 'メールアドレス',
    tel: 'TEL',
    workstyle: 'ワークスタイル',
    employment: '雇用形態',
    timeRange: '希望稼働時間帯',
    monthlyHours: '希望月間稼働時間',
    monthlyUnit: '希望月間稼働単価',
    expRoles: '経験職種',
    expIndustries: '経験業界',
    skills: 'スキル',
  };
  const missing = Object.entries(required).filter(([k]) => !payload[k] && payload[k] !== 0);
  if (missing.length) {
    const names = missing.map(([, label]) => label).join('、');
    throw new Error('未入力の必須項目があります：' + names);
  }

  // 形式チェック
  const timeOk = /^([01]\d|2[0-3]):[0-5]\d~([01]\d|2[0-3]):[0-5]\d$/.test(payload.timeRange);
  if (!timeOk) throw new Error('希望稼働時間帯は hh:mm~hh:mm の形式で入力してください（例: 09:00~18:00）');

  if (typeof payload.monthlyHours === 'number' && !(payload.monthlyHours > 0)) {
    throw new Error('希望月間稼働時間は 1 以上の数値で入力してください');
  }
  if (typeof payload.monthlyUnit === 'number' && !(payload.monthlyUnit >= 0)) {
    throw new Error('希望月間稼働単価は 0 以上の数値で入力してください');
  }

  // 正規化（サーバ側で canonical に寄せる）
  const dic = getDictionaries(); // maps: {rolls, industries, skills}
  payload.expRoles      = _canonicalizeCsvServer_(payload.expRoles,      dic.maps.rolls);
  payload.expIndustries = _canonicalizeCsvServer_(payload.expIndustries, dic.maps.industries);
  payload.skills        = _canonicalizeCsvServer_(payload.skills,        dic.maps.skills);

  // 書き込み
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName(DB_SHEET_NAME);
  if (!db) throw new Error('「人材DB」シートが見つかりません');

  // 1始まりカラムに対応した配列（A=1 … P=16）。B列(2)は仕様上未使用のため空。
  const row = new Array(16).fill('');
  row[0]  = payload.name;                 // (1) A 氏名
  // row[1] = ''                          // (2) B 未使用
  row[2]  = payload.dob;                  // (3) C 生年月日
  row[3]  = payload.email;                // (4) D メール
  row[4]  = payload.tel;                  // (5) E TEL
  row[5]  = payload.workstyle;            // (6) F ワークスタイル
  row[6]  = payload.employment;           // (7) G 雇用形態
  row[7]  = payload.timeRange;            // (8) H 希望稼働時間帯
  row[8]  = payload.monthlyHours;         // (9) I 希望月間稼働時間
  row[9]  = payload.monthlyUnit;          // (10) J 希望月間稼働単価
  row[10] = payload.acceptLower ? true : false; // (11) K 許容フラグ
  row[11] = payload.desiredFree;          // (12) L 希望職種（自由）
  row[12] = payload.expRoles;             // (13) M 経験職種（CSV, canonical）
  row[13] = payload.expIndustries;        // (14) N 経験業界（CSV, canonical）
  row[14] = payload.skills;               // (15) O スキル（CSV, canonical）
  row[15] = payload.certsFree;            // (16) P 保有資格（自由）

  db.appendRow(row);

  return { status: 'ok' }; // フロントの withSuccessHandler を確実に起動
}

/** 全角→半角(ASCII)/trim/小文字化 */
function _normKey_(s) {
  if (!s) return '';
  s = String(s).trim()
    .replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\u3000/g, ' ');
  return s.toLowerCase();
}

/** category 正規化（roles→rolls の吸収など） */
function _normCat_(s) {
  s = (s || '').toString().trim().toLowerCase();
  return s === 'roles' ? 'rolls' : s;
}

/** 「DICT」シートを読み出し
 *  - datalist 用の canonical 候補（rolls / industries / skills）
 *  - 正規化マップ（alias/表記ゆれ → canonical）
 *  - サジェストは canonical のみ
 */
function getDictionaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DIC_SHEET_NAME);
  if (!sh) throw new Error('「' + DIC_SHEET_NAME + '」シートが見つかりません');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    return { rolls: [], industries: [], skills: [], maps: { rolls:{}, industries:{}, skills:{} } };
  }

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(v => String(v).trim().toLowerCase());

  const idx = {
    canonical: header.findIndex(h => h === 'canonical'),
    category : header.findIndex(h => h === 'category'),
    aliases  : header.findIndex(h => h.startsWith('aliases')),
  };
  if (idx.canonical === -1 || idx.category === -1) {
    throw new Error('「' + DIC_SHEET_NAME + '」シートのヘッダに canonical / category が必要です');
  }

  const setR = new Set(), setI = new Set(), setS = new Set(); // サジェスト（canonicalのみ）
  const mapR = {}, mapI = {}, mapS = {};                       // 正規化マップ

  const addCanonical = (cat, canon) => {
    const c = String(canon || '').trim(); if (!c) return;
    const key = _normKey_(c); if (!key) return;
    if (cat === 'rolls')          { setR.add(c); mapR[key] = c; }
    else if (cat === 'industries'){ setI.add(c); mapI[key] = c; }
    else if (cat === 'skills')    { setS.add(c); mapS[key] = c; }
  };

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const cat = _normCat_(row[idx.category]);
    const canon = row[idx.canonical];
    if (!canon || !['rolls','industries','skills'].includes(cat)) continue;

    // 1) サジェスト：canonical のみ
    addCanonical(cat, canon);

    // 2) 正規化マップ：aliases → canonical
    if (idx.aliases !== -1 && row[idx.aliases]) {
      String(row[idx.aliases])
        .split(/[,，、;；]/)
        .map(s => s.trim())
        .filter(Boolean)
        .forEach(alias => {
          const akey = _normKey_(alias);
          if (!akey) return;
          const cval = String(canon).trim();
          if (cat === 'rolls')           mapR[akey] = cval;
          else if (cat === 'industries') mapI[akey] = cval;
          else if (cat === 'skills')     mapS[akey] = cval;
        });
    }
  }

  const sortJa = (a,b) => a.localeCompare(b, 'ja');

  return {
    rolls: Array.from(setR).sort(sortJa),
    industries: Array.from(setI).sort(sortJa),
    skills: Array.from(setS).sort(sortJa),
    maps: { rolls: mapR, industries: mapI, skills: mapS }
  };
}

/** CSV を canonical に寄せて整形（サーバ側で最終チェック用） */
function _canonicalizeCsvServer_(csv, mapsForCat) {
  const out = [];
  const seen = new Set();
  (String(csv || ''))
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .forEach(tok => {
      const key = _normKey_(tok);
      const canon = (mapsForCat && mapsForCat[key]) || tok; // マップに無ければそのまま
      if (!seen.has(canon)) { seen.add(canon); out.push(canon); }
    });
  return out.join(', ');
}
