/****************************************************
 * 求人3軸 × 人材3軸 のヒット数＆内訳（文付き）出力スクリプト
 * DICT対応版：tagsをDICTで正規化しcanonical一致で突合
 * フィルタ：人材DBのS列「ステータス」が「面談済み」のみ出力
 * 使い方：
 * 1) 求人票シートで採点したい行を選択
 * 2) scoreForActiveJobRowDICT() を実行
 *    （または addMatchingMenu() を1回実行してメニューから）
 ****************************************************/

/***** 設定：シート・列 *****/

// 求人票の3軸（前段のタグ化出力を想定：B=職種, C=業界, D=スキル）
const JOB_SHEET_NAME = '案件入力窓';
const JOB_START_ROW  = 2;
const JOB_COL_ROLES  = 3;  // B: 職種（環境に合わせて）
const JOB_COL_INDS   = 4;  // C: 業界
const JOB_COL_SKILLS = 5;  // D: スキル

// 人材DB（列は実シートの配置に合わせる）
const CAND_SHEET_NAME = '人材DB';
const CAND_START_ROW  = 2;
const CAND_COL_NAME   = 1;   // A: 氏名
const CAND_COL_RESUME = 2;   // B: 職務経歴書（必要に応じて変更）

// ★ご指定の追加出力項目（「人材DB」上の列）
const CAND_COL_WORKSTYLE     = 6;   // F: ワークスタイル
const CAND_COL_EMPLOYMENT    = 7;   // G: 雇用形態
const CAND_COL_PREF_TIME     = 8;   // H: 希望時間帯
const CAND_COL_MONTHLY_HOURS = 9;   // I: 月間稼働時間
const CAND_COL_MONTHLY_RATE  = 10;  // J: 月間稼働単価
const CAND_COL_ALLOW_BELOW   = 11;  // K: 単価以下許容
const CAND_COL_DESIRED_ROLES = 12;  // L: 希望職種
const CAND_COL_ROLES         = 13;  // M: 経験職種
const CAND_COL_INDS          = 14;  // N: 経験業界
const CAND_COL_SKILLS        = 15;  // O: スキル
const CAND_COL_CERTS         = 16;  // P: 保有資格

// ★フィルタ用：S列（ステータス）
const CAND_COL_STATUS        = 19;  // S: ステータス（「面談済み」以外は除外）

// 結果出力シート（毎回新規作成＋ローテーション）
const OUT_SHEET_PREFIX   = 'マッチング結果';   // ベース名
const OUT_SHEET_KEEP_MAX = 5;                 // 保持する最新シート枚数（調整可）
const OUT_SHEET_TIMEZONE = Session.getScriptTimeZone() || 'Asia/Tokyo';

/***** 表示調整 *****/
const MAX_SENT_LEN = 255; // 一文の最大表示長（超えると…で省略）

/***** ユーティリティ（このファイル専用。名称は *_DICT で衝突回避） *****/

// 全角→半角（英数記号）/全角スペース→半角/最低限のゆれ吸収→小文字化
function _normToken_DICT_(s) {
  if (!s) return '';
  s = String(s)
    .replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\u3000/g, ' ');
  // 最低限の表記ゆれ（DICT側regex/aliasesが主役）
  s = s.replace(/(node ?\.?js)/ig, 'node.js')
       .replace(/c[#＃]/ig, 'c#')
       .replace(/ｖue\.?js/ig,'vue.js')
       .replace(/ｗordpress|ワードプレス/ig,'wordpress')
       .replace(/ＥＣ|ｅｃ/ig,'EC')
       .replace(/route\s*53/ig,'route 53')
       .replace(/ＷＥＢ|ｗｅｂ|WEB/ig,'Web')
       .replace(/ＩＴ/ig,'IT');
  return s.toLowerCase().trim();
}
function _escapeRe_DICT_(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function _splitTags_DICT_(raw){
  if (!raw) return [];
  return String(raw).split(/[,\u3001\/\n]/).map(t=>_normToken_DICT_(t)).filter(Boolean);
}
function _clipSentence_DICT_(s) {
  if (!s) return '';
  let t = String(s).replace(/\s+/g, ' ').trim();
  if (t.length > MAX_SENT_LEN) t = t.slice(0, MAX_SENT_LEN) + '…';
  return t;
}

// 出力シートの用意
// 新しい結果シートを毎回作る（重複しないタイムスタンプ名）＋古い結果を整理
function _createOutSheetNew_DICT_() {
  const ss = SpreadsheetApp.getActive();

  // タイムスタンプ付き名前：マッチング結果_YYYYMMDD_HHmmss
  const ts = Utilities.formatDate(new Date(), OUT_SHEET_TIMEZONE, 'yyyyMMdd_HHmmss');
  const base = `${OUT_SHEET_PREFIX}_${ts}`;

  let sheetName = base;
  let suffix = 1;
  while (ss.getSheetByName(sheetName)) {
    suffix += 1;
    sheetName = `${base}_${suffix}`;
  }

  const out = ss.insertSheet(sheetName);
  out.setFrozenRows(1);
  // 古い結果シートのガーベジコレクション
  _gcOldOutSheets_();
  return out;
}

// 余分な「マッチング結果_*」シートを自動削除（新しい順でKEEP_MAXだけ残す）
function _gcOldOutSheets_() {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheets().map(s => s.getName());

  // 対象は prefix 一致（無印 'マッチング結果' も古い扱い）
  const targets = all
    .filter(n => n === OUT_SHEET_PREFIX || n.startsWith(OUT_SHEET_PREFIX + '_'))
    .sort((a, b) => a.localeCompare(b)); // 文字列昇順＝古い→新しい

  const over = Math.max(0, targets.length - OUT_SHEET_KEEP_MAX);
  for (let i = 0; i < over; i++) {
    const name = targets[i];
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  }
}

/***** DICT 読み込み＆テスター構築 *****/
function _normalizeCategory_DICT_(raw){
  const t = String(raw||'').trim().toLowerCase();
  if(['roles','role','職種','経験職種'].includes(t)) return 'roles';
  if(['industries','industry','業界','経験業界'].includes(t)) return 'industries';
  if(['skills','skill','スキル'].includes(t)) return 'skills';
  return '';
}
function _parseAliases_DICT_(raw){
  const s = String(raw||'').trim();
  if(!s) return [];
  return s.split(/[,\u3001\/\n]/).map(x=>_normToken_DICT_(x)).filter(Boolean);
}
function _buildDictTester_DICT_(dictSheetName){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(dictSheetName);
  if(!sh) throw new Error('DICTシート('+dictSheetName+')が見つかりません');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const values  = sh.getRange(1,1,lastRow,lastCol).getDisplayValues();
  if(values.length < 2) return { roles:[], industries:[], skills:[] };

  const header = values[0].map(h => String(h||'').trim().toLowerCase());
  const iCanonical = header.indexOf('canonical');
  const iCategory  = header.indexOf('category');
  const iAliases   = header.indexOf('aliases');
  const iRegex     = header.indexOf('regex');

  if(iCanonical === -1 || iCategory === -1){
    throw new Error('DICTには canonical / category が必須です');
  }

  const buckets = { roles:[], industries:[], skills:[] };

  for(let r=1; r<values.length; r++){
    const row = values[r];
    const canon = String(row[iCanonical]||'').trim();
    const cat   = _normalizeCategory_DICT_(row[iCategory]||'');
    if(!canon || !cat) continue;

    const aliases = (iAliases>=0) ? _parseAliases_DICT_(row[iAliases]) : [];
    const regex   = (iRegex>=0) ? String(row[iRegex]||'').trim() : '';

    let testers = [];
    // canonical 自身も一致対象に含める
    testers.push(new RegExp(`\\b${_escapeRe_DICT_(_normToken_DICT_(canon))}\\b`, 'i'));
    aliases.forEach(a=>{
      if(a) testers.push(new RegExp(`\\b${_escapeRe_DICT_(a)}\\b`, 'i'));
    });
    if(regex){
      try{
        testers.push(new RegExp(regex, 'i'));
      }catch(e){
        // 不正な正規表現は無視
      }
    }
    buckets[cat].push({ canonical: canon, testers });
  }
  return buckets;
}

function _tagsToCanonicalSet_DICT_(tagsRaw, dictBucket, cat){
  const tokens = _splitTags_DICT_(tagsRaw);
  if(tokens.length === 0) return new Set();
  const set = new Set();
  for(const token of tokens){
    let matched = false;
    for(const ent of dictBucket){
      if(ent.testers.some(re => re.test(token))){
        set.add(ent.canonical);
        matched = true;
        break;
      }
    }
    if(!matched){
      // 未収載語は「そのまま」も入れておく（将来の辞書補完用）
      set.add(token);
    }
  }
  return set;
}

function _detailFromResume_DICT_(resume, hitWords){
  if(!resume) return '';
  const text = String(resume).replace(/\r?\n/g,' ');
  // 1文単位にざっくり分割（。．.!? など）
  const sentences = text.split(/(?<=[。．.!?？])\s*/);
  const lower = text.toLowerCase();
  const lowerHits = hitWords.map(w=>String(w).toLowerCase());
  // 最初に見つかったヒット語を含む文を返す
  for(const sent of sentences){
    for(const w of lowerHits){
      if(String(sent).toLowerCase().includes(w)){
        return `${w}｜${_clipSentence_DICT_(sent)}`;
      }
    }
  }
  // 見つからなければ先頭から要約
  return (sentences.length ? `-｜${_clipSentence_DICT_(sentences[0])}` : '');
}

/***** 本体：求人票の指定行に対して人材DBを採点（DICT版） *****/
function scoreCandidatesForJobRowDICT(jobRow){
  const ss = SpreadsheetApp.getActive();
  const jobSh  = ss.getSheetByName(JOB_SHEET_NAME);
  const candSh = ss.getSheetByName(CAND_SHEET_NAME);
  if(!jobSh || !candSh) throw new Error('必要なシートが見つかりません。');

  if(!jobRow || jobRow < JOB_START_ROW) throw new Error('求人票の対象行を正しく選択してください。');

  // 求人側（3軸）
  const jobRolesOrig  = jobSh.getRange(jobRow, JOB_COL_ROLES).getDisplayValue();
  const jobIndsOrig   = jobSh.getRange(jobRow, JOB_COL_INDS).getDisplayValue();
  const jobSkillsOrig = jobSh.getRange(jobRow, JOB_COL_SKILLS).getDisplayValue();

  // DICT の読み込み
  const DICT = _buildDictTester_DICT_('DICT'); // ←辞書シート名に合わせる
  const jobRolesCanon  = _tagsToCanonicalSet_DICT_(jobRolesOrig,  DICT.roles,      'roles');
  const jobIndsCanon   = _tagsToCanonicalSet_DICT_(jobIndsOrig,   DICT.industries, 'industries');
  const jobSkillsCanon = _tagsToCanonicalSet_DICT_(jobSkillsOrig, DICT.skills,     'skills');

  // 人材側（全件）
  const lastRow = candSh.getLastRow();
  if (lastRow < CAND_START_ROW) throw new Error('人材DBにデータがありません。');

  const rowCount = lastRow - CAND_START_ROW + 1;

  // 基本フィールド
  const names        = candSh.getRange(CAND_START_ROW, CAND_COL_NAME,   rowCount, 1).getDisplayValues().flat();
  const rolesV       = candSh.getRange(CAND_START_ROW, CAND_COL_ROLES,  rowCount, 1).getDisplayValues().flat();
  const indsV        = candSh.getRange(CAND_START_ROW, CAND_COL_INDS,   rowCount, 1).getDisplayValues().flat();
  const skillsV      = candSh.getRange(CAND_START_ROW, CAND_COL_SKILLS, rowCount, 1).getDisplayValues().flat();
  const resumeV      = candSh.getRange(CAND_START_ROW, CAND_COL_RESUME, rowCount, 1).getDisplayValues().flat();

  // 追加フィールド
  const workstyleV     = candSh.getRange(CAND_START_ROW, CAND_COL_WORKSTYLE,     rowCount, 1).getDisplayValues().flat();
  const employmentV    = candSh.getRange(CAND_START_ROW, CAND_COL_EMPLOYMENT,    rowCount, 1).getDisplayValues().flat();
  const prefTimeV      = candSh.getRange(CAND_START_ROW, CAND_COL_PREF_TIME,     rowCount, 1).getDisplayValues().flat();
  const monthlyHoursV  = candSh.getRange(CAND_START_ROW, CAND_COL_MONTHLY_HOURS, rowCount, 1).getDisplayValues().flat();
  const monthlyRateV   = candSh.getRange(CAND_START_ROW, CAND_COL_MONTHLY_RATE,  rowCount, 1).getDisplayValues().flat();
  const allowBelowV    = candSh.getRange(CAND_START_ROW, CAND_COL_ALLOW_BELOW,   rowCount, 1).getDisplayValues().flat();
  const desiredRolesV  = candSh.getRange(CAND_START_ROW, CAND_COL_DESIRED_ROLES, rowCount, 1).getDisplayValues().flat();
  const certsV         = candSh.getRange(CAND_START_ROW, CAND_COL_CERTS,         rowCount, 1).getDisplayValues().flat();

  // フィルタ用：ステータス（S列）
  const statusV        = candSh.getRange(CAND_START_ROW, CAND_COL_STATUS,        rowCount, 1).getDisplayValues().flat();

  const rows = [];

  for (let i=0; i<rowCount; i++) {
    const status = String(statusV[i]||'').trim();
    if (status !== '面談済') continue; // フィルタ：面談済みのみ

    const candRolesCanon  = _tagsToCanonicalSet_DICT_(rolesV[i],  DICT.roles,      'roles');
    const candIndsCanon   = _tagsToCanonicalSet_DICT_(indsV[i],   DICT.industries, 'industries');
    const candSkillsCanon = _tagsToCanonicalSet_DICT_(skillsV[i], DICT.skills,     'skills');

    const rolesInter  = [...candRolesCanon].filter(x => jobRolesCanon.has(x));
    const indsInter   = [...candIndsCanon].filter(x => jobIndsCanon.has(x));
    const skillsInter = [...candSkillsCanon].filter(x => jobSkillsCanon.has(x));

    const hitWords = [...new Set([...rolesInter, ...indsInter, ...skillsInter])];
    if (hitWords.length === 0) continue;

    rows.push({
      name: names[i],
      rolesHit:  rolesInter.length,
      indsHit:   indsInter.length,
      skillsHit: skillsInter.length,
      detail: _detailFromResume_DICT_(resumeV[i], hitWords),
      workstyle:    workstyleV[i],
      employment:   employmentV[i],
      prefTime:     prefTimeV[i],
      monthlyHours: monthlyHoursV[i],
      monthlyRate:  monthlyRateV[i],
      allowBelow:   allowBelowV[i],
      desiredRoles: desiredRolesV[i],
      expRolesRaw:  rolesV[i],
      expIndsRaw:   indsV[i],
      expSkillsRaw: skillsV[i],
      certs:        certsV[i],
    });
  }

  // 並び：総ヒットの降順→職種→業界→スキル
  rows.sort((a,b)=>
    (b.rolesHit + b.indsHit + b.skillsHit) - (a.rolesHit + a.indsHit + a.skillsHit) ||
    b.rolesHit - a.rolesHit || b.indsHit - a.indsHit || b.skillsHit - a.skillsHit
  );

  // 出力（ご指定の順序）
  const header = [
    '氏名',
    '職種ヒット数','業界ヒット数','スキルヒット数',
    'ヒット内訳（語｜職務経歴書の一文）',
    'ワークスタイル','雇用形態','希望時間帯',
    '月間稼働時間','月間稼働単価','単価以下許容',
    '希望職種','経験職種','経験業界','スキル','保有資格'
  ];

  const values = rows.map(r => [
    r.name,
    r.rolesHit, r.indsHit, r.skillsHit,
    r.detail,
    r.workstyle, r.employment, r.prefTime,
    r.monthlyHours, r.monthlyRate, r.allowBelow,
    r.desiredRoles, r.expRolesRaw, r.expIndsRaw, r.expSkillsRaw, r.certs
  ]);

  const outSheet = _createOutSheetNew_DICT_(); // 新規作成＋古い結果整理
  outSheet.clearContents();
  outSheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
  if (values.length) outSheet.getRange(2,1,values.length, header.length).setValues(values);
  outSheet.autoResizeColumns(1, header.length);
}

/***** アクティブ行を使うランチャー（DICT版） *****/
function scoreForActiveJobRowDICT() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== JOB_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('先に求人票シート（' + JOB_SHEET_NAME + '）で対象行を選択してください。');
    return;
  }
  const row = sh.getActiveCell().getRow();
  scoreCandidatesForJobRowDICT(row);
}

/***** メニュー（任意） *****/
function addMatchingMenu() {
  SpreadsheetApp.getUi()
    .createMenu('マッチング')
    .addItem('求人票と人材をマッチング', 'scoreForActiveJobRowDICT')
    .addSeparator()
    .addItem('古い結果を整理する', '_gcOldOutSheets_')
    .addToUi();
}
// 一度これを実行するとメニューが出ます（以後はメニューから実行可）
// function showMenuNowDICT() { addMatchingMenu(); }
