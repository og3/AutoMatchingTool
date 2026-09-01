/***** 設定 *****/
const SHEET_NAME = '案件入力窓';
const START_ROW  = 2;

// 列: A=求人票原文 → B=必須スキル → C=職種, D=業界, E=スキル
const COL_SRC          = 1;
const COL_OUT_REQUIRED = 2;
const COL_OUT_ROLES    = 3;
const COL_OUT_INDS     = 4;
const COL_OUT_SKILLS   = 5;

const DICT_SHEET = 'DICT';

/***** 抽出ルール *****/
const START_KEYWORDS = ['必須スキル','必須条件','応募資格','応募要件','求めるスキル'];
const STOP_KEYWORDS  = ['歓迎','望ましい','あると尚可','あれば尚可','あると望ましい','求める人物像','人物像','プラス'];
const SKILL_HINTS    = ['経験','スキル','知識','資格','対応可能','できる','歴','使える','年'];

/***** ユーティリティ *****/
function normalize_(s){
  if(!s) return '';
  return String(s)
    .replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0))
    .replace(/\u3000/g, ' ')
    .toLowerCase();
}
function escapeRegExp_(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function boundaryRegex_(key){ return new RegExp('(^|[^a-z0-9#\\+])'+escapeRegExp_(key)+'($|[^a-z0-9#\\+])','i'); }

function normalizeCategory_(raw){
  const t = String(raw||'').trim().toLowerCase();
  if(['roles','role','職種','経験職種'].includes(t)) return 'roles';
  if(['industries','industry','業界','経験業界'].includes(t)) return 'industries';
  if(['skills','skill','スキル'].includes(t)) return 'skills';
  return '';
}

/***** DICT 読み込み（canonical, category, aliases, regex, notes）*****/
function loadDict_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(DICT_SHEET);
  if(!sh) throw new Error('DICTシートが見つかりません');

  const values = sh.getDataRange().getValues();
  if(values.length < 2) return { roles:[], industries:[], skills:[] };

  const header = values[0].map(h => String(h||'').trim().toLowerCase());
  const iCanonical = header.indexOf('canonical');
  const iCategory  = header.indexOf('category');
  const iAliases   = header.indexOf('aliases');
  const iRegex     = header.indexOf('regex');

  if(iCanonical === -1 || iCategory === -1){
    throw new Error('DICTに必須列 canonical / category がありません');
  }

  const buckets = { roles:[], industries:[], skills:[] };

  for(let r=1; r<values.length; r++){
    const row = values[r];
    const canon = String(row[iCanonical]||'').trim();
    const cat   = normalizeCategory_(row[iCategory]);
    if(!canon || !cat) continue; // 不正データはスキップ

    const aliases = iAliases>-1 ? String(row[iAliases]||'').split(/[,，、]/).map(s=>s.trim()).filter(Boolean) : [];
    const regexRaw= iRegex  >-1 ? String(row[iRegex]  ||'').trim() : '';

    const testers = [];

    // 1) regex（任意）— 無効な正規表現は握りつぶす
    if(regexRaw){
      try{
        testers.push({ kind:'regex', re:new RegExp(regexRaw,'i') });
      }catch(e){
        // 無効なregexは無視（必要ならnotesで管理）
      }
    }

    // 2) canonical / aliases を簡易境界で
    const keys = [canon, ...aliases].map(w=>normalize_(w)).filter(Boolean);
    const seen = new Set();
    for(const k of keys){
      if(seen.has(k)) continue; seen.add(k);
      testers.push({ kind:'word', re: boundaryRegex_(k), key:k });
    }

    buckets[cat].push({ canon, testers });
  }

  return buckets;
}

/***** カテゴリ別の前処理：industriesは「〜の業界／〜業界」を除去してから判定 *****/
function preprocForCategory_(text, cat){
  let t = String(text||'');
  if(cat === 'industries'){
    t = t.replace(/の業界/g, ''); // 例：半導体の業界 → 半導体
    t = t.replace(/業界\b/g, ''); // 例：自動車業界 → 自動車
  }
  return t;
}

/***** 求人票（A列）→ 必須スキル（B列）抽出 *****/
function extractRequiredFromSrc_(raw){
  if(typeof raw !== 'string' || !raw.trim()) return '';
  const lines = raw.split(/[\n。・\-■●]/).map(s=>s.trim()).filter(Boolean);

  let inBlock = false;
  const out = [];
  for(const line of lines){
    if(STOP_KEYWORDS.some(w => line.includes(w))) break;
    if(!inBlock && START_KEYWORDS.some(w => line.includes(w))){
      inBlock = true; continue; // 見出し行はスキップ
    }
    if(inBlock && SKILL_HINTS.some(k => line.includes(k)) && line.length <= 100){
      out.push('・' + line);
    }
  }
  return out.join('\n');
}

/***** 必須スキル（B列）→ CDE 仕分け *****/
function classifyFromRequired_(requiredText, dict){
  const text = String(requiredText || '');
  const roles = matchCategory_(text, dict.roles, 'roles');
  const inds  = matchCategory_(text, dict.industries, 'industries');
  const skills= matchCategory_(text, dict.skills, 'skills');
  return { roles, inds, skills };
}

function matchCategory_(text, entries, cat){
  if(!text || !entries || !entries.length) return [];
  const pre  = preprocForCategory_(text, cat);
  const norm = normalize_(pre);
  const hits = [];
  for(const {canon, testers} of entries){
    let ok = false;
    for(const t of testers){
      if(t.kind === 'regex'){
        if(t.re.test(pre) || t.re.test(norm)){ ok = true; break; }
      }else{
        if(t.re.test(norm)){ ok = true; break; }
      }
    }
    if(ok) hits.push(canon);
  }
  return Array.from(new Set(hits));
}

/***** 一括実行：A→B 抽出 → B→CDE 仕分け *****/
function runExtractAndTagAll(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if(!sh) throw new Error('シート「'+SHEET_NAME+'」が見つかりません');

  const lastRow = sh.getLastRow();
  if(lastRow < START_ROW) return;

  const dict = loadDict_();

  const num  = lastRow - START_ROW + 1;
  const srcs = sh.getRange(START_ROW, COL_SRC, num, 1).getValues();

  const outReq = [];
  const outR = [];
  const outI = [];
  const outS = [];

  for(const [src] of srcs){
    const req = extractRequiredFromSrc_(src);
    outReq.push([req]);

    const {roles, inds, skills} = classifyFromRequired_(req, dict);
    outR.push([roles.join(', ')]);
    outI.push([inds.join(', ')]);
    outS.push([skills.join(', ')]);
  }

  sh.getRange(START_ROW, COL_OUT_REQUIRED, outReq.length, 1).setValues(outReq);
  sh.getRange(START_ROW, COL_OUT_ROLES,    outR.length,   1).setValues(outR);
  sh.getRange(START_ROW, COL_OUT_INDS,     outI.length,   1).setValues(outI);
  sh.getRange(START_ROW, COL_OUT_SKILLS,   outS.length,   1).setValues(outS);
}

function addRequiredSkillsMenu(){
  SpreadsheetApp.getUi()
    .createMenu('求人票パーサー')
    .addItem('求人票→必須スキル抽出→3軸仕分け', 'runExtractAndTagAll')
    .addToUi();
}
