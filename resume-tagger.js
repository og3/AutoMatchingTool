/***** 履歴書→（職種/業界/スキル） 抽出：DICT参照版 *****/
const SHEET_NAME_RESUME = '人材DB';
const START_ROW_RESUME  = 2;  // 見出しの次
const SRC_COL_RESUME    = 2;  // A=履歴書原文
const ROLE_COL_RESUME   = 13;  // B=経験職種
const IND_COL_RESUME    = 14;  // C=経験業界
const SKL_COL_RESUME    = 15;  // D=スキル

const DICT_SHEET_NAME   = 'DICT';

/* --------- 共通ユーティリティ（衝突回避のため Resume 接尾辞） --------- */
function normalizeResume_(s){
  if(!s) return '';
  return String(s)
    .replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0)-0xFEE0))
    .replace(/\u3000/g, ' ')
    .toLowerCase();
}
function escapeRegExpResume_(s){ return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function boundaryRegexResume_(key){
  return new RegExp('(^|[^a-z0-9#\\+])' + escapeRegExpResume_(key) + '($|[^a-z0-9#\\+])','i');
}
function normalizeCategoryResume_(raw){
  const t = String(raw||'').trim().toLowerCase();
  if(['roles','role','職種','経験職種'].includes(t)) return 'roles';
  if(['industries','industry','業界','経験業界'].includes(t)) return 'industries';
  if(['skills','skill','スキル'].includes(t)) return 'skills';
  return '';
}

/* --------- DICT 読み込み --------- */
function loadDictResume_(){
  const sh = SpreadsheetApp.getActive().getSheetByName(DICT_SHEET_NAME);
  if(!sh) throw new Error('DICTシートが見つかりません: ' + DICT_SHEET_NAME);
  const values = sh.getDataRange().getValues();
  if(values.length < 2) return { roles:[], industries:[], skills:[] };

  const header = values[0].map(h => String(h||'').trim().toLowerCase());
  const iCanonical = header.indexOf('canonical');
  const iCategory  = header.indexOf('category');
  const iAliases   = header.indexOf('aliases');
  const iRegex     = header.indexOf('regex');

  if(iCanonical === -1 || iCategory === -1){
    throw new Error('DICT には canonical / category が必須です');
  }

  const buckets = { roles:[], industries:[], skills:[] };

  for(let r=1; r<values.length; r++){
    const row = values[r];
    const canon = String(row[iCanonical]||'').trim();
    const cat   = normalizeCategoryResume_(row[iCategory]);
    if(!canon || !cat) continue;

    const aliasRaw = iAliases>-1 ? String(row[iAliases]||'') : '';
    const aliasList = aliasRaw.split(/[,，、]/).map(s=>s.trim()).filter(Boolean);
    const regexRaw = iRegex>-1 ? String(row[iRegex]||'').trim() : '';

    const testers = [];
    if(regexRaw){
      try{
        testers.push({ kind:'regex', re:new RegExp(regexRaw,'i') });
      }catch(e){
        // 無効な正規表現は無視
      }
    }
    const keys = [canon, ...aliasList].map(w=>normalizeResume_(w)).filter(Boolean);
    const seen = new Set();
    for(const k of keys){
      if(seen.has(k)) continue; seen.add(k);
      testers.push({ kind:'word', re: boundaryRegexResume_(k), key:k });
    }
    buckets[cat].push({ canon, testers });
  }
  return buckets;
}

/* --------- カテゴリ別 前処理 --------- */
// industries は「〜の業界」「〜業界」を除去してから判定
function preprocForCategoryResume_(text, cat){
  let t = String(text||'');
  if(cat === 'industries'){
    t = t.replace(/の業界/g, '');
    t = t.replace(/業界\b/g, '');
  }
  return t;
}

/* --------- 照合 --------- */
function matchCategoryResume_(text, entries, cat){
  if(!text || !entries || !entries.length) return [];
  const pre  = preprocForCategoryResume_(text, cat);
  const norm = normalizeResume_(pre);
  const hits = [];
  for(const {canon, testers} of entries){
    let ok=false;
    for(const t of testers){
      if(t.kind === 'regex'){
        if(t.re.test(pre) || t.re.test(norm)){ ok=true; break; }
      }else{
        if(t.re.test(norm)){ ok=true; break; }
      }
    }
    if(ok) hits.push(canon);
  }
  return Array.from(new Set(hits));
}

/* --------- 1セル処理 --------- */
function parseResumeCellWithDict_(text, dict){
  const roles  = matchCategoryResume_(text, dict.roles,      'roles');
  const inds   = matchCategoryResume_(text, dict.industries, 'industries');
  const skills = matchCategoryResume_(text, dict.skills,     'skills');
  return { roles, industries: inds, skills };
}

/* --------- 全行処理（A→B/C/D） --------- */
function runParseAllResume(){
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_RESUME);
  if(!sh) throw new Error('シートが見つかりません: ' + SHEET_NAME_RESUME);
  const lastRow = sh.getLastRow();
  if(lastRow < START_ROW_RESUME) return;

  const dict = loadDictResume_();

  const num = lastRow - START_ROW_RESUME + 1;
  const values = sh.getRange(START_ROW_RESUME, SRC_COL_RESUME, num, 1).getValues();

  const outRoles = [], outInds = [], outSkls = [];
  for(const [raw] of values){
    const {roles, industries, skills} = parseResumeCellWithDict_(raw||'', dict);
    outRoles.push([roles.join(', ')]);
    outInds.push([industries.join(', ')]);
    outSkls.push([skills.join(', ')]);
  }

  sh.getRange(START_ROW_RESUME, ROLE_COL_RESUME, outRoles.length, 1).setValues(outRoles);
  sh.getRange(START_ROW_RESUME, IND_COL_RESUME,  outInds.length,  1).setValues(outInds);
  sh.getRange(START_ROW_RESUME, SKL_COL_RESUME,  outSkls.length,  1).setValues(outSkls);
}

/* --------- 編集行のみ更新 --------- */
function onEditResume(e){
  try{
    const sh = e.range.getSheet();
    if(sh.getName() !== SHEET_NAME_RESUME) return;
    if(e.range.getColumn() !== SRC_COL_RESUME) return;
    const row = e.range.getRow();
    if(row < START_ROW_RESUME) return;

    const dict = loadDictResume_();
    const text = e.range.getValue();
    const {roles, industries, skills} = parseResumeCellWithDict_(text, dict);

    sh.getRange(row, ROLE_COL_RESUME).setValue(roles.join(', '));
    sh.getRange(row, IND_COL_RESUME ).setValue(industries.join(', '));
    sh.getRange(row, SKL_COL_RESUME ).setValue(skills.join(', '));
  }catch(err){
    console.error(err);
  }
}

/* --------- メニュー（衝突回避のため onOpen は定義しない） --------- */
function addResumeMenu(){
  SpreadsheetApp.getUi()
    .createMenu('職務経歴書パーサー')
    .addItem('職務経歴書コピペ→3軸分け', 'runParseAllResume')
    .addToUi();
}
