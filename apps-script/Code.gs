/**
 * 키친매니저 — 공용 저장소 (구글 스프레드시트 + Apps Script 웹앱)
 *
 * 이 파일은 그대로 붙여 넣으면 됩니다. 고칠 곳이 없습니다.
 * 직원 로그인은 스프레드시트의 '직원' 시트에서 관리합니다.
 *   A열 이름, B열 비밀번호, C열 메모(자유), D열 회사, E열 권한('관리자'면 모든 회사 현장을 봄)
 *   - 공사 마진 현장은 회사별로 나뉩니다. 일반 직원은 자기 회사 현장만 보고 저장합니다.
 *   - 필름·생산도어(재고, 견적 기록, 거래처)는 FILM_ORG 회사 직원과 관리자만 씁니다.
 *   - 직원 추가: 한 줄 추가
 *   - 퇴사/차단: 그 줄 삭제
 *   - 비밀번호 변경: B열 수정
 * 배포: 배포 관리 → 앱이 쓰는 배포 선택 → 연필(수정) → 버전: 새 버전 → 배포
 *       (유형: 웹 앱, 실행 사용자: 나, 액세스 권한: 모든 사용자)
 */

const DEFAULT_PRICES = { p15: 70000, p18: 80000, pura: 39000, rate: 5 };

/* ---------- 시트 정의 ----------
 * 현장(마진계산) 시트는 사람이 보기 쉽게 모든 값을 열로 저장합니다.
 * 재고·재고이력·견적기록은 주요 값 열 + 원본 데이터(JSON) 열로 저장합니다.
 */
const PROJECT_SHEET = '현장';
const PROJECT_COLS = [
  ['id', 'ID'], ['date', '날짜'], ['company', '업체'], ['site', '현장명'],
  ['contract', '총공사금액'],
  ['q15', '15T 소요량'], ['q18', '18T 소요량'], ['qura', '우라 소요량'],
  ['p15', '15T 단가'], ['p18', '18T 단가'], ['pura', '우라 단가'],
  ['pb', 'PB 금액'], ['rate', '기본기물 비율(%)'], ['hw', '기본기물'],
  ['md', 'MD'], ['opt', '선택기물'], ['marble', '대리석'], ['labor', '시공비'],
  ['cost', '비용 총합계'], ['margin', '순마진'], ['marginRate', '이익률(%)'],
  ['author', '작성자'], ['updatedAt', '저장일시'], ['org', '회사']
];
const PROJECT_NUMS = ['contract','q15','q18','qura','p15','p18','pura','pb','rate','hw','md','opt','marble','labor','cost','margin','marginRate'];

const COLLECTIONS = {
  stock: { sheet: '재고', cols: [['catName', '대분류'], ['modelName', '색상'], ['qty', '잔량(m)'], ['warn', '경고기준(m)']] },
  stockLog: { sheet: '재고이력', cols: [['date', '날짜'], ['type', '구분'], ['modelName', '색상'], ['qty', '수량(m)'], ['memo', '메모']] },
  sales: { sheet: '견적기록', cols: [['date', '날짜'], ['type', '유형'], ['client', '거래처'], ['site', '현장'], ['total', '금액']] }
};
const KV_SHEET = '공유값';
const STAFF_SHEET = '직원';
const FILM_ORG = '오션메이드'; // 필름·생산도어 업무를 하는 회사
const SETTINGS_SHEET = '설정';

/* ---------- 요청 처리 ---------- */
function doGet(e) {
  // 데이터는 로그인(POST)으로만 내보냅니다. 여기서는 동작 여부만 알려 줍니다.
  staffSheet();
  return out({ ok: true, app: 'kitchen-manager' });
}

/* ---------- 직원 로그인 ---------- */
const STAFF_HEADERS = ['이름', '비밀번호', '메모', '회사', '권한'];
function staffSheet() {
  const sh = getSheet(STAFF_SHEET, STAFF_HEADERS);
  // 예전 시트(이름·비밀번호·메모만 있던 것)에 회사·권한 머리글 추가
  const head = sh.getRange(1, 1, 1, STAFF_HEADERS.length).getValues()[0];
  if (String(head[3]) !== '회사' || String(head[4]) !== '권한') {
    sh.getRange(1, 4, 1, 2).setValues([['회사', '권한']]);
    sh.getRange(1, 1, 1, STAFF_HEADERS.length).setFontWeight('bold');
  }
  return sh;
}
function readStaff() {
  const sh = staffSheet();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, STAFF_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), pw: String(r[1]).trim(), company: String(r[3]).trim(), admin: String(r[4]).trim() === '관리자' }));
}
// 이름과 비밀번호가 '직원' 시트의 한 줄과 같으면 {name, company, admin}
function checkLogin(name, pw) {
  name = String(name || '').trim(); pw = String(pw || '').trim();
  if (!name || !pw) return null;
  const staff = readStaff();
  const me = staff.find(s => s.name === name && s.pw === pw);
  if (!me) return null;
  const company = me.company || '(회사 미지정)';
  return { name: me.name, company: company, admin: me.admin, film: me.admin || company === FILM_ORG, staff: staff };
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_request' }); }
  const me = checkLogin(body && body.name, body && body.pw);
  if (!me) { Utilities.sleep(800); return out({ ok: false, error: 'bad_login' }); }
  const author = me.name;
  if (body.action === 'load') return out(Object.assign({ ok: true }, readAll(me)));

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    switch (body.action) {
      case 'save': // 마진계산 현장 저장 (여러 건 가능)
        (Array.isArray(body.projects) ? body.projects : [body.project]).forEach(p => {
          if (p && !(p.onlyIfNewer && p.author)) p.author = author; // 저장한 사람은 로그인한 이름으로
          saveProject(p, me);
        });
        break;
      case 'delete': // 마진계산 현장 삭제
        deleteProject(String(body.id || ''), me);
        break;
      case 'prices':
        writePrices(body.prices || {}, me);
        break;
      case 'sync': // 재고·재고이력·견적기록: 바뀐 것만 추가/수정/삭제
        if (!me.film) return out({ ok: false, error: 'forbidden' });
        syncCollection(body.collection, body.upserts || [], body.deletes || [], author, !!body.onlyNew);
        break;
      case 'kv': // 거래처 목록 같은 공유값
        if (!me.film || String(body.key || '').indexOf('prices:') === 0) return out({ ok: false, error: 'forbidden' });
        writeKV(String(body.key || ''), body.value);
        break;
      default:
        return out({ ok: false, error: 'bad_action' });
    }
    return out(Object.assign({ ok: true }, readAll(me)));
  } finally {
    lock.releaseLock();
  }
}

function readAll(me) {
  const all = readProjects(me.staff);
  const projects = me.admin ? all : all.filter(p => p.org === me.company);
  const res = {
    me: { name: me.name, company: me.company, admin: me.admin, film: me.film },
    projects: projects,
    prices: readPrices(me.company),
    // 필름·생산도어 데이터는 해당 회사 직원과 관리자에게만
    stock: me.film ? readCollection('stock') : [],
    stockLog: me.film ? readCollection('stockLog') : [],
    sales: me.film ? readCollection('sales') : [],
    kv: me.film ? publicKV() : {}
  };
  if (me.admin) {
    const orgs = {};
    me.staff.forEach(s => { if (s.company) orgs[s.company] = 1; });
    all.forEach(p => { if (p.org) orgs[p.org] = 1; });
    res.orgs = Object.keys(orgs).sort();
    res.orgPrices = {};
    res.orgs.forEach(o => res.orgPrices[o] = readPrices(o));
  }
  return res;
}

/* ---------- 공통 ---------- */
function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}
// 글자는 앞에 ' 를 붙여 그대로 저장 (날짜 자동변환·수식 실행 방지)
function txt(v) { return v == null || v === '' ? '' : "'" + String(v); }
function asText(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  return v == null ? '' : String(v);
}
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 마진계산 현장 ---------- */
function projectSheet() { return getSheet(PROJECT_SHEET, PROJECT_COLS.map(c => c[1])); }

function projectSheetReady() {
  const sh = projectSheet();
  // 예전 시트에 '회사' 머리글 추가
  const w = PROJECT_COLS.length;
  if (String(sh.getRange(1, w).getValue()) !== PROJECT_COLS[w - 1][1]) sh.getRange(1, w).setValue(PROJECT_COLS[w - 1][1]).setFontWeight('bold');
  return sh;
}
// 회사가 비어 있는 예전 현장은 작성자의 소속 회사로 봄
function readProjects(staff) {
  const sh = projectSheetReady();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const orgOf = {};
  (staff || []).forEach(s => { if (s.company) orgOf[s.name] = s.company; });
  return sh.getRange(2, 1, n, PROJECT_COLS.length).getValues().filter(r => r[0]).map(r => {
    const o = {};
    PROJECT_COLS.forEach((c, i) => {
      let v = r[i];
      if (c[0] === 'date' && v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[c[0]] = PROJECT_NUMS.indexOf(c[0]) >= 0 ? Number(v) || 0 : asText(v);
    });
    if (!o.org) o.org = orgOf[o.author] || '(회사 미지정)';
    return o;
  });
}
const COL_UPDATED = PROJECT_COLS.findIndex(c => c[0] === 'updatedAt') + 1;
const COL_ORG = PROJECT_COLS.findIndex(c => c[0] === 'org') + 1;
const COL_AUTHOR = PROJECT_COLS.findIndex(c => c[0] === 'author') + 1;
// 그 줄의 회사 (비어 있으면 작성자의 회사)
function rowOrg(sh, idx, staff) {
  const org = asText(sh.getRange(idx, COL_ORG).getValue());
  if (org) return org;
  const author = asText(sh.getRange(idx, COL_AUTHOR).getValue());
  const s = (staff || []).find(x => x.name === author);
  return (s && s.company) || '(회사 미지정)';
}

function saveProject(p, me) {
  if (!p || !p.id) return;
  const sh = projectSheetReady();
  const idx0 = findRow(sh, p.id);
  if (!me.admin) {
    if (idx0 > 0 && rowOrg(sh, idx0, me.staff) !== me.company) return; // 다른 회사 현장은 고칠 수 없음
    p.org = me.company;
  } else {
    p.org = String(p.org || '').trim() || (idx0 > 0 ? rowOrg(sh, idx0, me.staff) : me.company);
  }
  p.marginRate = Number(p.contract) > 0 ? Math.round(Number(p.margin) / Number(p.contract) * 1000) / 10 : 0;
  const row = PROJECT_COLS.map(c => PROJECT_NUMS.indexOf(c[0]) >= 0 ? Number(p[c[0]]) || 0 : txt(p[c[0]]));
  const idx = idx0;
  if (idx > 0) {
    const cur = asText(sh.getRange(idx, COL_UPDATED).getValue());
    if (p.onlyIfNewer && cur && p.updatedAt && cur > p.updatedAt) return;
    sh.getRange(idx, 1, 1, PROJECT_COLS.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
}

function deleteProject(id, me) {
  if (!id) return;
  const sh = projectSheetReady();
  const idx = findRow(sh, id);
  if (idx > 0 && (me.admin || rowOrg(sh, idx, me.staff) === me.company)) sh.deleteRow(idx);
}

function findRow(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/* ---------- PB 단가 ---------- */
function settingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange(1, 1, 5, 2).setValues([
      ['항목', '값'],
      ['15T 장당 단가', DEFAULT_PRICES.p15],
      ['18T 장당 단가', DEFAULT_PRICES.p18],
      ['우라 장당 단가', DEFAULT_PRICES.pura],
      ['기본기물 비율(%)', DEFAULT_PRICES.rate]
    ]);
    sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  return sh;
}
// 회사별 단가는 공유값 시트에 'prices:회사'로 저장, 없으면 설정 시트의 기본값
function readPrices(org) {
  if (org) {
    const o = readKV()['prices:' + org];
    if (o && typeof o === 'object') return Object.assign(readPrices(), o);
  }
  const v = settingsSheet().getRange(2, 2, 4, 1).getValues();
  return {
    p15: Number(v[0][0]) || DEFAULT_PRICES.p15,
    p18: Number(v[1][0]) || DEFAULT_PRICES.p18,
    pura: Number(v[2][0]) || DEFAULT_PRICES.pura,
    rate: v[3][0] === '' ? DEFAULT_PRICES.rate : Number(v[3][0])
  };
}
function writePrices(p, me) {
  const org = me.admin ? (String(p.org || '').trim() || me.company) : me.company;
  const cur = readPrices(org);
  const next = {
    p15: Number(p.p15) || cur.p15, p18: Number(p.p18) || cur.p18,
    pura: Number(p.pura) || cur.pura, rate: p.rate === undefined ? cur.rate : Number(p.rate) || 0
  };
  writeKV('prices:' + org, next);
}

/* ---------- 재고·재고이력·견적기록 ---------- */
function collectionHeaders(def) {
  return ['ID'].concat(def.cols.map(c => c[1]), ['작성자', '저장일시', '원본데이터(수정 금지)']);
}
function collectionSheet(name) {
  const def = COLLECTIONS[name];
  if (!def) throw new Error('unknown collection ' + name);
  return getSheet(def.sheet, collectionHeaders(def));
}
function readCollection(name) {
  const def = COLLECTIONS[name];
  const sh = collectionSheet(name);
  const width = def.cols.length + 4;
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const rows = sh.getRange(2, 1, n, width).getValues();
  const list = [];
  rows.forEach(r => {
    if (!r[0]) return;
    let o;
    try { o = JSON.parse(r[width - 1]); } catch (err) { o = null; }
    if (!o || typeof o !== 'object') o = {};
    // 시트에서 사람이 직접 고친 주요 값(예: 재고 잔량)을 우선 반영
    def.cols.forEach((c, i) => {
      const v = r[i + 1];
      if (v === '' || v == null) return;
      o[c[0]] = typeof o[c[0]] === 'number' ? Number(v) : (v instanceof Date ? asText(v).slice(0, 10) : v);
    });
    o.id = isNaN(Number(r[0])) ? String(r[0]) : Number(r[0]);
    list.push(o);
  });
  return list;
}
function syncCollection(name, upserts, deletes, author, onlyNew) {
  const def = COLLECTIONS[name];
  if (!def) return;
  const sh = collectionSheet(name);
  const width = def.cols.length + 4;
  const n = sh.getLastRow() - 1;
  let rows = n > 0 ? sh.getRange(2, 1, n, width).getValues() : [];
  const index = {};
  rows.forEach((r, i) => { if (r[0] !== '') index[String(r[0])] = i; });

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  const toRow = o => [txt(o.id)].concat(
    def.cols.map(c => typeof o[c[0]] === 'number' ? o[c[0]] : txt(o[c[0]])),
    [txt(author || o.author || ''), txt(now), JSON.stringify(o)]
  );

  upserts.forEach(o => {
    if (!o || o.id == null || o.id === '') return;
    const key = String(o.id);
    if (key in index) { if (!onlyNew) rows[index[key]] = toRow(o); }
    else { index[key] = rows.length; rows.push(toRow(o)); }
  });
  if (deletes.length) {
    const del = {};
    deletes.forEach(id => del[String(id)] = true);
    rows = rows.filter(r => !del[String(r[0])]);
  }

  if (n > 0) sh.getRange(2, 1, n, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

/* ---------- 공유값 (거래처 목록 등) ---------- */
function publicKV() {
  const o = readKV();
  Object.keys(o).forEach(k => { if (k.indexOf('prices:') === 0) delete o[k]; });
  return o;
}
function kvSheet() { return getSheet(KV_SHEET, ['항목', '값(수정 금지)']); }
function readKV() {
  const sh = kvSheet();
  const n = sh.getLastRow() - 1;
  const o = {};
  if (n < 1) return o;
  sh.getRange(2, 1, n, 2).getValues().forEach(r => {
    if (!r[0]) return;
    try { o[String(r[0])] = JSON.parse(r[1]); } catch (err) {}
  });
  return o;
}
function writeKV(key, value) {
  if (!key) return;
  const sh = kvSheet();
  const n = sh.getLastRow() - 1;
  const json = JSON.stringify(value == null ? null : value);
  if (n > 0) {
    const keys = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === key) { sh.getRange(i + 2, 2).setValue(json); return; }
    }
  }
  sh.appendRow([key, json]);
}
