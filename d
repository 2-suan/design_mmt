/** @OnlyCurrentDoc */
/**
 * MOMENTUM 카페 [스터디] 하위 게시판 → 월별 활동 정리 (카드뉴스 "월말결산" 용)
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여넣고 저장
 *       → ⚙️ 프로젝트 설정 → "편집기에서 appsscript.json 매니페스트 파일 표시" 체크
 *       → appsscript.json 내용을 함께 제공된 파일로 교체 (권한을 최소한으로 제한)
 *       → 함수 목록에서 setupMonthlyTrigger 를 한 번 실행(권한 승인)
 * 권한: 이 시트 1개만 / 외부(네이버) 접속 / 자동 실행 / 시트 메뉴 표시
 *       (구글 드라이브, 다른 시트, 지메일 등에는 접근하지 않음)
 * 결과: 매월 1일 오전 9시에 "지난달" 탭(예: 2026-11)이 자동 생성/갱신됩니다.
 * 수동: 시트 상단 메뉴 [모멘텀 결산] → "이번 달 정리" / "월 지정해서 정리"
 *
 * 활동 날짜 찾는 순서: ① 제목 → ② 본문 글
 *   둘 다 없으면 날짜 칸에 "날짜 직접 기입"(노란색)으로 표시 → 사진 등을 보고 직접 입력
 *   (이 경우 어느 달 탭에 넣을지는 글 작성일 기준으로 정함)
 */

const CAFE_ID = 31397081;
const CAFE_URL = 'smumomentum';
const STUDY_FOLDER_NAME = '스터디';                 // 카페 왼쪽 메뉴의 폴더 이름
const EXCLUDE_BOARDS = ['스터디/소모임 신청 게시판']; // 제외할 게시판
const SETTINGS_SHEET = '스터디 설정';
const MANUAL_DATE = '날짜 직접 기입';
// 한 행 = 스터디 1개의 활동 최대 2개 (카드 1장). 활동이 3개 이상이면 같은 스터디로 다음 행이 이어짐
const HEADERS = ['subtitle', 'title', 'date1', 'context1', 'image1', 'date2', 'context2', 'image2', 'link1', 'link2'];
const PER_ROW = 2;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('모멘텀 결산')
    .addItem('이번 달 정리', 'runThisMonth')
    .addItem('지난 달 정리', 'runLastMonth')
    .addItem('월 지정해서 정리', 'runPromptMonth')
    .addSeparator()
    .addItem('매월 자동 실행 켜기', 'setupMonthlyTrigger')
    .addToUi();
}

/* ---------- 실행 진입점 ---------- */
function runThisMonth() { const d = new Date(); buildMonth(d.getFullYear(), d.getMonth() + 1); }
function runLastMonth() { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); buildMonth(d.getFullYear(), d.getMonth() + 1); }
function runPromptMonth() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('정리할 월을 입력하세요 (예: 2026-11)');
  const m = /^(\d{4})-(\d{1,2})$/.exec(r.getResponseText().trim());
  if (!m) return ui.alert('형식이 올바르지 않아요. 예: 2026-11');
  buildMonth(+m[1], +m[2]);
}
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'runLastMonth').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runLastMonth').timeBased().onMonthDay(1).atHour(9).inTimezone('Asia/Seoul').create();
}

/* ---------- 메인 ---------- */
function buildMonth(year, month) {
  const ss = SpreadsheetApp.getActive();
  const boards = getStudyBoards();
  const settings = readSettings(ss, boards);
  const acts = collectMonthActivities(boards, settings, year, month);
  const rows = groupIntoRows(acts);

  const table = rows.map(r => {
    const slot = i => r.acts[i] || {};
    return [r.sub, r.title,
            slot(0).date || '', slot(0).context || '', slot(0).image || '',
            slot(1).date || '', slot(1).context || '', slot(1).image || '',
            slot(0).link || '', slot(1).link || ''];
  });

  const sh = ss.insertSheet(newSheetName(ss, year, month));   // 기존 탭은 절대 지우거나 덮어쓰지 않음
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#2F5597').setFontColor('#FFFFFF');
  sh.getRange('C:C').setNumberFormat('@');                     // 26.11.09 가 숫자/날짜로 바뀌지 않게
  sh.getRange('F:F').setNumberFormat('@');
  if (table.length) sh.getRange(2, 1, table.length, HEADERS.length).setValues(table);
  rows.forEach((r, i) => r.acts.forEach((a, k) => {             // 날짜를 못 찾은 칸: 노란색 + 메모로 작성일 안내
    if (a.date !== MANUAL_DATE) return;
    sh.getRange(i + 2, k === 0 ? 3 : 6).setBackground('#FFF2CC').setFontColor('#B45F06')
      .setNote(`제목·본문에 날짜 없음. 글 작성일: ${a.writeDateText}\n사진(image${k + 1})을 보고 활동 날짜를 입력하세요.`);
  }));
  sh.getRange('D:D').setWrap(true);
  sh.getRange('G:G').setWrap(true);
  sh.setColumnWidth(4, 360);
  sh.setColumnWidth(7, 360);
  sh.setFrozenRows(1);
}

/* 스터디별로 날짜순 정렬 후 2개씩 묶기 (날짜를 못 찾은 글은 작성일 순서로 끼워 넣음) */
function groupIntoRows(acts) {
  const byStudy = new Map();
  acts.forEach(a => {
    if (!byStudy.has(a.title)) byStudy.set(a.title, []);
    byStudy.get(a.title).push(a);
  });
  const rows = [];
  [...byStudy.keys()].sort((x, y) => x.localeCompare(y)).forEach(title => {
    const list = byStudy.get(title).sort((x, y) => x.sortKey - y.sortKey);
    for (let i = 0; i < list.length; i += PER_ROW) {
      rows.push({ sub: list[0].sub, title, acts: list.slice(i, i + PER_ROW) });
    }
  });
  return rows;
}

/* 탭 이름: 처음엔 "2026-11", 이미 있으면 "2026-11_최신",
   그것도 있으면 기존 "_최신" 탭 이름에 생성 시각을 붙여 보관한 뒤 새 "_최신" 탭을 만듦 */
function newSheetName(ss, year, month) {
  const base = `${year}-${String(month).padStart(2, '0')}`;
  if (!ss.getSheetByName(base)) return base;
  const latest = `${base}_최신`;
  const prev = ss.getSheetByName(latest);
  if (prev) {
    const stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd_HHmm');
    prev.setName(`${base}_이전_${stamp}`);
  }
  return latest;
}

/* [스터디] 폴더 아래 게시판 목록 (구분선/다음 폴더 전까지) */
function getStudyBoards() {
  const d = getJson(`https://apis.naver.com/cafe-web/cafe2/SideMenuList?cafeId=${CAFE_ID}`);
  const menus = d.message.result.menus;
  const out = [];
  let inside = false;
  for (const m of menus) {
    if (m.menuType === 'F' && m.menuName === STUDY_FOLDER_NAME) { inside = true; continue; }
    if (inside && (m.menuType === 'F' || m.menuType === 'S')) break;
    if (inside && m.menuType === 'B') {
      const name = decodeHtml(m.menuName);
      if (!EXCLUDE_BOARDS.includes(name)) out.push({ id: m.menuId, name });
    }
  }
  return out;
}

/* "스터디 설정" 탭: 게시판별 카드 제목/해시태그를 사람이 수정할 수 있게 */
function readSettings(ss, boards) {
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['게시판 이름(카페)', '#title (카드 제목)', '#sub_title (해시태그)']]).setFontWeight('bold');
  }
  const vals = sh.getDataRange().getValues().slice(1);
  const map = {};
  vals.forEach(r => { if (r[0]) map[r[0]] = { title: r[1] || r[0], sub: r[2] || '#스터디' }; });
  const missing = boards.filter(b => !map[b.name]);   // 새 스터디 게시판 자동 추가
  if (missing.length) {
    sh.getRange(sh.getLastRow() + 1, 1, missing.length, 3).setValues(missing.map(b => [b.name, b.name, '#스터디']));
    missing.forEach(b => map[b.name] = { title: b.name, sub: '#스터디' });
  }
  return map;
}

function collectMonthActivities(boards, settings, year, month) {
  const acts = [];
  for (const b of boards) {
    for (const a of listArticles(b.id)) {
      const writeDate = new Date(a.writeDateTimestamp);
      // 활동 월보다 한참 이전 글은 건너뜀 (작성일이 대상 월 시작 7일 전보다 이르면 skip)
      if (writeDate < new Date(year, month - 1, 1 - 7)) continue;
      const { text: body, image } = articleContent(a.articleId);
      const found = activityDate(a.subject, body, writeDate);
      const monthDate = found || writeDate;                 // 날짜를 못 찾으면 작성일 기준으로 월 판단
      if (monthDate.getFullYear() !== year || monthDate.getMonth() + 1 !== month) continue;
      const s = settings[b.name];
      acts.push({
        sub: s.sub, title: s.title,
        date: found ? fmtDate(found) : MANUAL_DATE,
        context: activityContent(body),
        image,
        link: `https://cafe.naver.com/${CAFE_URL}/${a.articleId}`,
        sortKey: monthDate.getTime(),
        writeDateText: fmtDate(writeDate),
      });
    }
  }
  return acts;
}

/* ---------- 네이버 카페 API ---------- */
function listArticles(menuId) {
  const out = [];
  for (let p = 1; p <= 20; p++) {
    const d = getJson(`https://apis.naver.com/cafe-web/cafe-boardlist-api/v1/cafes/${CAFE_ID}/menus/${menuId}/articles?page=${p}&pageSize=50&sortBy=TIME&viewType=L`);
    const list = d.result.articleList;
    list.filter(x => x.type === 'ARTICLE').forEach(x => out.push(x.item));
    if (list.length < 50) break;
  }
  return out;
}
function articleContent(id) {
  try {
    const d = getJson(`https://apis.naver.com/cafe-web/cafe-articleapi/v2.1/cafes/${CAFE_ID}/articles/${id}?useCafeId=true&requestFrom=A`);
    const html = d.result.article.contentHtml || '';
    return { text: htmlToText(html), image: firstImage(html) };
  } catch (e) { return { text: '', image: '' }; } // 회원 전용 글 등
}

/* 본문에 첨부된 첫 번째 사진 URL (스티커·이모티콘 제외) */
function firstImage(html) {
  for (const [tag] of html.matchAll(/<img\b[^>]*>/gi)) {
    const m = /\bsrc="([^"]+)"/i.exec(tag);
    if (!m) continue;
    const url = decodeHtml(m[1]);
    const isPhoto = /se-image-resource/.test(tag) || /\/\/cafeptthumb-phinf\.pstatic\.net\//.test(url);
    if (isPhoto && !/storep-phinf|sticker/i.test(url)) return url;
  }
  return '';
}
function getJson(url) {
  const res = UrlFetchApp.fetch(url, { headers: { Referer: 'https://cafe.naver.com/' }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error(`${res.getResponseCode()} ${url}`);
  return JSON.parse(res.getContentText());
}

/* ---------- 파싱 ---------- */
function htmlToText(h) {
  h = h.replace(/<(script|style)[\s\S]*?<\/\1>/gi, '')
       .replace(/<br\s*\/?>|<\/p>|<\/div>|<\/li>|<\/h\d>/gi, '\n')
       .replace(/<[^>]+>/g, '');
  return decodeHtml(h).replace(/\u200b/g, '').replace(/[ \t\u00a0]+/g, ' ')
                      .replace(/\n\s*\n+/g, '\n').trim();
}
function decodeHtml(s) {
  return s.replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(n)).replace(/&amp;/g, '&');
}

/* 활동 날짜: 제목 → 본문(출석 항목 전까지). 못 찾으면 null (→ "날짜 직접 기입") */
function activityDate(subject, body, writeDate) {
  return findDate(subject, writeDate, true) || findDate(body.split(/\n2\.\s/)[0], writeDate, false);
}

/* 날짜 후보를 모두 찾은 뒤, 작성일 기준 30일 전 ~ 작성일 사이인 첫 날짜를 채택 (엉뚱한 날짜 방지)
   지원 형식: 2026년 11월 9일 / 26년 11월 9일 / 11월 9일 / 2026.11.9 / 2026-11-09 / 2026/11/9 / 2026. 11. 9.
              26.11.09 / 11/9 / 20261109 / 261109 (제목에서만) */
const NB = '(?<!\\d)', NA = '(?!\\d)';   // 앞뒤에 다른 숫자가 붙어 있으면 날짜로 보지 않음
const SEP = '\\s*[.\\-\\/_]\\s*';
const DATE_PATTERNS = [
  // [정규식, 연도 그룹 종류]  'full' = 4자리, 'short' = 2자리, 'none' = 연도 없음(작성일 연도로 추정)
  [new RegExp(`${NB}(\\d{4})\\s*년\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일`, 'g'), 'full'],   // 2026년 11월 9일
  [new RegExp(`${NB}(\\d{2})\\s*년\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일`, 'g'), 'short'],  // 26년 11월 9일
  [new RegExp(`${NB}(\\d{4})${SEP}(\\d{1,2})${SEP}(\\d{1,2})${NA}`, 'g'), 'full'],             // 2026.11.9 / 2026-11-09
  [new RegExp(`${NB}(\\d{2})${SEP}(\\d{1,2})${SEP}(\\d{1,2})${NA}`, 'g'), 'short'],            // 26.11.09
  [new RegExp(`${NB}(20\\d{2})(\\d{2})(\\d{2})${NA}`, 'g'), 'full'],                            // 20261109
  [new RegExp(`(?<![\\d년]\\s*)(?<!\\d)(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일`, 'g'), 'none'],           // 11월 9일
  [new RegExp(`(?<![\\d\\/]\\s*)(?<!\\d)(\\d{1,2})\\s*\\/\\s*(\\d{1,2})${NA}(?!\\s*\\/)`, 'g'), 'none'],           // 11/9
];
const COMPACT_SHORT = new RegExp(`${NB}(\\d{2})(\\d{2})(\\d{2})${NA}`, 'g');                    // 261109 (제목에서만)

function findDate(txt, writeDate, allowCompact) {
  if (!txt) return null;
  const t = txt.replace(/\s+/g, ' ');
  const cands = [];                                   // [index, year(4자리), month, day]
  const wy = writeDate.getFullYear();
  const pats = allowCompact ? [...DATE_PATTERNS, [COMPACT_SHORT, 'short']] : DATE_PATTERNS;
  for (const [re, kind] of pats) {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(t))) {
      if (kind === 'none') {                          // 연도 없음 → 작성일 연도 (1월 글의 "12월 30일"은 전년도)
        cands.push([m.index, wy, +m[1], +m[2]], [m.index, wy - 1, +m[1], +m[2]]);
      } else {
        cands.push([m.index, kind === 'full' ? +m[1] : 2000 + +m[1], +m[2], +m[3]]);
      }
    }
  }
  cands.sort((a, b) => a[0] - b[0]);                  // 글에서 먼저 나온 날짜 우선
  for (const [, y, mo, d] of cands) {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) continue;
    const dt = new Date(y, mo - 1, d);
    if (dt.getMonth() !== mo - 1) continue;           // 2월 30일 같은 날짜 제외
    const diff = (writeDate - dt) / 864e5;
    if (diff > -1 && diff < 30) return dt;
  }
  return null;
}

/* "3. 간단하게 활동 내역을 설명해 주세요." 아래 내용만 추출 */
function activityContent(body) {
  const m = /3\.\s*간단하게 활동 내역[^\n]*\n([\s\S]*)/.exec(body);
  return (m ? m[1] : body).split('\n').map(l => l.trim().replace(/^(\d+\)|-|•|·)\s*/, '').trim())
          .filter(l => l && !/^(\d+\)|-)$/.test(l)).join('\n');
}
function fmtDate(d) {
  const p = n => String(n).padStart(2, '0');
  return `${p(d.getFullYear() % 100)}.${p(d.getMonth() + 1)}.${p(d.getDate())}`;
}
