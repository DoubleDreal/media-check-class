/**
 * 매체 사용 습관 점검 · 수집 서버 (Google Apps Script 웹 앱)
 *
 * 학생은 로그인 없이 '학생용 점검 도구(HTML)'에서 [제출]을 누르고,
 * 이 스크립트가 교사 구글 스프레드시트의 R_ 탭에 자동으로 저장합니다.
 * 저장된 자료는 '교사용_매체사용점검_자료'의 개인분석(교사용)·학급집계 탭이 바로 집계합니다.
 *
 * [설치 순서]
 * 1) 교사용 자료(xlsx)를 구글 스프레드시트로 연다. (드라이브에 올린 뒤 '구글 스프레드시트로 열기' → 저장)
 * 2) 확장 프로그램 > Apps Script 에 이 코드를 통째로 붙여넣는다.
 * 3) 아래 CLASS_CODE 를 수업 코드(예: 숫자 4자리)로 바꾼다. 학생에게는 이 코드만 알려 준다.
 * 4) 함수 'setupSheets' 를 선택해 ▶ 실행한다. (권한 승인 1회, R_ 탭이 만들어진다)
 * 5) 배포 > 새 배포 > 유형 '웹 앱'
 *      · 다음 사용자 인증으로 실행: 나
 *      · 액세스 권한이 있는 사용자: 모든 사용자
 *    → 배포 후 나오는 '웹 앱 URL(…/exec)'을 복사한다.
 * 6) 학생용 점검 도구(HTML)의 맨 위쪽 ENDPOINT 값에 그 URL을 붙여넣고 학생에게 배부한다.
 *
 * ※ 코드를 고친 뒤에는 '배포 관리 > 수정 > 새 버전'으로 다시 배포해야 반영됩니다.
 * ※ 학교 관리자가 '모든 사용자' 웹 앱 배포를 막아 두었다면 5)에서 선택지가 보이지 않습니다.
 * ※ 스크린타임 캡처 사진은 내 드라이브의 'CAPTURE_FOLDER' 폴더에 저장되며, 나만 볼 수 있습니다.
 */

const CLASS_CODE = '1234';            // ← 수업 코드로 바꾸세요
const CAPTURE_FOLDER = '매체사용_캡처';
const MAX_IMG_CHARS = 3500000;        // 사진(base64) 최대 길이(약 2.5MB)

const WHO_H = ['학년','반','이름'];   // 모든 탭(피드백 제외)의 맨 오른쪽에 붙는 열
const DEF = {
  pre:   { name: 'R_사전', headers: ['타임스탬프','코드번호','평일 예상','주말 예상','가장 많이 쓰는 매체','주로 하는 것','잠자기 직전 사용','나의 사용 평가(1~5)','조절 경험','조절 방법','알아보고 싶은 것','기록 방법'] },
  daily: { name: 'R_일일', headers: ['타임스탬프','코드번호','날짜','총 사용(분)','영상 보기(분)','게임(분)','메신저·SNS(분)','공부·검색(분)','기타(분)','기분','계획보다 오래','잠자기 30분 전','딴 앱도 열었나','한 줄 소감','스크린타임 캡처'] },
  goal:  { name: 'R_목표', headers: ['타임스탬프','코드번호','분석문','예상과 실제','고치고 싶은 점','목표','방해 요인','만약~하면~한다','대체 활동','도움 요청','실천 기간','자신감(1~5)'] },
  act:   { name: 'R_실천', headers: ['타임스탬프','코드번호','날짜','목표 지켰나','사용 시간(분)','도움이 된 것','어려웠던 것','계획 수정','스크린타임 캡처'] },
  post:  { name: 'R_사후', headers: ['타임스탬프','코드번호','평일 사용 시간','주말 사용 시간','잠자기 직전 사용','나의 사용 평가(1~5)','목표 지킨 정도(1~5)','도움이 된 방법','어려웠던 점','새로 알게 된 점','이어 갈 목표'] },
  fb:    { name: 'R_피드백', headers: ['타임스탬프','발표자 코드','잘한 점','궁금한 점·질문','가져올 방법'] }
};
['pre','daily','goal','act','post'].forEach(function (k) { DEF[k].headers = DEF[k].headers.concat(WHO_H); });

/* ───────── 설치용 ───────── */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  Object.keys(DEF).forEach(function (k) {
    let sh = ss.getSheetByName(DEF[k].name);
    if (!sh) sh = ss.insertSheet(DEF[k].name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, DEF[k].headers.length).setValues([DEF[k].headers])
        .setFontWeight('bold').setBackground('#1F5C7A').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
  });
  const d = ss.getSheetByName(DEF.daily.name); d.getRange('C2:C').setNumberFormat('yyyy-mm-dd');
  const a = ss.getSheetByName(DEF.act.name);   a.getRange('C2:C').setNumberFormat('yyyy-mm-dd');
  ['pre','daily','goal','act','post','fb'].forEach(function (k) { ss.getSheetByName(DEF[k].name).getRange('A2:A').setNumberFormat('m/d hh:mm'); });
  Logger.log('R_ 탭 준비 완료. 이제 [배포 > 새 배포 > 웹 앱]을 진행하세요.');
}

/* ───────── 웹 앱 ───────── */
function doGet() {
  return out_({ ok: true, service: 'media-habit-collector' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (String(body.classCode || '').trim() !== String(CLASS_CODE)) return out_({ ok: false, error: 'class_code' });
    const code = String(body.code || '').trim();
    if (!/^\d{1,6}$/.test(code)) return out_({ ok: false, error: 'student_code' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const type = body.type;
    if (!DEF[type]) return out_({ ok: false, error: 'type' });
    const sh = ss.getSheetByName(DEF[type].name);
    if (!sh) return out_({ ok: false, error: 'no_sheet' });
    const tz = ss.getSpreadsheetTimeZone();
    const now = new Date();
    let saved = 0;
    const who = [s_(body.grade).slice(0, 4), s_(body.klass).slice(0, 4), s_(body.name).slice(0, 20)];
    if (type !== 'fb') ensureHeader_(sh, DEF[type].headers);

    if (type === 'pre') {
      const d = body.data || {};
      upsert_(sh, [1], [now, Number(code), s_(d.wk), s_(d.we), s_(d.media), s_(d.purp), s_(d.sleep), n_(d.rate), s_(d.tried), s_(d.how), s_(d.want), s_(d.method)].concat(who), tz);
      saved = 1;
    } else if (type === 'goal') {
      const d = body.data || {};
      upsert_(sh, [1], [now, Number(code), s_(d.analysis), s_(d.cmp), s_(d.fix), s_(d.goal), s_(d.hard), s_(d.ifthen), s_(d.alt), s_(d.help), s_(d.days), n_(d.conf)].concat(who), tz);
      saved = 1;
    } else if (type === 'post') {
      const d = body.data || {};
      upsert_(sh, [1], [now, Number(code), s_(d.wk), s_(d.we), s_(d.sleep), n_(d.rate), n_(d.goalScore), s_(d.helped), s_(d.hard), s_(d.learned), s_(d.next)].concat(who), tz);
      saved = 1;
    } else if (type === 'fb') {
      const d = body.data || {};
      const sp = String(d.speaker || '').trim();
      if (!/^\d{1,6}$/.test(sp)) return out_({ ok: false, error: 'speaker_code' });
      sh.appendRow([now, Number(sp), s_(d.good), s_(d.question), s_(d.bring)]);
      saved = 1;
    } else if (type === 'daily' || type === 'act') {
      const rows = (body.rows || []).slice(0, 20);
      rows.forEach(function (r) {
        const date = parseDate_(r.date, tz);
        if (!date) return;
        const nm = code + '_' + String(r.date);
        let img = '';
        if (r.img && r.img.data && String(r.img.data).length <= MAX_IMG_CHARS) img = saveImage_(r.img, nm + (type === 'act' ? '_실천' : ''));
        if (type === 'daily') {
          const row = [now, Number(code), date, m_(r.total), m_(r.v), m_(r.g), m_(r.m), m_(r.s), m_(r.e), s_(r.mood), s_(r.over), s_(r.sleep), s_(r.app), s_(r.note), img].concat(who);
          upsert_(sh, [1, 2], row, tz, 14, img);
        } else {
          const row = [now, Number(code), date, s_(r.ok), m_(r.t), s_(r.help), s_(r.hard), s_(r.change), img].concat(who);
          upsert_(sh, [1, 2], row, tz, 8, img);
        }
        saved++;
      });
    }
    return out_({ ok: true, saved: saved });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ───────── 도우미 ───────── */
function ensureHeader_(sh, headers) {
  if (sh.getRange(1, headers.length).getValue() === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1F5C7A').setFontColor('#FFFFFF');
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function s_(v) { return v == null ? '' : String(v).replace(/^[=+\-@]/, "'$&").slice(0, 600); }   // 수식 입력 방지 + 길이 제한
function n_(v) { const x = parseFloat(v); return isFinite(x) ? x : ''; }
function m_(v) { const x = parseFloat(v); return isFinite(x) && x >= 0 ? Math.min(x, 1440) : ''; }
function parseDate_(str, tz) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(str || ''))) return null;
  return Utilities.parseDate(String(str), tz, 'yyyy-MM-dd');
}
function norm_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v).trim();
}
// 같은 키(코드번호[, 날짜])의 행이 있으면 그 행을 덮어쓰고, 없으면 맨 아래에 추가한다.
function upsert_(sh, keyIdx, row, tz, imgIdx, newImg) {
  const last = sh.getLastRow();
  let at = 0;
  if (last >= 2) {
    const vals = sh.getRange(2, 1, last - 1, row.length).getValues();
    for (let i = 0; i < vals.length && !at; i++) {
      let same = true;
      for (let j = 0; j < keyIdx.length; j++) {
        if (norm_(vals[i][keyIdx[j]], tz) !== norm_(row[keyIdx[j]], tz)) { same = false; break; }
      }
      if (same) {
        at = i + 2;
        if (imgIdx != null && !newImg) row[imgIdx] = vals[i][imgIdx];   // 새 사진이 없으면 기존 사진 링크 유지
      }
    }
  }
  if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}
function saveImage_(img, name) {
  const mime = /^image\/(jpeg|png)$/.test(img.mime) ? img.mime : 'image/jpeg';
  const it = DriveApp.getFoldersByName(CAPTURE_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CAPTURE_FOLDER);
  const blob = Utilities.newBlob(Utilities.base64Decode(img.data), mime, name + (mime === 'image/png' ? '.png' : '.jpg'));
  return folder.createFile(blob).getUrl();
}
