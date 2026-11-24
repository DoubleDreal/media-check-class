/**
 * 매체 사용 습관 점검 프로젝트 · 구글 폼 자동 생성 스크립트
 *
 * [사용법]
 * 1) '교사용_매체사용점검_자료'를 구글 스프레드시트로 연 뒤  확장 프로그램 > Apps Script
 * 2) 이 코드를 통째로 붙여넣고 저장 → 함수 'createAllForms' 선택 후 ▶ 실행 (권한 승인 1회)
 *    ※ 한 번만 실행하세요. 다시 실행하면 폼이 중복으로 만들어집니다.
 * 3) 실행이 끝나면 '폼 링크' 탭에 학생용 링크가 정리됩니다.
 * 4) [꼭 직접 해야 하는 일] 스크린타임 캡처 제출 칸 추가
 *    - '하루 기록'과 '실천 체크' 폼을 편집 모드로 열고, 마지막에 질문 추가 → 유형 '파일 업로드'
 *    - 제목: 스크린타임 캡처 사진 (선택)   / 허용 파일 형식: 이미지 / 최대 파일 수: 1 / 최대 크기: 10MB
 *    - 마지막 질문으로 넣어야 응답 시트의 맨 끝 열(하루 기록 O열, 실천 체크 I열)에 링크가 들어갑니다.
 *    - 파일 업로드 질문이 있으면 학생이 구글(학교) 계정으로 로그인해야 하고, 사진은 폼 소유자의 드라이브에 저장됩니다.
 *      (Apps Script로는 파일 업로드 질문을 만들 수 없습니다.)
 *
 * 응답 탭 이름은 R_사전, R_일일, R_목표, R_실천, R_사후, R_피드백 로 자동 변경됩니다.
 * 이 이름을 바꾸지 마세요. (분석 탭이 이 이름으로 자료를 읽습니다)
 */

const TIME_CHOICES = ['1시간 미만', '1~2시간', '2~3시간', '3~4시간', '4시간 이상']; // '설정' 탭과 글자가 같아야 함
const SLEEP_CHOICES = ['거의 매일', '자주', '가끔', '거의 안 함'];

function createAllForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const links = [];

  links.push(buildPre(ss));
  links.push(buildDaily(ss));
  links.push(buildGoal(ss));
  links.push(buildAction(ss));
  links.push(buildPost(ss));
  links.push(buildFeedback(ss));

  writeLinkSheet(ss, links);
  Logger.log('완료! [폼 링크] 탭을 확인하세요. 스크린타임 캡처 업로드 질문은 폼에서 직접 추가해야 합니다.');
}

/* ───────── 폼 정의 ───────── */

function buildPre(ss) {
  const f = newForm('[1차시] 매체 사용 습관 사전 설문',
    '정답은 없습니다. 지금 내 생각을 솔직하게 답해 주세요. (약 5분)');
  addCode(f);
  addChoice(f, '평일 하루 동안 매체(스마트폰·태블릿·TV·게임기·컴퓨터)를 쓰는 시간은 얼마나 될 것 같나요? (예상)', TIME_CHOICES, true);
  addChoice(f, '주말 하루는 얼마나 될 것 같나요? (예상)', TIME_CHOICES, true);
  addCheck(f, '가장 많이 쓰는 매체는 무엇인가요? (모두 선택)', ['스마트폰', '태블릿', 'TV', '게임기', '컴퓨터', '기타'], true);
  addCheck(f, '주로 무엇을 하나요? (모두 선택)', ['영상 보기', '게임', '메신저·SNS', '공부·검색', '음악', '웹툰·책', '기타'], true);
  addChoice(f, '잠자기 직전(30분 안)에 매체를 쓰나요?', SLEEP_CHOICES, true);
  addScale(f, '나의 매체 사용은 지금 어떤 것 같나요?', 1, 5, '너무 많다', '적당하다', true);
  addChoice(f, '매체 사용을 스스로 조절해 보려고 한 적이 있나요?', ['있다', '없다'], true);
  addPara(f, '있다면 어떻게 했나요? (없으면 비워 두세요)', false);
  addText(f, '이번 점검에서 특히 알아보고 싶은 것은 무엇인가요?', false);
  addChoice(f, '기록은 어떤 방법으로 할까요?', ['종이 기록표', '구글 폼', '목실감 앱', '기기 스크린타임 보고 폼에 입력'], true);
  f.setConfirmationMessage('응답해 줘서 고마워요! 내일부터 일주일 동안 기록을 시작해 봅시다.');
  return finish(ss, f, 'R_사전', '사전 설문 (1차시)');
}

function buildDaily(ss) {
  const f = newForm('[기록 주간] 오늘의 매체 사용 기록',
    '하루가 끝난 저녁에 한 번만 쓰세요. (1~2분) 기기의 스크린타임이나 종이 기록표를 보고 분 단위로 적습니다. ' +
    '기록하려고 폰을 켠 김에 다른 앱을 열었다면, 그것도 솔직하게 체크해 주세요.');
  addCode(f);
  addDate(f, '날짜');
  addNumber(f, '오늘 매체 사용 시간은 모두 몇 분인가요? (숫자만)', true);
  f.addSectionHeaderItem().setTitle('목적별로 나누면 몇 분인가요?').setHelpText('모르면 0을 쓰거나 비워 두세요. 숫자만 씁니다.');
  addNumber(f, '영상 보기 (분)', false);
  addNumber(f, '게임 (분)', false);
  addNumber(f, '메신저·SNS (분)', false);
  addNumber(f, '공부·검색 (분)', false);
  addNumber(f, '기타 (분)', false);
  addChoice(f, '오늘 사용하고 난 뒤 기분은 대체로 어땠나요?', ['좋았다', '그저 그랬다', '아쉬웠다'], true);
  addChoice(f, '계획한 것보다 오래 썼나요?', ['그렇다', '아니다', '계획이 없었다'], true);
  addChoice(f, '잠자기 30분 전에도 썼나요?', ['예', '아니요'], true);
  addChoice(f, '기록하려고 폰을 열었다가 다른 앱도 열었나요?', ['예', '아니요'], true);
  addText(f, '오늘 사용하면서 좋았던 점이나 아쉬웠던 점을 한 줄로 써 보세요.', false);
  f.setConfirmationMessage('오늘 기록 완료! 수고했어요.');
  return finish(ss, f, 'R_일일', '하루 기록 (1주) ※ 마지막에 스크린타임 캡처 업로드 질문 직접 추가');
}

function buildGoal(ss) {
  const f = newForm('[2차시] 나의 목표·실천 계획 카드',
    '내 기록을 분석한 결과를 바탕으로 목표와 실천 계획을 세웁니다.');
  addCode(f);
  addPara(f, '내 기록을 보니 ( )를 알 수 있다. 그 근거는 ( )이다. (그래프의 숫자를 한 가지 이상 넣어 쓰세요)', true);
  addChoice(f, '예상한 시간과 실제 기록을 비교하면 어땠나요?', ['예상보다 많았다', '비슷했다', '예상보다 적었다'], true);
  addChoice(f, '가장 고치고 싶은 점은 무엇인가요?', ['전체 사용 시간', '잠자기 전 사용', '사용 목적 바꾸기', '사용 시간대', '쓰고 난 뒤 기분', '기타'], true);
  addPara(f, '나의 목표는? (언제, 어디서, 얼마나가 들어가게 쓰세요)', true);
  addPara(f, '목표를 지키기 어렵게 만드는 것은 무엇일까요?', true);
  addPara(f, '만약 ( ) 하면, ( ) 한다. (방해 요인이 생겼을 때의 대처 계획)', true);
  addCheck(f, '매체 대신 할 수 있는 활동은? (모두 선택)', ['운동·산책', '책 읽기', '그림·만들기', '가족·친구와 놀기', '악기·취미', '기타'], true);
  addText(f, '나를 도와줄 사람은 누구이고, 어떻게 도와 달라고 할까요?', true);
  addChoice(f, '실천 기간은 며칠로 할까요?', ['7일', '10일', '14일'], true);
  addScale(f, '목표를 지킬 수 있다는 자신감은 어느 정도인가요?', 1, 5, '자신 없다', '자신 있다', true);
  f.setConfirmationMessage('계획이 완성되었어요. 오늘부터 실천을 시작해 봅시다!');
  return finish(ss, f, 'R_목표', '목표·실천 계획 카드 (2차시)');
}

function buildAction(ss) {
  const f = newForm('[실천 기간] 오늘의 실천 체크',
    '주 2~3회, 저녁에 1분만 쓰세요. 못 지킨 날도 솔직하게 쓰는 것이 가장 도움이 됩니다.');
  addCode(f);
  addDate(f, '날짜');
  addChoice(f, '오늘 목표를 지켰나요?', ['지켰다', '반쯤 지켰다', '못 지켰다'], true);
  addNumber(f, '오늘 매체 사용 시간은 몇 분인가요? (숫자만)', true);
  addText(f, '도움이 된 것은 무엇인가요?', false);
  addText(f, '어려웠던 것은 무엇인가요?', false);
  addChoice(f, '계획을 바꾸고 싶나요?', ['아니요', '조금 바꾸고 싶다', '많이 바꾸고 싶다'], true);
  f.setConfirmationMessage('체크 완료! 계속 이어 가 봅시다.');
  return finish(ss, f, 'R_실천', '실천 체크 (실천 기간) ※ 마지막에 스크린타임 캡처 업로드 질문 직접 추가');
}

function buildPost(ss) {
  const f = newForm('[3차시] 사후 설문과 소감',
    '실천을 마친 지금의 생각을 답해 주세요. 사전 설문과 같은 질문이 일부 있습니다.');
  addCode(f);
  addChoice(f, '평일 하루 동안 매체 사용 시간은 얼마나 되나요?', TIME_CHOICES, true);
  addChoice(f, '주말 하루는 얼마나 되나요?', TIME_CHOICES, true);
  addChoice(f, '잠자기 직전(30분 안)에 매체를 쓰나요?', SLEEP_CHOICES, true);
  addScale(f, '나의 매체 사용은 지금 어떤 것 같나요?', 1, 5, '너무 많다', '적당하다', true);
  addScale(f, '내가 세운 목표를 얼마나 지켰나요?', 1, 5, '거의 못 지켰다', '아주 잘 지켰다', true);
  addPara(f, '가장 도움이 된 방법은 무엇인가요?', true);
  addPara(f, '가장 어려웠던 점은 무엇인가요?', true);
  addPara(f, '이번 활동으로 새롭게 알게 된 점은 무엇인가요?', true);
  addText(f, '앞으로 이어 가고 싶은 목표 한 가지는 무엇인가요?', true);
  f.setConfirmationMessage('끝까지 해냈어요. 수고했습니다!');
  return finish(ss, f, 'R_사후', '사후 설문 (3차시)');
}

function buildFeedback(ss) {
  const f = newForm('[4차시] 발표 피드백',
    '친구의 발표를 듣고 남기는 글입니다. 비교하거나 평가하지 말고, 배울 점과 궁금한 점을 써 주세요.');
  addCode(f, '발표한 친구의 코드번호');
  addText(f, '발표에서 잘한 점', true);
  addText(f, '궁금한 점이나 질문', false);
  addText(f, '내 생활에 가져와 볼 만한 방법', true);
  f.setConfirmationMessage('피드백을 남겨 줘서 고마워요.');
  return finish(ss, f, 'R_피드백', '발표 피드백 (4차시)');
}

/* ───────── 도우미 ───────── */

function newForm(title, desc) {
  const f = FormApp.create(title);
  f.setDescription(desc);
  try { f.setCollectEmail(false); } catch (e) {}
  try { f.setAllowResponseEdits(false); } catch (e) {}
  try { f.setShowLinkToRespondAgain(true); } catch (e) {}
  return f;
}

function addCode(f, title) {
  const it = f.addTextItem().setTitle(title || '코드번호 (이름 대신 번호만 쓰세요)').setRequired(true);
  it.setValidation(FormApp.createTextValidation().requireNumber().setHelpText('숫자로 입력하세요.').build());
}
function addText(f, title, req) { f.addTextItem().setTitle(title).setRequired(!!req); }
function addPara(f, title, req) { f.addParagraphTextItem().setTitle(title).setRequired(!!req); }
function addDate(f, title) { f.addDateItem().setTitle(title).setRequired(true); }
function addNumber(f, title, req) {
  const it = f.addTextItem().setTitle(title).setRequired(!!req);
  it.setValidation(FormApp.createTextValidation().requireNumberBetween(0, 1440)
    .setHelpText('0~1440 사이의 숫자(분)로 입력하세요.').build());
}
function addChoice(f, title, choices, req) { f.addMultipleChoiceItem().setTitle(title).setChoiceValues(choices).setRequired(!!req); }
function addCheck(f, title, choices, req) { f.addCheckboxItem().setTitle(title).setChoiceValues(choices).setRequired(!!req); }
function addScale(f, title, lo, hi, left, right, req) {
  f.addScaleItem().setTitle(title).setBounds(lo, hi).setLabels(left, right).setRequired(!!req);
}

// 응답을 이 스프레드시트에 연결하고, 새로 생긴 응답 탭의 이름을 지정한다.
function finish(ss, f, sheetName, purpose) {
  const before = ss.getSheets().map(function (s) { return s.getSheetId(); });
  f.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  const added = ss.getSheets().filter(function (s) { return before.indexOf(s.getSheetId()) === -1; });
  let note = '';
  if (added.length === 1 && !ss.getSheetByName(sheetName)) {
    added[0].setName(sheetName);
  } else {
    note = '※ 응답 탭 이름을 자동으로 바꾸지 못했습니다. 새 탭 이름을 ' + sheetName + ' 으로 직접 바꿔 주세요.';
    Logger.log(f.getTitle() + ': ' + note);
  }
  return [purpose, f.getPublishedUrl(), f.getEditUrl(), sheetName, note];
}

function writeLinkSheet(ss, rows) {
  let sh = ss.getSheetByName('폼 링크');
  if (!sh) sh = ss.insertSheet('폼 링크');
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([['용도', '학생용 링크(배부)', '편집 링크(교사용)', '응답 탭', '비고']])
    .setFontWeight('bold').setBackground('#1F5C7A').setFontColor('#FFFFFF');
  sh.getRange(2, 1, rows.length, 5).setValues(rows);
  sh.setColumnWidths(1, 1, 360); sh.setColumnWidths(2, 2, 320); sh.setColumnWidths(4, 2, 140);
  sh.setFrozenRows(1);
}
