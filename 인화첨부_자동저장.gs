/**********************************************************************
 *  현스냅 · 인화 첨부파일 자동 저장
 *  ──────────────────────────────────────────────────────────────────
 *  제목에 "날짜(260220) + 한글이름(홍길동)" 이 들어있는 지메일을 찾아
 *  그 제목으로 구글 드라이브 폴더를 만들고 첨부파일을 저장합니다.
 *
 *  ✅ 하는 일 : 메일 읽기 / 드라이브 폴더 생성 / 첨부파일 저장
 *  🚫 안 하는 일 : 파일·폴더 삭제, 이동, 덮어쓰기 (절대 안 함)
 *
 *  처음 쓰는 순서
 *   1) 아래 '상위폴더ID' 에 저장할 드라이브 폴더 ID 붙여넣기
 *   2) 함수 [미리보기] 실행 → 실행기록(로그)에서 어떤 메일이 잡히는지 확인
 *   3) 괜찮으면 함수 [트리거설치] 한 번 실행 → 10분마다 자동 실행됨
 **********************************************************************/


/* ===================== 설정 (여기만 손대면 됩니다) ===================== */
var 설정 = {
  // 저장할 상위 폴더 ID.
  //   드라이브에서 폴더 열기 → 주소창 .../folders/  뒤의 글자들이 ID
  상위폴더ID: "여기에_드라이브_폴더ID_붙여넣기",

  // 처리한 메일에 붙일 라벨 (같은 메일 두 번 저장 방지)
  처리완료라벨: "인화저장완료",

  // 최근 며칠 메일만 검사할지 (서버 부담 줄이기용)
  검색기간: "newer_than:30d",

  // (선택) 제목에 이 단어가 꼭 있어야만 저장하고 싶으면 적기. 비워두면 무시.
  //   예: "인화"  /  안전하게 더 좁히고 싶을 때만 사용
  필수단어: ""
};
/* ===================================================================== */


/* ---------- 메인: 자동 저장 (트리거가 이걸 실행) ---------- */
function 첨부파일자동저장() {
  var 상위 = DriveApp.getFolderById(설정.상위폴더ID);
  var 라벨 = 라벨가져오기(설정.처리완료라벨);
  var 검색어 = 'has:attachment ' + 설정.검색기간 + ' -label:' + 설정.처리완료라벨;

  var 메일목록 = GmailApp.search(검색어, 0, 50);
  var 저장건수 = 0;

  for (var t = 0; t < 메일목록.length; t++) {
    var 스레드 = 메일목록[t];
    var 메시지들 = 스레드.getMessages();
    var 이스레드저장함 = false;

    for (var m = 0; m < 메시지들.length; m++) {
      var 메시지 = 메시지들[m];
      var 제목 = 메시지.getSubject();

      if (!제목매치(제목)) continue;                 // 날짜+한글이름 아니면 건너뜀

      // 첨부파일만 (메일 서명에 박힌 로고 등 인라인 이미지는 제외)
      var 첨부 = 메시지.getAttachments({ includeInlineImages: false, includeAttachments: true });
      if (첨부.length === 0) continue;

      var 폴더 = 폴더가져오기또는생성(상위, 폴더명만들기(제목));

      for (var a = 0; a < 첨부.length; a++) {
        var 이름 = 첨부[a].getName();
        // 이미 같은 이름 파일이 있으면 건너뜀 (덮어쓰기·삭제 절대 안 함)
        if (폴더.getFilesByName(이름).hasNext()) {
          Logger.log('건너뜀(이미있음): ' + 폴더.getName() + ' / ' + 이름);
          continue;
        }
        폴더.createFile(첨부[a].copyBlob()).setName(이름);
        저장건수++;
        Logger.log('저장: ' + 폴더.getName() + ' / ' + 이름);
      }
      이스레드저장함 = true;
    }

    if (이스레드저장함) 스레드.addLabel(라벨);          // 처리 표시 → 다음엔 안 잡힘
  }

  Logger.log('완료 · 새로 저장한 파일 ' + 저장건수 + '개');
}


/* ---------- 미리보기: 저장 안 하고 결과를 '구글 시트'에 적어줌 ---------- */
//  폰에서도 시트만 열면 결과를 볼 수 있습니다. (실제 저장은 안 함)
function 미리보기() {
  var 검색어 = 'has:attachment ' + 설정.검색기간 + ' -label:' + 설정.처리완료라벨;
  var 메일목록 = GmailApp.search(검색어, 0, 50);

  // '미리보기결과' 시트 준비 (없으면 만들고, 있으면 비우고 다시 씀)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('미리보기결과') || ss.insertSheet('미리보기결과');
  sh.clear();
  sh.getRange(1, 1, 1, 4)
    .setValues([['결과', '폴더명(저장될 이름)', '첨부수', '메일제목']])
    .setFontWeight('bold').setBackground('#f0ece6');

  var 행 = [];
  for (var t = 0; t < 메일목록.length; t++) {
    var 메시지들 = 메일목록[t].getMessages();
    for (var m = 0; m < 메시지들.length; m++) {
      var 제목 = 메시지들[m].getSubject();
      var 첨부수 = 메시지들[m].getAttachments({ includeInlineImages: false }).length;
      if (제목매치(제목)) {
        행.push(['✅ 저장대상', 폴더명만들기(제목), 첨부수, 제목]);
      } else {
        행.push(['— 제외', '', 첨부수, 제목]);
      }
    }
  }

  if (행.length > 0) sh.getRange(2, 1, 행.length, 4).setValues(행);
  sh.getRange(행.length + 3, 1).setValue('※ 미리보기는 실제로 저장하지 않습니다. 검사한 메일 ' + 메일목록.length + '개.');
  sh.autoResizeColumns(1, 4);
}


/* ---------- 트리거 설치: 10분마다 자동 실행 (한 번만 실행) ---------- */
function 트리거설치() {
  var 기존 = ScriptApp.getProjectTriggers();
  for (var i = 0; i < 기존.length; i++) {
    if (기존[i].getHandlerFunction() === '첨부파일자동저장') ScriptApp.deleteTrigger(기존[i]);
  }
  ScriptApp.newTrigger('첨부파일자동저장').timeBased().everyMinutes(10).create();
  Logger.log('설치 완료 · 이제 10분마다 자동으로 돕니다.');
}


/* ===================== 아래는 손대지 않아도 됩니다 ===================== */

// 제목이 "날짜 + 한글이름" 조건에 맞는지
function 제목매치(제목) {
  if (!제목) return false;
  if (설정.필수단어 && 제목.indexOf(설정.필수단어) === -1) return false;
  return 날짜있나(제목) && 한글이름있나(제목);
}

// 날짜처럼 보이는 토큰: YYMMDD 또는 YYYYMMDD (구분자 . - / 공백 허용, 월·일 범위 검증)
function 날짜있나(s) {
  return /(^|[^0-9])(\d{2}|\d{4})[.\-\/ ]?(0[1-9]|1[0-2])[.\-\/ ]?(0[1-9]|[12]\d|3[01])([^0-9]|$)/.test(s);
}

// 한글 이름(2~4글자) 포함 여부
function 한글이름있나(s) {
  return /[가-힣]{2,4}/.test(s);
}

// 제목 → 폴더 이름 (Re:/전달: 제거, 폴더에 못 쓰는 문자 정리)
function 폴더명만들기(제목) {
  var s = String(제목).replace(/^\s*(re|fwd?|회신|전달|답장)\s*:\s*/gi, '').trim();
  s = s.replace(/[\\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim();
  return s || '제목없음';
}

// 같은 이름 폴더가 있으면 그대로 쓰고, 없으면 새로 만듦 (중복 폴더 방지)
function 폴더가져오기또는생성(상위, 이름) {
  var it = 상위.getFoldersByName(이름);
  return it.hasNext() ? it.next() : 상위.createFolder(이름);
}

// 라벨 가져오기(없으면 생성)
function 라벨가져오기(이름) {
  return GmailApp.getUserLabelByName(이름) || GmailApp.createLabel(이름);
}
