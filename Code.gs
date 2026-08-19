/**
 * 가계부 앱 — 카드 결제 문자 웹훅 (Google Apps Script)
 *
 * 하는 일:
 *  1) 아이폰 단축어가 카드 결제 문자를 받으면 이 스크립트로 POST 요청을 보낸다.
 *  2) 문자 내용에서 금액/가맹점/날짜를 파싱해서 이 문서의 시트에 저장한다.
 *  3) 가계부 웹앱이 이 스크립트에 GET 요청을 보내면, 아직 앱에 반영 안 된
 *     "확인 대기" 내역을 JSON(JSONP)으로 돌려준다.
 *  4) 웹앱에서 "확정" 버튼을 누르면 이 스크립트로 다시 알려줘서 시트 상태를
 *     confirmed 로 바꾼다 (다음 조회 때 중복으로 다시 오지 않게).
 *
 * 설치 방법은 SETUP_GUIDE.md 파일을 참고하세요.
 */

// ⚠️ 반드시 아래 값을 아무나 못 맞출 임의의 문자열로 바꾸세요.
// 이 토큰이 곧 "비밀번호" 역할을 합니다. 단축어 설정과 웹앱 설정에도 똑같이 넣어야 합니다.
var SECRET_TOKEN = '여기에-나만-아는-임의의-문자열-입력';

// 이 카드로만 쓸 거라면 그냥 고정값으로 둬도 됩니다. 여러 카드를 쓰게 되면
// 문자 발신번호별로 분기하도록 나중에 확장하면 됩니다.
var DEFAULT_METHOD = '신한 체크';

var SHEET_NAME = 'transactions';
var HEADERS = ['id', 'date', 'amount', 'merchant', 'category', 'method', 'status', 'rawText', 'createdAt'];

/* ---------------- 시트 준비 ---------------- */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

/* ---------------- 카테고리 추정 (웹앱과 동일한 규칙) ---------------- */
var KEYWORD_MAP = [
  [/스벅|스타벅스|커피|카페|이디야|투썸|메가엠지씨|빽다방/, 'cafe'],
  [/지하철|버스|택시|주유|따릉이|교통/, 'transport'],
  [/gs25|cu|세븐일레븐|이마트24|편의점/i, 'store'],
  [/쿠팡|마켓컬리|올리브영|무신사|지그재그|쇼핑/, 'shopping'],
  [/넷플릭스|유튜브|멜론|왓챠|스포티파이|구독/, 'sub'],
  [/영화|팝콘|전시|공연|씨지브이|메가박스/, 'culture'],
  [/김밥|국밥|식당|맥도날드|버거|치킨|배달|점심|저녁|아침/, 'food']
];
function guessCategory_(text) {
  for (var i = 0; i < KEYWORD_MAP.length; i++) {
    if (KEYWORD_MAP[i][0].test(text)) return KEYWORD_MAP[i][1];
  }
  return 'etc';
}

/* ---------------- 카드 승인 문자 파싱 ----------------
 * 예: "신한카드(체크)승인 김민수님 5,600원 일시불 08/19 09:12 스타벅스 강남R점 누적1,234,000원"
 * 카드사 문자 형식은 조금씩 다를 수 있어서, 실제 문자를 보고 정규식을 같이 다듬어야 할 수 있습니다.
 */
function cleanMerchant_(s, amountStr) {
  if (!s) return '';
  var t = s;
  if (amountStr) t = t.split(amountStr).join('');
  t = t.replace(/일시불|할부\s*\d*\s*개월?|승인|신한카드|신한체크|카드체크|원/g, '');
  t = t.replace(/[^가-힣a-zA-Z0-9()·\-\s]/g, '');
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}

function parseCardSms_(text) {
  var original = text;
  // "누적 ...원" 이후는 이번 결제 금액이 아니라 카드 한도/누적액이므로 잘라낸다.
  var main = text.replace(/누적[\d,]*\s*원[^]*$/, '');

  var amountMatch = main.match(/(\d[\d,]{2,})\s*원/);
  var amount = amountMatch ? parseInt(amountMatch[1].replace(/,/g, ''), 10) : 0;

  var dateMatch = main.match(/(\d{1,2})[\/.](\d{1,2})\s+(\d{1,2}):(\d{2})/);
  var iso = new Date().toISOString();
  if (dateMatch) {
    var now = new Date();
    var d = new Date(now.getFullYear(), parseInt(dateMatch[1], 10) - 1, parseInt(dateMatch[2], 10),
                      parseInt(dateMatch[3], 10), parseInt(dateMatch[4], 10));
    iso = d.toISOString();
  }

  // 가맹점명은 보통 "날짜/시간" 뒤에 오거나(문자 끝), 그게 비어있으면
  // "금액"과 "날짜/시간" 사이에 온다. 두 후보를 순서대로 시도한다.
  var amountEnd = amountMatch ? main.indexOf(amountMatch[0]) + amountMatch[0].length : 0;
  var dateStart = dateMatch ? main.indexOf(dateMatch[0]) : -1;
  var dateEnd = dateMatch ? dateStart + dateMatch[0].length : -1;

  var candidateAfterDate = dateEnd > -1 ? main.slice(dateEnd) : '';
  var candidateBetween = (dateStart > amountEnd) ? main.slice(amountEnd, dateStart) : '';
  var candidateAfterAmount = main.slice(amountEnd);
  var amtStr = amountMatch ? amountMatch[0] : null;

  var merchant = cleanMerchant_(candidateAfterDate, amtStr) ||
                 cleanMerchant_(candidateBetween, amtStr) ||
                 cleanMerchant_(candidateAfterAmount, amtStr);
  if (!merchant) merchant = '카드결제';

  var category = guessCategory_(merchant);
  return { amount: amount, merchant: merchant, date: iso, category: category, raw: original };
}

/* ---------------- doPost: 단축어 -> 시트 저장 / 확정 처리 ---------------- */
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid body' });
  }

  if (body.token !== SECRET_TOKEN) {
    return jsonOut_({ ok: false, error: 'unauthorized' });
  }

  // 웹앱에서 "확정" 눌렀을 때 오는 요청: { token, action:'confirm', id }
  if (body.action === 'confirm' && body.id) {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(body.id)) {
        sheet.getRange(r + 1, HEADERS.indexOf('status') + 1).setValue('confirmed');
        break;
      }
    }
    return jsonOut_({ ok: true });
  }

  // 단축어에서 오는 요청: { token, text: "문자 원문" }
  if (body.text) {
    var parsed = parseCardSms_(body.text);
    if (!parsed.amount) {
      return jsonOut_({ ok: false, error: 'amount not parsed', raw: body.text });
    }
    var id = Utilities.getUuid();
    var now2 = new Date().toISOString();
    var sh = getSheet_();
    sh.appendRow([id, parsed.date, parsed.amount, parsed.merchant, parsed.category,
                  DEFAULT_METHOD, 'pending', parsed.raw, now2]);
    return jsonOut_({ ok: true, id: id, parsed: parsed });
  }

  return jsonOut_({ ok: false, error: 'unknown request' });
}

/* ---------------- doGet: 웹앱 -> 대기중 내역 조회 (JSONP) ---------------- */
function doGet(e) {
  var token = e.parameter.token;
  if (token !== SECRET_TOKEN) {
    return respond_(e, { ok: false, error: 'unauthorized' });
  }

  var since = e.parameter.since ? new Date(e.parameter.since) : null;
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var status = row[HEADERS.indexOf('status')];
    var createdAt = new Date(row[HEADERS.indexOf('createdAt')]);
    if (status === 'pending' && (!since || createdAt > since)) {
      rows.push({
        id: row[HEADERS.indexOf('id')],
        date: row[HEADERS.indexOf('date')],
        amount: row[HEADERS.indexOf('amount')],
        merchant: row[HEADERS.indexOf('merchant')],
        category: row[HEADERS.indexOf('category')],
        method: row[HEADERS.indexOf('method')],
        status: status
      });
    }
  }
  return respond_(e, { ok: true, items: rows, serverTime: new Date().toISOString() });
}

function respond_(e, obj) {
  if (e.parameter.callback) {
    return ContentService
      .createTextOutput(e.parameter.callback + '(' + JSON.stringify(obj) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut_(obj);
}
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 스크립트 편집기에서 바로 테스트하고 싶을 때 ---------------- */
function _test() {
  var sample = '신한카드(체크)승인 김민수님 5,600원 일시불 08/19 09:12 스타벅스 강남R점 누적1,234,000원';
  Logger.log(JSON.stringify(parseCardSms_(sample)));
}
