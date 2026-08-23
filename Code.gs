/**
 * 골프 라운딩 대시보드 백엔드 (Apps Script)
 *
 * 배포 방법
 * 1. https://script.google.com/ 에서 새 프로젝트 생성 후 이 파일 내용을 Code.gs에 붙여넣기
 * 2. 아래 PIN_SALT 값을 원하는 임의의 문자열로 바꾸기 (PIN 해시에 사용되는 값, 외부에 노출되면 안 됨)
 * 3. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 후 생성되는 웹 앱 URL을 index.html의 APPS_SCRIPT_URL 상수에 붙여넣기
 *
 * 동작 방식
 * - 이 스크립트는 최초 실행 시 "골프기록_회원DB"라는 스프레드시트를 자동 생성해
 *   회원 정보(Users 시트: 이름/전화번호/PIN 해시/전용 스프레드시트ID/세션토큰)를 관리한다.
 * - 회원가입 시 사용자 전용 스프레드시트("이름님의 골프 기록")를 자동 생성하고,
 *   그 스프레드시트 ID를 Users 시트에 연결해둔다.
 * - 로그인/회원가입에 성공하면 세션 토큰을 발급하고, 이후 모든 기록 조회/추가/삭제는
 *   그 토큰으로 사용자를 식별해 해당 사용자의 전용 스프레드시트에서만 처리한다.
 */

var PIN_SALT = "CHANGE_THIS_SALT_VALUE";
var MAX_PHOTOS = 6;

/* ---------- 진입점 ---------- */

function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === "list") return handleList_(e.parameter.token);
    return jsonOut_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: "invalid_json" });
  }
  try {
    switch (body.action) {
      case "signup": return handleSignup_(body);
      case "login": return handleLogin_(body);
      case "add": return handleAdd_(body);
      case "delete": return handleDelete_(body);
      default: return jsonOut_({ ok: false, error: "unknown_action" });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- 회원가입 / 로그인 ---------- */

function handleSignup_(body) {
  var name = (body.name || "").toString().trim();
  var phone = normalizePhone_(body.phone);
  var pin = (body.pin || "").toString().trim();

  if (!name) return jsonOut_({ ok: false, error: "name_required" });
  if (!phone) return jsonOut_({ ok: false, error: "phone_required" });
  if (pin.length < 4) return jsonOut_({ ok: false, error: "pin_too_short" });

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var usersSheet = getUsersSheet_();
    if (findUserRowByPhone_(usersSheet, phone) !== -1) {
      return jsonOut_({ ok: false, error: "phone_already_registered" });
    }

    var userSpreadsheet = createUserSpreadsheet_(name);
    var token = Utilities.getUuid();
    usersSheet.appendRow([name, phone, hashPin_(phone, pin), userSpreadsheet.getId(), token, new Date()]);

    return jsonOut_({ ok: true, token: token, name: name });
  } finally {
    lock.releaseLock();
  }
}

function handleLogin_(body) {
  var phone = normalizePhone_(body.phone);
  var pin = (body.pin || "").toString().trim();
  if (!phone || !pin) return jsonOut_({ ok: false, error: "missing_fields" });

  var usersSheet = getUsersSheet_();
  var rowIdx = findUserRowByPhone_(usersSheet, phone);
  if (rowIdx === -1) return jsonOut_({ ok: false, error: "invalid_credentials" });

  var rowVals = usersSheet.getRange(rowIdx, 1, 1, 6).getValues()[0];
  var storedHash = rowVals[2];
  if (hashPin_(phone, pin) !== storedHash) {
    return jsonOut_({ ok: false, error: "invalid_credentials" });
  }

  var token = Utilities.getUuid();
  usersSheet.getRange(rowIdx, 5).setValue(token);
  return jsonOut_({ ok: true, token: token, name: rowVals[0] });
}

/* ---------- 라운딩 기록 CRUD (로그인된 사용자 전용 시트에서 처리) ---------- */

function handleList_(token) {
  var user = resolveUser_(token);
  var sheet = getRecordsSheet_(user.spreadsheetId);
  var lastRow = sheet.getLastRow();
  var rows = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    values.forEach(function (v, i) {
      if (!v[0] && !v[1]) return;
      var photos = String(v[7] || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean);
      rows.push({
        row: i + 2,
        date: formatDate_(v[0]),
        course: v[1],
        score: v[2],
        par: v[3],
        weather: v[4],
        companions: v[5],
        memo: v[6],
        photos: photos
      });
    });
  }
  return jsonOut_({ ok: true, rows: rows });
}

function handleAdd_(body) {
  var user = resolveUser_(body.token);
  var sheet = getRecordsSheet_(user.spreadsheetId);
  var record = body.record || {};

  var photos = Array.isArray(body.photos) ? body.photos : (body.photo ? [body.photo] : []);
  var photoUrls = [];
  photos.slice(0, MAX_PHOTOS).forEach(function (photo) {
    if (!photo || !photo.data) return;
    var bytes = Utilities.base64Decode(photo.data);
    var blob = Utilities.newBlob(bytes, photo.mimeType || "image/jpeg", photo.filename || "round.jpg");
    var file = DriveApp.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // 공유 설정 실패해도 기록 저장은 계속 진행
    }
    photoUrls.push(file.getUrl());
  });

  sheet.appendRow([
    record.date || "", record.course || "", record.score || "", record.par || "",
    record.weather || "", record.companions || "", record.memo || "", photoUrls.join(",")
  ]);
  return jsonOut_({ ok: true });
}

function handleDelete_(body) {
  var user = resolveUser_(body.token);
  var sheet = getRecordsSheet_(user.spreadsheetId);
  var rowNum = parseInt(body.row, 10);
  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return jsonOut_({ ok: false, error: "invalid_row" });
  }
  sheet.deleteRow(rowNum);
  return jsonOut_({ ok: true });
}

/* ---------- 사용자/시트 조회 및 생성 ---------- */

function resolveUser_(token) {
  token = (token || "").toString().trim();
  if (!token) throw new Error("not_logged_in");

  var usersSheet = getUsersSheet_();
  var lastRow = usersSheet.getLastRow();
  if (lastRow >= 2) {
    var values = usersSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][4] === token) {
        return { name: values[i][0], phone: values[i][1], spreadsheetId: values[i][3] };
      }
    }
  }
  throw new Error("invalid_token");
}

function findUserRowByPhone_(usersSheet, phone) {
  var lastRow = usersSheet.getLastRow();
  if (lastRow < 2) return -1;
  var phones = usersSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < phones.length; i++) {
    if (String(phones[i][0]) === phone) return i + 2;
  }
  return -1;
}

function getUsersSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("USERS_SPREADSHEET_ID");
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create("골프기록_회원DB");
    props.setProperty("USERS_SPREADSHEET_ID", ss.getId());
  }

  var sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Name", "Phone", "PinHash", "SpreadsheetId", "SessionToken", "CreatedAt"]);
  }
  // 기존 시트를 재사용하는 경우에도 매번 적용해야 신규 행에 텍스트 서식이 유지된다.
  sheet.getRange("B:B").setNumberFormat("@"); // 전화번호가 숫자로 변환되지 않도록 텍스트 서식 고정
  return sheet;
}

function createUserSpreadsheet_(name) {
  var ss = SpreadsheetApp.create(name + "님의 골프 기록");
  var sheet = ss.getActiveSheet();
  sheet.setName("Records");
  sheet.appendRow(["Date", "Course", "Score", "Par", "Weather", "Companions", "Memo", "Photos"]);
  sheet.getRange("A:A").setNumberFormat("@"); // 날짜가 Date 타입으로 자동 변환되지 않도록 텍스트 서식 고정
  return ss;
}

function getRecordsSheet_(spreadsheetId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName("Records");
  if (!sheet) {
    sheet = ss.insertSheet("Records");
    sheet.appendRow(["Date", "Course", "Score", "Par", "Weather", "Companions", "Memo", "Photos"]);
    sheet.getRange("A:A").setNumberFormat("@");
  }
  return sheet;
}

/* ---------- 유틸 ---------- */

function normalizePhone_(phone) {
  return (phone || "").toString().replace(/[^0-9]/g, "");
}

function hashPin_(phone, pin) {
  var raw = phone + ":" + pin + ":" + PIN_SALT;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function formatDate_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
