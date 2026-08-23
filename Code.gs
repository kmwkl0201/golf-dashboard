/**
 * 골프 라운딩 추억 앨범 백엔드 (Apps Script)
 *
 * ==========================================================================
 * 처음 배포하는 경우 (계정이 하나도 없는 완전히 새 프로젝트)
 * ==========================================================================
 * 1. https://script.google.com/ 에서 새 프로젝트 생성 후 이 파일 내용을 Code.gs에 붙여넣기
 * 2. 아래 PIN_SALT 값을 원하는 임의의 문자열로 한 번만 바꾸기
 *    (PIN 해시에 사용되는 값, 외부에 노출되면 안 됨)
 * 3. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나(본인)
 *    - 액세스 권한: 모든 사용자
 * 4. 배포 후 생성되는 웹 앱 URL을 index.html의 APPS_SCRIPT_URL 상수에 붙여넣기
 *
 * ==========================================================================
 * 이미 회원이 있는 상태에서 이 파일로 코드만 업데이트하는 경우 (★ 지금 이 상황)
 * ==========================================================================
 * 아래 PIN_SALT는 반드시 "지금 실제로 배포되어 있는 스크립트에 적혀 있는 값 그대로"
 * 붙여넣어야 합니다. 이 값이 조금이라도 다르면 이미 가입된 모든 회원의 PIN이 전부
 * 틀린 것으로 판정되어 아무도 로그인할 수 없게 됩니다.
 * 자세한 절차는 이 파일 하단의 "배포 절차" 안내, 또는 대화창의 안내 메시지를 참고하세요.
 *
 * ==========================================================================
 * 동작 방식
 * ==========================================================================
 * - 이 스크립트는 최초 실행 시 "골프기록_회원DB"라는 스프레드시트를 자동 생성해
 *   회원 정보(Users 시트: 이름/전화번호/PIN 해시/전용 스프레드시트ID/세션토큰)를 관리한다.
 * - 회원가입 시 사용자 전용 스프레드시트("이름님의 골프 기록")를 자동 생성하고,
 *   그 스프레드시트 ID를 Users 시트에 연결해둔다.
 * - 로그인/회원가입에 성공하면 세션 토큰을 발급하고, 이후 모든 기록 조회/추가/수정/삭제는
 *   그 토큰으로 사용자를 식별해 해당 사용자의 전용 스프레드시트에서만 처리한다.
 * - 라운딩 사진은 여러 장(최대 MAX_PHOTOS장) Drive에 올린 뒤 URL을 콤마로 이어붙여
 *   Records 시트의 Photos 열 한 칸에 저장한다.
 */

// ⚠️ 이후 절대 변경 금지 ⚠️
// 이 값은 회원 PIN 해시 계산에 쓰인다. 배포 이후 이 값을 바꾸면 그 순간부터
// 기존에 가입된 모든 회원이 올바른 PIN을 입력해도 로그인에 실패하게 된다.
// (새 프로젝트를 맨 처음 배포할 때 딱 한 번만 원하는 값으로 정하고, 그 다음부터는
//  코드를 아무리 다시 정리하거나 배포를 새로 해도 이 줄만은 절대 건드리지 말 것.)
var PIN_SALT = "CHANGE_THIS_SALT_VALUE";

var MAX_PHOTOS = 6;

/* ========================================================================
 * 진입점
 * ======================================================================== */

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
      case "changePin": return handleChangePin_(body);
      case "withdraw": return handleWithdraw_(body);
      case "add": return handleAdd_(body);
      case "update": return handleUpdate_(body);
      case "delete": return handleDelete_(body);
      default: return jsonOut_({ ok: false, error: "unknown_action" });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err.message || err) });
  }
}

/* ========================================================================
 * 회원가입 / 로그인
 * ======================================================================== */

function handleSignup_(body) {
  var name = (body.name || "").toString().trim();
  var phone = normalizePhone_(body.phone);
  var pin = (body.pin || "").toString().trim();
  var recoveryPin = (body.recoveryPin || "").toString().trim();

  if (!name) return jsonOut_({ ok: false, error: "name_required" });
  if (!phone) return jsonOut_({ ok: false, error: "phone_required" });
  if (pin.length < 4) return jsonOut_({ ok: false, error: "pin_too_short" });

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var usersSheet = getUsersSheet_();
    var rowIdx = findUserRowByPhone_(usersSheet, phone);

    if (rowIdx !== -1) {
      var status = usersSheet.getRange(rowIdx, 7).getValue();
      if (status !== "withdrawn") {
        return jsonOut_({ ok: false, error: "phone_already_registered" });
      }

      // 탈퇴한 계정과 같은 번호로 재가입을 시도하는 경우: 예전 PIN으로 본인 확인이
      // 되는 경우에만 기존 행(=기존 라운딩 기록 시트)을 그대로 재사용해 복구한다.
      // 전화번호+이름만 아는 제3자는 예전 PinHash를 절대 통과할 수 없다.
      var existingHash = usersSheet.getRange(rowIdx, 3).getValue();
      if (!recoveryPin || hashPin_(phone, recoveryPin) !== existingHash) {
        return jsonOut_({ ok: false, error: "account_withdrawn_recovery_needed" });
      }

      var recoveredToken = Utilities.getUuid();
      usersSheet.getRange(rowIdx, 1).setValue(name);                       // Name
      usersSheet.getRange(rowIdx, 3).setValue(hashPin_(phone, pin));       // PinHash
      usersSheet.getRange(rowIdx, 5).setValue(recoveredToken);             // SessionToken
      usersSheet.getRange(rowIdx, 7).setValue("active");                   // Status
      // Phone(B), SpreadsheetId(D), CreatedAt(F)는 그대로 둔다 — 기존 기록 유지.

      return jsonOut_({ ok: true, token: recoveredToken, name: name });
    }

    var userSpreadsheet = createUserSpreadsheet_(name);
    var token = Utilities.getUuid();
    usersSheet.appendRow([name, phone, hashPin_(phone, pin), userSpreadsheet.getId(), token, new Date(), "active"]);

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

  var rowVals = usersSheet.getRange(rowIdx, 1, 1, 7).getValues()[0];
  var storedHash = rowVals[2];
  if (hashPin_(phone, pin) !== storedHash) {
    return jsonOut_({ ok: false, error: "invalid_credentials" });
  }
  if (rowVals[6] === "withdrawn") {
    return jsonOut_({ ok: false, error: "account_withdrawn" });
  }

  var token = Utilities.getUuid();
  usersSheet.getRange(rowIdx, 5).setValue(token);
  return jsonOut_({ ok: true, token: token, name: rowVals[0] });
}

/* ========================================================================
 * 계정 관리 (로그인된 본인만 — 세션 토큰으로 신원 확인)
 * ======================================================================== */

function handleChangePin_(body) {
  var user = resolveUser_(body.token); // 유효하지 않거나 탈퇴한 계정이면 여기서 에러를 던진다.
  var phone = normalizePhone_(body.phone || user.phone);
  var currentPin = (body.currentPin || "").toString().trim();
  var newPin = (body.newPin || "").toString().trim();

  if (!currentPin || !newPin) return jsonOut_({ ok: false, error: "missing_fields" });
  if (newPin.length < 4) return jsonOut_({ ok: false, error: "pin_too_short" });

  var usersSheet = getUsersSheet_();
  var storedHash = usersSheet.getRange(user.row, 3).getValue();
  if (hashPin_(phone, currentPin) !== storedHash) {
    return jsonOut_({ ok: false, error: "current_pin_invalid" });
  }

  usersSheet.getRange(user.row, 3).setValue(hashPin_(phone, newPin)); // PinHash만 갱신
  return jsonOut_({ ok: true });
}

function handleWithdraw_(body) {
  var user = resolveUser_(body.token);
  var usersSheet = getUsersSheet_();
  usersSheet.getRange(user.row, 7).setValue("withdrawn"); // Status만 변경. 행/SpreadsheetId는 그대로 둔다.
  return jsonOut_({ ok: true });
}

/* ========================================================================
 * 라운딩 기록 CRUD (로그인된 사용자 전용 시트에서 처리)
 * ======================================================================== */

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

function uploadPhotos_(photos, limit) {
  var urls = [];
  (photos || []).forEach(function (photo) {
    if (urls.length >= limit) return;
    if (!photo || !photo.data) return;
    var bytes = Utilities.base64Decode(photo.data);
    var blob = Utilities.newBlob(bytes, photo.mimeType || "image/jpeg", photo.filename || "round.jpg");
    var file = DriveApp.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // 공유 설정 실패해도 기록 저장은 계속 진행
    }
    urls.push(file.getUrl());
  });
  return urls;
}

function handleAdd_(body) {
  var user = resolveUser_(body.token);
  var sheet = getRecordsSheet_(user.spreadsheetId);
  var record = body.record || {};

  var photos = Array.isArray(body.photos) ? body.photos : (body.photo ? [body.photo] : []);
  var photoUrls = uploadPhotos_(photos, MAX_PHOTOS);

  sheet.appendRow([
    record.date || "", record.course || "", record.score || "", record.par || "",
    record.weather || "", record.companions || "", record.memo || "", photoUrls.join(",")
  ]);
  return jsonOut_({ ok: true });
}

function handleUpdate_(body) {
  var user = resolveUser_(body.token);
  var sheet = getRecordsSheet_(user.spreadsheetId);
  var rowNum = parseInt(body.row, 10);
  if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
    return jsonOut_({ ok: false, error: "invalid_row" });
  }
  var record = body.record || {};

  var existingPhotos = Array.isArray(body.existingPhotos) ? body.existingPhotos.filter(Boolean) : [];
  var newUrls = uploadPhotos_(body.photos, Math.max(0, MAX_PHOTOS - existingPhotos.length));
  var photoUrls = existingPhotos.concat(newUrls).slice(0, MAX_PHOTOS);

  sheet.getRange(rowNum, 1, 1, 8).setValues([[
    record.date || "", record.course || "", record.score || "", record.par || "",
    record.weather || "", record.companions || "", record.memo || "", photoUrls.join(",")
  ]]);
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

/* ========================================================================
 * 사용자/시트 조회 및 생성
 * ======================================================================== */

function resolveUser_(token) {
  token = (token || "").toString().trim();
  if (!token) throw new Error("not_logged_in");

  var usersSheet = getUsersSheet_();
  var lastRow = usersSheet.getLastRow();
  if (lastRow >= 2) {
    var values = usersSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < values.length; i++) {
      if (values[i][4] === token) {
        if (values[i][6] === "withdrawn") throw new Error("account_withdrawn");
        return { row: i + 2, name: values[i][0], phone: values[i][1], spreadsheetId: values[i][3] };
      }
    }
  }
  throw new Error("invalid_token");
}

function findUserRowByPhone_(usersSheet, phone) {
  var lastRow = usersSheet.getLastRow();
  if (lastRow < 2) return -1;
  var targetKey = phoneKeyForCompare_(phone);
  if (!targetKey) return -1;
  var phones = usersSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  for (var i = 0; i < phones.length; i++) {
    if (phoneKeyForCompare_(phones[i][0]) === targetKey) return i + 2;
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
    sheet.appendRow(["Name", "Phone", "PinHash", "SpreadsheetId", "SessionToken", "CreatedAt", "Status"]);
  }
  // 기존 시트를 재사용하는 경우에도 매번 적용해야 신규 행에 텍스트 서식이 유지된다.
  sheet.getRange("B:B").setNumberFormat("@"); // 전화번호가 숫자로 변환되지 않도록 텍스트 서식 고정
  ensureStatusColumn_(sheet); // Status 열이 없던 옛 시트에도 열을 추가하고 기존 회원을 active로 채운다.
  return sheet;
}

// Status 열(G)이 없는 예전 Users 시트에도 헤더를 추가하고, 기존 회원 행은 전부
// "active"로 채워 넣는다. 이미 Status 열이 있으면 아무것도 바꾸지 않는다(멱등).
function ensureStatusColumn_(sheet) {
  var header = sheet.getRange(1, 7).getValue();
  if (header !== "Status") {
    sheet.getRange(1, 7).setValue("Status");
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var statusRange = sheet.getRange(2, 7, lastRow - 1, 1);
  var values = statusRange.getValues();
  var changed = false;
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      values[i][0] = "active";
      changed = true;
    }
  }
  if (changed) statusRange.setValues(values);
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

/* ========================================================================
 * 유틸
 * ======================================================================== */

function normalizePhone_(phone) {
  return (phone || "").toString().replace(/[^0-9]/g, "");
}

// 전화번호를 숫자만 남긴 뒤 맨 앞의 0들까지 제거해 "비교용 키"로 만든다.
// 구글 시트가 Phone 열을 숫자로 잘못 인식해 맨 앞 0이 사라진 과거 데이터가 있어도
// (예: "01023793399" -> 1023793399) 같은 번호로 인식해서 로그인/중복가입 체크가
// 깨지지 않도록 하기 위한 안전장치다.
function phoneKeyForCompare_(phone) {
  return normalizePhone_(phone).replace(/^0+/, "");
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

/* ========================================================================
 * [관리자 전용] 회원 PIN 재설정
 * ========================================================================
 * 이 함수는 웹앱(doGet/doPost)에서는 절대 호출되지 않는다. Apps Script 편집기에서
 * 관리자가 직접 선택해서 실행하는 용도로만 존재한다 — 그래서 외부에서 URL로 접근해도
 * 아무도 이 함수를 실행시킬 수 없다.
 *
 * 사용법
 * 1. 아래 PHONE(전화번호)과 NEW_PIN(새 PIN, 숫자 4자리 이상)을 원하는 값으로 바꾼다.
 *    PHONE은 회원이 로그인할 때 입력하는 것과 같은 형식(맨 앞 0 포함)으로 적으면 된다.
 * 2. Apps Script 편집기 상단에서 함수 목록을 "resetMemberPin"으로 선택하고 실행한다.
 * 3. 실행 로그(보기 > 실행 기록)에서 성공 메시지를 확인한다.
 * 4. 새 PIN으로 실제 로그인이 되는지 확인한다.
 *
 * 이 함수는 해당 회원의 PinHash 칸만 바꾸고, SpreadsheetId(개인 라운딩 기록 시트 연결)는
 * 절대 건드리지 않으므로 기존 라운딩 기록은 그대로 유지된다.
 */
function resetMemberPin() {
  var PHONE = "01000000000"; // 재설정할 회원의 전화번호로 바꾸세요.
  var NEW_PIN = "0000";      // 새로 설정할 PIN(숫자 4자리 이상)으로 바꾸세요.

  var pin = (NEW_PIN || "").toString().trim();
  if (pin.length < 4) {
    Logger.log("NEW_PIN은 4자리 이상 숫자여야 합니다. 현재 값: " + NEW_PIN);
    return;
  }

  var usersSheet = getUsersSheet_();
  var rowIdx = findUserRowByPhone_(usersSheet, PHONE);
  if (rowIdx === -1) {
    Logger.log("해당 전화번호의 회원을 찾을 수 없습니다: " + PHONE);
    return;
  }

  var phoneForHash = normalizePhone_(PHONE);
  var newHash = hashPin_(phoneForHash, pin);
  usersSheet.getRange(rowIdx, 3).setValue(newHash); // PinHash(C열)만 갱신. SpreadsheetId(D열)는 그대로 둔다.

  var name = usersSheet.getRange(rowIdx, 1).getValue();
  Logger.log("완료: " + name + "(" + PHONE + ")의 PIN이 재설정됐습니다. 새 PIN(" + pin + ")으로 로그인해보세요.");
}
