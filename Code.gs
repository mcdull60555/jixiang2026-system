/**
 * 基翔小學通：學生資料 / 派發項目雲端資料庫（含登入驗證）
 *
 * 儲存方式：
 *   __APP_DATA__ 工作表
 *   A 欄 = JSON 分段內容（避免單一儲存格 50,000 字元限制）
 *
 * 安全設計：
 *   - 密碼以「加鹽 SHA-256」雜湊後儲存，網頁端永遠拿不到任何密碼。
 *   - 未登入只能取得「登入選單用的學生名單」，看不到作業、密碼或其他資料。
 *   - 登入成功後後端發一組憑證（token），之後每個請求都要帶；
 *     學生憑證只能讀寫自己的資料，老師憑證才能整份儲存與審核。
 *   - 連續輸錯 5 次，該帳號鎖定 1 分鐘，防止暴力猜密碼。
 *   - 老師端可以看到學生密碼：另外存一份「可還原」的 pwdView（以指令碼屬性裡的隨機金鑰加密），
 *     只在老師憑證的回應裡解開成 pwdPlain；學生與未登入的回應完全拿不到。登入驗證仍只看雜湊。
 *     這是可還原的，安全性低於純雜湊；舊的學生密碼（只有雜湊）無法還原，重設或學生自己改一次就會出現。
 *   - 電子聯絡簿 contactBook（日期/星期/內容）隨整份資料儲存；學生憑證可以讀、不能改。
 *
 * API：
 *   GET  ?action=get       公開：登入選單用的學生名單（姓名/學校/年級/班級）
 *   POST ?action=login     { role, pwd, studentId? } → { token, data }
 *   POST ?action=data      { token } → 依身分回傳可看的資料
 *   POST ?action=logout    { token }
 *   POST ?action=save      { token, data }  老師：整份儲存（密碼欄位有填才會更新）
 *   POST ?action=patch     { token, ops }   逐筆更新審核欄位
 *   POST ?action=changePwd { token, newPwd } 學生修改自己的密碼
 *
 * 第一次部署請在編輯器執行一次 migratePasswords()，把現有明碼密碼改成雜湊。
 * 忘記老師密碼時，可在編輯器執行 resetTeacherPasswordTo1234()。
 */

// 後端版本：用瀏覽器開 /exec 網址可看到，確認部署的是新版
const BACKEND_VERSION = "2026-10-09-v4-contactbook";
const CONTACT_BOOK_KEEP_DAYS = 10; // 聯絡簿日期超過這麼多天就自動刪除
const APP_TZ = "Asia/Taipei"; // 週的計算（週一～週日）固定用台灣時間
const DATA_SHEET_NAME = "__APP_DATA__";
const CHUNK_SIZE = 45000;

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    sheet.getRange("A1").setValue("");
    sheet.hideSheet();
  }

  return sheet;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readStoredJson_() {
  const sheet = getDataSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 1) return "";

  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  return values
    .map(row => row[0] == null ? "" : String(row[0]))
    .join("");
}

function writeStoredJson_(json) {
  const sheet = getDataSheet_();

  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push([json.substring(i, i + CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks.push([""]);

  const prevLastRow = sheet.getLastRow();

  // 只寫入需要的列（純文字格式，避免以 = 開頭的片段被當成公式）
  const range = sheet.getRange(1, 1, chunks.length, 1);
  range.setNumberFormat("@");
  range.setValues(chunks);

  // 新資料較短時，只清除多出來的舊列，不再清整張表
  if (prevLastRow > chunks.length) {
    sheet.getRange(chunks.length + 1, 1, prevLastRow - chunks.length, 1).clearContent();
  }
  SpreadsheetApp.flush();
}

function validateAppData_(data) {
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("收到的資料格式不是物件");
  }

  if (!Array.isArray(data.students)) {
    throw new Error("缺少 students 陣列");
  }

  if (!Array.isArray(data.tasks)) {
    throw new Error("缺少 tasks 陣列");
  }

  if (!Array.isArray(data.subjects)) {
    throw new Error("缺少 subjects 陣列");
  }

  if (!Array.isArray(data.grades)) {
    throw new Error("缺少 grades 陣列");
  }

  // 電子聯絡簿：可以沒有（舊資料），有的話一定要是陣列
  if (data.contactBook !== undefined && !Array.isArray(data.contactBook)) {
    throw new Error("contactBook 必須是陣列");
  }

  return true;
}


/* ==================================================================
 * 密碼、憑證、防暴力破解
 * ================================================================== */
const DEFAULT_TEACHER_PWD = "1234";
const SESSION_TTL_SEC = 21600;   // 憑證有效 6 小時（CacheService 上限）；每次使用都會延長
const MAX_LOGIN_FAILS = 5;
const LOGIN_LOCK_SEC = 60;

function randomHex_(n) {
  let s = "";
  while (s.length < n) s += Utilities.getUuid().replace(/-/g, "");
  return s.substring(0, n);
}

function toHex_(bytes) {
  return bytes.map(function (b) {
    const v = b < 0 ? b + 256 : b;
    return (v < 16 ? "0" : "") + v.toString(16);
  }).join("");
}

function hashWithSalt_(plain, salt) {
  return toHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + ":" + plain,
    Utilities.Charset.UTF_8
  ));
}

function makePwdHash_(plain) {
  const salt = randomHex_(16);
  return "sha256$" + salt + "$" + hashWithSalt_(String(plain), salt);
}

function isHashed_(v) {
  return typeof v === "string" && v.indexOf("sha256$") === 0;
}

/* ------------------------------------------------------------------
 * 老師可查看的「還原用密碼」（pwdView）
 *
 * 登入驗證仍然只看加鹽雜湊（pwd）。pwdView 是另外存的一份可還原密碼，
 * 只在「老師憑證」的回應裡解開成 pwdPlain 給老師端顯示；學生憑證與未登入
 * 的回應完全不會包含它。
 *
 * 做法：用存在「指令碼屬性」裡的一把隨機金鑰，以 HMAC-SHA256 產生金鑰流，
 * 對密碼做 XOR 後存成 base64，因此試算表裡看不到明碼。
 * 注意：這是可還原的，安全性低於純雜湊；金鑰與資料都在同一個 Google 帳號下，
 * 請只讓信任的人擁有試算表與 Apps Script 專案的編輯權限。
 * ------------------------------------------------------------------ */
function getPwdViewSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("PWD_VIEW_SECRET");
  if (!secret) {
    secret = randomHex_(64);
    props.setProperty("PWD_VIEW_SECRET", secret);
  }
  return secret;
}

// v2 金鑰流：只用 computeDigest(SHA-256)（登入雜湊也用同一個函式，已證實可用），
// 每個 SHA-256 區塊 32 bytes 切成 16 個 16 位元數字，配合密碼的每個字元（UTF-16 單位）做 XOR。
function pwdViewKeystreamV2_(secret, iv, nUnits) {
  const out = [];
  let counter = 0;
  while (out.length < nUnits) {
    const block = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      secret + ":" + iv + ":" + counter,
      Utilities.Charset.UTF_8
    ).map(function (b) { return b < 0 ? b + 256 : b; });
    for (let i = 0; i + 1 < block.length && out.length < nUnits; i += 2) {
      out.push(block[i] * 256 + block[i + 1]);
    }
    counter++;
  }
  return out;
}

function encryptPwdView_(plain) {
  const secret = getPwdViewSecret_();
  const iv = randomHex_(16);
  const str = String(plain);
  const ks = pwdViewKeystreamV2_(secret, iv, str.length);
  let hex = "";
  for (let i = 0; i < str.length; i++) {
    const v = (str.charCodeAt(i) ^ ks[i]) & 0xffff;
    hex += ("0000" + v.toString(16)).slice(-4);
  }
  return "v2$" + iv + "$" + hex;
}

// 解不開（沒有這份資料、金鑰變了、格式不對）一律回空字串；原因會寫進執行記錄方便查
function decryptPwdView_(stored) {
  try {
    if (typeof stored !== "string" || stored.indexOf("v2$") !== 0) return "";
    const parts = stored.split("$");
    if (parts.length !== 3 || parts[2].length % 4 !== 0) return "";
    const n = parts[2].length / 4;
    const ks = pwdViewKeystreamV2_(getPwdViewSecret_(), parts[1], n);
    let out = "";
    for (let i = 0; i < n; i++) {
      out += String.fromCharCode(parseInt(parts[2].substr(i * 4, 4), 16) ^ ks[i]);
    }
    return out;
  } catch (err) {
    console.error("decryptPwdView_ 失敗：" + err);
    return "";
  }
}

// 在編輯器手動執行：檢查還原用密碼功能是否正常（看「執行記錄」）
function diagnosePwdView() {
  const t = encryptPwdView_("0000");
  const back = decryptPwdView_(t);
  Logger.log("加密結果：" + t);
  Logger.log("解回：" + back + (back === "0000" ? "（正常）" : "（失敗！）"));
  const data = loadData_();
  if (data) {
    let ok = 0, bad = 0, none = 0;
    data.students.forEach(function (s) {
      if (!s.pwdView) none++; else if (decryptPwdView_(s.pwdView)) ok++; else bad++;
    });
    Logger.log("學生：可顯示密碼 " + ok + " 位、解不開 " + bad + " 位（舊格式或金鑰不符）、沒有還原資料 " + none + " 位");
  }
}

function safeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

// 回傳 { ok, legacy }；legacy=true 代表資料庫裡還是明碼（登入成功後會自動升級成雜湊）
function checkPwd_(stored, plain) {
  if (stored == null || stored === "") return { ok: false, legacy: false };
  plain = String(plain == null ? "" : plain);

  if (isHashed_(stored)) {
    const parts = String(stored).split("$");
    if (parts.length !== 3) return { ok: false, legacy: false };
    return { ok: safeEquals_(hashWithSalt_(plain, parts[1]), parts[2]), legacy: false };
  }
  return { ok: safeEquals_(String(stored), plain), legacy: true };
}

function createSession_(role, studentId) {
  const token = randomHex_(48);
  CacheService.getScriptCache().put(
    "tok:" + token,
    JSON.stringify({ role: role, studentId: studentId || null }),
    SESSION_TTL_SEC
  );
  return token;
}

function getSession_(token) {
  if (!token || typeof token !== "string" || token.length < 32 || token.length > 128) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get("tok:" + token);
  if (!raw) return null;
  try {
    cache.put("tok:" + token, raw, SESSION_TTL_SEC); // 有在使用就延長
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function failKey_(role, studentId) {
  return "fail:" + role + ":" + (studentId || "");
}
function isLoginLocked_(key) {
  return Number(CacheService.getScriptCache().get(key) || 0) >= MAX_LOGIN_FAILS;
}
function recordLoginFail_(key) {
  const cache = CacheService.getScriptCache();
  cache.put(key, String(Number(cache.get(key) || 0) + 1), LOGIN_LOCK_SEC);
}
function clearLoginFails_(key) {
  CacheService.getScriptCache().remove(key);
}

function unauthorized_() {
  return jsonResponse_({ status: "unauthorized", message: "登入已過期或沒有權限，請重新登入" });
}

/* ==================================================================
 * 資料讀取 / 依身分裁切
 * ================================================================== */
function loadData_() {
  const json = readStoredJson_();
  if (!json) return null;
  const data = JSON.parse(json);
  validateAppData_(data);
  return data;
}

function defaultData_() {
  return {
    teacherPwd: makePwdHash_(DEFAULT_TEACHER_PWD),
    subjects: ["國語", "數學", "自然", "社會", "英語"],
    grades: ["一年級", "二年級", "三年級", "四年級", "五年級", "六年級"],
    students: [],
    tasks: [],
    contactBook: []
  };
}

// 任何身分都不會拿到 pwd（雜湊）與 pwdView（加密的還原用密碼）
function stripStudent_(s) {
  const o = {};
  Object.keys(s).forEach(function (k) {
    if (k !== "pwd" && k !== "pwdView" && k !== "pwdPlain") o[k] = s[k];
  });
  return o;
}

/* ==================================================================
 * 跑馬燈：上週（週一～週日，台灣時間）所有項目都被老師審核「已完成」的學生，依完成快慢排序
 * 完成時間＝該學生上週項目裡最晚的「學生送出時間／審核通過時間」；沒有時間紀錄的排最後。
 * ================================================================== */
function ymdAdd_(ymd, days) {
  const d = new Date(ymd + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().substring(0, 10);
}

function lastWeekRange_(now) {
  const today = Utilities.formatDate(now, APP_TZ, "yyyy-MM-dd");
  const dow = Number(Utilities.formatDate(now, APP_TZ, "u")); // 1=週一 … 7=週日
  const thisMonday = ymdAdd_(today, -(dow - 1));
  return { start: ymdAdd_(thisMonday, -7), end: ymdAdd_(thisMonday, -1), today: today, dow: dow };
}

function computeChampions_(data, now) {
  const range = lastWeekRange_(now || new Date());
  const list = [];
  (data.students || []).forEach(function (st) {
    const tasks = (data.tasks || []).filter(function (t) {
      return t.studentId === st.id && t.date >= range.start && t.date <= range.end;
    });
    if (tasks.length === 0) return;
    if (!tasks.every(function (t) { return t.teacherStatus === "已完成"; })) return;
    let finishedAt = "";
    tasks.forEach(function (t) {
      const stamp = String(t.submittedAt || t.doneAt || "");
      if (stamp > finishedAt) finishedAt = stamp;
    });
    list.push({ id: st.id, name: st.name, school: st.school, grade: st.grade, avatar: st.avatar || "", finishedAt: finishedAt });
  });
  list.sort(function (a, b) {
    if (!a.finishedAt !== !b.finishedAt) return a.finishedAt ? -1 : 1; // 有時間紀錄的在前
    if (a.finishedAt !== b.finishedAt) return a.finishedAt < b.finishedAt ? -1 : 1; // 早完成的在前
    return String(a.name).localeCompare(String(b.name));
  });
  return { weekStart: range.start, weekEnd: range.end, list: list };
}

/* ==================================================================
 * 聯絡簿自動清除：日期早於「今天 − 10 天」的整筆刪除（台灣時間）
 * ================================================================== */
function contactBookCutoff_(now) {
  return ymdAdd_(Utilities.formatDate(now, APP_TZ, "yyyy-MM-dd"), -CONTACT_BOOK_KEEP_DAYS);
}
function contactBookKeep_(list, now) {
  const cutoff = contactBookCutoff_(now);
  return (Array.isArray(list) ? list : []).filter(function (x) {
    return x && typeof x.date === "string" && x.date >= cutoff;
  });
}
// 在鎖內重新讀取、清除、寫回；拿不到鎖就略過（下次再清，顯示端本來就會過濾）
function purgeContactBookStored_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return 0;
  try {
    const data = loadData_();
    if (!data || !Array.isArray(data.contactBook)) return 0;
    const kept = contactBookKeep_(data.contactBook, new Date());
    const removed = data.contactBook.length - kept.length;
    if (removed > 0) {
      data.contactBook = kept;
      writeStoredJson_(JSON.stringify(data));
    }
    return removed;
  } catch (err) {
    return 0;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function teacherView_(data) {
  const o = {};
  Object.keys(data).forEach(function (k) { if (k !== "teacherPwd") o[k] = data[k]; });
  // 只有老師憑證的回應才會解開顯示密碼；舊資料沒有 pwdView 的會是空字串
  o.students = data.students.map(function (s) {
    const v = stripStudent_(s);
    v.pwdPlain = decryptPwdView_(s.pwdView);
    return v;
  });
  o.contactBook = contactBookKeep_(o.contactBook, new Date());
  o.champions = computeChampions_(data, new Date());
  return o;
}

function studentView_(data, studentId) {
  const me = data.students.filter(function (s) { return s.id === studentId; })[0];
  if (!me) return null;
  return {
    subjects: data.subjects,
    grades: data.grades,
    students: [stripStudent_(me)],
    tasks: data.tasks.filter(function (t) { return t.studentId === studentId; }),
    contactBook: contactBookKeep_(data.contactBook, new Date()),
    champions: computeChampions_(data, new Date())
  };
}

function viewForSession_(data, session) {
  return session.role === "teacher" ? teacherView_(data) : studentView_(data, session.studentId);
}

function parseJsonBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("沒有收到前端 POST 資料");
  }
  try {
    const body = JSON.parse(String(e.postData.contents));
    if (!body || typeof body !== "object") throw new Error("格式不是物件");
    return body;
  } catch (parseError) {
    throw new Error("前端傳送的資料不是有效 JSON：" + parseError.toString());
  }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  // 拿不到鎖不算錯誤：回 busy，讓前端自動退避重試。
  if (!lock.tryLock(20000)) {
    return jsonResponse_({ status: "busy", message: "伺服器忙碌中（鎖定逾時），請稍後重試" });
  }
  try {
    return fn();
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ==================================================================
 * 路由
 * ================================================================== */
function doGet(e) {
  try {
    const data = loadData_() || defaultData_();
    // 公開資訊只有登入選單需要的欄位，不含作業、密碼或其他資料
    return jsonResponse_({
      status: "success",
      public: true,
      version: BACKEND_VERSION,
      students: data.students.map(function (s) {
        return { id: s.id, name: s.name, school: s.school, grade: s.grade, className: s.className || "" };
      }),
      subjects: data.subjects,
      grades: data.grades
    });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : "";

  if (action === "login") return handleLogin_(e);
  if (action === "data") return handleData_(e);
  if (action === "logout") return handleLogout_(e);
  if (action === "save") return handleSave_(e);
  if (action === "patch") return handlePatch_(e);
  if (action === "changePwd") return handleChangePwd_(e);

  return jsonResponse_({ status: "error", message: "未知的 action：" + action });
}

/* ==================================================================
 * 登入 / 讀取 / 登出
 * ================================================================== */
// 在鎖內轉換並寫回；拿不到鎖就略過（下次登入再試）
function lock_tryMigrate_(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    const fresh = loadData_();
    if (!fresh) return false;
    const n = migrateDataPasswords_(fresh);
    if (n > 0) writeStoredJson_(JSON.stringify(fresh));
    // 讓本次回應也用轉換後的資料
    Object.keys(data).forEach(function (k) { delete data[k]; });
    Object.keys(fresh).forEach(function (k) { data[k] = fresh[k]; });
    return true;
  } catch (err) {
    return false;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function handleLogin_(e) {
  let body;
  try { body = parseJsonBody_(e); } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }

  const role = body.role;
  if (role !== "teacher" && role !== "student") {
    return jsonResponse_({ status: "error", message: "role 不正確" });
  }
  const studentId = role === "student" ? String(body.studentId || "").substring(0, 100) : "";
  if (role === "student" && !studentId) {
    return jsonResponse_({ status: "error", message: "缺少 studentId" });
  }
  const plain = String(body.pwd == null ? "" : body.pwd);

  const key = failKey_(role, studentId);
  if (isLoginLocked_(key)) {
    return jsonResponse_({ status: "locked", message: "嘗試次數過多，請 1 分鐘後再試" });
  }

  try {
    const data = loadData_() || defaultData_();

    let stored = null;
    if (role === "teacher") {
      stored = data.teacherPwd ? data.teacherPwd : DEFAULT_TEACHER_PWD;
    } else {
      const st = data.students.filter(function (s) { return s.id === studentId; })[0];
      stored = st ? st.pwd : null;
    }

    const check = checkPwd_(stored, plain);
    if (!check.ok) {
      recordLoginFail_(key);
      return jsonResponse_({ status: "denied", message: "帳號或密碼錯誤" });
    }
    clearLoginFails_(key);

    // 老師登入時，順手把還沒轉換的明碼密碼（例如剛從舊版匯入的資料）轉成雜湊並存下可查看版本
    if (role === "teacher") {
      const needs = data.students.some(function (s) { return s.pwd && !isHashed_(s.pwd); });
      if (needs && lock_tryMigrate_(data)) { /* data 已就地更新 */ }
      purgeContactBookStored_();
    }

    if (check.legacy) upgradePwdHash_(role, studentId, plain);

    const session = { role: role, studentId: studentId || null };
    return jsonResponse_({
      status: "success",
      token: createSession_(role, studentId),
      role: role,
      studentId: studentId || null,
      data: viewForSession_(data, session)
    });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

// 登入成功且資料庫還是明碼時，順手升級成雜湊；失敗不影響登入
function upgradePwdHash_(role, studentId, plain) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const data = loadData_();
    if (!data) return;

    let changed = false;
    if (role === "teacher") {
      if (!isHashed_(data.teacherPwd)) {
        data.teacherPwd = makePwdHash_(plain);
        changed = true;
      }
    } else {
      const st = data.students.filter(function (s) { return s.id === studentId; })[0];
      if (st && !isHashed_(st.pwd) && String(st.pwd) === plain) {
        st.pwd = makePwdHash_(plain);
        st.pwdView = encryptPwdView_(plain);
        changed = true;
      }
    }
    if (changed) writeStoredJson_(JSON.stringify(data));
  } catch (err) {
    // 忽略：下次登入會再嘗試
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function handleData_(e) {
  let body;
  try { body = parseJsonBody_(e); } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
  const session = getSession_(body.token);
  if (!session) return unauthorized_();

  try {
    if (session.role === "teacher") purgeContactBookStored_();
    const data = loadData_() || (session.role === "teacher" ? defaultData_() : null);
    const view = data ? viewForSession_(data, session) : null;
    if (!view) return unauthorized_(); // 例如該學生已被老師刪除
    return jsonResponse_({ status: "success", data: view });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

function handleLogout_(e) {
  try {
    const body = parseJsonBody_(e);
    if (body.token && typeof body.token === "string") {
      CacheService.getScriptCache().remove("tok:" + body.token);
    }
  } catch (err) { /* 忽略 */ }
  return jsonResponse_({ status: "success" });
}

/* ==================================================================
 * 整份儲存（老師專用）
 *   學生資料裡的 pwd、以及 teacherPwd「有填才更新」（填的是新的明碼，由後端雜湊）；
 *   沒填就沿用資料庫原本的雜湊。網頁端因此永遠不需要持有任何密碼。
 * ================================================================== */
function mergeIncoming_(stored, incoming) {
  validateAppData_(incoming);

  const storedById = {};
  stored.students.forEach(function (s) { storedById[s.id] = s; });

  const out = {};
  Object.keys(incoming).forEach(function (k) {
    if (k !== "teacherPwd" && k !== "students" && k !== "champions") out[k] = incoming[k]; // champions 是即時算出來的，不存
  });

  out.students = incoming.students.map(function (s) {
    const copy = {};
    // 前端送來的 pwd / pwdPlain / pwdView 一律不信任，只認下面「有填新密碼」或沿用資料庫舊值
    Object.keys(s).forEach(function (k) {
      if (k !== "pwd" && k !== "pwdPlain" && k !== "pwdView") copy[k] = s[k];
    });

    const old = storedById[s.id];
    const plain = (typeof s.pwd === "string" || typeof s.pwd === "number") ? String(s.pwd).trim() : "";

    if (plain) {
      if (isHashed_(plain) || plain.length > 64) throw new Error("學生密碼格式不正確");
      copy.pwd = makePwdHash_(plain);
      copy.pwdView = encryptPwdView_(plain);
    } else if (old && old.pwd) {
      copy.pwd = old.pwd;
      if (old.pwdView) copy.pwdView = old.pwdView;
    } else {
      throw new Error("學生「" + (s.name || s.id) + "」缺少密碼");
    }
    return copy;
  });

  const tp = typeof incoming.teacherPwd === "string" ? incoming.teacherPwd.trim() : "";
  if (tp) {
    if (isHashed_(tp) || tp.length > 64) throw new Error("老師密碼格式不正確");
    out.teacherPwd = makePwdHash_(tp);
  } else {
    out.teacherPwd = stored.teacherPwd || makePwdHash_(DEFAULT_TEACHER_PWD);
  }

  return out;
}

function handleSave_(e) {
  let body, session;
  try {
    body = parseJsonBody_(e);
    session = getSession_(body.token);
    if (!session || session.role !== "teacher") return unauthorized_();
    validateAppData_(body.data); // 先在鎖外驗證，縮短持鎖時間
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }

  return withLock_(function () {
    const stored = loadData_() || defaultData_();
    const merged = mergeIncoming_(stored, body.data);
    writeStoredJson_(JSON.stringify(merged));

    return jsonResponse_({
      status: "success",
      message: "資料儲存成功",
      studentCount: merged.students.length,
      taskCount: merged.tasks.length,
      savedAt: new Date().toISOString()
    });
  });
}

/* ==================================================================
 * 學生修改自己的密碼
 * ================================================================== */
function handleChangePwd_(e) {
  let body, session;
  try {
    body = parseJsonBody_(e);
    session = getSession_(body.token);
    if (!session || session.role !== "student") return unauthorized_();
    if (!/^\d{4}$/.test(String(body.newPwd == null ? "" : body.newPwd))) {
      throw new Error("新密碼必須是 4 位數字");
    }
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }

  return withLock_(function () {
    const data = loadData_();
    const st = data && data.students.filter(function (s) { return s.id === session.studentId; })[0];
    if (!st) return unauthorized_();

    st.pwd = makePwdHash_(String(body.newPwd));
    st.pwdView = encryptPwdView_(String(body.newPwd)); // 讓老師端看到的是學生改過後的新密碼
    writeStoredJson_(JSON.stringify(data));
    return jsonResponse_({ status: "success", message: "密碼已更新" });
  });
}

/* ==================================================================
 * 逐筆更新（patch）
 *
 *   { "token": "...", "ops": [
 *       { "op": "studentSubmit",   "studentId": "s1", "ids": ["t1","t2"] },
 *       { "op": "studentWithdraw", "studentId": "s1", "ids": ["t3"] },
 *       { "op": "teacherReview",   "status": "已完成", "ids": ["t1","t2"] }
 *   ] }
 *
 *   學生憑證：studentId 一律以憑證為準（傳別人的也沒用），且不能做 teacherReview。
 *   老師憑證：可代任何學生操作，也可審核。
 *
 *   studentSubmit   ：studentChecked=true,  pendingReview=true
 *   studentWithdraw ：studentChecked=false, pendingReview=false
 *     兩者只處理「屬於該學生」且「老師尚未審核為已完成」的項目，其餘列入 skipped。
 *   teacherReview   ：teacherStatus=status, pendingReview=false,
 *                     studentChecked=(status==="已完成")
 *
 * 回傳 { status, changed, skipped, tasks }；學生憑證的 tasks 只含自己的項目。
 * ================================================================== */
const REVIEW_STATUSES = ["未完成", "待訂正", "已完成"];
const MAX_PATCH_OPS = 50;
const MAX_PATCH_IDS = 500;

function validatePatchOps_(rawOps, session) {
  if (!Array.isArray(rawOps) || rawOps.length === 0) throw new Error("缺少 ops 陣列");
  if (rawOps.length > MAX_PATCH_OPS) throw new Error("一次最多 " + MAX_PATCH_OPS + " 個操作");

  return rawOps.map(function (op) {
    if (!op || typeof op !== "object") throw new Error("ops 內含無效的操作");
    if (!Array.isArray(op.ids) || op.ids.length === 0 || op.ids.length > MAX_PATCH_IDS) {
      throw new Error("ids 必須是 1~" + MAX_PATCH_IDS + " 個項目編號");
    }

    if (op.op === "studentSubmit" || op.op === "studentWithdraw") {
      const sid = session.role === "student" ? session.studentId : op.studentId;
      if (!sid) throw new Error(op.op + " 缺少 studentId");
      return { op: op.op, studentId: String(sid), ids: op.ids };
    }

    if (op.op === "teacherReview") {
      if (session.role !== "teacher") throw new Error("NOT_ALLOWED");
      if (REVIEW_STATUSES.indexOf(op.status) < 0) {
        throw new Error("teacherReview 的 status 不正確：" + op.status);
      }
      return { op: op.op, status: op.status, ids: op.ids };
    }

    throw new Error("未知的 op：" + op.op);
  });
}

function setTaskFields_(task, fields) {
  let changed = false;
  Object.keys(fields).forEach(function (key) {
    if (task[key] !== fields[key]) {
      task[key] = fields[key];
      changed = true;
    }
  });
  return changed;
}

function applyPatchOps_(data, ops) {
  const byId = {};
  data.tasks.forEach(function (t) { byId[String(t.id)] = t; });

  let changed = 0;
  const skipped = [];

  ops.forEach(function (op) {
    op.ids.forEach(function (rawId) {
      const id = String(rawId);
      const task = byId[id];

      if (!task) { skipped.push(id); return; }

      if (op.op === "teacherReview") {
        const wasDone = task.teacherStatus === "已完成";
        const didChange = setTaskFields_(task, {
          teacherStatus: op.status,
          pendingReview: false,
          studentChecked: op.status === "已完成"
        });
        if (didChange) changed++;
        // 記下審核通過的時間（跑馬燈排序用）；改成其他狀態就清掉
        if (op.status === "已完成" && !wasDone) task.doneAt = new Date().toISOString();
        else if (op.status !== "已完成") delete task.doneAt;
        return;
      }

      if (task.studentId !== op.studentId || task.teacherStatus === "已完成") {
        skipped.push(id);
        return;
      }

      const submit = (op.op === "studentSubmit");
      const didChange = setTaskFields_(task, { studentChecked: submit, pendingReview: submit });
      if (didChange) changed++;
      // 記下學生送出的時間（跑馬燈「完成快慢」排序用）；撤回就清掉
      if (submit && didChange) task.submittedAt = new Date().toISOString();
      else if (!submit) delete task.submittedAt;
    });
  });

  return { changed: changed, skipped: skipped };
}

function handlePatch_(e) {
  let session, ops;
  try {
    const body = parseJsonBody_(e);
    session = getSession_(body.token);
    if (!session) return unauthorized_();
    ops = validatePatchOps_(body.ops, session);
  } catch (err) {
    if (String(err.message) === "NOT_ALLOWED") return unauthorized_();
    return jsonResponse_({ status: "error", message: err.toString() });
  }

  return withLock_(function () {
    const data = loadData_();
    if (!data) return jsonResponse_({ status: "error", message: "雲端尚無資料，無法逐筆更新" });

    const result = applyPatchOps_(data, ops);
    if (result.changed > 0) writeStoredJson_(JSON.stringify(data));

    const tasks = session.role === "teacher"
      ? data.tasks
      : data.tasks.filter(function (t) { return t.studentId === session.studentId; });

    return jsonResponse_({
      status: "success",
      message: "逐筆更新成功",
      changed: result.changed,
      skipped: result.skipped,
      tasks: tasks,
      savedAt: new Date().toISOString()
    });
  });
}

/* ==================================================================
 * 維護用：在 Apps Script 編輯器手動執行
 * ================================================================== */

// 把資料裡還是明碼的密碼改成雜湊，同時存下老師端可查看的還原用密碼。回傳轉換的組數。
function migrateDataPasswords_(data) {
  let n = 0;
  if (!data.teacherPwd) {
    data.teacherPwd = makePwdHash_(DEFAULT_TEACHER_PWD); n++;
  } else if (!isHashed_(data.teacherPwd)) {
    data.teacherPwd = makePwdHash_(String(data.teacherPwd)); n++;
  }
  data.students.forEach(function (s) {
    if (s.pwd && !isHashed_(s.pwd)) {
      s.pwdView = encryptPwdView_(String(s.pwd)); // 明碼還在的這一刻順便存起還原用密碼
      s.pwd = makePwdHash_(String(s.pwd));
      n++;
    }
  });
  return n;
}

// 第一次部署新版後執行一次：把現有明碼密碼全部改成雜湊
function migratePasswords() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = loadData_();
    if (!data) { Logger.log("尚無資料，不需要轉換"); return; }
    const n = migrateDataPasswords_(data);
    if (n > 0) writeStoredJson_(JSON.stringify(data));
    Logger.log("已把 " + n + " 組密碼轉成雜湊");
  } finally {
    lock.releaseLock();
  }
}

// 從舊版部署把全部資料搬進這份新試算表（在編輯器手動執行一次）
// 用法：把下面 OLD_EXEC_URL 換成「舊版」網頁應用程式的 /exec 網址，選 importFromOldDeployment 按執行。
// 新試算表若已經有學生資料，會中止以免覆蓋；確定要覆蓋請把 OVERWRITE_EXISTING 改成 true。
function importFromOldDeployment() {
  const OLD_EXEC_URL = "請貼上舊版的 /exec 網址";
  const OVERWRITE_EXISTING = false;

  if (OLD_EXEC_URL.indexOf("https://script.google.com/") !== 0) {
    throw new Error("請先把 OLD_EXEC_URL 換成舊版的 /exec 網址");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const current = loadData_();
    if (current && current.students && current.students.length > 0 && !OVERWRITE_EXISTING) {
      throw new Error("新資料庫已有 " + current.students.length + " 位學生，為避免覆蓋已中止。");
    }
    const res = UrlFetchApp.fetch(OLD_EXEC_URL, { followRedirects: true, muteHttpExceptions: true });
    const old = JSON.parse(res.getContentText());
    if (old && old.status === "error") throw new Error("舊版回傳錯誤：" + old.message);
    validateAppData_(old);
    if (old.public) throw new Error("取得的是新版公開資料，請確認貼的是『舊版』網址");

    if (!Array.isArray(old.contactBook)) old.contactBook = [];
    writeStoredJson_(JSON.stringify(old));
    Logger.log("已匯入：學生 " + old.students.length + " 位、項目 " + old.tasks.length + " 筆");
  } finally {
    lock.releaseLock();
  }
  migratePasswords(); // 明碼 → 雜湊，並存下老師端可查看的密碼
}

// 把所有學生的登入密碼一次設成 0000（同時存下老師端可查看的版本）。在編輯器手動執行一次。
function resetAllStudentPasswordsTo0000() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = loadData_();
    if (!data) { Logger.log("尚無資料"); return; }
    data.students.forEach(function (s) {
      s.pwd = makePwdHash_("0000");
      s.pwdView = encryptPwdView_("0000");
    });
    writeStoredJson_(JSON.stringify(data));
    Logger.log("已把 " + data.students.length + " 位學生的密碼都設為 0000");
  } finally {
    lock.releaseLock();
  }
}

// 忘記老師密碼時執行：把老師密碼重設為 1234（登入後請立刻修改）
function resetTeacherPasswordTo1234() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = loadData_() || defaultData_();
    data.teacherPwd = makePwdHash_(DEFAULT_TEACHER_PWD);
    writeStoredJson_(JSON.stringify(data));
    Logger.log("老師密碼已重設為 1234");
  } finally {
    lock.releaseLock();
  }
}
