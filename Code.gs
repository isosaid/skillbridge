// SkillBridge — сервер на Google Apps Script. Все данные хранятся в этой Google-таблице.
//
// Установка (подробно — в README):
//   1. Создайте Google-таблицу → Расширения → Apps Script → вставьте этот файл целиком → Сохранить.
//   2. Выберите функцию setup → Выполнить → разрешите доступ. Появятся листы, в том числе «Настройки».
//   3. На листе «Настройки» впишите пароль администратора (и, по желанию, токен бота и ваш chat id).
//   4. Начать развёртывание → Новое развёртывание → Веб-приложение:
//      «Выполнять как: Я», «У кого есть доступ: Все» → Развернуть → скопируйте ссылку /exec.
//   5. Вставьте ссылку в файл config.js на сайте.
//
// Курсы (COURSES, PROGRAMS) в конце файла собираются из src/courses.txt командой python3 build.py.

var CATEGORIES = ["dev", "marketing", "design", "data", "lang", "cloud", "video", "text"];
var TRACK_RU = { it: "IT-разработка", marketing: "Цифровой маркетинг", design: "UX/UI-дизайн", data: "Данные и аналитика", english: "Профессиональный английский", cloud: "Cloud и DevOps" };
var STATUSES = ["new", "accepted", "declined", "done", "cancelled"];
var STATUS_RU = { new: "новый", accepted: "принят", declined: "отклонён", done: "выполнен", cancelled: "отменён" };
var PASS_SCORE = 0.7;
var EXAM_ATTEMPTS_PER_DAY = 5;
var MAX_FAILS = 8;               // неудачных входов за 15 минут на один логин
var HASH_ROUNDS = 300;
var SESSION_DAYS = 30;

// Листы таблицы и их столбцы. Первая строка листа — названия столбцов, их не меняйте.
var TABLES = {
  "Пользователи": ["id", "login", "name", "salt", "hash", "created_at"],
  "Анкеты": ["user_id", "name", "title", "category", "skills", "price", "city", "portfolio", "status", "updated_at"],
  "Заказы": ["id", "client_id", "specialist_id", "task", "budget", "deadline", "contact", "status", "created_at"],
  "Заявки": ["id", "kind", "name", "contact", "track", "level", "english", "city", "company", "message", "status", "created_at"],
  "Прогресс": ["user_id", "course", "lesson", "score", "passed", "updated_at"],
  "Экзамены": ["id", "user_id", "course", "score", "passed", "created_at"],
  "Сертификаты": ["code", "user_id", "kind", "ref", "title", "full_name", "hours", "score", "revoked", "issued_at"],
  "Входы": ["key", "at"]
};
var SETTINGS = [
  ["ADMIN_LOGIN", "admin", "Логин администратора для входа на сайт"],
  ["ADMIN_PASSWORD", "", "Впишите пароль администратора (от 10 символов). После первого запроса сайта он сохранится в зашифрованном виде, а ячейка очистится."],
  ["BOT_TOKEN", "", "Необязательно: токен Telegram-бота от @BotFather — для уведомлений о заявках"],
  ["TELEGRAM_CHAT_ID", "", "Необязательно: ваш Telegram ID от @userinfobot — сюда придут уведомления (сначала нажмите Start у своего бота)"],
  ["SITE_URL", "https://isosaid.github.io/skillbridge/", "Адрес сайта — для ссылок в уведомлениях"]
];

// ---------- таблица ----------
function book() {
  var id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  props.setProperty("SHEET_ID", ss.getId());
  if (!props.getProperty("SECRET")) props.setProperty("SECRET", Utilities.getUuid() + Utilities.getUuid());
  Object.keys(TABLES).forEach(function (name) { sheet(name); });
  var st = ss.getSheetByName("Настройки");
  if (!st) {
    st = ss.insertSheet("Настройки", 0);
    st.getRange(1, 1, 1, 3).setValues([["Параметр", "Значение", "Пояснение"]]).setFontWeight("bold");
    st.getRange(2, 1, SETTINGS.length, 3).setNumberFormat("@").setValues(SETTINGS);
    st.setFrozenRows(1);
    st.setColumnWidth(1, 170); st.setColumnWidth(2, 320); st.setColumnWidth(3, 560);
  }
  var first = ss.getSheetByName("Лист1") || ss.getSheetByName("Sheet1");
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  return "Готово: листы созданы. Заполните лист «Настройки» и разверните веб-приложение.";
}
function sheet(name) {
  var ss = book(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var cols = TABLES[name];
    sh.getRange(1, 1, 1000, cols.length).setNumberFormat("@");   // всё храним как текст
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}
var _cache = {};
function rows(name) {
  if (_cache[name]) return _cache[name];
  var sh = sheet(name), cols = TABLES[name], n = sh.getLastRow() - 1, out = [];
  if (n > 0) {
    var vals = sh.getRange(2, 1, n, cols.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (vals[i].join("") === "") continue;
      var o = { _row: i + 2 };
      for (var j = 0; j < cols.length; j++) o[cols[j]] = vals[i][j] === "" ? "" : String(vals[i][j]);
      out.push(o);
    }
  }
  return (_cache[name] = out);
}
function cell(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "1" : "0";
  v = String(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;   // чтобы текст пользователя не стал формулой
}
function insert(name, obj) {
  var sh = sheet(name), cols = TABLES[name];
  sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length).setNumberFormat("@")
    .setValues([cols.map(function (c) { return cell(obj[c]); })]);
  delete _cache[name];
}
function update(name, row, patch) {
  var sh = sheet(name), cols = TABLES[name];
  var cur = sh.getRange(row._row, 1, 1, cols.length).getValues()[0];
  cols.forEach(function (c, j) { if (c in patch) cur[j] = cell(patch[c]); else cur[j] = cell(cur[j]); });
  sh.getRange(row._row, 1, 1, cols.length).setValues([cur]);
  delete _cache[name];
}
function find(name, pred) { var r = rows(name); for (var i = 0; i < r.length; i++) if (pred(r[i])) return r[i]; return null; }
function filter(name, pred) { return rows(name).filter(pred); }
function nextId(name) { var m = 0; rows(name).forEach(function (r) { m = Math.max(m, Number(r.id) || 0); }); return m + 1; }
function now() { return new Date().toISOString().replace("T", " ").slice(0, 19); }

function settings() {
  if (_cache._settings) return _cache._settings;
  var s = {}, sh = book().getSheetByName("Настройки");
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r, i) { s[String(r[0]).trim()] = { v: String(r[1]).trim(), row: i + 2 }; });
  }
  var get = function (k, d) { return s[k] && s[k].v ? s[k].v : (d || ""); };
  var out = {
    adminLogin: get("ADMIN_LOGIN", "admin").toLowerCase(), botToken: get("BOT_TOKEN"), chatId: get("TELEGRAM_CHAT_ID"),
    site: get("SITE_URL", "").replace(/\/?$/, "/"), _raw: s, _sheet: sh
  };
  return (_cache._settings = out);
}
// Пароль администратора из листа «Настройки» превращается в учётную запись (хэш), ячейка очищается.
function syncAdmin() {
  var s = settings(), p = s._raw.ADMIN_PASSWORD;
  if (!p || !p.v || p.v.indexOf("✓") === 0) return;
  if (p.v.length < 10) return;
  var u = find("Пользователи", function (r) { return r.login === s.adminLogin; });
  var h = hashPassword(p.v);
  if (u) update("Пользователи", u, { salt: h.salt, hash: h.hash });
  else insert("Пользователи", { id: nextId("Пользователи"), login: s.adminLogin, name: "Администратор", salt: h.salt, hash: h.hash, created_at: now() });
  s._sheet.getRange(p.row, 2).setValue("✓ пароль сохранён " + now().slice(0, 10) + " (чтобы сменить — впишите новый)");
}

// ---------- криптография ----------
function secret() {
  var props = PropertiesService.getScriptProperties(), s = props.getProperty("SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty("SECRET", s); }
  return s;
}
function bytes(s) { return Utilities.newBlob(String(s)).getBytes(); }
function b64(b) { return Utilities.base64EncodeWebSafe(b).replace(/=+$/, ""); }
function hmac(msg, key) { return b64(Utilities.computeHmacSha256Signature(bytes(msg), bytes(key))); }
function hashPassword(password, salt) {
  salt = salt || b64(bytes(Utilities.getUuid() + Utilities.getUuid())).slice(0, 32);
  var key = bytes(secret() + ":" + salt), h = Utilities.computeHmacSha256Signature(bytes(password), key);
  for (var i = 0; i < HASH_ROUNDS; i++) h = Utilities.computeHmacSha256Signature(h, key);
  return { salt: salt, hash: b64(h) };
}
function safeEqual(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length !== b.length) return false;
  var r = 0; for (var i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function makeToken(uid) {
  var body = uid + "." + (Math.floor(Date.now() / 1000) + SESSION_DAYS * 86400);
  return body + "." + hmac(body, "session:" + secret());
}
function readToken(t) {
  var p = String(t || "").split(".");
  if (p.length !== 3 || Number(p[1]) < Date.now() / 1000) return null;
  if (!safeEqual(hmac(p[0] + "." + p[1], "session:" + secret()), p[2])) return null;
  var u = find("Пользователи", function (r) { return r.id === p[0]; });
  return u ? Number(u.id) : null;
}
function certCode() {
  var abc = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789", b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now()), s = "";
  for (var i = 0; i < 8; i++) s += abc[(b[i] + 256) % abc.length];
  return "SB-" + new Date().getUTCFullYear() + "-" + s;
}

// ---------- помощники ----------
function ok(data) { return { status: 200, data: data }; }
function fail(msg, status) { return { status: status || 400, data: { error: msg } }; }
function clip(v, n) { return String(v == null ? "" : v).trim().slice(0, n); }
function isAdmin(uid) {
  if (!uid) return false;
  var u = find("Пользователи", function (r) { return Number(r.id) === uid; });
  return !!u && u.login === settings().adminLogin;
}
function userById(id) { return find("Пользователи", function (r) { return Number(r.id) === Number(id); }); }
function publicUser(u) { return u ? { id: Number(u.id), first_name: u.name, last_name: "", username: null, photo_url: "", login: u.login } : null; }
function notify(text) {
  var s = settings();
  if (!s.botToken || !s.chatId) return;
  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + s.botToken + "/sendMessage", {
      method: "post", contentType: "application/json", muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: s.chatId, text: text, disable_web_page_preview: true })
    });
  } catch (e) { /* уведомление не обязательно */ }
}
function siteLink(path) { return settings().site + (path ? "#" + path : ""); }
function num(v) { return v === "" || v == null ? null : Number(v); }
function profileOut(p) {
  return p ? { user_id: Number(p.user_id), name: p.name, title: p.title, category: p.category, skills: p.skills, price: p.price,
    city: p.city, portfolio: p.portfolio, status: p.status, updated_at: p.updated_at } : null;
}
function orderOut(o) {
  var spec = o.specialist_id ? find("Анкеты", function (p) { return p.user_id === o.specialist_id; }) : null;
  var cl = userById(o.client_id);
  return { id: Number(o.id), client_id: Number(o.client_id), specialist_id: num(o.specialist_id), task: o.task, budget: o.budget,
    deadline: o.deadline, contact: o.contact, status: o.status, created_at: o.created_at,
    specialist_name: spec ? spec.name : null, specialist_title: spec ? spec.title : null,
    client_first: cl ? cl.name : "", client_last: "", client_username: null };
}
function byIdDesc(a, b) { return Number(b.id) - Number(a.id); }

// ---------- обучение ----------
function courseBySlug(slug) { for (var i = 0; i < COURSES.length; i++) if (COURSES[i].slug === slug) return COURSES[i]; return null; }
function stripQuiz(quiz) { return quiz.map(function (q, i) { return { i: i, q: q.q, options: q.options }; }); }
function courseCard(c) {
  return { slug: c.slug, title: c.title, title_en: c.title_en, track: c.track, level: c.level, hours: c.hours,
    summary: c.summary, lessons: c.lessons.length, exam: c.exam.length };
}
function grade(quiz, answers) {
  var results = quiz.map(function (q, i) { return answers != null && answers[i] != null && Number(answers[i]) === q.correct; });
  var right = results.filter(Boolean).length;
  return { results: results, score: quiz.length ? right / quiz.length : 1, right: right, total: quiz.length };
}
function passedLessons(uid, slug) {
  return filter("Прогресс", function (r) { return Number(r.user_id) === uid && r.course === slug && r.passed === "1"; })
    .map(function (r) { return Number(r.lesson); });
}
function myCounts(uid) {
  var my = {};
  filter("Прогресс", function (r) { return Number(r.user_id) === uid && r.passed === "1"; }).forEach(function (r) { my[r.course] = (my[r.course] || 0) + 1; });
  return my;
}
function liveCert(uid, kind, ref) {
  return find("Сертификаты", function (c) { return Number(c.user_id) === uid && c.kind === kind && c.ref === ref && c.revoked !== "1"; });
}
function issueProgramCerts(uid, fullName) {
  var issued = [];
  PROGRAMS.forEach(function (pr) {
    if (liveCert(uid, "program", pr.slug)) return;
    var have = pr.courses.map(function (s) { return liveCert(uid, "course", s); });
    if (have.some(function (x) { return !x; })) return;
    var hours = pr.courses.reduce(function (n, s) { var c = courseBySlug(s); return n + (c ? c.hours : 0); }, 0);
    var score = Math.round(have.reduce(function (n, c) { return n + Number(c.score); }, 0) / have.length);
    var code = certCode();
    insert("Сертификаты", { code: code, user_id: uid, kind: "program", ref: pr.slug, title: pr.title, full_name: fullName, hours: hours, score: score, revoked: 0, issued_at: now() });
    issued.push(code);
  });
  return issued;
}

function learningApi(p, method, uid, body) {
  var m;
  if (p === "/api/courses" && method === "GET") return ok({ courses: COURSES.map(courseCard), programs: PROGRAMS, my: uid ? myCounts(uid) : {} });

  if ((m = p.match(/^\/api\/courses\/(\w+)$/)) && method === "GET") {
    var c = courseBySlug(m[1]);
    if (!c) return fail("not found", 404);
    var my = null;
    if (uid) { var ct = liveCert(uid, "course", c.slug); my = { passed: passedLessons(uid, c.slug), certificate: ct ? ct.code : null }; }
    var card = courseCard(c); card.about = c.about;
    card.outline = c.lessons.map(function (l, i) { return { n: i + 1, title: l.title, quiz: l.quiz.length }; });
    return ok({ course: card, programs: PROGRAMS.filter(function (pr) { return pr.courses.indexOf(c.slug) >= 0; }), my: my });
  }

  if ((m = p.match(/^\/api\/courses\/(\w+)\/lessons\/(\d+)$/))) {
    if (!uid) return fail("login required", 401);
    var c2 = courseBySlug(m[1]), n = Number(m[2]), l = c2 && c2.lessons[n - 1];
    if (!l) return fail("not found", 404);
    var row = find("Прогресс", function (r) { return Number(r.user_id) === uid && r.course === c2.slug && Number(r.lesson) === n; });
    if (method === "GET") {
      return ok({ course: courseCard(c2), lesson: { n: n, title: l.title, body: l.body, quiz: stripQuiz(l.quiz) },
        passed: !!(row && row.passed === "1"), total: c2.lessons.length,
        outline: c2.lessons.map(function (x, i) { return { n: i + 1, title: x.title }; }), done: passedLessons(uid, c2.slug) });
    }
    var g = grade(l.quiz, body.answers), passed = g.score >= PASS_SCORE;
    if (row) update("Прогресс", row, { score: Math.max(Number(row.score) || 0, g.score), passed: row.passed === "1" || passed, updated_at: now() });
    else insert("Прогресс", { user_id: uid, course: c2.slug, lesson: n, score: g.score, passed: passed, updated_at: now() });
    g.passed = passed;
    return ok(g);
  }

  if ((m = p.match(/^\/api\/courses\/(\w+)\/exam$/))) {
    if (!uid) return fail("login required", 401);
    var c3 = courseBySlug(m[1]);
    if (!c3) return fail("not found", 404);
    var done = passedLessons(uid, c3.slug);
    var ready = c3.lessons.every(function (_, i) { return done.indexOf(i + 1) >= 0; });
    var since = new Date(Date.now() - 86400000).toISOString().replace("T", " ").slice(0, 19);
    var att = filter("Экзамены", function (r) { return Number(r.user_id) === uid && r.course === c3.slug && r.created_at > since; }).length;
    var left = Math.max(0, EXAM_ATTEMPTS_PER_DAY - att);
    var cert = liveCert(uid, "course", c3.slug);
    if (method === "GET") return ok({ course: courseCard(c3), ready: ready, attemptsLeft: left, certificate: cert ? cert.code : null, questions: ready ? stripQuiz(c3.exam) : [] });
    if (!ready) return fail("finish all lessons first", 403);
    if (!left) return fail("no attempts left today", 429);
    var fullName = clip(body.full_name, 100);
    if (fullName.length < 3) return fail("full name required");
    var ge = grade(c3.exam, body.answers), ok3 = ge.score >= PASS_SCORE;
    insert("Экзамены", { id: nextId("Экзамены"), user_id: uid, course: c3.slug, score: ge.score, passed: ok3, created_at: now() });
    var code = cert ? cert.code : null, programs = [];
    if (ok3 && !code) {
      code = certCode();
      insert("Сертификаты", { code: code, user_id: uid, kind: "course", ref: c3.slug, title: c3.title, full_name: fullName, hours: c3.hours, score: Math.round(ge.score * 100), revoked: 0, issued_at: now() });
      programs = issueProgramCerts(uid, fullName);
      notify("Выдан сертификат " + code + "\n" + fullName + " — «" + c3.title + "», " + Math.round(ge.score * 100) + "%" +
        (programs.length ? "\nИ удостоверение: " + programs.join(", ") : "") + "\n" + siteLink("/cert/" + code));
    }
    return ok({ score: ge.score, right: ge.right, total: ge.total, passed: ok3, certificate: ok3 ? code : null, programs: programs, attemptsLeft: left - 1 });
  }

  if ((m = p.match(/^\/api\/certs\/([A-Z0-9-]{6,30})$/)) && method === "GET") {
    var ce = find("Сертификаты", function (r) { return r.code === m[1]; });
    if (!ce) return fail("not found", 404);
    var out = { code: ce.code, kind: ce.kind, ref: ce.ref, title: ce.title, full_name: ce.full_name, hours: Number(ce.hours),
      score: Number(ce.score), revoked: ce.revoked === "1" ? 1 : 0, issued_at: ce.issued_at };
    if (ce.kind === "program") {
      var pr = PROGRAMS.filter(function (x) { return x.slug === ce.ref; })[0];
      out.courses = pr ? pr.courses.map(function (s) { var c = courseBySlug(s); return c ? c.title : null; }).filter(Boolean) : [];
    }
    return ok({ certificate: out });
  }

  if (p === "/api/learning" && method === "GET") {
    if (!uid) return fail("login required", 401);
    var counts = myCounts(uid), progress = [];
    Object.keys(counts).forEach(function (s) { var c = courseBySlug(s); if (c) { var k = courseCard(c); k.passed = counts[s]; progress.push(k); } });
    var certs = filter("Сертификаты", function (r) { return Number(r.user_id) === uid && r.revoked !== "1"; })
      .map(function (r) { return { code: r.code, kind: r.kind, ref: r.ref, title: r.title, score: Number(r.score), hours: Number(r.hours), issued_at: r.issued_at }; })
      .sort(function (a, b) { return a.issued_at < b.issued_at ? 1 : -1; });
    return ok({ progress: progress, certificates: certs });
  }
  return null;
}

// ---------- вход ----------
function tooManyFails(keys) {
  var since = Math.floor(Date.now() / 1000) - 900;
  return keys.some(function (k) { return filter("Входы", function (r) { return r.key === k && Number(r.at) >= since; }).length >= MAX_FAILS; });
}
function addFail(keys) {
  var t = Math.floor(Date.now() / 1000);
  keys.forEach(function (k) { insert("Входы", { key: k, at: t }); });
  var sh = sheet("Входы");
  if (sh.getLastRow() > 2000) {          // чистим старые записи
    var keep = rows("Входы").filter(function (r) { return Number(r.at) >= t - 900; });
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 2).setValues(keep.map(function (r) { return [r.key, r.at]; }));
    delete _cache["Входы"];
  }
}
function authApi(p, uid, body) {
  var s = settings();
  if (p === "/api/auth/register") {
    var login = clip(body.login, 32).toLowerCase(), password = String(body.password || ""), name = clip(body.name, 80);
    if (!/^[a-z0-9_.-]{3,32}$/.test(login)) return fail("bad login");
    if (password.length < 8 || password.length > 200) return fail("weak password");
    if (name.length < 2) return fail("name required");
    if (login === s.adminLogin || find("Пользователи", function (r) { return r.login === login; })) return fail("login taken", 409);
    if (tooManyFails(["register"])) return fail("too many attempts", 429);
    var recent = filter("Пользователи", function (r) { return r.created_at > new Date(Date.now() - 600000).toISOString().replace("T", " ").slice(0, 19); });
    if (recent.length >= 30) return fail("too many attempts", 429);   // защита от массовой регистрации ботами
    var id = nextId("Пользователи"), h = hashPassword(password);
    insert("Пользователи", { id: id, login: login, name: name, salt: h.salt, hash: h.hash, created_at: now() });
    return ok({ ok: true, token: makeToken(id) });
  }
  if (p === "/api/auth/login") {
    var lg = clip(body.login, 32).toLowerCase(), pw = String(body.password || "");
    var keys = ["login:" + lg];
    if (tooManyFails(keys)) return fail("too many attempts", 429);
    var u = find("Пользователи", function (r) { return r.login === lg; });
    var hh = hashPassword(pw, u ? u.salt : null);
    if (!u || !safeEqual(hh.hash, u.hash)) { addFail(keys); return fail("wrong login or password", 401); }
    return ok({ ok: true, token: makeToken(Number(u.id)) });
  }
  if (p === "/api/auth/password") {
    if (!uid) return fail("login required", 401);
    var me = userById(uid);
    if (!me) return fail("login required", 401);
    if (tooManyFails(["pw:" + uid])) return fail("too many attempts", 429);
    if (!safeEqual(hashPassword(String(body.old_password || ""), me.salt).hash, me.hash)) { addFail(["pw:" + uid]); return fail("wrong password", 403); }
    var np = String(body.new_password || "");
    if (np.length < 8 || np.length > 200) return fail("weak password");
    var nh = hashPassword(np);
    update("Пользователи", me, { salt: nh.salt, hash: nh.hash });
    return ok({ ok: true });
  }
  return null;
}

// ---------- API ----------
function api(p, method, uid, body) {
  if (p.indexOf("/api/auth/") === 0 && method === "POST") { var a = authApi(p, uid, body); if (a) return a; }
  var learned = learningApi(p, method, uid, body);
  if (learned) return learned;
  var m;

  if (p === "/api/me" && method === "GET") {
    if (!uid) return ok({ user: null });
    var u = userById(uid);
    return ok({ user: publicUser(u), profile: profileOut(find("Анкеты", function (r) { return Number(r.user_id) === uid; })), isAdmin: isAdmin(uid) });
  }

  if (p === "/api/specialists" && method === "GET") {
    return ok({ specialists: filter("Анкеты", function (r) { return r.status === "approved"; })
      .sort(function (a, b) { return a.updated_at < b.updated_at ? 1 : -1; }).slice(0, 500).map(profileOut) });
  }

  if (p === "/api/profile" && method === "POST") {
    if (!uid) return fail("login required", 401);
    var f = { name: clip(body.name, 80), title: clip(body.title, 80), category: clip(body.category, 20), skills: clip(body.skills, 1500),
      price: clip(body.price, 60), city: clip(body.city, 60), portfolio: clip(body.portfolio, 300) };
    if (!f.name || !f.title || !f.skills) return fail("name, title and skills are required");
    if (CATEGORIES.indexOf(f.category) < 0) return fail("bad category");
    if (f.portfolio && !/^https?:\/\//i.test(f.portfolio)) f.portfolio = "https://" + f.portfolio;
    f.status = "pending"; f.updated_at = now();
    var ex = find("Анкеты", function (r) { return Number(r.user_id) === uid; });
    if (ex) update("Анкеты", ex, f); else { f.user_id = uid; insert("Анкеты", f); }
    var me = userById(uid);
    notify("Анкета специалиста на проверку\n\n" + f.name + " — " + f.title + "\nЛогин: " + (me ? me.login : "—") +
      "\nНавыки: " + f.skills + (f.price ? "\nЦена: " + f.price : "") + (f.city ? "\nГород: " + f.city : "") +
      (f.portfolio ? "\nРаботы: " + f.portfolio : "") + "\n\nОдобрить: " + siteLink("/cabinet"));
    return ok({ profile: profileOut(find("Анкеты", function (r) { return Number(r.user_id) === uid; })) });
  }

  if (p === "/api/leads" && method === "POST") {
    if (body.website) return ok({ ok: true });                  // ловушка для спам-ботов
    var kind = body.kind === "b2b" ? "b2b" : "learn";
    var L = { name: clip(body.name, 80), contact: clip(body.contact, 120), track: clip(body.track, 20), level: clip(body.level, 40),
      english: clip(body.english, 20), city: clip(body.city, 60), company: clip(body.company, 120), message: clip(body.message, 1500) };
    if (!L.name || !L.contact) return fail("name and contact are required");
    if (L.track && !/^[a-z0-9_]{1,30}$/.test(L.track)) return fail("bad track");
    var hour = new Date(Date.now() - 3600000).toISOString().replace("T", " ").slice(0, 19);
    if (filter("Заявки", function (r) { return r.created_at > hour; }).length >= 60) return fail("too many attempts", 429);
    var trackLabel = clip(body.track_name, 80) || TRACK_RU[L.track] || L.track;
    L.id = nextId("Заявки"); L.kind = kind; L.status = "new"; L.created_at = now();
    insert("Заявки", L);
    notify((kind === "b2b" ? "Заявка от компании" : "Запись на обучение") + " №" + L.id + "\n\nИмя: " + L.name + "\nКонтакт: " + L.contact +
      (L.company ? "\nКомпания: " + L.company : "") + (L.track ? "\nНаправление: " + trackLabel : "") +
      (L.level ? "\nУровень: " + L.level : "") + (L.english ? "\nАнглийский: " + L.english : "") +
      (L.city ? "\nГород: " + L.city : "") + (L.message ? "\nКомментарий: " + L.message : "") + "\n\n" + siteLink("/cabinet"));
    return ok({ ok: true });
  }

  if (p === "/api/orders" && method === "POST") {
    if (!uid) return fail("login required", 401);
    var task = clip(body.task, 3000), contact = clip(body.contact, 120);
    if (!task || !contact) return fail("task and contact are required");
    var specId = body.specialist_id ? Number(body.specialist_id) : null, spec = null;
    if (specId) {
      spec = find("Анкеты", function (r) { return Number(r.user_id) === specId && r.status === "approved"; });
      if (!spec) return fail("specialist not found", 404);
      if (specId === uid) return fail("cannot order from yourself");
    }
    var id = nextId("Заказы");
    insert("Заказы", { id: id, client_id: uid, specialist_id: specId || "", task: task, budget: clip(body.budget, 60), deadline: clip(body.deadline, 60),
      contact: contact, status: "new", created_at: now() });
    var cl = userById(uid);
    notify("Новый заказ на SkillBridge\nЗаказ №" + id + "\n\nКлиент: " + (cl ? cl.name + " (" + cl.login + ")" : "—") + "\nКонтакт: " + contact +
      "\nСпециалист: " + (spec ? spec.name + " — " + spec.title : "не выбран, нужно подобрать") + "\nЗадача: " + task +
      (body.budget ? "\nБюджет: " + clip(body.budget, 60) : "") + (body.deadline ? "\nСрок: " + clip(body.deadline, 60) : "") + "\n\n" + siteLink("/cabinet"));
    return ok({ id: id });
  }

  if (p === "/api/orders" && method === "GET") {
    if (!uid) return fail("login required", 401);
    return ok({ mine: filter("Заказы", function (o) { return Number(o.client_id) === uid; }).sort(byIdDesc).slice(0, 200).map(orderOut),
      forme: filter("Заказы", function (o) { return Number(o.specialist_id) === uid; }).sort(byIdDesc).slice(0, 200).map(orderOut) });
  }

  if ((m = p.match(/^\/api\/orders\/(\d+)\/status$/)) && method === "POST") {
    if (!uid) return fail("login required", 401);
    var oid = Number(m[1]), status = String(body.status || "");
    if (STATUSES.indexOf(status) < 0) return fail("bad status");
    var o = find("Заказы", function (r) { return Number(r.id) === oid; });
    if (!o) return fail("not found", 404);
    var admin = isAdmin(uid), asSpec = Number(o.specialist_id) === uid, asClient = Number(o.client_id) === uid;
    var allowed = admin ||
      (asSpec && ((o.status === "new" && (status === "accepted" || status === "declined")) || (o.status === "accepted" && status === "done"))) ||
      (asClient && (o.status === "new" || o.status === "accepted") && status === "cancelled");
    if (!allowed) return fail("not allowed", 403);
    update("Заказы", o, { status: status });
    if (!admin) notify("Заказ №" + oid + ": статус «" + STATUS_RU[status] + "»\n" + clip(o.task, 200) + "\n\n" + siteLink("/cabinet"));
    return ok({ ok: true });
  }

  if (p.indexOf("/api/admin") === 0) {
    if (!isAdmin(uid)) return fail("admin only", 403);
    if (p === "/api/admin" && method === "GET") {
      var users = rows("Пользователи");
      var nameOf = function (id) { var x = userById(id); return x ? x.name : ""; };
      var loginOf = function (id) { var x = userById(id); return x ? x.login : ""; };
      return ok({
        pending: filter("Анкеты", function (r) { return r.status === "pending"; }).map(function (r) {
          var x = profileOut(r); x.first_name = nameOf(r.user_id); x.last_name = ""; x.username = null; x.login = loginOf(r.user_id); return x; }),
        approved: filter("Анкеты", function (r) { return r.status === "approved"; }).map(function (r) { return { user_id: Number(r.user_id), name: r.name, title: r.title }; })
          .sort(function (a, b) { return a.name < b.name ? -1 : 1; }),
        orders: rows("Заказы").slice().sort(byIdDesc).slice(0, 200).map(orderOut),
        leads: rows("Заявки").slice().sort(byIdDesc).slice(0, 200).map(function (r) {
          var x = {}; TABLES["Заявки"].forEach(function (c) { x[c] = r[c]; }); x.id = Number(r.id); return x; }),
        certs: rows("Сертификаты").slice().sort(function (a, b) { return a.issued_at < b.issued_at ? 1 : -1; }).slice(0, 100).map(function (c) {
          return { code: c.code, kind: c.kind, title: c.title, full_name: c.full_name, score: Number(c.score), revoked: c.revoked === "1" ? 1 : 0, issued_at: c.issued_at, username: null }; }),
        counts: { users: users.length, specialists: filter("Анкеты", function (r) { return r.status === "approved"; }).length, orders: rows("Заказы").length,
          leads: filter("Заявки", function (r) { return r.status === "new"; }).length, certs: filter("Сертификаты", function (r) { return r.revoked !== "1"; }).length }
      });
    }
    if ((m = p.match(/^\/api\/admin\/profiles\/(\d+)$/)) && method === "POST") {
      var st = String(body.status || "");
      if (["approved", "rejected", "pending"].indexOf(st) < 0) return fail("bad status");
      var pr = find("Анкеты", function (r) { return Number(r.user_id) === Number(m[1]); });
      if (pr) update("Анкеты", pr, { status: st });
      return ok({ ok: true });
    }
    if ((m = p.match(/^\/api\/admin\/certs\/([A-Z0-9-]{6,30})$/)) && method === "POST") {
      var ce = find("Сертификаты", function (r) { return r.code === m[1]; });
      if (ce) update("Сертификаты", ce, { revoked: body.revoked ? 1 : 0 });
      return ok({ ok: true });
    }
    if ((m = p.match(/^\/api\/admin\/leads\/(\d+)$/)) && method === "POST") {
      var ls = String(body.status || "");
      if (["new", "contacted", "archived"].indexOf(ls) < 0) return fail("bad status");
      var ld = find("Заявки", function (r) { return Number(r.id) === Number(m[1]); });
      if (ld) update("Заявки", ld, { status: ls });
      return ok({ ok: true });
    }
    if ((m = p.match(/^\/api\/admin\/orders\/(\d+)$/)) && method === "POST") {
      var sid = Number(body.specialist_id);
      var sp = find("Анкеты", function (r) { return Number(r.user_id) === sid && r.status === "approved"; });
      if (!sp) return fail("specialist not found", 404);
      var od = find("Заказы", function (r) { return Number(r.id) === Number(m[1]); });
      if (!od) return fail("not found", 404);
      update("Заказы", od, { specialist_id: sid, status: "new" });
      return ok({ ok: true });
    }
  }
  return fail("not found", 404);
}

// ---------- точка входа веб-приложения ----------
function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return respond({ ok: true, service: "SkillBridge", hint: "Сервер работает. Вставьте эту ссылку в config.js на сайте." });
}
function doPost(e) {
  var req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (err) { return respond(fail("bad json")); }
  var path = String(req.path || ""), method = req.method === "POST" ? "POST" : "GET", body = req.body && typeof req.body === "object" ? req.body : {};
  if (path.indexOf("/api/") !== 0) return respond(fail("not found", 404));
  var lock = null;
  try {
    if (method === "POST") { lock = LockService.getScriptLock(); lock.waitLock(25000); }
    _cache = {};
    syncAdmin();
    var uid = readToken(req.token);
    return respond(api(path, method, uid, body));
  } catch (err) {
    return respond({ status: 500, data: { error: "server error" } });
  } finally {
    if (lock) try { lock.releaseLock(); } catch (x) {}
  }
}

// ---------- курсы (собрано из src/courses.txt) ----------
var COURSES = [{"slug": "freelance", "lessons": [{"title": "Как устроен фриланс", "body": "Фриланс — это работа на заказ без постоянного трудоустройства: клиент платит за результат или за время, а вы сами выбираете проекты.\n\n### Три модели оплаты\n* **Проектная** — фиксированная цена за результат: логотип, лендинг, перевод документа.\n* **Почасовая** — оплата за отработанные часы. Подходит для задач, объём которых сложно оценить заранее.\n* **Ретейнер** — ежемесячная оплата за постоянный объём работы: ведение соцсетей, поддержка сайта.\n\n### Где искать заказы\n* Международные площадки: Upwork, Fiverr, Freelancer — оплата в долларах, нужен английский.\n* Русскоязычные: Kwork, FL.ru, Freelance.ru.\n* Telegram-каналы и чаты с вакансиями, LinkedIn, знакомые и бывшие работодатели.\n\n### Чего ждёт клиент\nКлиенту важно не ваше резюме, а уверенность, что задача будет сделана в срок и так, как нужно. Поэтому на старте решают три вещи: понятное описание услуги, примеры работ и быстрый вежливый ответ.", "quiz": [{"q": "Как называется модель, когда клиент платит фиксированную сумму каждый месяц за постоянный объём работы?", "options": ["Проектная оплата", "Почасовая оплата", "Ретейнер"], "correct": 2}, {"q": "Что важнее всего клиенту, когда он выбирает исполнителя без отзывов?", "options": ["Диплом престижного вуза", "Понятная услуга, примеры работ и быстрый ответ", "Самая низкая цена на площадке"], "correct": 1}]}, {"title": "Выбор ниши и упаковка услуги", "body": "Новичку проще продавать одну понятную услугу, чем «всё понемногу». «Делаю дизайн» — размыто. «Делаю обложки для Reels для кафе и салонов красоты» — понятно и запоминается.\n\n### Формула услуги\n**Что делаю + для кого + какой результат получит клиент.**\n\nПример: «Настрою таргетированную рекламу в Instagram для местного бизнеса, чтобы получать заявки в директ».\n\n### Как выбрать нишу\n1. Выпишите навыки, которые у вас уже есть или которые вы осваиваете.\n2. Посмотрите на площадках, какие заказы по ним публикуют чаще всего.\n3. Выберите 1–2 услуги, где спрос есть, а вход для новичка реален.\n\n### Цена на старте\nНачните с нижней части рыночного диапазона, но не демпингуйте. После 3–5 хороших отзывов повышайте цену. Слишком низкая цена отпугивает часть клиентов: она выглядит как признак низкого качества.", "quiz": [{"q": "Какая формулировка услуги лучше подходит для профиля новичка?", "options": ["«Делаю всё, что связано с компьютером»", "«Дизайн, SMM, сайты, переводы, тексты»", "«Делаю карточки товаров для Wildberries, которые повышают кликабельность»"], "correct": 2}, {"q": "Когда обычно стоит повышать цену?", "options": ["Сразу после регистрации на площадке", "После нескольких хороших отзывов и выполненных заказов", "Никогда: низкая цена — главное преимущество"], "correct": 1}]}, {"title": "Портфолио без опыта", "body": "Без примеров работ клиенту трудно вам поверить. Хорошая новость: первые работы можно сделать самостоятельно.\n\n### Откуда взять первые работы\n* **Учебные проекты** — задания из курсов, доведённые до аккуратного вида.\n* **Редизайн** — улучшите сайт, баннер или текст реальной компании и покажите «было — стало».\n* **Работа для знакомых** — бесплатно или недорого, но в обмен на отзыв.\n\n### Как оформить кейс\nКаждая работа в портфолио — это короткая история:\n1. **Задача** — что нужно было сделать и для кого.\n2. **Решение** — что вы сделали и почему.\n3. **Результат** — что получилось: цифры, отзыв, скриншоты.\n\n### Где хранить\nBehance или Dribbble — для дизайнеров, GitHub — для разработчиков, Google Drive или Notion — для всех остальных. Главное, чтобы ссылка открывалась без регистрации.", "quiz": [{"q": "Из каких частей состоит хорошо оформленный кейс?", "options": ["Задача, решение, результат", "Название, цена, дата", "Только финальная картинка без описания"], "correct": 0}, {"q": "Что делать, если у вас ещё нет реальных заказов?", "options": ["Ждать первого заказа, а потом делать портфолио", "Сделать учебные проекты и редизайны и оформить их как кейсы", "Взять чужие работы из интернета"], "correct": 1}]}, {"title": "Профиль и отклик клиенту", "body": "### Профиль\n* **Фото** — ваше лицо, нейтральный фон, хорошее освещение.\n* **Заголовок** — ваша услуга, а не должность: «UX/UI-дизайнер мобильных приложений».\n* **Описание** — для кого вы работаете, какие задачи решаете, как проходит работа.\n\n### Отклик, на который отвечают\nКлиент получает десятки откликов. Шаблонное «Здравствуйте, готов выполнить ваш заказ» он пропускает. Хороший отклик короткий и показывает, что вы прочитали задачу:\n1. Обращение и одна фраза о том, как вы поняли задачу.\n2. Как будете решать — 2–3 шага.\n3. Ссылка на похожую работу.\n4. Уточняющий вопрос.\n5. Срок и цена.\n\n**Пример:** «Здравствуйте, Анна! Вам нужен лендинг для курсов английского с формой записи. Сделаю макет в Figma, сверстаю на Tilda и подключу заявки в Telegram. Похожая работа: (ссылка). Есть ли у вас фирменные цвета и логотип? Срок — 5 дней, стоимость — 150 $».", "quiz": [{"q": "Что лучше написать в заголовке профиля?", "options": ["«Ищу работу»", "Конкретную услугу, например «Настройка таргетированной рекламы для малого бизнеса»", "Только имя и город"], "correct": 1}, {"q": "Зачем в отклике задавать уточняющий вопрос?", "options": ["Он показывает, что вы вникли в задачу, и начинает диалог", "Чтобы потянуть время", "Вопросы в отклике лучше не задавать"], "correct": 0}]}, {"title": "Работа с заказом и безопасность", "body": "### Перед стартом\n* Уточните задачу письменно: что нужно сделать, в каком виде сдать, срок, сколько правок входит в цену.\n* Договоритесь об оплате: предоплата, оплата по этапам или безопасная сделка площадки.\n\n### Во время работы\nСообщайте о прогрессе без напоминаний клиента. Если что-то пошло не по плану — пишите сразу и предлагайте решение.\n\n### Сдача и отзыв\nСдавайте работу в оговорённом виде, коротко поясните, что сделано. После приёмки вежливо попросите отзыв — на старте это ваш главный капитал.\n\n### Признаки мошенничества\n* Просят заплатить за «тестовое задание», «доступ к заказам» или «обучение».\n* Уговаривают уйти с площадки до оплаты.\n* Присылают ссылки на «оплату», где нужно ввести данные карты.\n* Обещают оплату намного выше рынка за простую работу.", "quiz": [{"q": "Какой способ оплаты защищает исполнителя лучше всего?", "options": ["Полная оплата после сдачи, без договорённостей", "Предоплата, оплата по этапам или безопасная сделка площадки", "Оплата «как получится»"], "correct": 1}, {"q": "Заказчик просит вас перевести 20 $ за «доступ к проекту». Что это?", "options": ["Обычная практика", "Признак мошенничества", "Предоплата исполнителя"], "correct": 1}]}], "exam": [{"q": "Какая модель оплаты подходит для ежемесячного ведения соцсетей?", "options": ["Проектная", "Ретейнер", "Разовая"], "correct": 1}, {"q": "Какая формулировка услуги сильнее?", "options": ["«Дизайнер, сделаю всё»", "«Создаю обложки для YouTube, которые повышают кликабельность»", "«Работаю за любые деньги»"], "correct": 1}, {"q": "Что лучше всего заменяет отзывы на старте?", "options": ["Скидка 90%", "Портфолио из учебных проектов и редизайнов", "Длинное резюме"], "correct": 1}, {"q": "Как начать хороший отклик?", "options": ["«Здравствуйте, готов выполнить заказ»", "С обращения и одной фразы о том, как вы поняли задачу", "С рассказа о своём образовании"], "correct": 1}, {"q": "Что нужно согласовать письменно до начала работы?", "options": ["Задачу, формат сдачи, срок, число правок и порядок оплаты", "Только цену", "Ничего: так быстрее"], "correct": 0}, {"q": "Клиент предлагает продолжить общение в мессенджере и оплатить «потом». Как поступить?", "options": ["Согласиться сразу", "Оставаться на площадке или договориться о предоплате и безопасной сделке", "Попросить оплату криптовалютой без договора"], "correct": 1}, {"q": "Когда разумно повышать цену?", "options": ["После нескольких выполненных заказов с хорошими отзывами", "На следующий день после регистрации", "Цену нельзя менять"], "correct": 0}, {"q": "Что делать, если вы не успеваете к сроку?", "options": ["Молчать до последнего", "Сразу сообщить клиенту и предложить решение", "Сдать недоделанную работу без объяснений"], "correct": 1}], "title": "Старт на фрилансе: от профиля до первого заказа", "title_en": "Freelance start: from profile to first order", "level": "Начальный", "hours": 4, "summary": "Как выбрать услугу, собрать портфолио без опыта, написать отклик и безопасно провести первый заказ.", "about": "Курс для тех, кто хочет зарабатывать удалённо, но не знает, с чего начать. Вы разберётесь, как устроены фриланс-площадки, упакуете свою услугу, соберёте первые работы в портфолио и научитесь писать отклики, на которые отвечают клиенты."}, {"slug": "excel", "lessons": [{"title": "Интерфейс и данные", "body": "Таблица состоит из **ячеек**. У каждой ячейки есть адрес: буква столбца и номер строки — например, `B3`.\n\n### Типы данных\n* **Числа** — с ними можно считать. Выравниваются по правому краю.\n* **Текст** — выравнивается по левому краю. Если число «прижалось» влево, скорее всего, оно сохранено как текст и не будет считаться.\n* **Даты** — внутри хранятся как числа, поэтому их можно вычитать и сравнивать.\n\n### Полезные приёмы\n* **Автозаполнение** — потяните за правый нижний угол ячейки, чтобы продолжить ряд: 1, 2, 3… или январь, февраль…\n* **Закрепление строки** — «Вид → Закрепить», чтобы заголовки оставались на экране при прокрутке.\n* **Формат ячеек** — денежный, процентный, дата. Формат меняет вид числа, но не само значение.\n\n### Важно про язык\nВ русском Excel функции пишутся по-русски (`СУММ`), а аргументы разделяются точкой с запятой. В Google Таблицах работают английские названия (`SUM`). В курсе мы приводим оба варианта.", "quiz": [{"q": "Что обозначает адрес ячейки C5?", "options": ["Столбец C, строка 5", "Строка C, столбец 5", "Пятый лист таблицы"], "correct": 0}, {"q": "Число в ячейке прижато к левому краю и не участвует в сумме. Почему?", "options": ["Ячейка слишком узкая", "Число сохранено как текст", "В таблице слишком много строк"], "correct": 1}]}, {"title": "Формулы и ссылки", "body": "Любая формула начинается со знака `=`. Например, `=B2*C2` перемножает значения двух ячеек.\n\n### Базовые функции\n* `=СУММ(B2:B10)` / `=SUM(B2:B10)` — сумма диапазона.\n* `=СРЗНАЧ(B2:B10)` / `=AVERAGE(B2:B10)` — среднее.\n* `=МАКС()` и `=МИН()` / `=MAX()` и `=MIN()` — наибольшее и наименьшее.\n* `=ОКРУГЛ(B2;2)` / `=ROUND(B2,2)` — округление до двух знаков.\n\n### Относительные и абсолютные ссылки\nКогда вы копируете формулу вниз, ссылки **сдвигаются**: `=B2*C2` превращается в `=B3*C3`. Это относительные ссылки.\n\nЕсли ссылка должна оставаться на месте — например, на ячейку с курсом валюты, — закрепите её знаком доллара: `=B2*$F$1`. Это абсолютная ссылка. Быстро поставить знаки `$` можно клавишей **F4**.", "quiz": [{"q": "Формулу `=B2*$F$1` скопировали на строку ниже. Что получится?", "options": ["`=B3*$F$2`", "`=B3*$F$1`", "`=B2*$F$1`"], "correct": 1}, {"q": "С какого символа начинается любая формула?", "options": ["С двоеточия", "Со знака равенства", "С кавычки"], "correct": 1}]}, {"title": "Логика и поиск данных", "body": "### ЕСЛИ\n`=ЕСЛИ(B2>=100;\"Да\";\"Нет\")` / `=IF(B2>=100,\"Да\",\"Нет\")` — проверяет условие и возвращает одно из двух значений.\n\n### Подсчёт по условию\n* `=СЧЁТЕСЛИ(C2:C100;\"Душанбе\")` / `=COUNTIF(...)` — сколько раз встречается значение.\n* `=СУММЕСЛИ(C2:C100;\"Душанбе\";D2:D100)` / `=SUMIF(...)` — сумма по условию.\n\n### ВПР — поиск значения в другой таблице\n`=ВПР(A2;Прайс!A:B;2;ЛОЖЬ)` / `=VLOOKUP(A2,Прайс!A:B,2,FALSE)` ищет значение из `A2` в первом столбце диапазона и возвращает значение из второго столбца.\n\nПоследний аргумент `ЛОЖЬ` (или `0`) означает **точное совпадение** — почти всегда нужен именно он. В новых версиях есть удобная замена — `ПРОСМОТРX` / `XLOOKUP`.", "quiz": [{"q": "Что вернёт `=ЕСЛИ(5>3;\"А\";\"Б\")`?", "options": ["А", "Б", "Ошибку"], "correct": 0}, {"q": "Зачем в ВПР указывают последний аргумент ЛОЖЬ (0)?", "options": ["Чтобы искать приблизительно", "Чтобы искать точное совпадение", "Чтобы скрыть ошибки"], "correct": 1}]}, {"title": "Сортировка, фильтры и оформление", "body": "### Сортировка и фильтр\nВыделите таблицу с заголовками и включите фильтр («Данные → Фильтр»). В заголовках появятся стрелки: по ним можно сортировать и скрывать лишние строки. Фильтр не удаляет данные — он только временно их прячет.\n\n### Условное форматирование\nПравила, которые сами раскрашивают ячейки: например, просроченные даты — красным, суммы больше 1000 — зелёным. Удобно, чтобы сразу видеть важное.\n\n### Проверка данных\n«Данные → Проверка данных» позволяет ограничить ввод: только числа, только даты или выбор из выпадающего списка. Так меньше ошибок, когда таблицу заполняют несколько человек.", "quiz": [{"q": "Что делает фильтр со строками, которые не подходят под условие?", "options": ["Удаляет их навсегда", "Временно скрывает", "Перемещает на другой лист"], "correct": 1}, {"q": "Как сделать, чтобы в ячейке можно было выбрать значение только из списка?", "options": ["Условное форматирование", "Проверка данных с выпадающим списком", "Закрепление строки"], "correct": 1}]}, {"title": "Сводные таблицы и диаграммы", "body": "### Сводная таблица\nПозволяет за минуту получить итог по большим данным: например, сумму продаж по каждому городу и месяцу. «Вставка → Сводная таблица», затем перетащите поля:\n* в **Строки** — то, по чему группируем (город);\n* в **Столбцы** — второй разрез (месяц);\n* в **Значения** — что считаем (сумма продаж).\n\n### Какую диаграмму выбрать\n* **Линейчатая или столбчатая** — сравнить значения между категориями.\n* **График (линия)** — показать изменение во времени.\n* **Круговая** — доли целого, если категорий немного (до 5–6).\n\nДайте диаграмме понятное название и подпишите оси — так отчёт поймут без пояснений.", "quiz": [{"q": "Куда в сводной таблице перетаскивают поле, по которому нужно посчитать сумму?", "options": ["В «Фильтры»", "В «Строки»", "В «Значения»"], "correct": 2}, {"q": "Какая диаграмма лучше показывает изменение продаж по месяцам?", "options": ["График (линия)", "Круговая", "Таблица без диаграммы"], "correct": 0}]}], "exam": [{"q": "Как называется ссылка `$A$1`?", "options": ["Относительная", "Абсолютная", "Текстовая"], "correct": 1}, {"q": "Какая функция считает среднее значение?", "options": ["`СУММ` / `SUM`", "`СРЗНАЧ` / `AVERAGE`", "`СЧЁТ` / `COUNT`"], "correct": 1}, {"q": "Какая клавиша быстро добавляет знаки `$` в ссылку?", "options": ["F2", "F4", "F9"], "correct": 1}, {"q": "Как посчитать, сколько раз в столбце встречается «Худжанд»?", "options": ["`СЧЁТЕСЛИ` / `COUNTIF`", "`СУММ` / `SUM`", "`ОКРУГЛ` / `ROUND`"], "correct": 0}, {"q": "Что делает `ВПР` / `VLOOKUP`?", "options": ["Сортирует таблицу", "Находит значение в первом столбце диапазона и возвращает значение из нужного столбца", "Строит диаграмму"], "correct": 1}, {"q": "Что нужно, чтобы просроченные даты автоматически выделялись красным?", "options": ["Фильтр", "Условное форматирование", "Сводная таблица"], "correct": 1}, {"q": "Какой инструмент быстрее всего покажет сумму продаж по каждому городу?", "options": ["Сводная таблица", "Закрепление строки", "Проверка данных"], "correct": 0}, {"q": "Какая диаграмма подходит, чтобы показать доли 3–4 категорий в целом?", "options": ["График", "Круговая", "Точечная"], "correct": 1}], "title": "Excel и Google Таблицы для работы", "title_en": "Excel and Google Sheets for work", "track": "excel", "level": "Начальный", "hours": 6, "summary": "Формулы, ссылки, ЕСЛИ и ВПР, фильтры, условное форматирование, сводные таблицы и диаграммы.", "about": "Таблицы нужны почти в любой удалённой работе: отчёты, учёт заказов, аналитика, ввод данных. За курс вы научитесь уверенно работать с формулами, находить данные в больших таблицах и строить сводные отчёты. Примеры подходят и для Excel, и для Google Таблиц."}, {"slug": "html", "lessons": [{"title": "Как устроен веб", "body": "Когда вы открываете сайт, браузер запрашивает у **сервера** файлы и показывает их на экране.\n\n### Три языка веб-страницы\n* **HTML** — содержание и структура: заголовки, текст, ссылки, картинки.\n* **CSS** — внешний вид: цвета, шрифты, отступы, расположение.\n* **JavaScript** — поведение: реакция на клики, загрузка данных.\n\n### Рабочее место\n1. Установите редактор **VS Code**.\n2. Создайте папку проекта и файл `index.html`.\n3. Откройте файл в браузере двойным щелчком — и вы увидите свою страницу.", "quiz": [{"q": "Какой язык отвечает за внешний вид страницы: цвета, шрифты, отступы?", "options": ["HTML", "CSS", "JavaScript"], "correct": 1}, {"q": "Как обычно называют главный файл сайта?", "options": ["index.html", "main.css", "site.exe"], "correct": 0}]}, {"title": "Структура HTML-документа", "body": "HTML состоит из **тегов**. Большинство тегов парные: `<p>текст</p>`.\n\n### Каркас страницы\n```\n<!doctype html>\n<html lang=\"ru\">\n<head>\n  <meta charset=\"utf-8\">\n  <title>Моя страница</title>\n</head>\n<body>\n  <h1>Привет!</h1>\n  <p>Это моя первая страница.</p>\n</body>\n</html>\n```\nВ `head` — служебная информация, в `body` — всё, что видно на экране.\n\n### Основные теги\n* `<h1>`–`<h6>` — заголовки разных уровней. `h1` на странице обычно один.\n* `<p>` — абзац текста.\n* `<a href=\"https://...\">` — ссылка.\n* `<img src=\"photo.jpg\" alt=\"описание\">` — картинка. Атрибут `alt` описывает изображение для поисковиков и незрячих людей.\n* `<ul>` и `<ol>` со вложенными `<li>` — маркированный и нумерованный списки.", "quiz": [{"q": "В какой части документа находится содержимое, которое видит пользователь?", "options": ["В `<head>`", "В `<body>`", "В `<title>`"], "correct": 1}, {"q": "Для чего у тега `<img>` нужен атрибут `alt`?", "options": ["Задаёт размер картинки", "Описывает изображение текстом", "Делает картинку ссылкой"], "correct": 1}]}, {"title": "Семантика и формы", "body": "### Семантические теги\nВместо одних `<div>` используйте теги, которые описывают смысл блока:\n* `<header>` — шапка, `<nav>` — меню, `<main>` — основное содержимое;\n* `<section>` — смысловой раздел, `<footer>` — подвал.\n\nСемантика помогает поисковикам и экранным дикторам понимать страницу.\n\n### Формы\n```\n<form>\n  <label for=\"email\">E-mail</label>\n  <input id=\"email\" type=\"email\" required>\n  <button type=\"submit\">Отправить</button>\n</form>\n```\n* `<label for=\"...\">` связывается с полем по `id` — по клику на подпись курсор встаёт в поле.\n* Тип поля (`email`, `tel`, `number`, `date`) подсказывает телефону, какую клавиатуру показать.\n* `required` не даст отправить форму с пустым полем.", "quiz": [{"q": "Какой тег подходит для главного меню сайта?", "options": ["`<footer>`", "`<nav>`", "`<p>`"], "correct": 1}, {"q": "Как связать подпись `<label>` с полем ввода?", "options": ["Указать в `for` тот же текст, что в `id` поля", "Написать подпись внутри `<title>`", "Никак, они связываются сами"], "correct": 0}]}, {"title": "Основы CSS", "body": "Стили подключают отдельным файлом: `<link rel=\"stylesheet\" href=\"style.css\">` в `head`.\n\n### Селекторы\n* `p { }` — все абзацы;\n* `.card { }` — все элементы с классом `card`;\n* `#logo { }` — элемент с `id=\"logo\"`.\n\n### Блочная модель\nКаждый элемент — это прямоугольник:\n* **content** — содержимое;\n* **padding** — внутренний отступ (между содержимым и рамкой);\n* **border** — рамка;\n* **margin** — внешний отступ (до соседних элементов).\n\nДобавьте в начало стилей `* { box-sizing: border-box; }` — тогда `width` будет включать отступы и рамку, и считать размеры станет проще.", "quiz": [{"q": "Какой селектор выберет все элементы с классом `button`?", "options": ["`#button`", "`.button`", "`button()`"], "correct": 1}, {"q": "Как называется внутренний отступ между содержимым и рамкой?", "options": ["padding", "margin", "border"], "correct": 0}]}, {"title": "Раскладка, адаптивность и публикация", "body": "### Flexbox\n```\n.row {\n  display: flex;\n  gap: 16px;\n  justify-content: space-between;\n  align-items: center;\n}\n```\n`display: flex` выстраивает дочерние элементы в ряд, `gap` задаёт расстояние между ними, `justify-content` — выравнивание по горизонтали, `align-items` — по вертикали.\n\n### Адаптивность\nДобавьте в `head` строку `<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">` — без неё телефон покажет уменьшенную «десктопную» версию.\n\nМедиазапросы меняют стили на узких экранах:\n```\n@media (max-width: 600px) {\n  .row { flex-direction: column; }\n}\n```\n\n### Публикация\nБесплатно опубликовать страницу можно на **GitHub Pages** или **Netlify**: загрузите папку проекта — и получите ссылку, которую можно отправить клиенту или добавить в портфолио.", "quiz": [{"q": "Какое свойство задаёт расстояние между элементами во flex-контейнере?", "options": ["padding", "gap", "border"], "correct": 1}, {"q": "Что произойдёт без `<meta name=\"viewport\">` на телефоне?", "options": ["Страница отобразится уменьшенной, как на компьютере", "Страница не откроется", "Пропадут картинки"], "correct": 0}]}], "exam": [{"q": "Что описывает HTML?", "options": ["Структуру и содержание страницы", "Только цвета и шрифты", "Работу сервера"], "correct": 0}, {"q": "Какой тег создаёт ссылку?", "options": ["`<link>`", "`<a>`", "`<href>`"], "correct": 1}, {"q": "Какой тег подходит для подвала страницы?", "options": ["`<header>`", "`<main>`", "`<footer>`"], "correct": 2}, {"q": "Как выбрать элемент с `id=\"menu\"` в CSS?", "options": ["`#menu`", "`.menu`", "`menu`"], "correct": 0}, {"q": "Чем margin отличается от padding?", "options": ["Ничем", "margin — внешний отступ, padding — внутренний", "margin — цвет рамки, padding — её толщина"], "correct": 1}, {"q": "Что делает `box-sizing: border-box`?", "options": ["Включает padding и border в заданную ширину элемента", "Добавляет рамку всем элементам", "Скрывает переполнение"], "correct": 0}, {"q": "Какое свойство выстраивает дочерние элементы в ряд?", "options": ["`display: block`", "`display: flex`", "`position: absolute`"], "correct": 1}, {"q": "Где можно бесплатно опубликовать статичный сайт?", "options": ["Только на собственном сервере", "GitHub Pages или Netlify", "В Excel"], "correct": 1}], "title": "HTML и CSS: первая веб-страница", "title_en": "HTML and CSS: your first web page", "track": "frontend", "level": "Начальный", "hours": 6, "summary": "Структура HTML, семантика, формы, селекторы и блочная модель CSS, flexbox, адаптивность и публикация сайта.", "about": "Первый шаг во фронтенд-разработку. Вы поймёте, как устроены веб-страницы, сверстаете аккуратную адаптивную страницу и опубликуете её в интернете. Нужен только компьютер с браузером и бесплатным редактором VS Code."}];
var PROGRAMS = [{"slug": "digital_basic", "title": "Цифровой специалист: базовая квалификация", "title_en": "Digital specialist: basic qualification", "courses": ["freelance", "excel", "html"], "summary": "Три курса — фриланс, таблицы и вёрстка. После всех трёх сертификатов выдаётся удостоверение SkillBridge."}];
