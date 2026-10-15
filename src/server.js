// SkillBridge — биржа специалистов: каталог, вход через Telegram, кабинет, заказы.
// Cloudflare Worker + база D1.
//
// Настройки воркера (Settings → Variables and Secrets / Bindings):
//   DB            — привязка (binding) к базе D1
//   BOT_TOKEN     — токен бота от @BotFather (Secret)
//   BOT_USERNAME  — логин бота без @, например skillbridge_leads_bot
//   ADMIN_ID      — ваш числовой Telegram ID (от @userinfobot): вы — администратор

const CATEGORIES = ["dev", "marketing", "design", "data", "lang", "cloud", "video", "text"];
const TRACKS = ["it", "marketing", "design", "data", "english", "cloud"];
const TRACK_RU = { it: "IT-разработка", marketing: "Цифровой маркетинг", design: "UX/UI-дизайн", data: "Данные и аналитика", english: "Профессиональный английский", cloud: "Cloud и DevOps" };
const STATUSES = ["new", "accepted", "declined", "done", "cancelled"];
const STATUS_RU = { new: "новый", accepted: "принят", declined: "отклонён", done: "выполнен", cancelled: "отменён" };

const SCHEMA = [
  `CREATE TABLE IF NOT EXISTS users (
     id INTEGER PRIMARY KEY, first_name TEXT, last_name TEXT, username TEXT, photo_url TEXT,
     created_at TEXT DEFAULT (datetime('now')))`,
  `CREATE TABLE IF NOT EXISTS profiles (
     user_id INTEGER PRIMARY KEY, name TEXT, title TEXT, category TEXT, skills TEXT, price TEXT,
     city TEXT, portfolio TEXT, status TEXT DEFAULT 'pending', updated_at TEXT DEFAULT (datetime('now')))`,
  `CREATE TABLE IF NOT EXISTS orders (
     id INTEGER PRIMARY KEY AUTOINCREMENT, client_id INTEGER NOT NULL, specialist_id INTEGER,
     task TEXT NOT NULL, budget TEXT, deadline TEXT, contact TEXT, status TEXT DEFAULT 'new',
     created_at TEXT DEFAULT (datetime('now')))`,
  `CREATE TABLE IF NOT EXISTS progress (
     user_id INTEGER NOT NULL, course TEXT NOT NULL, lesson INTEGER NOT NULL, score REAL, passed INTEGER DEFAULT 0,
     updated_at TEXT DEFAULT (datetime('now')), PRIMARY KEY (user_id, course, lesson))`,
  `CREATE TABLE IF NOT EXISTS exam_attempts (
     id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, course TEXT NOT NULL, score REAL, passed INTEGER,
     created_at TEXT DEFAULT (datetime('now')))`,
  `CREATE TABLE IF NOT EXISTS certificates (
     code TEXT PRIMARY KEY, user_id INTEGER NOT NULL, kind TEXT NOT NULL, ref TEXT NOT NULL, title TEXT, full_name TEXT,
     hours INTEGER, score INTEGER, revoked INTEGER DEFAULT 0, issued_at TEXT DEFAULT (datetime('now')))`,
  `CREATE TABLE IF NOT EXISTS leads (
     id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT, name TEXT, contact TEXT, track TEXT, level TEXT,
     english TEXT, city TEXT, company TEXT, message TEXT, status TEXT DEFAULT 'new',
     created_at TEXT DEFAULT (datetime('now')))`,
];
let schemaReady = false;
async function ensureSchema(env) {
  if (schemaReady) return;
  await env.DB.batch(SCHEMA.map((s) => env.DB.prepare(s)));
  schemaReady = true;
}

// ---------- crypto / session ----------
const enc = new TextEncoder();
async function sha256(data) {
  return new Uint8Array(await crypto.subtle.digest("SHA-256", typeof data === "string" ? enc.encode(data) : data));
}
async function hmac(keyBytes, msg) {
  const key = await crypto.subtle.importKey("raw", keyBytes, { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  return new Uint8Array(await crypto.subtle.sign("HMAC", key, enc.encode(msg)));
}
const hex = (b) => [...b].map((x) => x.toString(16).padStart(2, "0")).join("");
function safeEqual(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

// https://core.telegram.org/widgets/login#checking-authorization
async function verifyTelegram(params, botToken) {
  const fields = ["auth_date", "first_name", "id", "last_name", "photo_url", "username"];
  const hash = params.get("hash");
  if (!hash || !params.get("id") || !params.get("auth_date")) return null;
  const check = fields.filter((f) => params.has(f)).map((f) => f + "=" + params.get(f)).join("\n");
  const calc = hex(await hmac(await sha256(botToken), check));
  if (!safeEqual(calc, hash)) return null;
  if (Date.now() / 1000 - Number(params.get("auth_date")) > 86400) return null;
  return {
    id: Number(params.get("id")),
    first_name: params.get("first_name") || "",
    last_name: params.get("last_name") || "",
    username: params.get("username") || "",
    photo_url: params.get("photo_url") || "",
  };
}

async function sessionKey(env) {
  return hmac(enc.encode("skillbridge-session-v1"), env.BOT_TOKEN);
}
async function makeSession(env, uid) {
  const exp = Math.floor(Date.now() / 1000) + 30 * 86400;
  const body = uid + "." + exp;
  return body + "." + hex(await hmac(await sessionKey(env), body));
}
async function readSession(request, env) {
  const m = (request.headers.get("Cookie") || "").match(/(?:^|;\s*)sb=([^;]+)/);
  if (!m) return null;
  const [uid, exp, sig] = m[1].split(".");
  if (!uid || !exp || !sig || Number(exp) < Date.now() / 1000) return null;
  const calc = hex(await hmac(await sessionKey(env), uid + "." + exp));
  return safeEqual(calc, sig) ? Number(uid) : null;
}
function cookie(value, maxAge) {
  return "sb=" + value + "; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=" + maxAge;
}

// ---------- helpers ----------
const json = (data, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });
const fail = (msg, status = 400) => json({ error: msg }, status);
const clip = (v, n) => String(v == null ? "" : v).trim().slice(0, n);
const isAdmin = (env, uid) => uid != null && String(uid) === String(env.ADMIN_ID || "").trim();

async function notify(env, chatId, text) {
  if (!chatId || !env.BOT_TOKEN) return;
  try {
    await fetch("https://api.telegram.org/bot" + env.BOT_TOKEN + "/sendMessage", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ chat_id: chatId, text, disable_web_page_preview: true }),
    });
  } catch (e) { /* уведомление не обязательно */ }
}
function who(u) {
  if (!u) return "—";
  const name = [u.first_name, u.last_name].filter(Boolean).join(" ");
  return name + (u.username ? " (@" + u.username + ")" : "");
}


// ---------- обучение: курсы, тесты, экзамены, сертификаты ----------
// COURSES и PROGRAMS собираются из src/courses.txt (см. build.py).
const PASS_SCORE = 0.7;          // доля верных ответов для зачёта урока и экзамена
const EXAM_ATTEMPTS_PER_DAY = 5;
const courseBySlug = (slug) => COURSES.find((c) => c.slug === slug);
const stripQuiz = (quiz) => quiz.map((q, i) => ({ i, q: q.q, options: q.options }));
function courseCard(c) {
  return { slug: c.slug, title: c.title, title_en: c.title_en, track: c.track, level: c.level, hours: c.hours,
    summary: c.summary, lessons: c.lessons.length, exam: c.exam.length };
}
function grade(quiz, answers) {
  const results = quiz.map((q, i) => Number(answers && answers[i]) === q.correct);
  const right = results.filter(Boolean).length;
  return { results, score: quiz.length ? right / quiz.length : 1, right, total: quiz.length };
}
function certCode() {
  const abc = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789", b = new Uint8Array(8);
  crypto.getRandomValues(b);
  return "SB-" + new Date().getUTCFullYear() + "-" + [...b].map((x) => abc[x % abc.length]).join("");
}
async function passedLessons(db, uid, slug) {
  const { results } = await db.prepare("SELECT lesson FROM progress WHERE user_id = ? AND course = ? AND passed = 1").bind(uid, slug).all();
  return results.map((r) => r.lesson);
}
async function issueProgramCerts(db, env, uid, fullName) {
  const issued = [];
  const { results } = await db.prepare("SELECT kind, ref, score FROM certificates WHERE user_id = ? AND revoked = 0").bind(uid).all();
  const have = (kind, ref) => results.find((r) => r.kind === kind && r.ref === ref);
  for (const pr of PROGRAMS) {
    if (have("program", pr.slug) || !pr.courses.every((s) => have("course", s))) continue;
    const hours = pr.courses.reduce((n, s) => n + (courseBySlug(s)?.hours || 0), 0);
    const score = Math.round(pr.courses.reduce((n, s) => n + have("course", s).score, 0) / pr.courses.length);
    const code = certCode();
    await db.prepare("INSERT INTO certificates (code, user_id, kind, ref, title, full_name, hours, score) VALUES (?, ?, 'program', ?, ?, ?, ?, ?)")
      .bind(code, uid, pr.slug, pr.title, fullName, hours, score).run();
    issued.push(code);
  }
  return issued;
}

async function learningApi(request, env, url, uid, body) {
  const p = url.pathname, method = request.method, db = env.DB;
  const needAuth = () => (uid ? null : fail("login required", 401));
  let m;

  if (p === "/api/courses" && method === "GET") {
    let my = {};
    if (uid) {
      const { results } = await db.prepare("SELECT course, COUNT(*) AS n FROM progress WHERE user_id = ? AND passed = 1 GROUP BY course").bind(uid).all();
      results.forEach((r) => (my[r.course] = r.n));
    }
    return json({ courses: COURSES.map(courseCard), programs: PROGRAMS, my });
  }

  if ((m = p.match(/^\/api\/courses\/(\w+)$/)) && method === "GET") {
    const c = courseBySlug(m[1]);
    if (!c) return fail("not found", 404);
    let my = null;
    if (uid) {
      const cert = await db.prepare("SELECT code FROM certificates WHERE user_id = ? AND kind = 'course' AND ref = ? AND revoked = 0").bind(uid, c.slug).first();
      my = { passed: await passedLessons(db, uid, c.slug), certificate: cert ? cert.code : null };
    }
    return json({ course: { ...courseCard(c), about: c.about, outline: c.lessons.map((l, i) => ({ n: i + 1, title: l.title, quiz: l.quiz.length })) },
      programs: PROGRAMS.filter((pr) => pr.courses.includes(c.slug)), my });
  }

  if ((m = p.match(/^\/api\/courses\/(\w+)\/lessons\/(\d+)$/))) {
    const denied = needAuth(); if (denied) return denied;
    const c = courseBySlug(m[1]), n = Number(m[2]), l = c && c.lessons[n - 1];
    if (!l) return fail("not found", 404);
    if (method === "GET") {
      const row = await db.prepare("SELECT passed, score FROM progress WHERE user_id = ? AND course = ? AND lesson = ?").bind(uid, c.slug, n).first();
      return json({ course: courseCard(c), lesson: { n, title: l.title, body: l.body, quiz: stripQuiz(l.quiz) },
        passed: !!(row && row.passed), total: c.lessons.length, outline: c.lessons.map((x, i) => ({ n: i + 1, title: x.title })),
        done: await passedLessons(db, uid, c.slug) });
    }
    if (method === "POST") {
      const g = grade(l.quiz, body.answers);
      const passed = g.score >= PASS_SCORE ? 1 : 0;
      await db.prepare(`INSERT INTO progress (user_id, course, lesson, score, passed, updated_at) VALUES (?, ?, ?, ?, ?, datetime('now'))
                        ON CONFLICT(user_id, course, lesson) DO UPDATE SET score = MAX(score, excluded.score),
                        passed = MAX(passed, excluded.passed), updated_at = datetime('now')`)
        .bind(uid, c.slug, n, g.score, passed).run();
      return json({ ...g, passed: !!passed });
    }
  }

  if ((m = p.match(/^\/api\/courses\/(\w+)\/exam$/))) {
    const denied = needAuth(); if (denied) return denied;
    const c = courseBySlug(m[1]);
    if (!c) return fail("not found", 404);
    const done = await passedLessons(db, uid, c.slug);
    const ready = c.lessons.every((_, i) => done.includes(i + 1));
    const att = await db.prepare("SELECT COUNT(*) AS n FROM exam_attempts WHERE user_id = ? AND course = ? AND created_at > datetime('now', '-1 day')").bind(uid, c.slug).first();
    const left = Math.max(0, EXAM_ATTEMPTS_PER_DAY - att.n);
    const cert = await db.prepare("SELECT code FROM certificates WHERE user_id = ? AND kind = 'course' AND ref = ? AND revoked = 0").bind(uid, c.slug).first();
    if (method === "GET") return json({ course: courseCard(c), ready, attemptsLeft: left, certificate: cert ? cert.code : null, questions: ready ? stripQuiz(c.exam) : [] });
    if (method === "POST") {
      if (!ready) return fail("finish all lessons first", 403);
      if (!left) return fail("no attempts left today", 429);
      const fullName = clip(body.full_name, 100);
      if (fullName.length < 3) return fail("full name required");
      const g = grade(c.exam, body.answers);
      const passed = g.score >= PASS_SCORE;
      await db.prepare("INSERT INTO exam_attempts (user_id, course, score, passed) VALUES (?, ?, ?, ?)").bind(uid, c.slug, g.score, passed ? 1 : 0).run();
      let code = cert ? cert.code : null, programs = [];
      if (passed && !code) {
        code = certCode();
        await db.prepare("INSERT INTO certificates (code, user_id, kind, ref, title, full_name, hours, score) VALUES (?, ?, 'course', ?, ?, ?, ?, ?)")
          .bind(code, uid, c.slug, c.title, fullName, c.hours, Math.round(g.score * 100)).run();
        programs = await issueProgramCerts(db, env, uid, fullName);
        await notify(env, env.ADMIN_ID, "Выдан сертификат " + code + "\n" + fullName + " — «" + c.title + "», " + Math.round(g.score * 100) + "%" +
          (programs.length ? "\nИ удостоверение: " + programs.join(", ") : "") + "\n" + url.origin + "/cert/" + code);
        await notify(env, uid, "Поздравляем! Вы сдали экзамен по курсу «" + c.title + "». Ваш сертификат: " + url.origin + "/cert/" + code +
          (programs.length ? "\nВам также выдано удостоверение SkillBridge: " + url.origin + "/cert/" + programs[0] : ""));
      }
      return json({ score: g.score, right: g.right, total: g.total, passed, certificate: passed ? code : null, programs, attemptsLeft: left - 1 });
    }
  }

  if ((m = p.match(/^\/api\/certs\/([A-Z0-9-]{6,30})$/)) && method === "GET") {
    const c = await db.prepare("SELECT code, kind, ref, title, full_name, hours, score, revoked, issued_at FROM certificates WHERE code = ?").bind(m[1]).first();
    if (!c) return fail("not found", 404);
    if (c.kind === "program") { const pr = PROGRAMS.find((x) => x.slug === c.ref); c.courses = pr ? pr.courses.map((s) => courseBySlug(s)?.title).filter(Boolean) : []; }
    return json({ certificate: c });
  }

  if (p === "/api/learning" && method === "GET") {
    const denied = needAuth(); if (denied) return denied;
    const { results } = await db.prepare("SELECT course, COUNT(*) AS n FROM progress WHERE user_id = ? AND passed = 1 GROUP BY course").bind(uid).all();
    const certs = await db.prepare("SELECT code, kind, ref, title, score, hours, issued_at FROM certificates WHERE user_id = ? AND revoked = 0 ORDER BY issued_at DESC").bind(uid).all();
    return json({ progress: results.map((r) => ({ ...courseCard(courseBySlug(r.course) || { slug: r.course, lessons: [], exam: [] }), passed: r.n })).filter((x) => x.title),
      certificates: certs.results });
  }
  return null;
}

// ---------- API ----------
async function api(request, env, url, uid) {
  const p = url.pathname;
  const method = request.method;
  let body = {};
  if (method === "POST") {
    if (!(request.headers.get("content-type") || "").includes("application/json")) return fail("bad content type", 415);
    try { body = await request.json(); } catch (e) { return fail("bad json"); }
  }
  const db = env.DB;
  const needAuth = () => (uid ? null : fail("login required", 401));

  const learned = await learningApi(request, env, url, uid, body);
  if (learned) return learned;

  if (p === "/api/me" && method === "GET") {
    if (!uid) return json({ user: null });
    const user = await db.prepare("SELECT id, first_name, last_name, username, photo_url FROM users WHERE id = ?").bind(uid).first();
    if (!user) return json({ user: null });
    const profile = await db.prepare("SELECT * FROM profiles WHERE user_id = ?").bind(uid).first();
    return json({ user, profile: profile || null, isAdmin: isAdmin(env, uid) });
  }

  if (p === "/api/specialists" && method === "GET") {
    const { results } = await db
      .prepare("SELECT user_id, name, title, category, skills, price, city, portfolio FROM profiles WHERE status = 'approved' ORDER BY updated_at DESC LIMIT 500")
      .all();
    return json({ specialists: results });
  }

  if (p === "/api/profile" && method === "POST") {
    const denied = needAuth(); if (denied) return denied;
    const f = {
      name: clip(body.name, 80), title: clip(body.title, 80), category: clip(body.category, 20),
      skills: clip(body.skills, 1500), price: clip(body.price, 60), city: clip(body.city, 60), portfolio: clip(body.portfolio, 300),
    };
    if (!f.name || !f.title || !f.skills) return fail("name, title and skills are required");
    if (!CATEGORIES.includes(f.category)) return fail("bad category");
    if (f.portfolio && !/^https?:\/\//i.test(f.portfolio)) f.portfolio = "https://" + f.portfolio;
    await db
      .prepare(`INSERT INTO profiles (user_id, name, title, category, skills, price, city, portfolio, status, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending', datetime('now'))
                ON CONFLICT(user_id) DO UPDATE SET name=excluded.name, title=excluded.title, category=excluded.category,
                  skills=excluded.skills, price=excluded.price, city=excluded.city, portfolio=excluded.portfolio,
                  status='pending', updated_at=datetime('now')`)
      .bind(uid, f.name, f.title, f.category, f.skills, f.price, f.city, f.portfolio)
      .run();
    const u = await db.prepare("SELECT * FROM users WHERE id = ?").bind(uid).first();
    await notify(env, env.ADMIN_ID,
      "Анкета специалиста на проверку\n\n" + f.name + " — " + f.title + "\nTelegram: " + who(u) +
      "\nНавыки: " + f.skills + (f.price ? "\nЦена: " + f.price : "") + (f.city ? "\nГород: " + f.city : "") +
      (f.portfolio ? "\nРаботы: " + f.portfolio : "") + "\n\nОдобрить: " + url.origin + "/cabinet");
    const profile = await db.prepare("SELECT * FROM profiles WHERE user_id = ?").bind(uid).first();
    return json({ profile });
  }

  if (p === "/api/leads" && method === "POST") {
    if (body.website) return json({ ok: true });                      // ловушка для спам-ботов
    const kind = body.kind === "b2b" ? "b2b" : "learn";
    const f = {
      name: clip(body.name, 80), contact: clip(body.contact, 120), track: clip(body.track, 20), level: clip(body.level, 40),
      english: clip(body.english, 20), city: clip(body.city, 60), company: clip(body.company, 120), message: clip(body.message, 1500),
    };
    if (!f.name || !f.contact) return fail("name and contact are required");
    if (f.track && !/^[a-z0-9_]{1,30}$/.test(f.track)) return fail("bad track");
    const trackLabel = clip(body.track_name, 80) || TRACK_RU[f.track] || f.track;
    const r = await db
      .prepare("INSERT INTO leads (kind, name, contact, track, level, english, city, company, message) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")
      .bind(kind, f.name, f.contact, f.track, f.level, f.english, f.city, f.company, f.message)
      .run();
    const head = kind === "b2b" ? "Заявка от компании" : "Запись на обучение";
    await notify(env, env.ADMIN_ID,
      head + " №" + r.meta.last_row_id + "\n\nИмя: " + f.name + "\nКонтакт: " + f.contact +
      (f.company ? "\nКомпания: " + f.company : "") + (f.track ? "\nНаправление: " + trackLabel : "") +
      (f.level ? "\nУровень: " + f.level : "") + (f.english ? "\nАнглийский: " + f.english : "") +
      (f.city ? "\nГород: " + f.city : "") + (f.message ? "\nКомментарий: " + f.message : "") + "\n\n" + url.origin + "/cabinet");
    return json({ ok: true });
  }

  if (p === "/api/orders" && method === "POST") {
    const denied = needAuth(); if (denied) return denied;
    const task = clip(body.task, 3000), contact = clip(body.contact, 120);
    if (!task || !contact) return fail("task and contact are required");
    let specId = body.specialist_id ? Number(body.specialist_id) : null;
    let spec = null;
    if (specId) {
      spec = await db.prepare("SELECT * FROM profiles WHERE user_id = ? AND status = 'approved'").bind(specId).first();
      if (!spec) return fail("specialist not found", 404);
      if (specId === uid) return fail("cannot order from yourself");
    }
    const r = await db
      .prepare("INSERT INTO orders (client_id, specialist_id, task, budget, deadline, contact) VALUES (?, ?, ?, ?, ?, ?)")
      .bind(uid, specId, task, clip(body.budget, 60), clip(body.deadline, 60), contact)
      .run();
    const id = r.meta.last_row_id;
    const client = await db.prepare("SELECT * FROM users WHERE id = ?").bind(uid).first();
    const text = "Заказ №" + id + "\n\nКлиент: " + who(client) + "\nКонтакт: " + contact +
      "\nСпециалист: " + (spec ? spec.name + " — " + spec.title : "не выбран, нужно подобрать") +
      "\nЗадача: " + task + (body.budget ? "\nБюджет: " + clip(body.budget, 60) : "") +
      (body.deadline ? "\nСрок: " + clip(body.deadline, 60) : "") + "\n\nКабинет: " + url.origin + "/cabinet";
    await notify(env, env.ADMIN_ID, "Новый заказ на SkillBridge\n" + text);
    if (spec && String(spec.user_id) !== String(env.ADMIN_ID)) await notify(env, spec.user_id, "Вам новый заказ на SkillBridge\n" + text);
    return json({ id });
  }

  if (p === "/api/orders" && method === "GET") {
    const denied = needAuth(); if (denied) return denied;
    const mine = await db
      .prepare(`SELECT o.*, p.name AS specialist_name, p.title AS specialist_title FROM orders o
                LEFT JOIN profiles p ON p.user_id = o.specialist_id WHERE o.client_id = ? ORDER BY o.id DESC LIMIT 200`)
      .bind(uid).all();
    const forme = await db
      .prepare(`SELECT o.*, u.first_name AS client_first, u.last_name AS client_last, u.username AS client_username FROM orders o
                JOIN users u ON u.id = o.client_id WHERE o.specialist_id = ? ORDER BY o.id DESC LIMIT 200`)
      .bind(uid).all();
    return json({ mine: mine.results, forme: forme.results });
  }

  let m = p.match(/^\/api\/orders\/(\d+)\/status$/);
  if (m && method === "POST") {
    const denied = needAuth(); if (denied) return denied;
    const id = Number(m[1]), status = String(body.status || "");
    if (!STATUSES.includes(status)) return fail("bad status");
    const o = await db.prepare("SELECT * FROM orders WHERE id = ?").bind(id).first();
    if (!o) return fail("not found", 404);
    const admin = isAdmin(env, uid);
    const asSpec = o.specialist_id === uid, asClient = o.client_id === uid;
    const allowed =
      admin ||
      (asSpec && ((o.status === "new" && (status === "accepted" || status === "declined")) || (o.status === "accepted" && status === "done"))) ||
      (asClient && (o.status === "new" || o.status === "accepted") && status === "cancelled");
    if (!allowed) return fail("not allowed", 403);
    await db.prepare("UPDATE orders SET status = ? WHERE id = ?").bind(status, id).run();
    const line = "Заказ №" + id + ": статус «" + STATUS_RU[status] + "»\n" + clip(o.task, 200) + "\n\n" + url.origin + "/cabinet";
    if (!asClient) await notify(env, o.client_id, line);
    if (o.specialist_id && !asSpec && status === "cancelled") await notify(env, o.specialist_id, line);
    if (!admin && (status === "declined" || status === "cancelled")) await notify(env, env.ADMIN_ID, line);
    return json({ ok: true });
  }

  // ----- admin -----
  if (p.startsWith("/api/admin")) {
    if (!isAdmin(env, uid)) return fail("admin only", 403);

    if (p === "/api/admin" && method === "GET") {
      const pending = await db
        .prepare(`SELECT p.*, u.username, u.first_name, u.last_name FROM profiles p JOIN users u ON u.id = p.user_id
                  WHERE p.status = 'pending' ORDER BY p.updated_at`)
        .all();
      const approved = await db.prepare("SELECT user_id, name, title FROM profiles WHERE status = 'approved' ORDER BY name").all();
      const orders = await db
        .prepare(`SELECT o.*, p.name AS specialist_name, u.first_name AS client_first, u.last_name AS client_last, u.username AS client_username
                  FROM orders o JOIN users u ON u.id = o.client_id LEFT JOIN profiles p ON p.user_id = o.specialist_id
                  ORDER BY o.id DESC LIMIT 200`)
        .all();
      const counts = await db
        .prepare(`SELECT (SELECT COUNT(*) FROM users) AS users, (SELECT COUNT(*) FROM profiles WHERE status='approved') AS specialists,
                         (SELECT COUNT(*) FROM orders) AS orders, (SELECT COUNT(*) FROM leads WHERE status='new') AS leads,
                         (SELECT COUNT(*) FROM certificates WHERE revoked = 0) AS certs`)
        .first();
      const leads = await db.prepare("SELECT * FROM leads ORDER BY id DESC LIMIT 200").all();
      const certs = await db.prepare(`SELECT c.code, c.kind, c.title, c.full_name, c.score, c.revoked, c.issued_at, u.username FROM certificates c
                                       LEFT JOIN users u ON u.id = c.user_id ORDER BY c.issued_at DESC LIMIT 100`).all();
      return json({ pending: pending.results, approved: approved.results, orders: orders.results, leads: leads.results, certs: certs.results, counts });
    }

    m = p.match(/^\/api\/admin\/profiles\/(\d+)$/);
    if (m && method === "POST") {
      const id = Number(m[1]), status = String(body.status || "");
      if (!["approved", "rejected", "pending"].includes(status)) return fail("bad status");
      await db.prepare("UPDATE profiles SET status = ? WHERE user_id = ?").bind(status, id).run();
      if (status === "approved") await notify(env, id, "Ваша анкета на SkillBridge одобрена и опубликована в каталоге.\n" + url.origin);
      if (status === "rejected") await notify(env, id, "Ваша анкета на SkillBridge не прошла проверку. Дополните её и отправьте снова: " + url.origin + "/cabinet");
      return json({ ok: true });
    }

    m = p.match(/^\/api\/admin\/certs\/([A-Z0-9-]{6,30})$/);
    if (m && method === "POST") {
      await db.prepare("UPDATE certificates SET revoked = ? WHERE code = ?").bind(body.revoked ? 1 : 0, m[1]).run();
      return json({ ok: true });
    }

    m = p.match(/^\/api\/admin\/leads\/(\d+)$/);
    if (m && method === "POST") {
      const status = String(body.status || "");
      if (!["new", "contacted", "archived"].includes(status)) return fail("bad status");
      await db.prepare("UPDATE leads SET status = ? WHERE id = ?").bind(status, Number(m[1])).run();
      return json({ ok: true });
    }

    m = p.match(/^\/api\/admin\/orders\/(\d+)$/);
    if (m && method === "POST") {
      const id = Number(m[1]), specId = Number(body.specialist_id);
      const spec = await db.prepare("SELECT * FROM profiles WHERE user_id = ? AND status = 'approved'").bind(specId).first();
      if (!spec) return fail("specialist not found", 404);
      const o = await db.prepare("SELECT * FROM orders WHERE id = ?").bind(id).first();
      if (!o) return fail("not found", 404);
      await db.prepare("UPDATE orders SET specialist_id = ?, status = 'new' WHERE id = ?").bind(specId, id).run();
      await notify(env, specId, "Вам назначен заказ №" + id + " на SkillBridge\n\nЗадача: " + o.task + "\n\n" + url.origin + "/cabinet");
      return json({ ok: true });
    }
  }

  return fail("not found", 404);
}

// ---------- router ----------
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const p = url.pathname;
    if (!env.DB) return new Response("Не подключена база D1 (binding DB). См. инструкцию.", { status: 500 });
    await ensureSchema(env);
    const uid = await readSession(request, env);

    if (p === "/auth/telegram") {
      const tg = env.BOT_TOKEN ? await verifyTelegram(url.searchParams, env.BOT_TOKEN) : null;
      let next = url.searchParams.get("next") || "/cabinet";
      if (!next.startsWith("/") || next.startsWith("//")) next = "/cabinet";
      if (!tg) return new Response(null, { status: 302, headers: { Location: "/cabinet?login=failed" } });
      await env.DB
        .prepare(`INSERT INTO users (id, first_name, last_name, username, photo_url) VALUES (?, ?, ?, ?, ?)
                  ON CONFLICT(id) DO UPDATE SET first_name=excluded.first_name, last_name=excluded.last_name,
                  username=excluded.username, photo_url=excluded.photo_url`)
        .bind(tg.id, tg.first_name, tg.last_name, tg.username, tg.photo_url)
        .run();
      return new Response(null, { status: 302, headers: { Location: next, "Set-Cookie": cookie(await makeSession(env, tg.id), 30 * 86400) } });
    }

    if (p === "/logout" && request.method === "POST") {
      return new Response(null, { status: 303, headers: { Location: "/", "Set-Cookie": cookie("", 0) } });
    }

    if (p.startsWith("/api/")) {
      try { return await api(request, env, url, uid); }
      catch (e) { return fail("server error", 500); }
    }

    const PAGES = /^\/(cabinet|courses|verify|course\/\w+|learn\/\w+\/\d+|exam\/\w+|cert\/[A-Z0-9-]{6,30})?$/;
    if (request.method === "GET" && PAGES.test(p)) {
      const html = PAGE.split("__BOT_USERNAME__").join(String(env.BOT_USERNAME || "").replace(/[^A-Za-z0-9_]/g, ""));
      return new Response(html, {
        headers: {
          "content-type": "text/html; charset=utf-8",
          "cache-control": "no-store",
          "x-frame-options": "SAMEORIGIN",
          "referrer-policy": "strict-origin-when-cross-origin",
        },
      });
    }
    return new Response("Not found", { status: 404 });
  },
};
