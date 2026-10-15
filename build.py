"""Сборка: dist/skillbridge-worker.js (сайт для Cloudflare) и dist/preview.html (демо).
Курсы берутся из src/courses.txt (формат описан в начале файла)."""
import json, re, pathlib
root = pathlib.Path(__file__).parent

def parse_courses(text):
    courses, programs = [], []
    cur = None; sec = None; q = None
    for raw in text.splitlines():
        line = raw.rstrip()
        if cur is None and line.startswith("#"): continue
        m = re.match(r"^=== (COURSE|PROGRAM) (\w+)$", line)
        if m:
            cur = {"slug": m.group(2)}
            if m.group(1) == "COURSE": cur.update(lessons=[], exam=[]); courses.append(cur)
            else: programs.append(cur)
            sec = None; q = None; continue
        if cur is None: continue
        m = re.match(r"^--- LESSON (.+)$", line)
        if m:
            sec = {"title": m.group(1).strip(), "body": [], "quiz": []}; cur["lessons"].append(sec); q = None; continue
        if line == "--- EXAM":
            sec = {"title": "EXAM", "body": [], "quiz": cur["exam"]}; q = None; continue
        if line.startswith("?? "):
            q = {"q": line[3:].strip(), "options": [], "correct": None}; sec["quiz"].append(q); continue
        if q is not None:
            if line.startswith("+ ") or line.startswith("- "):
                if line.startswith("+ "):
                    assert q["correct"] is None, "два верных ответа: " + q["q"]
                    q["correct"] = len(q["options"])
                q["options"].append(line[2:].strip()); continue
            if not line.strip(): continue
            raise ValueError("непонятная строка после вопроса: " + line)
        if sec is None:
            m = re.match(r"^(\w+): (.*)$", line)
            if m: cur[m.group(1)] = m.group(2).strip()
            continue
        sec["body"].append(line)
    for c in courses:
        c["hours"] = int(c.get("hours", 0))
        for l in c["lessons"]:
            l["body"] = "\n".join(l["body"]).strip()
        for qq in [x for l in c["lessons"] for x in l["quiz"]] + c["exam"]:
            assert qq["correct"] is not None and len(qq["options"]) >= 2, "проверьте вопрос: " + qq["q"]
        assert c["exam"], "нет экзамена в курсе " + c["slug"]
    slugs = {c["slug"] for c in courses}
    for p in programs:
        p["courses"] = [s.strip() for s in p["courses"].split(",")]
        assert all(s in slugs for s in p["courses"]), "неизвестный курс в программе " + p["slug"]
    return courses, programs

(root / "dist").mkdir(exist_ok=True)
courses, programs = parse_courses((root / "src/courses.txt").read_text(encoding="utf-8"))
data = "const COURSES = " + json.dumps(courses, ensure_ascii=False) + ";\nconst PROGRAMS = " + json.dumps(programs, ensure_ascii=False) + ";\n"
srv = (root / "src/server.js").read_text(encoding="utf-8")
page = (root / "src/page.html").read_text(encoding="utf-8")
import base64
page = page.replace("__SIGNATURE__", "data:image/png;base64," + base64.b64encode((root / "src/signature.png").read_bytes()).decode())
(root / "dist/skillbridge-worker.js").write_text(srv + "\n" + data + "const PAGE = " + json.dumps(page, ensure_ascii=False) + ";\n", encoding="utf-8")

head = re.search(r"<head>(.*)</head>", page, re.S).group(1)
body = re.search(r"<body>(.*)</body>", page, re.S).group(1)
head = re.sub(r'<meta charset="utf-8">\s*|<meta name="viewport"[^>]*>\s*', "", head)
demo = "<script>var DEMO_COURSES = " + json.dumps(courses, ensure_ascii=False).replace("</", "<\\/") + "; var DEMO_PROGRAMS = " + json.dumps(programs, ensure_ascii=False) + ";</script>\n"
prev = head.strip() + "\n" + demo + body.strip().replace("var DEMO = false;", "var DEMO = true;", 1)
assert "var DEMO = true;" in prev
(root / "dist/preview.html").write_text(prev + "\n", encoding="utf-8")
print("курсов:", len(courses), "уроков:", sum(len(c["lessons"]) for c in courses),
      "вопросов:", sum(len(l["quiz"]) for c in courses for l in c["lessons"]) + sum(len(c["exam"]) for c in courses), "программ:", len(programs))
