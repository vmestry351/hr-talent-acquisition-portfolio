#!/usr/bin/env python3
"""
Mini ATS - a rule-based Applicant Tracking System prototype.

Reads resumes (PDF, DOCX, TXT), runs a series of test cases against a job
profile, scores each candidate out of 100, decides SELECTED / HOLD / REJECTED
and stores everything in an Excel workbook.

Usage:
    python ats.py --list-jobs
    python ats.py --job junior_data_analyst --resumes sample_resumes
    python ats.py --job sales_executive --resumes my_resume.pdf --output results.xlsx

This is a screening AID, not a replacement for human judgement. Every
decision should be reviewed by a recruiter before a candidate is rejected.
"""
import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("Missing package. Run: pip install -r requirements.txt")

BASE_DIR = Path(__file__).resolve().parent
EDU_RANK = {"diploma": 1, "bachelor": 2, "master": 3, "phd": 4}
EDU_PATTERNS = [
    ("phd", [r"\bph\.?\s?d\b", r"\bdoctorate\b"]),
    ("master", [r"\bmba\b", r"\bpgdm\b", r"\bm\.?(?:com|sc|tech|ca)\b", r"\bm\.e\b",
                r"\bmasters?\s+(?:of|in|degree)\b", r"\bmaster's\b", r"\bpost\s?graduate\b", r"\bpostgraduate\b"]),
    ("bachelor", [r"\bb\.?\s?(?:com|sc|tech|ca)\b", r"\bbba\b", r"\bbms\b", r"\bbbi\b", r"\bb\.e\b",
                  r"\bb\.a\b", r"\bba\b", r"\bbachelor(?:'?s)?\b", r"\bgraduat(?:e|ion|ed)\b",
                  r"\bundergraduate\b"]),
    ("diploma", [r"\bdiploma\b"]),
]
SECTION_HEADS = {
    "experience": ["experience", "work experience", "professional experience", "employment",
                   "employment history", "work history", "internship", "internships",
                   "internship experience"],
    "education": ["education", "academic background", "academics", "qualification", "qualifications"],
    "skills": ["skills", "technical skills", "key skills", "core skills", "skill set"],
    "projects": ["projects", "academic projects", "personal projects", "project work"],
    "certifications": ["certifications", "certificates", "courses", "training"],
    "summary": ["summary", "profile", "objective", "career objective", "about me"],
}
MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
          "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
_MONTH = r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
_DATE = rf"(?:{_MONTH}\.?,?\s+\d{{4}}|\d{{1,2}}/\d{{4}}|\d{{4}})"
DATE_RANGE = re.compile(
    rf"({_DATE})\s*(?:-|\u2013|\u2014|to|till)\s*({_DATE}|present|current|now|till date|ongoing)",
    re.IGNORECASE)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
PHONE_RE = re.compile(r"(?<!\d)(?:\+?91[\s-]?)?([6-9]\d{4}[\s-]?\d{5})(?!\d)")
NAME_STOP = {"resume", "curriculum", "vitae", "cv", "profile", "summary", "objective", "contact", "address"}


# ----------------------------------------------------------------------------
# 1. Reading resumes
# ----------------------------------------------------------------------------
def read_resume(path: Path) -> str:
    """Return the text of a PDF, DOCX or TXT resume ('' if unreadable)."""
    suffix = path.suffix.lower()
    try:
        if suffix == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(str(path))
            return "\n".join((page.extract_text() or "") for page in reader.pages)
        if suffix == ".docx":
            import docx
            document = docx.Document(str(path))
            parts = [p.text for p in document.paragraphs]
            for table in document.tables:
                for row in table.rows:
                    parts.extend(cell.text for cell in row.cells)
            return "\n".join(parts)
        if suffix == ".txt":
            return path.read_text(encoding="utf-8", errors="ignore")
    except ImportError as exc:
        sys.exit(f"Missing package ({exc.name}). Run: pip install -r requirements.txt")
    except Exception as exc:  # corrupted or protected file
        print(f"  ! Could not read {path.name}: {exc}")
    return ""


# ----------------------------------------------------------------------------
# 2. Extracting information
# ----------------------------------------------------------------------------
def find_term(term: str, text: str) -> bool:
    pattern = rf"(?<![a-z0-9]){re.escape(term.lower())}(?![a-z0-9])"
    return re.search(pattern, text) is not None


def match_skills(skill_names, skills_db, text):
    found, missing = [], []
    for name in skill_names:
        aliases = skills_db.get(name, [name])
        (found if any(find_term(a, text) for a in aliases) else missing).append(name)
    return found, missing


def extract_name(text: str, fallback: str) -> str:
    for line in [l.strip() for l in text.splitlines() if l.strip()][:8]:
        words = line.replace(",", " ").split()
        if not 2 <= len(words) <= 4 or "@" in line or any(ch.isdigit() for ch in line):
            continue
        if any(w.lower().strip(".:") in NAME_STOP for w in words):
            continue
        if all(re.fullmatch(r"[A-Za-z.\-']+", w) for w in words):
            return " ".join(w.capitalize() if w.isupper() else w for w in words)
    return fallback


def extract_contact(text: str):
    email = EMAIL_RE.search(text)
    phone = PHONE_RE.search(text)
    digits = re.sub(r"\D", "", phone.group(1)) if phone else ""
    return (email.group(0).lower() if email else ""), digits


def detect_education(text_lower: str):
    for level, patterns in EDU_PATTERNS:  # ordered highest to lowest
        if any(re.search(p, text_lower) for p in patterns):
            return level
    return ""


def split_sections(text: str) -> dict:
    sections, current = {}, None
    for line in text.splitlines():
        cleaned = re.sub(r"[^a-z ]", "", line.lower()).strip()
        matched = None
        if cleaned and len(cleaned) <= 30:
            for key, heads in SECTION_HEADS.items():
                if cleaned in heads:
                    matched = key
                    break
        if matched:
            current = matched
            sections.setdefault(current, [])
        elif current:
            sections[current].append(line)
    return {k: "\n".join(v) for k, v in sections.items()}


def _to_index(token: str, is_end: bool):
    t = token.lower().strip()
    today = date.today()
    if t in ("present", "current", "now", "till date", "ongoing"):
        return today.year * 12 + today.month
    m = re.fullmatch(r"(\d{1,2})/(\d{4})", t)
    if m:
        return int(m.group(2)) * 12 + int(m.group(1))
    m = re.fullmatch(rf"({_MONTH})\.?,?\s+(\d{{4}})", t)
    if m:
        return int(m.group(2)) * 12 + MONTHS[m.group(1)[:3]]
    m = re.fullmatch(r"(\d{4})", t)
    if m:
        return int(m.group(1)) * 12 + (12 if is_end else 1)
    return None


def estimate_experience(text: str, sections: dict):
    """Return (years, source). Uses stated years, else dates in the experience section."""
    lower = text.lower()
    stated = 0.0
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*\+?\s*(?:years?|yrs?)\b", lower):
        window = lower[max(0, m.start() - 40): m.end() + 60]
        if "experience" in window:
            stated = max(stated, float(m.group(1)))
    for m in re.finditer(r"(\d+)\s*months?\b", lower):
        window = lower[max(0, m.start() - 40): m.end() + 60]
        if "experience" in window:
            stated = max(stated, int(m.group(1)) / 12)

    dated, intervals = 0.0, []
    exp_text = sections.get("experience", "")
    now_idx = date.today().year * 12 + date.today().month
    for m in DATE_RANGE.finditer(exp_text):
        start, end = _to_index(m.group(1), False), _to_index(m.group(2), True)
        if start and end:
            end = min(end, now_idx)
            if end >= start:
                intervals.append((start, end))
    if intervals:
        intervals.sort()
        merged = [list(intervals[0])]
        for s, e in intervals[1:]:
            if s <= merged[-1][1] + 1:
                merged[-1][1] = max(merged[-1][1], e)
            else:
                merged.append([s, e])
        dated = sum(e - s + 1 for s, e in merged) / 12

    if stated >= dated and stated > 0:
        return round(stated, 1), "stated in resume"
    if dated > 0:
        return round(dated, 1), "calculated from dates"
    return 0.0, "no experience found"


# ----------------------------------------------------------------------------
# 3. Test cases
# ----------------------------------------------------------------------------
WEIGHTS = {"contact": 5, "education": 10, "experience": 15, "must_have": 35,
           "nice_to_have": 10, "title": 10, "quality": 10, "location": 5}


def t(name, passed, score, maximum, knockout, detail):
    return {"test": name, "passed": passed, "score": round(score, 1), "max": maximum,
            "knockout": knockout, "detail": detail}


def run_tests(info: dict, job: dict, settings: dict, skills_db: dict, text_lower: str):
    tests = []

    # T1 contact details
    has_email, has_phone = bool(info["email"]), bool(info["phone"])
    score = WEIGHTS["contact"] if (has_email and has_phone) else (3 if (has_email or has_phone) else 0)
    tests.append(t("Contact details", has_email or has_phone, score, WEIGHTS["contact"], True,
                   f"Email: {'yes' if has_email else 'no'}; Phone: {'yes' if has_phone else 'no'}"))

    # T2 education
    required = job.get("min_education", "bachelor")
    level = info["education"]
    ok = bool(level) and EDU_RANK[level] >= EDU_RANK[required]
    tests.append(t("Minimum education", ok, WEIGHTS["education"] if ok else 0, WEIGHTS["education"], True,
                   f"Found: {level or 'none'}; required: {required}"))

    # T3 experience
    exp, src = info["experience"], info["experience_source"]
    lo, hi = job.get("min_experience", 0), job.get("max_experience", 99)
    if exp < lo:
        ok, score, note = False, WEIGHTS["experience"] * exp / lo if lo else 0, f"below minimum of {lo} yrs"
    elif exp > hi:
        ok, score, note = True, WEIGHTS["experience"] * 0.7, f"above maximum of {hi} yrs (possibly over-qualified)"
        info["overqualified"] = True
    else:
        ok, score, note = True, WEIGHTS["experience"], f"within {lo}-{hi} yrs"
    tests.append(t("Experience range", ok, score, WEIGHTS["experience"], lo > 0,
                   f"{exp} yrs ({src}); {note}"))

    # T4 must-have skills
    must = job.get("must_have_skills", [])
    found, missing = match_skills(must, skills_db, text_lower)
    ratio = len(found) / len(must) if must else 1
    need = job.get("must_have_min_ratio", 0.6)
    tests.append(t("Must-have skills", ratio >= need, WEIGHTS["must_have"] * ratio, WEIGHTS["must_have"], True,
                   f"Matched {len(found)}/{len(must)} (needs {int(need * 100)}%). Missing: {', '.join(missing) or 'none'}"))
    info["must_found"], info["must_missing"] = found, missing

    # T5 nice-to-have skills
    nice = job.get("nice_to_have_skills", [])
    nfound, nmissing = match_skills(nice, skills_db, text_lower)
    nratio = len(nfound) / len(nice) if nice else 1
    tests.append(t("Nice-to-have skills", nratio > 0, WEIGHTS["nice_to_have"] * nratio, WEIGHTS["nice_to_have"], False,
                   f"Matched {len(nfound)}/{len(nice)}: {', '.join(nfound) or 'none'}"))
    info["nice_found"] = nfound

    # T6 role relevance
    hits = [k for k in job.get("title_keywords", []) if find_term(k, text_lower)]
    score = WEIGHTS["title"] * min(len(hits), 2) / 2
    tests.append(t("Role relevance", bool(hits), score, WEIGHTS["title"], False,
                   f"Keywords found: {', '.join(hits[:4]) or 'none'}"))

    # T7 resume quality
    words = len(text_lower.split())
    secs = info["sections"]
    core = sum(1 for k in (("experience", "projects"), ("education",), ("skills",))
               if any(x in secs for x in k))
    min_words = settings.get("min_words", 150)
    score = WEIGHTS["quality"] * (0.5 * min(words / min_words, 1) + 0.5 * core / 3)
    tests.append(t("Resume quality", words >= min_words and core >= 2, score, WEIGHTS["quality"], False,
                   f"{words} words; core sections found: {core}/3"))

    # T8 location
    locs = job.get("locations", [])
    if locs:
        hit = [l for l in locs if find_term(l, text_lower)]
        tests.append(t("Location", bool(hit), WEIGHTS["location"] if hit else 0, WEIGHTS["location"],
                       bool(job.get("location_required")), f"Preferred: {', '.join(locs)}; found: {', '.join(hit) or 'none'}"))
    else:
        tests.append(t("Location", True, WEIGHTS["location"], WEIGHTS["location"], False, "No location requirement"))
    return tests


def decide(tests, settings, overqualified=False):
    total = round(sum(x["score"] for x in tests), 1)
    failed = [x for x in tests if x["knockout"] and not x["passed"]]
    if failed:
        reasons = "; ".join(f"FAILED {x['test']} ({x['detail']})" for x in failed)
        return total, "REJECTED", reasons
    weak = sorted(tests, key=lambda x: x["max"] - x["score"], reverse=True)[:2]
    weak_txt = ", ".join(f"{x['test']} ({x['score']}/{x['max']})" for x in weak if x["max"] - x["score"] > 0.05)
    if overqualified and total >= settings.get("hold_threshold", 55):
        return total, "HOLD", (f"Score {total}, but experience is above the role's maximum (possibly over-qualified). "
                               "Recruiter to confirm interest in this level, or consider a senior role.")
    if total >= settings.get("select_threshold", 70):
        return total, "SELECTED", f"Passed all knockout tests; score {total}. " + (f"Weakest: {weak_txt}" if weak_txt else "")
    if total >= settings.get("hold_threshold", 55):
        return total, "HOLD", f"Passed knockouts but score {total} is below {settings.get('select_threshold', 70)}. Weakest: {weak_txt}"
    return total, "REJECTED", f"Score {total} is below {settings.get('hold_threshold', 55)}. Weakest: {weak_txt}"


# ----------------------------------------------------------------------------
# 4. Processing one resume
# ----------------------------------------------------------------------------
def process_resume(path: Path, job: dict, settings: dict, skills_db: dict) -> dict:
    text = read_resume(path)
    result = {"file": path.name, "job": job["title"], "processed_on": datetime.now().strftime("%Y-%m-%d %H:%M")}
    if len(text.strip()) < 50:
        result.update(name=path.stem, email="", phone="", education="", experience=0, must_found=[], must_missing=[],
                      nice_found=[], score=0, decision="MANUAL REVIEW", tests=[],
                      reasons="Could not read text (scanned image, empty or protected file). Review manually.")
        return result
    lower = text.lower()
    sections = split_sections(text)
    email, phone = extract_contact(text)
    experience, source = estimate_experience(text, sections)
    info = {"email": email, "phone": phone, "education": detect_education(lower), "experience": experience,
            "experience_source": source, "sections": sections}
    tests = run_tests(info, job, settings, skills_db, lower)
    score, decision, reasons = decide(tests, settings, info.get("overqualified", False))
    result.update(name=extract_name(text, path.stem), email=email, phone=phone, education=info["education"] or "not found",
                  experience=experience, must_found=info["must_found"], must_missing=info["must_missing"],
                  nice_found=info["nice_found"], score=score, decision=decision, tests=tests, reasons=reasons)
    return result


# ----------------------------------------------------------------------------
# 5. Excel storage
# ----------------------------------------------------------------------------
CAND_HEAD = ["Candidate ID", "Processed On", "Job", "Resume File", "Name", "Email", "Phone", "Education",
             "Experience (yrs)", "Must-have Skills Matched", "Must-have Skills Missing",
             "Nice-to-have Skills Matched", "Score (/100)", "Decision", "Reasons / Notes"]
TEST_HEAD = ["Candidate ID", "Name", "Test", "Result", "Score", "Max", "Knockout test?", "Detail"]
NAVY = PatternFill("solid", fgColor="1F3864")
HFONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BFONT = Font(name="Arial", size=10)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header(ws, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, i, h)
        c.font, c.fill, c.border = HFONT, NAVY, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"


def open_workbook(path: Path):
    if path.exists():
        wb = load_workbook(path)
        if "Candidates" in wb.sheetnames and "Test_Results" in wb.sheetnames:
            return wb
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidates"
    _style_header(ws, CAND_HEAD, [13, 17, 26, 24, 22, 28, 13, 11, 11, 28, 28, 28, 10, 15, 70])
    ws.auto_filter.ref = "A1:O1"
    dv = DataValidation(type="list", formula1='"SELECTED,HOLD,REJECTED,MANUAL REVIEW"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("N2:N5000")
    for text, color in (("SELECTED", "C6EFCE"), ("HOLD", "FFEB9C"), ("REJECTED", "FFC7CE"), ("MANUAL REVIEW", "D9D9D9")):
        ws.conditional_formatting.add("N2:N5000", FormulaRule(
            formula=[f'$N2="{text}"'], fill=PatternFill("solid", bgColor=color, fgColor=color)))
    tr = wb.create_sheet("Test_Results")
    _style_header(tr, TEST_HEAD, [13, 22, 22, 9, 8, 8, 14, 80])
    tr.auto_filter.ref = "A1:H1"
    return wb


def existing_keys(ws):
    keys = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            keys.add((row[2], str(row[5] or "").lower(), str(row[6] or "")))
    return keys


def next_id(ws):
    top = 0
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        m = re.fullmatch(r"ATS-(\d+)", str(row[0] or ""))
        if m:
            top = max(top, int(m.group(1)))
    return top + 1


def rebuild_summary(wb, jobs):
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws = wb.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "ATS Summary (updates automatically from the Candidates sheet)"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="1F3864")
    heads = ["Job", "Total", "Selected", "Hold", "Rejected", "Manual Review", "Average Score", "Selection Rate"]
    for i, h in enumerate(heads, 1):
        c = ws.cell(3, i, h)
        c.font, c.fill, c.border = HFONT, NAVY, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    titles = [j["title"] for j in jobs.values()]
    for r, title in enumerate(titles, 4):
        ws.cell(r, 1, title)
        ws.cell(r, 2, f'=COUNTIFS(Candidates!$C:$C,$A{r})')
        for col, label in zip("CDEF", ("SELECTED", "HOLD", "REJECTED", "MANUAL REVIEW")):
            ws[f"{col}{r}"] = f'=COUNTIFS(Candidates!$C:$C,$A{r},Candidates!$N:$N,"{label}")'
        ws.cell(r, 7, f'=IFERROR(AVERAGEIFS(Candidates!$M:$M,Candidates!$C:$C,$A{r}),0)')
        ws.cell(r, 8, f'=IF(B{r}=0,0,C{r}/B{r})')
    last = 3 + len(titles)
    tot = last + 1
    ws.cell(tot, 1, "Total")
    for col in "BCDEF":
        ws[f"{col}{tot}"] = f"=SUM({col}4:{col}{last})"
    ws.cell(tot, 7, "=IFERROR(AVERAGE(Candidates!$M$2:$M$5000),0)")
    ws.cell(tot, 8, f"=IF(B{tot}=0,0,C{tot}/B{tot})")
    for r in range(4, tot + 1):
        for c in range(1, 9):
            cell = ws.cell(r, c)
            cell.border, cell.font = BOX, Font(name="Arial", size=10, bold=(r == tot))
            if c > 1:
                cell.alignment = Alignment(horizontal="center")
        ws.cell(r, 7).number_format = "0.0"
        ws.cell(r, 8).number_format = "0%"
    ws.column_dimensions["A"].width = 32
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 14


def save_results(results, jobs, out_path: Path):
    wb = open_workbook(out_path)
    cand, tests_ws = wb["Candidates"], wb["Test_Results"]
    seen, counter, added, skipped = existing_keys(cand), next_id(cand), [], []
    for r in results:
        key = (r["job"], r["email"], r["phone"])
        if (r["email"] or r["phone"]) and key in seen:
            skipped.append(r)
            continue
        seen.add(key)
        cid = f"ATS-{counter:04d}"
        counter += 1
        r["id"] = cid
        row = [cid, r["processed_on"], r["job"], r["file"], r["name"], r["email"], r["phone"], r["education"],
               r["experience"], ", ".join(r["must_found"]), ", ".join(r["must_missing"]), ", ".join(r["nice_found"]),
               r["score"], r["decision"], r["reasons"]]
        cand.append(row)
        rr = cand.max_row
        for c in range(1, len(row) + 1):
            cell = cand.cell(rr, c)
            cell.font, cell.border = BFONT, BOX
            cell.alignment = Alignment(vertical="top", wrap_text=c in (10, 11, 12, 15),
                                       horizontal="center" if c in (1, 9, 13, 14) else None)
        for tst in r["tests"]:
            tests_ws.append([cid, r["name"], tst["test"], "PASS" if tst["passed"] else "FAIL", tst["score"], tst["max"],
                             "Yes" if tst["knockout"] else "No", tst["detail"]])
            tr = tests_ws.max_row
            for c in range(1, 9):
                tests_ws.cell(tr, c).font, tests_ws.cell(tr, c).border = BFONT, BOX
            tests_ws.cell(tr, 4).font = Font(name="Arial", size=10, bold=True,
                                             color="006100" if tst["passed"] else "9C0006")
        added.append(r)
    cand.auto_filter.ref = f"A1:O{max(cand.max_row, 2)}"
    tests_ws.auto_filter.ref = f"A1:H{max(tests_ws.max_row, 2)}"
    rebuild_summary(wb, jobs)
    wb.save(out_path)
    return added, skipped


# ----------------------------------------------------------------------------
# 6. Command line
# ----------------------------------------------------------------------------
def collect_files(target: Path):
    if target.is_file():
        return [target]
    if target.is_dir():
        return sorted(p for p in target.iterdir() if p.suffix.lower() in (".pdf", ".docx", ".txt"))
    return []


def main():
    ap = argparse.ArgumentParser(description="Mini ATS: screen resumes against a job profile and save results to Excel.")
    ap.add_argument("--job", help="Job key from jobs.json (use --list-jobs to see them)")
    ap.add_argument("--resumes", default="resumes", help="A resume file or a folder of resumes (PDF/DOCX/TXT)")
    ap.add_argument("--output", default="ats_results.xlsx", help="Excel file to create or update")
    ap.add_argument("--config", default=str(BASE_DIR / "jobs.json"), help="Path to jobs.json")
    ap.add_argument("--list-jobs", action="store_true", help="Show available jobs and exit")
    ap.add_argument("--dry-run", action="store_true", help="Show results without writing the Excel file")
    args = ap.parse_args()

    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    jobs, settings, skills_db = config["jobs"], config.get("settings", {}), config.get("skills", {})

    if args.list_jobs or not args.job:
        print("Available jobs:")
        for key, j in jobs.items():
            print(f"  {key:22s} {j['title']}  (must-have: {', '.join(j['must_have_skills'])})")
        if not args.job:
            return
    if args.job not in jobs:
        sys.exit(f"Unknown job '{args.job}'. Use --list-jobs.")

    files = collect_files(Path(args.resumes))
    if not files:
        sys.exit(f"No PDF, DOCX or TXT resumes found at: {args.resumes}")
    job = jobs[args.job]
    print(f"\nScreening {len(files)} resume(s) for: {job['title']}\n")
    results = [process_resume(f, job, settings, skills_db) for f in files]

    if args.dry_run:
        added, skipped = results, []
    else:
        added, skipped = save_results(results, jobs, Path(args.output))

    print(f"{'File':30s} {'Name':22s} {'Score':>6s}  Decision")
    print("-" * 74)
    for r in added:
        print(f"{r['file'][:29]:30s} {r['name'][:21]:22s} {r['score']:>6}  {r['decision']}")
    for r in skipped:
        print(f"{r['file'][:29]:30s} {r['name'][:21]:22s} {'-':>6s}  DUPLICATE (already in Excel, skipped)")
    counts = {d: sum(1 for r in added if r["decision"] == d) for d in ("SELECTED", "HOLD", "REJECTED", "MANUAL REVIEW")}
    print("-" * 74)
    print("  ".join(f"{k}: {v}" for k, v in counts.items()))
    if not args.dry_run:
        print(f"\nSaved to {args.output}  (sheets: Summary, Candidates, Test_Results)")


if __name__ == "__main__":
    main()
