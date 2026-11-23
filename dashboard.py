"""
Exam Timetable Dashboard
------------------------
Streamlit front end for "Optimize Schedule.py". The optimizer file is imported as it is:
this dashboard never edits it, it only calls its functions:

    read_inputs()  ->  build_slots()  ->  solve()  ->  report()

Special requests (fixed date / days / hours, 3-4 h exams, no proctor) are typed in the dashboard.
While solving with requests, the dashboard wraps three of the optimizer's helper functions
(assignment_model, initial_schedule, assign_rooms) so the request rules are added to the same
ILP. The optimizer file on disk is never changed, and the originals are restored after each solve.

Run (from the folder that holds both files):
    streamlit run dashboard.py
"""
import contextlib
import importlib.util
import io
import math
import re
import tempfile
import time as clock_time
from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta
from itertools import combinations
from pathlib import Path

import altair as alt
import pandas as pd
import pulp as pl
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.cell import get_column_letter

HERE = Path(__file__).resolve().parent
OPTIMIZER_FILE = HERE / "Optimize Schedule.py"
STD, LONG = "Up to 2 h (1 slot)", "3–4 h (2 slots)"

st.set_page_config(page_title="Exam Timetable", page_icon="🎓", layout="wide")


# ------------------------------------------------------------------ load the optimizer (unchanged)
@st.cache_resource
def load_optimizer():
    spec = importlib.util.spec_from_file_location("optimize_schedule", OPTIMIZER_FILE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module._dash_orig = (module.assignment_model, module.initial_schedule, module.assign_rooms)
    return module


if not OPTIMIZER_FILE.exists():
    st.error(f'"Optimize Schedule.py" was not found next to the dashboard ({HERE}). '
             "Put both files in the same folder.")
    st.stop()
opt = load_optimizer()
ORIG_MODEL, ORIG_SEED, ORIG_ROOMS = opt._dash_orig
opt.assignment_model, opt.initial_schedule, opt.assign_rooms = opt._dash_orig  # always start clean


def clean_solver_files():
    """Delete the temporary files PuLP/CBC writes while solving (*-pulp.mps, .mst, .sol, .lp).
    They are normally removed after each solve, but stay behind if a solve is interrupted."""
    for folder in {HERE, Path.cwd()}:
        for pattern in ("*-pulp.mps", "*-pulp.mst", "*-pulp.sol", "*-pulp.lp"):
            for f in folder.glob(pattern):
                try:
                    f.unlink()
                except OSError:
                    pass


clean_solver_files()  # removes leftovers from any interrupted solve

# ------------------------------------------------------------------ styling
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Lora:wght@600;700&family=Source+Sans+3:wght@400;600;700&display=swap" rel="stylesheet">
<style>
:root{--maroon:#840132;--maroon-d:#5e0124;--tint:#F6F1F2;--line:#E7DADE;--ink:#2A1C21;--muted:#6E5F65;--lab:#1F6F78;--good:#2E7D4F;--warn:#B7791F;--bad:#B42318}
html,body,[class*="css"]{font-family:"Source Sans 3",system-ui,sans-serif}
.block-container{padding-top:3rem;max-width:1400px}
[data-testid="stSidebar"]{background:var(--tint);border-right:3px solid var(--maroon)}
[data-testid="stSidebar"] h2{color:var(--maroon);font-family:"Lora",Georgia,serif;font-size:1.05rem;border-bottom:1px solid var(--line);padding-bottom:4px}
.band{background:linear-gradient(120deg,#840132,#5E0124);color:#fff;border-radius:20px;padding:22px 26px;margin-bottom:18px;display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap}
.band .k{font-size:.78rem;letter-spacing:.12em;text-transform:uppercase;opacity:.8}
.band h1{font-family:"Lora",Georgia,serif;font-size:2.1rem;margin:4px 0 0;color:#fff;padding:0}
.band p{margin:6px 0 0;opacity:.85;font-size:.95rem;max-width:62ch}
.band .file{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.3);border-radius:999px;padding:6px 14px;font-size:.85rem;white-space:nowrap}
.kpi{min-height:104px;border-radius:18px;padding:14px 16px;border-top:4px solid var(--maroon);background:#fff;box-shadow:0 4px 14px rgba(132,1,50,.10)}
.kpi .v{font-size:clamp(1.25rem,2.1vw,1.9rem);font-weight:700;line-height:1.1;color:var(--maroon);white-space:nowrap;font-variant-numeric:tabular-nums}
.kpi .l{font-size:.82rem;color:var(--muted);margin-top:6px;line-height:1.3}
.kpi.good{border-top-color:var(--good)} .kpi.good .v{color:var(--good)}
.kpi.warn{border-top-color:var(--warn)} .kpi.warn .v{color:var(--warn)}
.kpi.bad{border-top-color:var(--bad)} .kpi.bad .v{color:var(--bad)}
.stTabs [data-baseweb="tab-list"]{gap:4px;border-bottom:2px solid var(--line)}
.stTabs [data-baseweb="tab"]{font-weight:600}
.cal{border-collapse:separate;border-spacing:6px;width:100%;min-width:1250px;table-layout:fixed}
.cal th{font-size:.8rem;font-weight:700;color:var(--maroon);text-align:left;padding:4px;border-bottom:2px solid var(--maroon)}
.cal td{vertical-align:top;padding:0}
.cal td.t{font-size:.78rem;color:var(--muted);white-space:nowrap;width:92px;padding-top:6px;font-weight:600}
.cell{min-height:46px;border-radius:12px;border:1px solid #EFE4E8;padding:4px;display:flex;flex-direction:column;gap:4px;background:#fff}
.chip{border-radius:10px;padding:5px 7px;font-size:.74rem;line-height:1.25;border-left:4px solid;word-break:normal;overflow-wrap:normal;hyphens:auto;color:var(--ink)}
.chip b{font-size:.78rem}
.chip.hall{background:rgba(132,1,50,.08);border-color:var(--maroon)}
.chip.lab{background:rgba(31,111,120,.12);border-color:var(--lab)}
.chip.cont{border-left-style:dashed;opacity:.75}
.chip.hit{outline:2px solid var(--warn)}
.scroll{overflow-x:auto}
.legend span{display:inline-block;margin-right:14px;font-size:.8rem;color:var(--muted)}
.dot{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:5px;vertical-align:middle}
[data-testid="stAppViewContainer"],[data-testid="stMain"],[data-testid="stHeader"]{background:#FAF7F8}
[data-testid="stAppDeployButton"],[data-testid="stMainMenu"]{display:none}
.stButton>button[kind="primary"],[data-testid="stBaseButton-primary"],.stDownloadButton>button[kind="primary"],[data-testid="stBaseButton-primary"]:hover{background:#840132!important;border-color:#840132!important;color:#fff!important}
[data-testid="stSlider"] [role="slider"]{background:#840132!important;border-color:#840132!important}
[data-testid="stSlider"] [data-testid="stThumbValue"]{color:#840132!important}
[data-baseweb="slider"]>div>div>div:first-child{background:#840132!important}
.stTabs [aria-selected="true"]{color:#840132!important}
.stTabs [data-baseweb="tab-highlight"]{background:#840132!important}
a{color:#840132}
</style>
""", unsafe_allow_html=True)


def kpi(col, value, label, tone=""):
    col.markdown(f'<div class="kpi {tone}"><div class="v">{value}</div><div class="l">{label}</div></div>',
                 unsafe_allow_html=True)


# ------------------------------------------------------------------ small parsers
def blank(v):
    if v is None:
        return True
    if isinstance(v, str):
        return not v.strip()
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def as_time(value):
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, time):
        return value
    if isinstance(value, (int, float)) and 0 <= value < 1:
        m = round(value * 1440)
        return time(m // 60, m % 60)
    s = str(value).strip().upper()
    if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?(\.\d+)?", s):
        s = s.split(".")[0]
    s = s.replace(".", ":")
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M%p", "%I %p", "%I%p", "%H"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    raise ValueError(f"'{value}' is not a time")


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and value > 20000:  # Excel serial number
        return date(1899, 12, 30) + timedelta(days=int(value))
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d.%m.%Y", "%d %b %Y", "%d %B %Y",
                "%a %d %b %Y", "%A %d %B %Y", "%A, %d %B %Y", "%a, %d %b %Y"):
        try:
            return datetime.strptime(s[:10] if fmt == "%Y-%m-%d" else s, fmt).date()
        except ValueError:
            pass
    return pd.to_datetime(s, dayfirst=True).date()


WEEKDAYS = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}


def parse_days(textval):
    """'Mon, Wed' / 'Tuesday or Thursday' / '18/12/2026' -> (weekday numbers, dates)."""
    wdays, dates_ = set(), set()
    for token in re.split(r"[,;/&]|\band\b|\bor\b", str(textval), flags=re.I):
        token = token.strip()
        if not token:
            continue
        key = token[:3].casefold()
        if key in WEEKDAYS and token.isalpha():
            wdays.add(WEEKDAYS[key])
            continue
        try:
            dates_.add(as_date(token))
        except Exception:
            raise ValueError(f"'{token}' is not a weekday or a date")
    return wdays, dates_


def norm(s):
    return re.sub(r"[\s\-_./]", "", str(s)).casefold()


def fmt_t(d, t):
    return f"{d:%a %d %b} {t:%H:%M}"


# ------------------------------------------------------------------ blank templates (built in, so dashboard.py stays one file)
import base64 as _b64
TEMPLATES = {  # name -> base64 of the .xlsx
    "Exam Timetabling Input – TEMPLATE.xlsx": "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",
    "Requests – TEMPLATE.xlsx": "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",
    "Registrar Schedule – TEMPLATE.xlsx": "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/f56vOri9KA9PXL69ADpw5fRhy9z/7d5egi+KfZH5YdvyvyPyrtW8cpniGRPkayNZG0kayO1ilc+Q6T5U6R5G2neRpq3kVrFK58h0uIp0qKNtGgjLdpIreKVzxBp+RRp2UZatpGWbaRW8cpniLR6irRqI63aSKs2Uqt45TNEWj9FWreR1m2kdRupVbzyGSJtniJt2kibNtKmjdQqXvkMkWQ2NfesDfUoVakepSpWL3npNQYLkyJ9MOmDSR+slbz0GoNNVS59l0tf5tK3eS956TUGmwpd+kaXvtKl7/Re8tJrDDbVuvS9Ln2xS9/sveSl1xhsKnfp2136epe+33vJS68x2FTx0ne89CUvfcv3kpdeY7Cp6KVveumrXvqu7yUvvcZgU91L3/fSF770jd9LXnqNwabSl771pa996Xu/l7z0Gn8inZpf++bXvvm1b/5e8tJrDDY1v/bNr33za9/8veSl1xgs/BAPP8XDj/Hwczz8IP9s8+vU/No3v/bNr33z95KXXmOwqfm1b37tm1/75u8lL73GYFPza9/82je/9s3fS156jcGm5te++bVvfu2bv5e89BqDTc2vffNr3/zaN38veek1BpuaX/vm1775tW/+XvLSaww2Nb/2za9982vf/L3kpdf4P/5T81vf/NY3v/XN30teeo3Bpua3vvmtb37rm7+XvPQag03Nb33zW9/81jd/L3npNQYLT+LAszjwNA48jwNP5Dzb/DY1v/XNb33zW9/8veSl1xhsan7rm9/65re++XvJS68x2NT81je/9c1vffP3kpdeY7Cp+a1vfuub3/rm7yUvvcZgU/Nb3/zWN7/1zd9LXnqNwabmt775rW9+65u/l7z0Gp9fnZp/3jf/vG/+ed/8veSl1xhsav553/zzvvnnffP3kpdeY7Cp+ed988/75p/3zd9LXnqNwabmn/fNP++bf943fy956TUGC0/iw7P48DQ+PI8PT+Q/2/zzqfnnffPP++af983fS156jcGm5p/3zT/vm3/eN38veek1Bpuaf943/7xv/nnf/L3kpdcYbGr+ed/88775533z95KXXmOwqfnnffPP++af983fS156jb/Gmpp/0Tf/om/+Rd/8veSl1xhsav5F3/yLvvkXffP3kpdeY7Cp+Rd98y/65l/0zd9LXnqNwabmX/TNv+ibf9E3fy956TUGm5p/0Tf/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",
}
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def template_button(where, name, label, key):
    where.download_button(label, _b64.b64decode(TEMPLATES[name]), name, XLSX, key=key, use_container_width=True)


# ------------------------------------------------------------------ header + input file
HEADER = st.empty()

with st.sidebar:
    st.header("1 · Input workbook")
    upload = st.file_uploader("Exam Timetabling Input.xlsx", type=["xlsx"])

if upload is None:
    HEADER.markdown(
        '''<div class="band"><div><div class="k">AUB Mediterraneo · Final Year Project</div>
        <h1>Final Exam Timetable</h1>
        <p>Upload the semester's input workbook (filled template) in the sidebar to start.</p></div></div>''',
        unsafe_allow_html=True)
    st.info("⬅️ Upload the semester's input workbook to start. Nothing is loaded until you do.")
    st.markdown("#### Start here: download the blank templates")
    st.caption("Fill them with the semester's data (Paste Special → Values, keep the data inside the tables). "
               "Upload the input workbook on the left to start; the two optional files are loaded later, inside "
               "the dashboard, where their templates are also available.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**1 · Input workbook** (required)  \nCourses, students, rooms, lab finals, settings.")
        template_button(st, "Exam Timetabling Input – TEMPLATE.xlsx", "⬇️ Input workbook template", "tpl_in")
    with c2:
        st.markdown("**2 · Special requests** (optional)  \nProfessors' dates, days, times, long exams, no proctor.")
        template_button(st, "Requests – TEMPLATE.xlsx", "⬇️ Requests template", "tpl_req")
    with c3:
        st.markdown("**3 · Registrar's schedule** (optional)  \nTo compare the registrar's own timetable with ours.")
        template_button(st, "Registrar Schedule – TEMPLATE.xlsx", "⬇️ Registrar schedule template", "tpl_reg")
    st.stop()
file_bytes = upload.getvalue()
file_name = upload.name


def _with_temp(data, fn):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(data)
        path = tmp.name
    try:
        return fn(path)
    finally:
        Path(path).unlink(missing_ok=True)


REQUIRED = {  # what the optimizer reads: sheet, table -> columns
    ("Settings", "tblSettings"): ["Setting", "Value"],
    ("Exams", "tblLabFinals"): ["LabFinalKey"],
    ("Courses", "tblCourses"): ["OfferingID", "CourseCode", "CourseTitle", "HasFinal", "ProfessorID", "FinalGroup"],
    ("Students", "tblStudents"): ["StudentID"] + [f"Course{i}" for i in range(1, 8)],
    ("Rooms", "tblRooms"): ["RoomID", "RoomName", "RoomType", "UsableExamSeats"],
}
SETTING_KEYS = ["WeeksToUse", "ExamDaysPerWeek", "PeriodStartDate", "WindowStart", "WindowEnd", "DefaultExamMinutes",
                "TurnaroundMinutes", "MaxExamsPerStudentPerDay", "SameDayWeight", "BackToBackWeight", "StudentsPerProctor"]


@st.cache_data(show_spinner=False)
def structure_check(data):
    """Sheets, tables, columns, settings, and rows pasted outside the tables (the solver ignores those)."""
    from openpyxl.utils.cell import range_boundaries

    def filled(v):
        return v not in (None, "") and not (isinstance(v, str) and (v.startswith("=") or not v.strip()))

    def go(path):
        errors, warnings = [], []
        try:
            book = load_workbook(path)
        except Exception as exc:
            return [f"The file can't be opened as an Excel workbook ({exc})."], []
        try:
            for (sheet, table), cols in REQUIRED.items():
                if sheet not in book.sheetnames:
                    errors.append(f"Sheet **{sheet}** is missing (sheets in your file: {', '.join(book.sheetnames)}). "
                                  "Use the template and keep its sheet names.")
                    continue
                ws = book[sheet]
                if table not in ws.tables:
                    errors.append(f"Sheet **{sheet}** has no table named **{table}** "
                                  f"(found: {', '.join(ws.tables) or 'no table'}). In Excel: click in the table → "
                                  "Table Design → Table Name.")
                    continue
                left, top, right, bottom = range_boundaries(ws.tables[table].ref)
                headers = [opt.text(c.value) for c in ws[top][left - 1:right]]
                missing = [c for c in cols if c not in headers]
                if missing:
                    errors.append(f"Table **{table}** ({sheet}) is missing the column(s) **{', '.join(missing)}**. "
                                  "Don't rename the header row.")
                body = [r for r in ws.iter_rows(min_row=top + 1, max_row=bottom, min_col=left, max_col=right,
                                                values_only=True) if any(filled(v) for v in r)]
                if not body and table != "tblLabFinals":
                    errors.append(f"Table **{table}** ({sheet}) is empty.")
                outside = sum(1 for r in ws.iter_rows(min_row=bottom + 1, max_row=ws.max_row, min_col=left,
                                                      max_col=right, values_only=True) if any(filled(v) for v in r))
                if outside:
                    warnings.append(f"**{sheet}**: {outside} filled row(s) are **below the table {table}**, so the "
                                    "solver ignores them. Fix: Table Design → Resize Table to include them.")
                if table == "tblSettings" and not missing:
                    have = {opt.text(r[headers.index("Setting")]) for r in body}
                    lack = [k for k in SETTING_KEYS if k not in have]
                    if lack:
                        errors.append(f"Settings is missing: **{', '.join(lack)}**.")
        finally:
            book.close()
        return errors, warnings
    return _with_temp(data, go)


@st.cache_data(show_spinner=False)
def read_workbook(data):
    return _with_temp(data, opt.read_inputs)


@st.cache_data(show_spinner=False)
def read_course_codes(data):
    """OfferingID -> (CourseCode, FinalGroup), used for readable labels and to match the registrar's file."""
    def go(path):
        book = load_workbook(path)
        try:
            return {opt.text(r["OfferingID"]).casefold(): (opt.text(r["CourseCode"]), opt.text(r["FinalGroup"]))
                    for _, r in opt.records(book, "Courses", "tblCourses")}
        finally:
            book.close()
    return _with_temp(data, go)


@st.cache_data(show_spinner=False)
def read_student_offerings(data):
    """StudentID -> the offerings (sections) they take, needed to split an exam by section."""
    def go(path):
        book = load_workbook(path)
        try:
            return {opt.text(r["StudentID"]): {opt.text(r[f"Course{i}"]).casefold() for i in range(1, 8)
                                               if opt.text(r[f"Course{i}"])}
                    for _, r in opt.records(book, "Students", "tblStudents")}
        finally:
            book.close()
    return _with_temp(data, go)


struct_errors, struct_warnings = structure_check(file_bytes)
if struct_errors:
    st.error("**The input workbook doesn't match the template, so nothing was solved:**\n\n"
             + "\n".join(f"- {e}" for e in struct_errors))
    st.stop()
try:
    base_settings, exams, students, rooms = read_workbook(file_bytes)
    course_codes = read_course_codes(file_bytes)
    student_offerings = read_student_offerings(file_bytes)
except KeyError as exc:
    st.error(f"The input workbook is missing **{exc}** (a setting, column or sheet). Use the template's names exactly.")
    st.stop()
except Exception as exc:  # the optimizer's own row-by-row checks (already say which row)
    st.error(f"**The input workbook has a problem:** {exc}")
    st.stop()

# reset per-file state when a different workbook is loaded
if st.session_state.get("_file") != (file_name, len(file_bytes)):
    for key in ("pick_days", "period", "exam_days", "weeks", "days_pw", "start_date", "win_start", "win_end",
                "result", "runs", "requests", "req_id", "req_cost", "_period_seen", "_req_loaded", "_req_notes",
                "fixes", "show_check"):
        st.session_state.pop(key, None)
    st.session_state["_file"] = (file_name, len(file_bytes))
st.session_state.setdefault("fixes", [])


def apply_fixes(exams_, students_, rooms_, fixes):
    """Room/exam fixes chosen in the Data check. Applied to a copy in memory: the Excel is never changed."""
    exams_ = {e: dict(x, offerings=list(x["offerings"]), students=list(x["students"])) for e, x in exams_.items()}
    students_, rooms_ = dict(students_), dict(rooms_)
    merged, split_group = {}, {}
    by_offerings = lambda key: next((e for e, x in exams_.items() if tuple(x["offerings"]) == tuple(key)), None)
    for f in fixes:
        if f["kind"] == "seats" and f["room"] in rooms_:
            name, kind, _ = rooms_[f["room"]]
            rooms_[f["room"]] = (name, kind, int(f["seats"]))
        elif f["kind"] == "merge" and all(r in rooms_ for r in f["rooms"]):
            kind = rooms_[f["rooms"][0]][1]
            new_id = "+".join(f["rooms"])
            rooms_[new_id] = (" + ".join(rooms_[r][0] for r in f["rooms"]), kind, sum(rooms_[r][2] for r in f["rooms"]))
            for r in f["rooms"]:
                del rooms_[r]
                merged[r] = new_id
        elif f["kind"] == "hall":
            e = by_offerings(f["exam"])
            if e:
                exams_[e]["type"] = opt.HALL
        elif f["kind"] == "split":
            e = by_offerings(f["exam"])
            if not e:
                continue
            x = exams_.pop(e)
            parts = {}
            for i, o in enumerate(x["offerings"], 1):
                pid = f"{e}-{i}"
                studs = [s_ for s_ in x["students"] if o.casefold() in student_offerings.get(s_, set())]
                exams_[pid] = dict(x, offerings=[o], students=studs)
                parts[pid] = set(studs)
                split_group[pid] = e
            for s_ in x["students"]:
                students_[s_] = tuple(sorted([y for y in students_[s_] if y != e]
                                             + [pid for pid, st_ in parts.items() if s_ in st_]))
    return exams_, students_, rooms_, merged, split_group


def fix_text(f):
    if f["kind"] == "seats":
        return f"room {f['room']} set to {f['seats']} seats"
    if f["kind"] == "merge":
        return f"rooms {' + '.join(f['rooms'])} merged"
    return f"{'/'.join(f['exam'])} " + ("split by section" if f["kind"] == "split" else "made a hall exam")


exams, students, rooms, MERGED, SPLIT_GROUP = apply_fixes(exams, students, rooms, st.session_state["fixes"])
FIX_SIG = tuple(fix_text(f) for f in st.session_state["fixes"])

headcount = {e: len(x["students"]) for e, x in exams.items()}
exam_codes = {}
for e, x in exams.items():
    codes = []
    for o in x["offerings"]:
        c = course_codes.get(o.casefold(), (o, ""))[0] or o
        if c not in codes:
            codes.append(c)
    exam_codes[e] = " / ".join(codes)
# two different finals with the same course code (e.g. GRLG 100 sec 1 vs sec 2+3): use their FinalGroup instead
dup = Counter(exam_codes.values())
for e, x in exams.items():
    if dup[exam_codes[e]] > 1:
        groups = {course_codes.get(o.casefold(), ("", ""))[1] for o in x["offerings"]} - {""}
        exam_codes[e] = next(iter(groups)) if len(groups) == 1 else ", ".join(x["offerings"])
exam_label = {e: f"{exam_codes[e]} · {exams[e]['name']}" for e in exams}
label_exam = {v: k for k, v in exam_label.items()}
exam_students = {e: set(x["students"]) for e, x in exams.items()}

def build_keymap():
    km = {}
    for e, x in exams.items():
        km[norm(e)] = e
        km[norm(exam_codes[e])] = e
        for o in x["offerings"]:
            km[norm(o)] = e
            code, group = course_codes.get(o.casefold(), ("", ""))
            if group:
                km[norm(group)] = e
            if code:
                km.setdefault(norm(code), e)
    return km


def find_exam(raw, km):
    e = km.get(norm(raw))
    if e is None:  # e.g. "ACCT 215-01" -> try the code without the section
        e = km.get(norm(re.split(r"[-–(]", str(raw))[0]))
    return e


# ------------------------------------------------------------------ scenario settings (default = workbook)
s0 = base_settings
ss = st.session_state
ss.setdefault("pick_days", False)
ss.setdefault("weeks", int(s0["WeeksToUse"]) if int(s0["WeeksToUse"]) in (1, 2) else 2)
ss.setdefault("days_pw", min(5, max(1, int(s0["ExamDaysPerWeek"]))))
ss.setdefault("start_date", as_date(s0["PeriodStartDate"]))
ss.setdefault("win_start", as_time(s0["WindowStart"]))
ss.setdefault("win_end", as_time(s0["WindowEnd"]))
first_default = as_date(s0["PeriodStartDate"])
ss.setdefault("period", (first_default, first_default + timedelta(days=13)))

with st.sidebar:
    st.header("2 · Exam period")
    pick_days = st.toggle("Pick the exam days myself", key="pick_days",
                          help="Choose any days, Saturdays included, starting on any weekday.")
    chosen_days = None
    if pick_days:
        period = st.date_input("From – to", key="period")
        if not isinstance(period, (tuple, list)) or len(period) != 2:
            st.info("Pick the last day of the period too.")
            st.stop()
        all_days = [period[0] + timedelta(days=i) for i in range((period[1] - period[0]).days + 1)]
        if ss.get("_period_seen") != tuple(period) or "exam_days" not in ss:
            ss["exam_days"] = [d for d in all_days if d.weekday() != 6]   # new period: all days but Sundays
            ss["_period_seen"] = tuple(period)
        ss["exam_days"] = [d for d in ss["exam_days"] if d in all_days]
        chosen_days = sorted(st.multiselect("Exam days", all_days, key="exam_days",
                                            format_func=lambda d: f"{d:%a %d %b}"))
        if not chosen_days:
            st.warning("Choose at least one exam day.")
            st.stop()
        weeks = 2 if (chosen_days[-1] - chosen_days[0]).days >= 7 else 1
        days = len(chosen_days)
        start_date = chosen_days[0]
    else:
        weeks = st.radio("Weeks", [1, 2], key="weeks", horizontal=True,
                         format_func=lambda w: f"{w} week" + ("s" if w > 1 else ""))
        days = st.slider("Exam days per week", 1, 5, key="days_pw")
        start_date = st.date_input("First exam day", key="start_date")
    c1, c2 = st.columns(2)
    win_start = c1.time_input("Day starts", key="win_start", step=900)
    win_end = c2.time_input("Day ends", key="win_end", step=900)
    c1, c2 = st.columns(2)
    minutes = c1.number_input("Exam length (min)", 15, 300, int(s0["DefaultExamMinutes"]), 15)
    gap = c2.number_input("Break (min)", 0, 180, int(s0["TurnaroundMinutes"]), 5)

    st.header("3 · Rules and penalties")
    cap = st.number_input("Max exams per student per day", 1, 5, int(s0["MaxExamsPerStudentPerDay"]))
    alpha = st.number_input("α · penalty for 2 exams on the same day", 0.0, 20.0, float(s0["SameDayWeight"]), 0.5)
    beta = st.number_input("β · extra penalty for back-to-back exams", 0.0, 20.0, float(s0["BackToBackWeight"]), 0.5)
    ratio = st.number_input("Students per proctor", 1, 200, int(s0["StudentsPerProctor"]))
    limit = st.slider("Max solving time (seconds)", 10, 600, int(opt.TIME_LIMIT), 10)

settings = dict(s0, WeeksToUse=weeks, ExamDaysPerWeek=min(days, 5), PeriodStartDate=datetime.combine(start_date, time()),
                WindowStart=win_start.strftime("%H:%M"), WindowEnd=win_end.strftime("%H:%M"),
                DefaultExamMinutes=int(minutes), TurnaroundMinutes=int(gap), MaxExamsPerStudentPerDay=int(cap),
                SameDayWeight=float(alpha), BackToBackWeight=float(beta), StudentsPerProctor=int(ratio))


def slots_for_days(settings, chosen):
    """Same slot format as build_slots, but for the exact days picked in the sidebar."""
    opening, closing = opt.clock(settings["WindowStart"]), opt.clock(settings["WindowEnd"])
    duration, gap = int(settings["DefaultExamMinutes"]), int(settings["TurnaroundMinutes"])
    if opening >= closing or duration < 1 or gap < 0:
        raise ValueError("Check the daily window, exam length and break.")
    slots = {}
    for d in chosen:
        start, stop, session = datetime.combine(d, opening), datetime.combine(d, closing), 1
        while start + timedelta(minutes=duration) <= stop:
            end = start + timedelta(minutes=duration)
            slots[f"{d.isoformat()}@{start:%H:%M}"] = (d, start.time(), end.time(), session)
            start += timedelta(minutes=duration + gap)
            session += 1
    return slots


def make_slots(settings, chosen):
    return slots_for_days(settings, chosen) if chosen else opt.build_slots(settings, int(settings["WeeksToUse"]))


try:
    slots = make_slots(settings, chosen_days)
except Exception as exc:
    st.error(f"Exam period settings: {exc}")
    st.stop()


def next_slots(slots_):
    """slot -> the next slot on the same day (h -> h+1)."""
    by_day = defaultdict(list)
    for t, (d, _, _, n) in slots_.items():
        by_day[d].append((n, t))
    nxt = {}
    for lst in by_day.values():
        lst.sort()
        for (_, a), (_, b) in zip(lst, lst[1:]):
            nxt[a] = b
    return nxt


# ------------------------------------------------------------------ special requests
REQ_COLS = ["Exam", "On date", "Days", "Start at", "Not before", "Finish by", "Length", "No proctor", "Note"]


def empty_requests():
    return pd.DataFrame({c: pd.Series([], dtype="bool" if c == "No proctor" else "object") for c in REQ_COLS})


def compile_requests(df, slots_):
    """Turn the Requests table into rules.
    allowed[e] = slots where exam e may START; long = exams that take 2 back-to-back slots."""
    nxt = next_slots(slots_)
    rules, errors, rows = defaultdict(list), [], []
    long_, noproc = set(), set()
    for i, r in df.iterrows():
        if blank(r.get("Exam")):
            continue
        e = label_exam.get(r["Exam"])
        if e is None:
            errors.append(f"Row {i + 1}: unknown exam '{r['Exam']}'.")
            continue
        rule = {}
        try:
            if not blank(r.get("On date")):
                rule["on"] = as_date(r["On date"])
            if not blank(r.get("Days")):
                rule["wdays"], rule["dates"] = parse_days(r["Days"])
            for col, key in (("Start at", "at"), ("Not before", "after"), ("Finish by", "by")):
                if not blank(r.get(col)):
                    rule[key] = as_time(r[col])
        except Exception as exc:
            errors.append(f"Row {i + 1} ({exam_codes[e]}): {exc}.")
            continue
        if r.get("Length") == LONG:
            long_.add(e)
        if bool(r.get("No proctor")) and not blank(r.get("No proctor")):
            noproc.add(e)
        if rule:
            rules[e].append(rule)
        rows.append((e, rule, r.get("Length") == LONG, bool(r.get("No proctor")) and not blank(r.get("No proctor")),
                     "" if blank(r.get("Note")) else str(r["Note"])))

    allowed = {}
    for e in set(rules) | long_:
        ok = set()
        for t, (d, s, en, _) in slots_.items():
            if e in long_:
                if t not in nxt:
                    continue
                en = slots_[nxt[t]][2]
            if all(rule_ok(rule, d, s, en) for rule in rules.get(e, [])):
                ok.add(t)
        allowed[e] = ok
    return dict(allowed=allowed, long=long_, noproctor=noproc, rules=dict(rules), errors=errors, rows=rows)


def rule_ok(rule, d, s, en):
    if "on" in rule and d != rule["on"]:
        return False
    if ("wdays" in rule or "dates" in rule) and (rule.get("wdays") or rule.get("dates")):
        if d.weekday() not in rule.get("wdays", set()) and d not in rule.get("dates", set()):
            return False
    if "at" in rule and s != rule["at"]:
        return False
    if "after" in rule and s < rule["after"]:
        return False
    if "by" in rule and en > rule["by"]:
        return False
    return True


def describe(e, rule, is_long, no_p):
    bits = []
    if "on" in rule:
        bits.append(f"on {rule['on']:%a %d %b}")
    if rule.get("wdays") or rule.get("dates"):
        names = [k.title() for k, v in WEEKDAYS.items() if v in rule.get("wdays", set())]
        names += [f"{d:%d %b}" for d in sorted(rule.get("dates", set()))]
        bits.append("only " + "/".join(names))
    if "at" in rule:
        bits.append(f"starts {rule['at']:%H:%M}")
    if "after" in rule:
        bits.append(f"not before {rule['after']:%H:%M}")
    if "by" in rule:
        bits.append(f"finished by {rule['by']:%H:%M}")
    if is_long:
        bits.append("3–4 h (2 slots)")
    if no_p:
        bits.append("no extra proctor")
    return f"{exam_codes[e]}: " + (", ".join(bits) or "no rule")


def has_rules(req):
    return bool(req and (req["allowed"] or req["long"]))


def conflicts_of(e):
    """Exams that cannot overlap e: shared students or the same professor."""
    prof = exams[e]["professor"].casefold()
    return {f for f in exams if f != e and (exam_students[e] & exam_students[f]
                                            or exams[f]["professor"].casefold() == prof)}


def add_request_rules(model, x, slots_, capacities, req, relax=None):
    """Request rules as linear constraints on the optimizer's own x[e, t] variables.
    relax[e] (optional, binary) lets the diagnosis switch one exam's request off."""
    nxt = next_slots(slots_)
    prev = {b: a for a, b in nxt.items()}
    r = (lambda e: relax[e]) if relax else (lambda e: 0)
    # (R1) the exam may only start in an allowed slot
    for e, ok in req["allowed"].items():
        banned = [x[e, t] for t in slots_ if t not in ok]
        if banned:
            model += pl.lpSum(banned) <= r(e), f"req_slot_{e}"
    # (R2) a long exam also occupies the next slot: no clash with its students / professor there,
    #      and its students get no exam in the slot right after it ends
    for L in req["long"]:
        for f in conflicts_of(L):
            for t, t2 in nxt.items():
                model += x[L, t] + x[f, t2] <= 1 + r(L)
        for f in exams:
            if f != L and exam_students[L] & exam_students[f]:
                for t, t2 in nxt.items():
                    if t2 in nxt:
                        model += x[L, t] + x[f, nxt[t2]] <= 1 + r(L)
    # (R3) rooms: a long exam keeps its room in the second slot too
    if req["long"]:
        for kind in (opt.HALL, opt.LAB):
            sized = {e: headcount[e] for e in exams if exams[e]["type"] == kind}
            for thr in sorted(set(sized.values())):
                group = [e for e, n in sized.items() if n >= thr]
                longs = [L for L in group if L in req["long"]]
                count = sum(seats >= thr for seats in capacities[kind])
                if longs and len(group) + len(longs) > count:
                    for t in slots_:
                        terms = [x[e, t] for e in group] + [x[L, prev[t]] for L in longs if t in prev]
                        model += pl.lpSum(terms) <= count + pl.lpSum(r(L) for L in longs)


def rooms_with_long(chosen, exams_, rooms_, headcount_, long_, slots_):
    """Same rule as the optimizer's assign_rooms (largest exam -> smallest room that fits),
    but a long exam keeps its room for the next slot as well."""
    nxt = next_slots(slots_)
    held, out = defaultdict(set), {}
    for t in sorted(set(chosen.values()), key=lambda s: (slots_[s][0], slots_[s][1])):
        free = sorted((r for r in rooms_ if r not in held[t]), key=lambda r: rooms_[r][2])
        for e in sorted((e for e in chosen if chosen[e] == t), key=headcount_.get, reverse=True):
            room = next(r for r in free if rooms_[r][1] == exams_[e]["type"] and rooms_[r][2] >= headcount_[e])
            out[e] = (t, room)
            free.remove(room)
            if e in long_ and t in nxt:
                held[nxt[t]].add(room)
    return out


def repair_seed(seed, req, slots_):
    """Move only the requested exams of the optimizer's quick start into slots that respect their
    request, so CBC still gets a good starting timetable. Returns {} if no repair is found."""
    seed = dict(seed)
    nxt = next_slots(slots_)
    longs = req["long"]
    cap_ = int(settings["MaxExamsPerStudentPerDay"])
    caps = {k: sum(1 for r in rooms.values() if r[1] == k) for k in (opt.HALL, opt.LAB)}

    def occ(e, t):
        return {t, nxt[t]} if e in longs and t in nxt else {t}

    def fits(e, t):
        if e in req["allowed"] and t not in req["allowed"][e]:
            return False
        mine = occ(e, t)
        for f in conflicts_of(e):
            if f in seed and mine & occ(f, seed[f]):
                return False
        for f in exams:  # nothing for e's students right after a long exam (both directions)
            if f == e or f not in seed or not (exam_students[e] & exam_students[f]):
                continue
            if e in longs and t in nxt and nxt[t] in nxt and seed[f] == nxt[nxt[t]]:
                return False
            tf = seed[f]
            if f in longs and tf in nxt and nxt[tf] in nxt and t == nxt[nxt[tf]]:
                return False
        kind = exams[e]["type"]
        if sum(1 for f, tf in seed.items() if f != e and exams[f]["type"] == kind and mine & occ(f, tf)) >= caps[kind]:
            return False
        day = slots_[t][0]
        for stu in exams[e]["students"]:
            if sum(1 for f in students[stu] if f != e and f in seed and slots_[seed[f]][0] == day) + 1 > cap_:
                return False
        return True

    todo = sorted(set(req["allowed"]) | longs, key=lambda e: len(req["allowed"].get(e, slots_)))
    for e in todo:
        if e in seed and fits(e, seed[e]):
            continue
        options = [t for t in slots_ if fits(e, t)]
        if not options:
            return {}
        # prefer the slot that adds the fewest same-day exams for e's students
        day_load = lambda t: sum(1 for stu in exams[e]["students"] for f in students[stu]
                                 if f != e and f in seed and slots_[seed[f]][0] == slots_[t][0])
        seed[e] = min(options, key=day_load)
    return seed


@contextlib.contextmanager
def requests_applied(req, slots_):
    """Temporarily wrap 3 optimizer helpers so solve() builds the ILP with the request rules."""
    def model_fn(exams_, slots__, headcount_, capacities):
        model, x = ORIG_MODEL(exams_, slots__, headcount_, capacities)
        add_request_rules(model, x, slots__, capacities, req)
        return model, x

    def seed_fn(*args, **kwargs):
        seed = ORIG_SEED(*args, **kwargs)
        return repair_seed(seed, req, slots_) if seed else {}

    def rooms_fn(chosen, exams_, rooms_, headcount_):
        return rooms_with_long(chosen, exams_, rooms_, headcount_, req["long"], slots_)

    opt.assignment_model, opt.initial_schedule, opt.assign_rooms = model_fn, seed_fn, rooms_fn
    try:
        yield
    finally:
        opt.assignment_model, opt.initial_schedule, opt.assign_rooms = ORIG_MODEL, ORIG_SEED, ORIG_ROOMS


# ------------------------------------------------------------------ evaluation (same rules for our schedule and the registrar's)
def placements_from(assigned, slots_, req):
    nxt = next_slots(slots_)
    out = {}
    for e, (t, r) in assigned.items():
        d, s, en, _ = slots_[t]
        if req and e in req["long"] and t in nxt:
            en = slots_[nxt[t]][2]
        out[e] = dict(start=datetime.combine(d, s), end=datetime.combine(d, en), rooms=[r])
    return out


def evaluate(placed, set_, req, b2b_gap=60):
    """Rule check + penalty Z on real times, so any timetable can be scored the same way.
    same-day pair: 2 exams of one student on one date (α each);
    back-to-back: the second starts at most `b2b_gap` minutes after the first ends (β extra each)."""
    a, b = float(set_["SameDayWeight"]), float(set_["BackToBackWeight"])
    cap_, ratio_ = int(set_["MaxExamsPerStudentPerDay"]), int(set_["StudentsPerProctor"])
    noproc = req["noproctor"] if req else set()
    problems = []
    missing = [e for e in exams if e not in placed]
    problems += [f"{exam_codes[e]}: not scheduled." for e in missing]

    seat_short = wrong_type = unknown_rooms = 0
    booked = seated = 0
    room_minutes = 0.0
    by_room, by_prof = defaultdict(list), defaultdict(list)
    for e, p in placed.items():
        n = headcount[e]
        known = [r for r in p["rooms"] if r in rooms]
        unknown_rooms += len(p["rooms"]) - len(known)
        mins = (p["end"] - p["start"]).total_seconds() / 60
        if known:
            seats = sum(rooms[r][2] for r in known)
            booked += seats
            seated += min(n, seats)
            room_minutes += mins * len(known)
            if n > seats and len(known) == len(p["rooms"]):
                seat_short += 1
                problems.append(f"{exam_codes[e]}: {n} students but {seats} seats in {', '.join(known)}.")
            if exams[e]["type"] == opt.LAB and not any(rooms[r][1] == opt.LAB for r in known):
                wrong_type += 1
                problems.append(f"{exam_codes[e]}: needs a computer lab, placed in {', '.join(known)}.")
        for r in known:
            by_room[r].append((p["start"], p["end"], e))
        by_prof[exams[e]["professor"].casefold()].append((p["start"], p["end"], e))

    def overlaps(groups, what):
        count = 0
        for key, items in groups.items():
            for (s1, e1, x1), (s2, e2, x2) in combinations(sorted(items), 2):
                same_exam = SPLIT_GROUP.get(x1, x1) == SPLIT_GROUP.get(x2, x2)  # sections of one split exam
                if s1 < e2 and s2 < e1 and x1 != x2 and not (same_exam and what == "Room"):
                    count += 1
                    problems.append(f"{what} {key}: {exam_codes[x1]} and {exam_codes[x2]} overlap.")
        return count

    room_conf = overlaps(by_room, "Room")
    prof_conf = overlaps(by_prof, "Professor")

    same_day = b2b = clashes = two_days = three_days = over = 0
    std_same, std_b2b = set(), set()
    for s_id, es in students.items():
        by_day = defaultdict(list)
        for e in es:
            if e in placed:
                by_day[placed[e]["start"].date()].append((placed[e]["start"], placed[e]["end"], e))
        for d, items in by_day.items():
            n = len(items)
            two_days += n == 2
            three_days += n >= 3
            if n > cap_:
                over += 1
                problems.append(f"{s_id}: {n} exams on {d:%a %d %b} (max {cap_}).")
            if n > 1:
                std_same.add(s_id)
            for (s1, e1, x1), (s2, e2, x2) in combinations(sorted(items), 2):
                same_day += 1
                if s1 < e2 and s2 < e1:
                    clashes += 1
                    problems.append(f"{s_id}: {exam_codes[x1]} and {exam_codes[x2]} overlap on {d:%a %d %b}.")
                elif 0 <= (s2 - e1).total_seconds() / 60 <= b2b_gap:
                    b2b += 1
                    std_b2b.add(s_id)

    # proctors over time (the course professor counts as one; "no proctor" exams need only him/her)
    def proctors(e):
        return 1 if e in noproc else max(1, math.ceil(headcount[e] / ratio_))
    points = sorted({p["start"] for p in placed.values()})
    peak_p = peak_extra = peak_std = 0
    for tp in points:
        live = [e for e, p in placed.items() if p["start"] <= tp < p["end"]]
        peak_p = max(peak_p, sum(proctors(e) for e in live))
        peak_extra = max(peak_extra, sum(proctors(e) - 1 for e in live))
        peak_std = max(peak_std, sum(headcount[e] for e in live))

    # requests (fixed date/day/hour, and 3-4 h exams get at least 3 h)
    req_total = req_met = 0
    if req:
        for e in set(req["rules"]) | req["long"]:
            req_total += 1
            p = placed.get(e)
            ok = bool(p) and all(rule_ok(rule, p["start"].date(), p["start"].time(), p["end"].time())
                                 for rule in req["rules"].get(e, []))
            if p and e in req["long"] and (p["end"] - p["start"]).total_seconds() <= int(set_["DefaultExamMinutes"]) * 60:
                ok = False
            req_met += ok
            if not ok:
                problems.append(f"{exam_codes[e]}: special request not met.")

    days_used = sorted({p["start"].date() for p in placed.values()})
    per_day = Counter(p["start"].date() for p in placed.values())
    hard = len(missing) + clashes + prof_conf + room_conf + over + seat_short + wrong_type
    k = dict(
        Z=a * same_day + b * b2b, SameDayPairs=same_day, BackToBackPairs=b2b,
        StudentsSameDay=len(std_same), StudentsB2B=len(std_b2b), TwoExamDays=two_days, ThreePlusDays=three_days,
        StudentClashes=clashes, ProfessorClashes=prof_conf, RoomConflicts=room_conf, OverDailyMax=over,
        SeatShortages=seat_short, WrongRoomType=wrong_type, UnknownRooms=unknown_rooms, Missing=len(missing),
        Hard=hard, Scheduled=len(placed), ExamDays=len(days_used),
        Span=(days_used[-1] - days_used[0]).days + 1 if days_used else 0,
        SeatFill=seated / booked if booked else None, RoomsUsed=len(by_room) if booked else None,
        RoomHours=room_minutes / 60 if booked else None,
        PeakProctors=peak_p, PeakExtraProctors=peak_extra, PeakStudents=peak_std,
        RequestsMet=req_met, Requests=req_total,
    )
    return k, problems, per_day


def patch_sheets(sheets, assigned, slots_, req, set_):
    """Fix the optimizer's output tables for 2-slot exams and no-proctor exams."""
    if not req or not (req["long"] or req["noproctor"]) or not assigned:
        return sheets
    nxt = next_slots(slots_)
    sh = {k: v.copy() for k, v in sheets.items()}
    ratio_ = int(set_["StudentsPerProctor"])
    tt, ru, su, cal, sch = (sh["Exam Timetable"], sh["Room & Lab Use"], sh["Slot Utilization"],
                            sh["Calendar"], sh["Student Schedules"])
    tt["Length"] = [LONG if e in req["long"] else STD for e in tt["ExamID"]]
    for e, (t, r) in assigned.items():
        orig = max(1, math.ceil(headcount[e] / ratio_))
        if e in req["noproctor"]:
            tt.loc[tt["ExamID"] == e, ["ProctorsRequired", "AdditionalProctors"]] = [1, 0]
            su.loc[su["SlotID"] == t, "ProctorsRequired"] -= orig - 1
            su.loc[su["SlotID"] == t, "AdditionalProctors"] -= orig - 1
        if e in req["long"] and t in nxt:
            t2 = nxt[t]
            d, s2, e2, _ = slots_[t2]
            tt.loc[tt["ExamID"] == e, "EndTime"] = f"{e2:%H:%M}"
            if not sch.empty:
                sch.loc[sch["ExamID"] == e, "EndTime"] = f"{e2:%H:%M}"
            m = (ru["SlotID"] == t2) & (ru["RoomID"] == r)
            ru.loc[m, "ExamIDs"] = f"{e} (cont.)"
            ru.loc[m, "SeatsUsed"] = headcount[e]
            ru.loc[m, "SeatUtilization"] = headcount[e] / rooms[r][2] if rooms[r][2] else None
            m = su["SlotID"] == t2
            p = 1 if e in req["noproctor"] else orig
            su.loc[m, "Exams"] += 1
            su.loc[m, "SeatsUsed"] += headcount[e]
            su.loc[m, "ProctorsRequired"] += p
            su.loc[m, "AdditionalProctors"] += p - 1
            if exams[e]["type"] == opt.LAB:
                su.loc[m, "LabSeatsUsed"] += headcount[e]
            row = cal["Session"] == f"{s2:%H:%M}–{e2:%H:%M}"
            col = d.isoformat()
            if col in cal.columns:
                old = cal.loc[row, col].fillna("").astype(str)
                cal.loc[row, col] = (old + "\n").where(old != "", "") + f"{e} {exams[e]['name']} ({r}) (cont.)"
    su["SeatUtilization"] = su["SeatsUsed"] / su["SeatsAvailable"]
    su["LabUtilization"] = su["LabSeatsUsed"] / su["LabSeatsAvailable"].where(su["LabSeatsAvailable"] > 0)
    return sh


# ------------------------------------------------------------------ solving
def solve_once(set_, slots_, req, limit_):
    opt.TIME_LIMIT = int(limit_)
    log, t0 = io.StringIO(), clock_time.time()
    ctx = requests_applied(req, slots_) if has_rules(req) else contextlib.nullcontext()
    try:
        with ctx, contextlib.redirect_stdout(log):
            status, assigned, note, solver = opt.solve(set_, exams, students, rooms, slots_)
    except StopIteration:
        status, assigned, note, solver = "No timetable", {}, "Rooms could not be matched to the exams.", "CBC"
    finally:
        clean_solver_files()
    seconds = clock_time.time() - t0
    _, _, sheets = opt.report(set_, exams, students, rooms, slots_, assigned)
    sheets = patch_sheets(sheets, assigned, slots_, req, set_)
    ev, problems, per_day = evaluate(placements_from(assigned, slots_, req), set_, req)
    return dict(status=status, assigned=assigned, note=note, solver=solver, seconds=seconds, sheets=sheets,
                ev=ev, problems=problems, per_day=per_day, settings=set_, slots=slots_, req=req, log=log.getvalue())


def feasibility(set_, slots_, req, relax=True, limit_=60):
    """Fast check without penalties. relax=True: find the fewest requests that must be dropped.
    Returns (status, dropped exams)."""
    caps = {kind: sorted(s for _, rt, s in rooms.values() if rt == kind) for kind in (opt.HALL, opt.LAB)}
    for e in exams:
        if not any(s >= headcount[e] for s in caps[exams[e]["type"]]):
            return "Seats", [e]
    model, x = ORIG_MODEL(exams, slots_, headcount, caps)
    req_exams = set(req["allowed"]) | req["long"]
    rv = {e: pl.LpVariable(f"drop_{i}", cat="Binary") for i, e in enumerate(sorted(req_exams))} if relax else None
    add_request_rules(model, x, slots_, caps, req, rv)
    cap_ = int(set_["MaxExamsPerStudentPerDay"])
    cohorts = {tuple(sorted(es)) for es in students.values() if len(es) > 1}
    profs = defaultdict(list)
    for e, xx in exams.items():
        profs[xx["professor"].casefold()].append(e)
    dates_ = defaultdict(list)
    for t, (d, *_rest) in slots_.items():
        dates_[d].append(t)
    for es in cohorts | {tuple(v) for v in profs.values() if len(v) > 1}:
        for t in slots_:
            model += pl.lpSum(x[e, t] for e in es) <= 1
    for es in cohorts:
        if len(es) > cap_:
            for times in dates_.values():
                model += pl.lpSum(x[e, t] for e in es for t in times) <= cap_
    model += pl.lpSum(list(rv.values()) if rv else [])
    try:
        model.solve(pl.PULP_CBC_CMD(msg=False, timeLimit=limit_))
    finally:
        clean_solver_files()
    if model.status == pl.LpStatusInfeasible:
        return "Infeasible", []
    if model.sol_status not in (pl.LpSolutionOptimal, pl.LpSolutionIntegerFeasible):
        return "Unknown", []
    dropped = [e for e, v in (rv or {}).items() if v.value() and v.value() > 0.5]
    return ("Optimal" if model.sol_status == pl.LpSolutionOptimal else "Feasible"), dropped


def why(e, req):
    """Plain-language reasons a request is hard to keep."""
    out = []
    ok = req["allowed"].get(e)
    if ok is not None:
        out.append(f"only {len(ok)} session(s) fit this request")
        if not ok:
            return ["no session in this exam period matches the request"]
    for f in sorted(set(req["allowed"]) - {e}):
        if ok and req["allowed"][f] & ok and f in conflicts_of(e):
            shared = len(exam_students[e] & exam_students[f])
            same_p = exams[e]["professor"].casefold() == exams[f]["professor"].casefold()
            reason = ", ".join(([f"{shared} shared students"] if shared else []) + (["same professor"] if same_p else []))
            out.append(f"competes with {exam_codes[f]}'s request for the same session(s) ({reason})")
    return out


def scenario_label(n, set_, slots_, req, chosen):
    dd = sorted({v[0] for v in slots_.values()})
    per = len(slots_) // len(dd) if dd else 0
    period = (f"{len(dd)} picked days" if chosen else f"{set_['WeeksToUse']} wk × {set_['ExamDaysPerWeek']} days")
    nreq = len(set(req["rules"]) | req["long"]) if req else 0
    return (f"#{n} · {period} · {per}/day · α={set_['SameDayWeight']:g} β={set_['BackToBackWeight']:g} · "
            f"max {set_['MaxExamsPerStudentPerDay']}/day · " + (f"{nreq} requests" if nreq else "no requests"))


# ------------------------------------------------------------------ input overview
hall_rooms = [r for r in rooms.values() if r[1] == opt.HALL]
lab_rooms = [r for r in rooms.values() if r[1] == opt.LAB]
n_lab = sum(e["type"] == opt.LAB for e in exams.values())
dates = sorted({v[0] for v in slots.values()})
per_day_n = len(slots) // len(dates) if dates else 0

HEADER.markdown(
    f'''<div class="band"><div><div class="k">AUB Mediterraneo · Final Year Project</div>
    <h1>Final Exam Timetable</h1>
    <p>Upload the semester's input workbook, choose the exam period, add any special requests and click
    <b>Solve</b>. An integer programming model (solved with CBC) builds the timetable.</p></div>
    <div class="file">📄 {file_name}</div></div>''', unsafe_allow_html=True)
cols = st.columns(5)
kpi(cols[0], len(exams), f"finals ({n_lab} in a computer lab)")
kpi(cols[1], len(students), "students")
kpi(cols[2], len(rooms), f"rooms ({len(hall_rooms)} halls · {len(lab_rooms)} labs)")
kpi(cols[3], len(dates), "exam days")
kpi(cols[4], len(slots), f"sessions ({per_day_n} per day)")
if not slots:
    st.warning("The daily window is too short for even one exam. Make the day longer or the exam shorter.")


# ------------------------------------------------------------------ data check (before solving)
def add_fix(f):
    ss["fixes"] = ss["fixes"] + [f]
    ss["show_check"] = True


def undo_fix(i):
    ss["fixes"] = [f for j, f in enumerate(ss["fixes"]) if j != i]
    ss["show_check"] = True


def data_check():
    """Everything that can be checked before solving. Returns (blocking, warnings, passed, room_issues)."""
    blocking, warns, passed, room_issues = [], list(struct_warnings), [], []
    n_days = len({v[0] for v in slots.values()})
    cap_ = int(settings["MaxExamsPerStudentPerDay"])
    by_kind = {k: sorted((r for r in rooms if rooms[r][1] == k), key=lambda r: -rooms[r][2]) for k in (opt.HALL, opt.LAB)}

    # rooms
    if not by_kind[opt.HALL]:
        blocking.append("There are **no halls** in Rooms.")
    if any(x["type"] == opt.LAB for x in exams.values()) and not by_kind[opt.LAB]:
        blocking.append("Some finals need a computer lab, but there are **no computer labs** in Rooms.")
    for r, (name, kind, seats) in rooms.items():
        if seats == 0:
            warns.append(f"Room **{r}** has 0 usable seats, so it is never used.")
        elif seats > 1000:
            warns.append(f"Room **{r}** has {seats} seats. Is that a typo?")
    for e, x in sorted(exams.items(), key=lambda kv: -headcount[kv[0]]):
        kind, n = x["type"], headcount[e]
        if by_kind[kind] and rooms[by_kind[kind][0]][2] < n:
            room_issues.append(e)
            blocking.append(f"**{exam_codes[e]}** has **{n} students** but the biggest "
                            f"{'computer lab' if kind == opt.LAB else 'hall'} (**{by_kind[kind][0]}**) has "
                            f"{rooms[by_kind[kind][0]][2]} seats. The model puts each exam in one room.")
    if not room_issues:
        passed.append("Every final fits in at least one room of its type")
    for kind, label in ((opt.HALL, "hall"), (opt.LAB, "lab")):
        need = sum(1 for x in exams.values() if x["type"] == kind)
        have = len(by_kind[kind]) * len(slots)
        if need > have:
            blocking.append(f"{need} {label} finals but only {have} {label}-sessions "
                            f"({len(by_kind[kind])} rooms × {len(slots)} sessions). Add days, sessions or rooms.")
    # exams and students
    empty = [e for e in exams if headcount[e] == 0]
    if empty:
        warns.append(f"{len(empty)} final(s) have **no students**: {', '.join(exam_codes[e] for e in empty[:8])}"
                     + ("…" if len(empty) > 8 else "") + ". Check Students or HasFinal.")
    else:
        passed.append("Every final has students")
    none = sum(1 for es in students.values() if not es)
    if none:
        warns.append(f"{none} student(s) have no final at all (only courses without finals). That's fine if expected.")
    heavy = [(s_, len(es)) for s_, es in students.items() if len(es) > n_days * cap_]
    if heavy:
        blocking.append(f"{len(heavy)} student(s) have more finals than the period allows ({n_days} days × {cap_} a day), "
                        f"e.g. {heavy[0][0]} with {heavy[0][1]}. Add days or raise the max per day.")
    else:
        passed.append(f"No student has more finals than {n_days} days × {cap_} a day")
    profs = Counter(x["professor"].casefold() for x in exams.values())
    busy = [(p_, n) for p_, n in profs.items() if n > len(slots)]
    if busy:
        blocking.append(f"Professor {busy[0][0]} has {busy[0][1]} finals but there are only {len(slots)} sessions.")
    else:
        passed.append("No professor has more finals than sessions")
    if not struct_warnings:
        passed.append("All rows are inside the tables")
    passed.append(f"Sheets, tables and columns match the template")
    return blocking, warns, passed, room_issues


blocking, check_warns, check_passed, room_issues = data_check()
with st.sidebar:
    if st.button("🔎 Check my data", use_container_width=True):
        ss["show_check"] = True
open_check = bool(blocking or ss.get("show_check"))
title = ("❌ Data check: problems to fix before solving" if blocking else
         "⚠️ Data check: passed, with warnings" if check_warns else "✅ Data check: all good")
with st.expander(title, expanded=open_check):
    st.caption(f"{len(exams)} finals · {len(students)} students · {len(rooms)} rooms · "
               f"{sum(1 for x in exams.values() if x['type'] == opt.LAB)} lab finals · {len(slots)} sessions")
    for b in blocking:
        st.markdown(f"❌ {b}")
    for e in room_issues:
        x, n, kind = exams[e], headcount[e], exams[e]["type"]
        same = sorted((r for r in rooms if rooms[r][1] == kind), key=lambda r: -rooms[r][2])
        st.markdown(f"**Options for {exam_codes[e]}** (applied in the dashboard only; your Excel isn't changed):")
        cols = st.columns(4)
        big = same[0]
        cols[0].button(f"➕ {big}: {rooms[big][2]} → {n} seats", key=f"fx_seat_{e}", on_click=add_fix,
                       args=({"kind": "seats", "room": big, "seats": n},), help="e.g. add laptops or chairs")
        pick, total = [], 0
        for r in same:
            pick.append(r)
            total += rooms[r][2]
            if total >= n:
                break
        if len(pick) > 1 and total >= n:
            cols[1].button(f"🔗 Merge {' + '.join(pick)} ({total} seats)", key=f"fx_merge_{e}", on_click=add_fix,
                           args=({"kind": "merge", "rooms": pick},),
                           help="Used together as one room. Then only one exam at a time can use them.")
        if len(x["offerings"]) > 1:
            cols[2].button(f"✂️ Split into {len(x['offerings'])} sections", key=f"fx_split_{e}", on_click=add_fix,
                           args=({"kind": "split", "exam": tuple(x["offerings"])},),
                           help="Each section becomes its own exam (they may end up at different times).")
        if kind == opt.LAB and any(rooms[r][1] == opt.HALL and rooms[r][2] >= n for r in rooms):
            cols[3].button("🏛️ Make it a hall exam", key=f"fx_hall_{e}", on_click=add_fix,
                           args=({"kind": "hall", "exam": tuple(x["offerings"])},))
    for w in check_warns:
        st.markdown(f"⚠️ {w}")
    for ok in check_passed:
        st.markdown(f"✅ {ok}")
    if ss["fixes"]:
        st.markdown("**Fixes applied to this run:**")
        for i, f in enumerate(ss["fixes"]):
            c1, c2 = st.columns([6, 1])
            c1.markdown(f"- {fix_text(f)}")
            c2.button("↩️ Undo", key=f"undo_{i}", on_click=undo_fix, args=(i,))

# ------------------------------------------------------------------ requests editor (before solving)
ss.setdefault("requests", [])
ss.setdefault("req_id", 0)
DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
ANY = "Any"


def requests_df(items):
    """The list of requests as the table compile_requests() reads."""
    rows = [{"Exam": r["Exam"], "On date": r["On date"], "Days": ", ".join(r["Days"]) or None,
             "Start at": r["Start at"], "Not before": r["Not before"], "Finish by": r["Finish by"],
             "Length": r["Length"], "No proctor": r["No proctor"], "Note": None} for r in items]
    return pd.DataFrame(rows, columns=REQ_COLS) if rows else empty_requests()


def add_request():
    f = {k: ss.get(f"rq_{k}") for k in ("exam", "date", "days", "at", "after", "by", "long", "noproc")}
    if not f["exam"]:
        ss["rq_msg"] = ("error", "Choose an exam first.")
        return
    item = {"id": ss["req_id"], "Exam": f["exam"],
            "On date": None if f["date"] in (None, ANY) else f["date"],
            "Days": list(f["days"] or []),
            "Start at": None if f["at"] in (None, ANY) else f["at"],
            "Not before": None if f["after"] in (None, ANY) else f["after"],
            "Finish by": None if f["by"] in (None, ANY) else f["by"],
            "Length": LONG if f["long"] == LONG else STD, "No proctor": bool(f["noproc"])}
    if not (item["On date"] or item["Days"] or item["Start at"] or item["Not before"] or item["Finish by"]
            or item["Length"] == LONG or item["No proctor"]):
        ss["rq_msg"] = ("error", "This request is empty: pick at least one rule.")
        return
    ss["requests"].append(item)
    ss["req_id"] += 1
    ss["rq_msg"] = ("success", f"Added: {f['exam'].split(' · ')[0]}.")
    for k, v in (("exam", None), ("date", ANY), ("days", []), ("at", ANY), ("after", ANY), ("by", ANY),
                 ("long", STD), ("noproc", False)):
        ss[f"rq_{k}"] = v


def delete_request(rid):
    ss["requests"] = [r for r in ss["requests"] if r["id"] != rid]


st.markdown("#### 📌 Special requests from professors")
st.caption("Requests are hard rules: the solver keeps all of them, or tells you which ones can't fit together. "
           "Pick only what the professor asked for and leave the rest on *Any*.")
period_dates = sorted({v[0] for v in slots.values()})
starts = sorted({v[1] for v in slots.values()})
ends = sorted({v[2] for v in slots.values()})
with st.form("add_request", clear_on_submit=False, border=True):
    c1, c2, c3 = st.columns([2.2, 1.3, 1.6])
    c1.selectbox("Exam", sorted(exam_label.values()), index=None, placeholder="Choose the exam…", key="rq_exam")
    c2.selectbox("On a specific date", [ANY] + period_dates, key="rq_date",
                 format_func=lambda d: d if d == ANY else f"{d:%a %d %b}")
    c3.multiselect("Only on these weekdays", DAY_NAMES, key="rq_days", placeholder="Any day")
    c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1.3, 1])
    c1.selectbox("Starts exactly at", [ANY] + starts, key="rq_at", format_func=lambda t: t if t == ANY else f"{t:%H:%M}")
    c2.selectbox("Not before", [ANY] + starts, key="rq_after", format_func=lambda t: t if t == ANY else f"{t:%H:%M}")
    c3.selectbox("Finished by", [ANY] + ends, key="rq_by", format_func=lambda t: t if t == ANY else f"{t:%H:%M}")
    c4.radio("Exam length", [STD, LONG], key="rq_long")
    c5.checkbox("No proctor (only the Dr.)", key="rq_noproc")
    st.form_submit_button("➕ Add request", type="primary", on_click=add_request)
msg = ss.pop("rq_msg", None)
if msg:
    (st.error if msg[0] == "error" else st.success)(msg[1])

REQ_FILE_COLS = ["Course", "On date", "Days", "Start at", "Not before", "Finish by", "Duration (min)", "No proctor"]


def requests_file(items):
    rows = [{"Course": r["Exam"].split(" · ")[0], "On date": r["On date"], "Days": ", ".join(r["Days"]),
             "Start at": r["Start at"].strftime("%H:%M") if r["Start at"] else None,
             "Not before": r["Not before"].strftime("%H:%M") if r["Not before"] else None,
             "Finish by": r["Finish by"].strftime("%H:%M") if r["Finish by"] else None,
             "Duration (min)": 180 if r["Length"] == LONG else None,
             "No proctor": "Yes" if r["No proctor"] else None} for r in items]
    buf = io.BytesIO()
    pd.DataFrame(rows, columns=REQ_FILE_COLS).to_excel(buf, index=False, sheet_name="Requests")
    return buf.getvalue()


def load_requests(data, name):
    """Requests file -> list of requests. Columns: Course, On date, Days, Start at, Not before, Finish by,
    Duration (min) (> 120 = 2 slots), No proctor (Yes/No)."""
    df = pd.read_csv(io.BytesIO(data)) if name.lower().endswith(".csv") else pd.read_excel(io.BytesIO(data))
    cols = {norm(c): c for c in df.columns}
    get = lambda row, key: row[cols[norm(key)]] if norm(key) in cols else None
    km, items, notes = build_keymap(), [], []
    for i, row in df.iterrows():
        raw = get(row, "Course")
        if blank(raw):
            continue
        e = find_exam(raw, km)
        if e is None:
            notes.append(f"Row {i + 2}: '{raw}' is not a final in the input workbook (skipped).")
            continue
        try:
            days = []
            if not blank(get(row, "Days")):
                wd, _ = parse_days(get(row, "Days"))
                days = [DAY_NAMES[d] for d in sorted(wd)]
            t = {k: None if blank(get(row, k)) else as_time(get(row, k)) for k in ("Start at", "Not before", "Finish by")}
            dur = get(row, "Duration (min)")
            np_ = get(row, "No proctor")
            item = {"id": ss["req_id"] + len(items), "Exam": exam_label[e],
                    "On date": None if blank(get(row, "On date")) else as_date(get(row, "On date")),
                    "Days": days, "Start at": t["Start at"], "Not before": t["Not before"], "Finish by": t["Finish by"],
                    "Length": LONG if not blank(dur) and float(dur) > int(settings["DefaultExamMinutes"]) else STD,
                    "No proctor": (not blank(np_)) and str(np_).strip().casefold() in ("yes", "y", "true", "1", "x")}
        except Exception as exc:
            notes.append(f"Row {i + 2} ({raw}): {exc} (skipped).")
            continue
        items.append(item)
    return items, notes


with st.expander("📂 Load or save a requests file", expanded=False):
    c1, c2 = st.columns([3, 1])
    rf = c1.file_uploader("Requests file (Excel or CSV)", type=["xlsx", "csv"], key="req_file",
                          help="Columns: Course, On date, Days, Start at, Not before, Finish by, Duration (min), No proctor")
    template_button(c2, "Requests – TEMPLATE.xlsx", "⬇️ Blank template", "tpl_req_box")
    c2.download_button("⬇️ Save current requests", requests_file(ss["requests"]), "Requests.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       disabled=not ss["requests"])
    if rf is not None:
        sig = (rf.name, rf.size)
        if ss.get("_req_loaded") != sig:
            try:
                items, notes = load_requests(rf.getvalue(), rf.name)
                ss["requests"] = items
                ss["req_id"] += len(items) + 1
                ss["_req_loaded"] = sig
                ss["_req_notes"] = notes
                st.rerun()
            except Exception as exc:
                st.error(f"Couldn't read the requests file: {exc}")
        st.success(f"Loaded {len(ss['requests'])} requests from {rf.name}. You can still add or delete below.")
        for n in ss.get("_req_notes", []):
            st.caption("ℹ️ " + n)

edited = requests_df(ss["requests"])
req = compile_requests(edited, slots)
for err in req["errors"]:
    st.error(err)
if ss["requests"]:
    st.markdown(f"**Saved requests ({len(ss['requests'])})**")
    for item, (e, rule, lg, npv, _) in zip(ss["requests"], req["rows"]):
        ok = req["allowed"].get(e)
        c1, c2 = st.columns([12, 1])
        flag = " ⚠️ *no session in this exam period matches it*" if ok is not None and not ok else ""
        c1.markdown(f"- {describe(e, rule, lg, npv)}{flag}")
        c2.button("🗑️", key=f"del_{item['id']}", on_click=delete_request, args=(item["id"],), help="Delete this request")
    if st.button("Clear all requests"):
        ss["requests"] = []
        st.rerun()
else:
    st.caption("No requests yet.")

with st.sidebar:
    also_base = st.checkbox("Also solve without the requests (to show their cost)", value=True,
                            disabled=not has_rules(req))
    solve_clicked = st.button("Solve", type="primary", use_container_width=True, disabled=bool(blocking))
    if blocking:
        st.caption("❌ Fix the problems in **Data check** first.")

if solve_clicked:
    ss.setdefault("runs", [])
    with st.spinner("Solving the timetable… (CBC is searching for the lowest penalty)"):
        res = solve_once(settings, slots, req, limit)
        if has_rules(req) and also_base:
            base = solve_once(settings, slots, None, limit)
            res["baseline"] = base
    res["label"] = scenario_label(len(ss["runs"]) + 1, settings, slots, req, chosen_days) + \
        (" · " + "; ".join(FIX_SIG) if FIX_SIG else "")
    res["fix_sig"] = FIX_SIG
    res["chosen_days"] = chosen_days
    if res["status"] not in ("Optimal", "Feasible") and has_rules(req):
        with st.spinner("Finding which requests can't fit together…"):
            res["diagnosis"] = feasibility(settings, slots, req, relax=True)
    ss["runs"].append(res)
    ss["result"] = res
    ss.pop("req_cost", None)

result = ss.get("result")
if not result:
    st.info("Set the exam period on the left, add requests if any, and click **Solve**.")
    st.stop()

req_sig = lambda r: None if not r else (r["allowed"], r["long"], r["noproctor"])
if result["settings"] != settings or req_sig(result["req"]) != req_sig(req):
    st.warning("Settings or requests changed since the last solve. Click **Solve** to update the timetable.")

if result.get("fix_sig", ()) != FIX_SIG:
    st.warning("The data fixes changed since the last solve (rooms or exams are different now). Click **Solve**.")
    st.stop()
status, assigned = result["status"], result["assigned"]
slots_used, set_used, req_used = result["slots"], result["settings"], result["req"]


def fixes_panel(res):
    """Test simple fixes (one more session per day, one more exam day) with the requests kept."""
    set_, slots_, rq = res["settings"], res["slots"], res["req"]
    st.markdown("**Possible fixes for the registrar** (each one was tested with every request kept):")
    dd = sorted({v[0] for v in slots_.values()})
    end_now = opt.clock(set_["WindowEnd"])
    step = int(set_["DefaultExamMinutes"]) + int(set_["TurnaroundMinutes"])
    new_end = (datetime.combine(date.today(), end_now) + timedelta(minutes=step)).time()
    extra_day = dd[-1] + timedelta(days=1)
    while extra_day.weekday() == 6:
        extra_day += timedelta(days=1)
    options = [
        ("one more session per day", f"Day ends {new_end:%H:%M} instead of {end_now:%H:%M}",
         dict(set_, WindowEnd=new_end.strftime("%H:%M")), res["chosen_days"], dict(win_end=new_end)),
        ("one more exam day", f"Add {extra_day:%a %d %b}", set_, dd + [extra_day], dict(day=extra_day)),
    ]
    any_ok = False
    for name, how, s_new, chosen_new, apply in options:
        if new_end <= end_now and "win_end" in apply:
            continue
        try:
            slots_new = make_slots(s_new, chosen_new)
            rq_new = compile_requests(edited, slots_new)
            stt, _ = feasibility(s_new, slots_new, rq_new, relax=False, limit_=45)
        except Exception as exc:
            stt = f"error: {exc}"
        ok = stt in ("Optimal", "Feasible")
        c1, c2 = st.columns([4, 1])
        c1.markdown(f"{'✅' if ok else '❌' if stt == 'Infeasible' else '❔'} **{name.capitalize()}**: {how} → "
                    + ("all requests fit." if ok else "still doesn't fit." if stt == "Infeasible"
                       else "couldn't be decided in time."))
        if ok:
            any_ok = True
            c2.button("Apply", key=f"apply_{name}", on_click=apply_fix, args=(apply, dd))
    if not any_ok:
        st.info("More days or sessions don't help here (e.g. two requests ask for the exact same hour). "
                "Only changing one of the requests can fix it.")


def apply_fix(apply, dd):
    if "win_end" in apply:
        ss["win_end"] = apply["win_end"]
    if "day" in apply:
        days_new = sorted(set(dd) | {apply["day"]})
        ss["pick_days"] = True
        ss["period"] = (days_new[0], days_new[-1])
        ss["_period_seen"] = (days_new[0], days_new[-1])
        ss["exam_days"] = days_new


if status not in ("Optimal", "Feasible"):
    st.error(f"**{status}.** {result['note']}")
    diag = result.get("diagnosis")
    if diag:
        dstat, dropped = diag
        if dstat == "Seats":
            st.error(f"{exam_codes[dropped[0]]} has {headcount[dropped[0]]} students: no single room is big enough. "
                     "Split it into sections (FinalGroup) or add a bigger room.")
        elif dstat == "Infeasible":
            st.error("Even without the requests, no timetable fits these settings. Add exam days or sessions.")
        elif dstat in ("Optimal", "Feasible") and dropped:
            st.markdown("#### ⚠️ Conflicting requests")
            st.markdown("These requests can't all be kept together. Keeping all the others works if "
                        f"{'this one is' if len(dropped) == 1 else 'these are'} changed:")
            for e in dropped:
                reasons = why(e, req_used)
                st.markdown(f"- **{describe(e, req_used['rules'].get(e, [{}])[0] if req_used['rules'].get(e) else {}, e in req_used['long'], e in req_used['noproctor'])}**"
                            + (f": {'; '.join(reasons)}" if reasons else ""))
            st.caption("Ask the professor(s) to change the request, or use one of the fixes below.")
            fixes_panel(result)
        elif dstat in ("Optimal", "Feasible"):
            st.info("All requests can fit together. The solver just ran out of time: raise the max solving time.")
        else:
            st.info("The check couldn't finish in time. Raise the max solving time and solve again.")
    else:
        st.write("Try more exam days, a longer day, a shorter exam, or a higher max exams per day.")
    st.stop()

# ------------------------------------------------------------------ results header
ev, problems = result["ev"], result["problems"]
if status == "Optimal":
    st.success(f"**Optimal timetable** found in {result['seconds']:.0f} s: the lowest possible penalty for these settings"
               + (" and requests." if has_rules(req_used) else "."))
else:
    st.warning(f"**Feasible timetable** (time limit reached before proving it is the best). {result['note']}")
if problems:
    st.error("Validation problems:\n\n" + "\n".join(f"- {p}" for p in problems[:20]))

cols = st.columns(6)
kpi(cols[0], f'{ev["Z"]:g}', "penalty score Z (lower is better)", "good" if ev["Z"] == 0 else "warn")
kpi(cols[1], ev["Hard"], "rule violations (clashes, rooms)", "good" if ev["Hard"] == 0 else "bad")
kpi(cols[2], ev["TwoExamDays"], "student-days with 2 exams", "good" if ev["TwoExamDays"] == 0 else "warn")
kpi(cols[3], ev["SameDayPairs"], "same-day exam pairs (× α)", "good" if ev["SameDayPairs"] == 0 else "warn")
kpi(cols[4], ev["BackToBackPairs"], "back-to-back pairs (× β)", "good" if ev["BackToBackPairs"] == 0 else "warn")
if ev["Requests"]:
    kpi(cols[5], f'{ev["RequestsMet"]}/{ev["Requests"]}', "special requests met",
        "good" if ev["RequestsMet"] == ev["Requests"] else "bad")
else:
    kpi(cols[5], ev["PeakProctors"], "proctors needed at the busiest session")

st.write("")
tabs = st.tabs(["🗓️ Calendar", "📋 Timetable", "🎓 Student lookup", "🏫 Rooms & labs", "👥 Sessions & proctors",
                "🔀 Move an exam", "📌 Requests impact", "⚖️ Compare scenarios", "🏛️ Compare with registrar"])
timetable = result["sheets"]["Exam Timetable"]
nxt_used = next_slots(slots_used)

# ------------------------------------------------------------------ calendar
with tabs[0]:
    search = st.text_input("Highlight a course (code or name)", "", placeholder="e.g. INDE302 or Calculus")
    by_slot = defaultdict(list)
    for _, row in timetable.iterrows():
        by_slot[row["SlotID"]].append((row, False))
        if req_used and row["ExamID"] in req_used["long"] and row["SlotID"] in nxt_used:
            by_slot[nxt_used[row["SlotID"]]].append((row, True))
    sessions = sorted({(v[1], v[2]) for v in slots_used.values()})
    html = ['<div class="scroll"><table class="cal"><tr><th></th>']
    dates_used = sorted({v[0] for v in slots_used.values()})
    for d in dates_used:
        html.append(f"<th>{d:%a %d %b}</th>")
    html.append("</tr>")
    for start, end in sessions:
        html.append(f'<tr><td class="t">{start:%H:%M}–{end:%H:%M}</td>')
        for d in dates_used:
            sid = f"{d.isoformat()}@{start:%H:%M}"
            chips = []
            for row, cont in by_slot.get(sid, []):
                kind = "lab" if row["RoomType"] == opt.LAB else "hall"
                hit = " hit" if search and (search.lower() in str(row["ExamName"]).lower()
                                            or search.lower() in exam_codes[row["ExamID"]].lower()) else ""
                tag = " (cont.)" if cont else ""
                chips.append(f'<div class="chip {kind}{hit}{" cont" if cont else ""}"><b>{row["ExamName"]}{tag}</b><br>'
                             f'{exam_codes[row["ExamID"]]} · {row["RoomID"]} · {row["Students"]} std</div>')
            html.append(f'<td><div class="cell">{"".join(chips)}</div></td>')
        html.append("</tr>")
    html.append("</table></div>")
    st.markdown('<div class="legend"><span><i class="dot" style="background:#840132"></i>Hall</span>'
                '<span><i class="dot" style="background:#1F6F78"></i>Computer lab</span>'
                '<span>dashed = 2nd half of a 3–4 h exam</span></div>', unsafe_allow_html=True)
    st.markdown("".join(html), unsafe_allow_html=True)

# ------------------------------------------------------------------ timetable table
with tabs[1]:
    c1, c2 = st.columns([2, 1])
    q = c1.text_input("Search", "", placeholder="course, professor, room…")
    kinds = c2.multiselect("Room type", [opt.HALL, opt.LAB], default=[opt.HALL, opt.LAB])
    view = timetable[timetable["RoomType"].isin(kinds)].copy()
    view.insert(4, "Course", [exam_codes[e] for e in view["ExamID"]])
    if q:
        view = view[view.apply(lambda r: q.lower() in " ".join(map(str, r.values)).lower(), axis=1)]
    show = ["ExamDate", "StartTime", "EndTime", "Course", "ExamName", "CourseOfferingIDs", "ProfessorID",
            "RoomName", "RoomType", "Students", "RoomSeats", "ProctorsRequired", "AdditionalProctors"]
    if "Length" in view:
        show.insert(3, "Length")
    st.dataframe(view[show], hide_index=True, use_container_width=True, height=min(700, 40 + 35 * len(view)))

# ------------------------------------------------------------------ student lookup
with tabs[2]:
    sched = result["sheets"]["Student Schedules"]
    ids = sorted(students)
    sid = st.selectbox("Student ID", ids)
    mine = sched[sched["StudentID"] == sid] if not sched.empty else sched
    if mine.empty:
        st.info("This student has no finals.")
    else:
        per = mine.groupby("ExamDate").size()
        busy = per[per > 1]
        if busy.empty:
            st.success(f"{sid} has {len(mine)} finals, never more than one per day.")
        else:
            st.warning(f"{sid} has 2 exams on: {', '.join(busy.index)}")
        st.dataframe(mine[["ExamDate", "StartTime", "EndTime", "ExamName", "RoomName"]], hide_index=True,
                     use_container_width=True)
    st.caption("Every student's schedule is also in the downloaded Excel (Student Schedules sheet).")

# ------------------------------------------------------------------ rooms
with tabs[3]:
    ru = result["sheets"]["Room & Lab Use"].copy()
    ru["Session"] = ru["SlotID"].str.replace("@", " ")
    ru["Use %"] = (ru["SeatUtilization"].fillna(0) * 100).round(0)
    ru["Exam"] = ru["ExamIDs"].replace("", "—")
    heat = alt.Chart(ru).mark_rect(cornerRadius=3).encode(
        x=alt.X("Session:N", sort=list(slots_used), title=None, axis=alt.Axis(labelAngle=-60, labelLimit=140)),
        y=alt.Y("RoomName:N", title=None),
        color=alt.Color("Use %:Q", scale=alt.Scale(domain=[0, 100], range=["#F6F1F2", "#840132"]), title="Seats used %"),
        tooltip=["RoomName", "RoomType", "Session", "Exam", "SeatsUsed", "SeatsAvailable", "Use %"],
    ).properties(height=40 + 34 * len(rooms))
    st.altair_chart(heat, use_container_width=True)
    occ = ru.groupby(["RoomName", "RoomType"]).agg(Sessions_used=("ExamIDs", lambda s: (s != "").sum()),
                                                  Sessions=("ExamIDs", "size")).reset_index()
    occ["Occupancy %"] = (100 * occ["Sessions_used"] / occ["Sessions"]).round(0)
    st.dataframe(occ, hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ sessions + proctors
with tabs[4]:
    su = result["sheets"]["Slot Utilization"].copy()
    su["Session"] = su["SlotID"].str.replace("@", " ")
    bars = alt.Chart(su).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, color="#840132").encode(
        x=alt.X("Session:N", sort=list(slots_used), title=None, axis=alt.Axis(labelAngle=-60)),
        y=alt.Y("SeatsUsed:Q", title="Students sitting exams"),
        tooltip=["Session", "Exams", "SeatsUsed", "SeatsAvailable", "ProctorsRequired", "AdditionalProctors"])
    cap_line = alt.Chart(su).mark_rule(color="#2A1C21", strokeDash=[5, 4]).encode(y="max(SeatsAvailable):Q")
    st.markdown("**Students per session** (dashed line = all exam seats)")
    st.altair_chart(bars + cap_line, use_container_width=True)
    pro = alt.Chart(su).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, color="#1F6F78").encode(
        x=alt.X("Session:N", sort=list(slots_used), title=None, axis=alt.Axis(labelAngle=-60)),
        y=alt.Y("ProctorsRequired:Q", title="Proctors"),
        tooltip=["Session", "ProctorsRequired", "AdditionalProctors"])
    st.markdown(f"**Proctors per session** (1 per {set_used['StudentsPerProctor']} students, rounded up; "
                "the course professor counts as one; 'no proctor' exams need only the professor)")
    st.altair_chart(pro, use_container_width=True)
    st.dataframe(su.drop(columns=["Session"]), hide_index=True, use_container_width=True)

# ------------------------------------------------------------------ move an exam
with tabs[5]:
    st.caption("Try moving one exam by hand. The dashboard reassigns rooms and re-checks every rule "
               "and every special request before you apply it.")
    chosen = {e: t for e, (t, _) in assigned.items()}
    labels = {e: exam_label[e] for e in sorted(assigned, key=lambda e: exam_label[e])}
    c1, c2 = st.columns(2)
    e_sel = c1.selectbox("Exam", list(labels), format_func=labels.get)
    slot_ids = list(slots_used)
    t_sel = c2.selectbox("New session", slot_ids, index=slot_ids.index(chosen[e_sel]),
                         format_func=lambda t: fmt_t(slots_used[t][0], slots_used[t][1]))
    if t_sel != chosen[e_sel]:
        trial = dict(chosen, **{e_sel: t_sel})
        try:
            new_assigned = rooms_with_long(trial, exams, rooms, headcount, req_used["long"] if req_used else set(),
                                           slots_used)
            nev, nproblems, nper = evaluate(placements_from(new_assigned, slots_used, req_used), set_used, req_used)
        except StopIteration:
            new_assigned, nproblems = None, ["No free room of the right type and size in that session."]
        if nproblems:
            st.error("This move breaks a rule:\n\n" + "\n".join(f"- {p}" for p in nproblems[:10]))
        else:
            dz = nev["Z"] - ev["Z"]
            msg = f"Allowed. Penalty Z: {ev['Z']:g} → {nev['Z']:g} ({'+' if dz > 0 else ''}{dz:g})."
            (st.warning if dz > 0 else st.success)(msg + (" The optimizer's timetable was better." if dz > 0 else ""))
            if st.button("Apply this move"):
                _, _, nsheets = opt.report(set_used, exams, students, rooms, slots_used, new_assigned)
                result.update(assigned=new_assigned, ev=nev, problems=nproblems, per_day=nper,
                              sheets=patch_sheets(nsheets, new_assigned, slots_used, req_used, set_used),
                              status="Feasible", note="Edited by hand after solving.")
                st.rerun()


# ------------------------------------------------------------------ comparison helpers
METRICS = [  # (key, label, lower is better?)
    ("Z", "Penalty score Z", True),
    ("Hard", "Rule violations (total)", True),
    ("StudentClashes", "Student clashes (2 exams at once)", True),
    ("ProfessorClashes", "Professor clashes", True),
    ("RoomConflicts", "Room double-bookings", True),
    ("SeatShortages", "Exams with too few seats", True),
    ("WrongRoomType", "Lab exams outside a lab", True),
    ("OverDailyMax", "Students over the daily max", True),
    ("Missing", "Exams not scheduled", True),
    ("SameDayPairs", "Same-day exam pairs (× α)", True),
    ("BackToBackPairs", "Back-to-back pairs (× β)", True),
    ("StudentsSameDay", "Students with 2+ exams on a day", True),
    ("StudentsB2B", "Students with a back-to-back", True),
    ("TwoExamDays", "Student-days with 2 exams", True),
    ("ThreePlusDays", "Student-days with 3+ exams", True),
    ("ExamDays", "Exam days used", None),
    ("RoomsUsed", "Rooms used", None),
    ("RoomHours", "Room-hours booked", True),
    ("SeatFill", "Seat fill (students ÷ seats booked)", False),
    ("PeakStudents", "Peak students at the same time", None),
    ("PeakProctors", "Peak proctors at the same time", True),
    ("PeakExtraProctors", "Peak extra proctors (besides the Drs.)", True),
    ("RequestsMet", "Special requests met", False),
]


def fmt_v(key, v):
    if v is None:
        return "—"
    if key == "SeatFill":
        return f"{100 * v:.0f}%"
    if key == "RoomHours":
        return f"{v:.0f}"
    return f"{v:g}" if isinstance(v, float) else str(v)


def compare_table(a, b, name_a, name_b, req_count=0):
    rows = []
    for key, label, lower in METRICS:
        if key == "RequestsMet" and not req_count:
            continue
        va, vb = a.get(key), b.get(key)
        if lower is None or va is None or vb is None or va == vb:
            better = "same" if va == vb else ""
        else:
            better = name_a if (va < vb) == lower else name_b
        rows.append({"Metric": label, name_a: fmt_v(key, va), name_b: fmt_v(key, vb), "Better": better})
    return pd.DataFrame(rows)


def per_day_chart(series):
    data = pd.DataFrame([{"Date": d.isoformat(), "Day": f"{d:%a %d %b}", "Schedule": name, "Exams": n}
                         for name, pdays in series.items() for d, n in pdays.items()])
    if data.empty:
        return
    st.altair_chart(alt.Chart(data).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
        x=alt.X("Day:N", sort=alt.EncodingSortField("Date"), title=None, axis=alt.Axis(labelAngle=-45)),
        xOffset="Schedule:N", y=alt.Y("Exams:Q", title="Exams that day"),
        color=alt.Color("Schedule:N", scale=alt.Scale(range=["#840132", "#1F6F78", "#B7791F"]), title=None),
        tooltip=["Day", "Schedule", "Exams"]), use_container_width=True)


# ------------------------------------------------------------------ requests impact
with tabs[6]:
    if not req_used or not (req_used["rules"] or req_used["long"]):
        st.info("This solve had no special requests. Add some above the results and solve again to see their cost.")
    else:
        st.markdown("**Same settings, with vs without the professors' requests.** "
                    "Requests don't add penalties directly, but they can force students into same-day or "
                    "back-to-back exams. This is what you can show a professor.")
        base = result.get("baseline")
        if base is None or base["status"] not in ("Optimal", "Feasible"):
            st.info("No 'without requests' solve for this run. Tick *Also solve without the requests* and solve again.")
        else:
            st.dataframe(compare_table(base["ev"], ev, "Without requests", "With requests", ev["Requests"]),
                         hide_index=True, use_container_width=True)
            dz = ev["Z"] - base["ev"]["Z"]
            (st.success if dz <= 0 else st.warning)(
                f"The requests change the penalty by **{'+' if dz > 0 else ''}{dz:g}** "
                f"({base['ev']['Z']:g} → {ev['Z']:g}). Students with 2 exams on a day: "
                f"{base['ev']['StudentsSameDay']} → {ev['StudentsSameDay']}.")
            if base["status"] != "Optimal" or status != "Optimal":
                st.caption("One of the two solves stopped at the time limit (Feasible), so the difference is approximate.")
            per_day_chart({"Without requests": base["per_day"], "With requests": result["per_day"]})

        st.markdown("**Cost of each request** (re-solves once per request, keeping all the others)")
        if st.button("Calculate the cost of each request"):
            out = []
            ids = sorted(set(req_used["rules"]) | req_used["long"], key=lambda e: exam_codes[e])
            bar = st.progress(0.0)
            for i, e in enumerate(ids):
                sub_df = edited[edited["Exam"] != exam_label[e]]
                sub = compile_requests(sub_df, slots_used)
                r2 = solve_once(set_used, slots_used, sub if has_rules(sub) else None, limit)
                rule = req_used["rules"].get(e, [{}])[0] if req_used["rules"].get(e) else {}
                if r2["status"] in ("Optimal", "Feasible"):
                    e2 = r2["ev"]
                    if e2["Z"] > ev["Z"]:  # dropping a request can't make it worse: that solve ran out of time
                        e2, note = ev, "time limit: raise the solving time for an exact number"
                    else:
                        note = "exact" if r2["status"] == "Optimal" else "approximate (time limit)"
                    out.append({"Request": describe(e, rule, e in req_used["long"], e in req_used["noproctor"]),
                                "Z without this request": e2["Z"],
                                "Extra penalty caused": ev["Z"] - e2["Z"],
                                "Extra students with 2 exams/day": ev["StudentsSameDay"] - e2["StudentsSameDay"],
                                "Extra back-to-backs": ev["BackToBackPairs"] - e2["BackToBackPairs"],
                                "Result": note})
                bar.progress((i + 1) / len(ids))
            ss["req_cost"] = pd.DataFrame(out)
        if "req_cost" in ss and not ss["req_cost"].empty:
            st.dataframe(ss["req_cost"], hide_index=True, use_container_width=True)
            st.caption("Positive = this request makes the timetable worse for students. "
                       "If a solve hit the time limit (Feasible), its number is approximate.")


# ------------------------------------------------------------------ scenario comparison
with tabs[7]:
    runs = ss.get("runs", [])
    st.caption("Every Solve is added here, so you can compare exam periods, penalty weights and requests.")
    if runs:
        sc = pd.DataFrame([{
            "Scenario": r["label"], "Status": r["status"],
            "Penalty Z": r["ev"]["Z"] if r["assigned"] else None,
            "Students with 2 exams on a day": r["ev"]["StudentsSameDay"] if r["assigned"] else None,
            "Same-day pairs": r["ev"]["SameDayPairs"] if r["assigned"] else None,
            "Back-to-back pairs": r["ev"]["BackToBackPairs"] if r["assigned"] else None,
            "Requests met": f'{r["ev"]["RequestsMet"]}/{r["ev"]["Requests"]}' if r["ev"]["Requests"] else "—",
            "Peak proctors": r["ev"]["PeakProctors"] if r["assigned"] else None,
            "Seconds": round(r["seconds"], 1)} for r in runs])
        st.dataframe(sc, hide_index=True, use_container_width=True)
        ok = sc[sc["Penalty Z"].notna()]
        if len(ok) > 1:
            st.altair_chart(alt.Chart(ok).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, color="#840132").encode(
                x=alt.X("Scenario:N", sort=None, title=None, axis=alt.Axis(labelAngle=0, labelLimit=260)),
                y=alt.Y("Penalty Z:Q"), tooltip=list(ok.columns)), use_container_width=True)
        if st.button("Clear comparison"):
            ss["runs"] = [result]
            st.rerun()


# ------------------------------------------------------------------ registrar's schedule
ALIASES = {
    "course": ["course", "coursecode", "code", "exam", "examid", "offeringid", "offering", "finalgroup", "subject"],
    "date": ["date", "examdate", "day"],
    "start": ["start", "starttime", "from", "time", "examtime"],
    "end": ["end", "endtime", "to", "until"],
    "room": ["room", "rooms", "roomid", "roomname", "location", "venue", "hall"],
}


def registrar_template():
    buf = io.BytesIO()
    pd.DataFrame([{"Course": "ACCT 215", "Date": date(2026, 12, 16), "Start": "09:00", "End": "11:00",
                   "Room": "B101, B102"}]).to_excel(buf, index=False, sheet_name="Registrar Schedule")
    return buf.getvalue()


def read_registrar(data, name):
    """Registrar's schedule -> {exam: placement} with the same format as ours.
    Columns (any order, any case): Course, Date, Start, End, Room (rooms separated by commas)."""
    df = pd.read_csv(io.BytesIO(data)) if name.lower().endswith(".csv") else pd.read_excel(io.BytesIO(data))
    cols = {norm(c): c for c in df.columns}
    pick = {k: next((cols[a] for a in al if a in cols), None) for k, al in ALIASES.items()}
    missing_cols = [k for k in ("course", "date", "start") if pick[k] is None]
    if missing_cols:
        raise ValueError(f"Missing column(s): {', '.join(missing_cols)}. Use the template: Course, Date, Start, End, Room.")
    keymap = build_keymap()
    roommap = {norm(old): new for old, new in MERGED.items()}  # a merged lab still matches "117" / "GF15"
    for r, (rname, _, _) in rooms.items():
        roommap[norm(r)] = r
        roommap.setdefault(norm(rname), r)
    placed, notes = {}, []
    dur = int(settings["DefaultExamMinutes"])
    for i, row in df.iterrows():
        raw = row[pick["course"]]
        if blank(raw):
            continue
        e = find_exam(raw, keymap)
        if e is None:
            notes.append(f"Row {i + 2}: '{raw}' is not a final in the input workbook (ignored).")
            continue
        try:
            d = as_date(row[pick["date"]])
            sv = row[pick["start"]]
            ev_ = row[pick["end"]] if pick["end"] else None
            if isinstance(sv, str) and re.search(r"[-–]|\bto\b", sv) and blank(ev_):
                sv, ev_ = re.split(r"\s*(?:[-–]|\bto\b)\s*", sv, maxsplit=1)
            s = as_time(sv)
            en = as_time(ev_) if not blank(ev_) else (datetime.combine(d, s) + timedelta(minutes=dur)).time()
        except Exception as exc:
            notes.append(f"Row {i + 2} ({raw}): {exc} (ignored).")
            continue
        rms = []
        if pick["room"] and not blank(row[pick["room"]]):
            for token in re.split(r"[,;+/&]|\band\b", str(row[pick["room"]])):
                if token.strip():
                    rms.append(roommap.get(norm(token), token.strip()))
        rms = list(dict.fromkeys(rms))  # "117, GF15" -> one merged lab, counted once
        p = dict(start=datetime.combine(d, s), end=datetime.combine(d, en), rooms=rms)
        if e in placed:
            if placed[e]["start"] == p["start"]:
                placed[e]["rooms"] = list(dict.fromkeys(placed[e]["rooms"] + rms))  # another section, same time
                placed[e]["end"] = max(placed[e]["end"], p["end"])
            else:
                notes.append(f"{exam_codes[e]}: listed at two different times; the first one is used.")
            continue
        placed[e] = p
        if e in SPLIT_GROUP:  # the registrar sat all sections together: same time and rooms for every part
            for part, orig in SPLIT_GROUP.items():
                if orig == SPLIT_GROUP[e] and part not in placed:
                    placed[part] = dict(p, rooms=list(p["rooms"]))
    unknown = sorted({r for p in placed.values() for r in p["rooms"] if r not in rooms})
    if unknown:
        notes.append(f"Rooms not in the input workbook (seats unknown, not checked): {', '.join(unknown)}.")
    return placed, notes


with tabs[8]:
    st.markdown("**Compare the registrar's own Fall schedule with one of our solved scenarios.** "
                "Both are scored with exactly the same rules and the same α and β.")
    c1, c2 = st.columns([3, 1])
    reg_file = c1.file_uploader("Registrar's exam schedule (Excel or CSV)", type=["xlsx", "csv"], key="reg_file")
    template_button(c2, "Registrar Schedule – TEMPLATE.xlsx", "⬇️ Blank template", "tpl_reg_tab")
    c2.caption("Columns: Course, Date, Start, End, Room")
    ok_runs = [r for r in ss.get("runs", []) if r["status"] in ("Optimal", "Feasible")
               and r.get("fix_sig", ()) == FIX_SIG]
    if reg_file is not None and ok_runs:
        try:
            reg_placed, reg_notes = read_registrar(reg_file.getvalue(), reg_file.name)
        except Exception as exc:
            st.error(f"Couldn't read the registrar's file: {exc}")
            reg_placed = None
        if reg_placed is not None:
            names = [r["label"] for r in ok_runs]
            pick_run = st.selectbox("Compare with our scenario", names, index=names.index(result["label"])
                                    if result["label"] in names else len(names) - 1)
            run = ok_runs[names.index(pick_run)]
            b2b_gap = st.number_input("Count as back-to-back if the break between two exams is at most (min)",
                                      0, 180, max(30, 2 * int(run["settings"]["TurnaroundMinutes"])), 15)
            reg_ev, reg_problems, reg_per_day = evaluate(reg_placed, run["settings"], run["req"], b2b_gap)
            our_ev, _, our_per_day = evaluate(placements_from(run["assigned"], run["slots"], run["req"]),
                                              run["settings"], run["req"], b2b_gap)

            # same-settings check
            reg_days = sorted({p["start"].date() for p in reg_placed.values()})
            our_days = sorted({v[0] for v in run["slots"].values()})
            reg_starts = sorted({p["start"].time() for p in reg_placed.values()})
            our_starts = sorted({v[1] for v in run["slots"].values()})
            diffs = []
            if set(reg_days) - set(our_days):
                diffs.append("the registrar uses days outside this scenario's period: "
                             + ", ".join(f"{d:%a %d %b}" for d in sorted(set(reg_days) - set(our_days))))
            if len(reg_days) != len(our_days):
                diffs.append(f"{len(reg_days)} exam days (registrar) vs {len(our_days)} available (ours)")
            if set(reg_starts) - set(our_starts):
                diffs.append("different start times: registrar " + ", ".join(f"{t:%H:%M}" for t in reg_starts)
                             + " · ours " + ", ".join(f"{t:%H:%M}" for t in our_starts))
            if diffs:
                st.warning("⚠️ Settings don't fully match, so the comparison is less fair:\n\n"
                           + "\n".join(f"- {d}" for d in diffs)
                           + "\n\nTip: solve a scenario with the same exam days and session times as the registrar.")
            else:
                st.success("Same exam days and session times: a fair comparison.")
            for n in reg_notes:
                st.caption("ℹ️ " + n)

            cols = st.columns(4)
            dz = reg_ev["Z"] - our_ev["Z"]
            kpi(cols[0], f'{reg_ev["Z"]:g}', "registrar's penalty Z", "warn" if dz > 0 else "good")
            kpi(cols[1], f'{our_ev["Z"]:g}', "our penalty Z", "good" if dz >= 0 else "warn")
            kpi(cols[2], f'{100 * dz / reg_ev["Z"]:.0f}%' if reg_ev["Z"] else "—", "lower penalty with our model",
                "good" if dz > 0 else "warn")
            kpi(cols[3], f'{reg_ev["StudentsSameDay"] - our_ev["StudentsSameDay"]:+d}',
                "fewer students with 2 exams on a day", "good" if reg_ev["StudentsSameDay"] > our_ev["StudentsSameDay"] else "")
            st.write("")
            st.dataframe(compare_table(reg_ev, our_ev, "Registrar", "Our model", our_ev["Requests"]),
                         hide_index=True, use_container_width=True, height=36 * (len(METRICS) + 1))
            per_day_chart({"Registrar": reg_per_day, "Our model": our_per_day})
            if reg_problems:
                with st.expander(f"Rule problems in the registrar's schedule ({len(reg_problems)})"):
                    st.write("\n".join(f"- {p}" for p in reg_problems[:300]))

            ours = placements_from(run["assigned"], run["slots"], run["req"])
            side = pd.DataFrame([{
                "Course": exam_codes[e], "Exam": exams[e]["name"], "Students": headcount[e],
                "Registrar": f"{reg_placed[e]['start']:%a %d %b %H:%M}" if e in reg_placed else "—",
                "Registrar room(s)": ", ".join(reg_placed[e]["rooms"]) if e in reg_placed else "",
                "Our model": f"{ours[e]['start']:%a %d %b %H:%M}" if e in ours else "—",
                "Our room": ", ".join(ours[e]["rooms"]) if e in ours else "",
            } for e in sorted(exams, key=lambda e: exam_codes[e])])
            with st.expander("Exam by exam"):
                st.dataframe(side, hide_index=True, use_container_width=True)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                compare_table(reg_ev, our_ev, "Registrar", "Our model", our_ev["Requests"]).to_excel(w, sheet_name="Comparison", index=False)
                side.to_excel(w, sheet_name="Exam by exam", index=False)
                pd.DataFrame({"Registrar rule problems": reg_problems}).to_excel(w, sheet_name="Registrar problems", index=False)
            st.download_button("⬇️ Download the comparison (Excel)", buf.getvalue(), "Registrar vs Model.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif reg_file is not None:
        st.info("Solve at least one scenario first.")


# ------------------------------------------------------------------ download (same sheets and layout as the optimizer's Excel)
def results_excel(res):
    kk = res["ev"]
    summary = dict(Scenario=res["label"], Solver=res["solver"], Status=res["status"], Slots=len(res["slots"]),
                   PenaltyScore=kk["Z"], **{k: v for k, v in kk.items() if k != "Z"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame([summary]).T.rename(columns={0: "Value"}).rename_axis("Metric").to_excel(writer, sheet_name="Summary")
        for name, frame in res["sheets"].items():
            frame.to_excel(writer, sheet_name=name, index=False)
        if res["req"] and res["req"]["rows"]:
            pd.DataFrame([{"Request": describe(e, rule, lg, npv), "Note": note}
                          for e, rule, lg, npv, note in res["req"]["rows"]]).to_excel(writer, sheet_name="Requests", index=False)
        for ws in writer.sheets.values():
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                ws.column_dimensions[get_column_letter(cell.column)].width = max(14, min(48, len(str(cell.value or "")) + 3))
        cal = writer.sheets["Calendar"]
        writer.sheets["Summary"].column_dimensions["A"].width = 25
        for col in range(2, cal.max_column + 1):
            cal.column_dimensions[get_column_letter(col)].width = 36
        for row in cal.iter_rows(min_row=2, min_col=1):
            lines = max((str(c.value or "").count("\n") + 1 for c in row), default=1)
            cal.row_dimensions[row[0].row].height = min(409, max(19, 20 * lines))
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=True)
    return buf.getvalue()


st.divider()
st.download_button("⬇️ Download results (Excel)", results_excel(result), "Exam Timetable Results.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
