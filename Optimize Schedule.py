from collections import Counter, defaultdict
from datetime import date, datetime, time, timedelta
from itertools import combinations
from math import isfinite
from pathlib import Path
from random import Random
from time import perf_counter
from uuid import uuid4

import pandas as pd
import pulp as pl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.cell import get_column_letter, range_boundaries


INPUT = Path("Exam Timetabling Input.xlsx")
OUTPUT = Path("Exam Timetable Results.xlsx")
TIME_LIMIT = 120  # Seconds per solve

HALL, LAB = "Hall", "Computer Lab"


def text(value):
    return "" if value is None else str(value).strip()


def clock(value):
    return value if isinstance(value, time) else time.fromisoformat(text(value))


def day(value):
    return value.date() if isinstance(value, datetime) else date.fromisoformat(str(value)[:10])


def records(book, sheet, table):
    ws = book[sheet]
    left, top, right, bottom = range_boundaries(ws.tables[table].ref)
    rows = ws.iter_rows(
        min_row=top, max_row=bottom, min_col=left, max_col=right, values_only=True
    )
    headers = next(rows)
    for row_number, row in enumerate(rows, top + 1):
        if any(text(value) for value in row):
            yield row_number, dict(zip(headers, row))


def read_inputs(path):
    book = load_workbook(path, data_only=False)
    try:
        settings = {
            r["Setting"]: r["Value"]
            for _, r in records(book, "Settings", "tblSettings")
        }
        labs = {
            text(r["LabFinalKey"]).casefold()
            for _, r in records(book, "Exams", "tblLabFinals")
        }
        exams, offerings, groups, known_keys = {}, {}, {}, set()

        for row, r in records(book, "Courses", "tblCourses"):
            offering = text(r["OfferingID"])
            key = (text(r["FinalGroup"]) or text(r["CourseCode"])).casefold()
            flag = text(r["HasFinal"]).casefold()
            if (
                not offering or not key or offering.casefold() in offerings
                or flag not in {"yes", "no"}
            ):
                raise ValueError(
                    f"Courses row {row}: enter a unique OfferingID, course code, and Yes or No."
                )

            known_keys.add(key)
            offerings[offering.casefold()] = None
            if flag == "yes":
                professor, name = text(r["ProfessorID"]), text(r["CourseTitle"])
                if not professor or not name:
                    raise ValueError(
                        f"Courses row {row}: a final needs CourseTitle and ProfessorID."
                    )
                exam_id = groups.setdefault(key, f"E{len(groups) + 1:02d}")
                exam = exams.setdefault(exam_id, {
                    "name": name,
                    "professor": professor,
                    "type": LAB if key in labs else HALL,
                    "offerings": [],
                    "students": [],
                })
                if exam["professor"].casefold() != professor.casefold():
                    raise ValueError(
                        f"Courses row {row}: shared final {exam_id} has different professor IDs."
                    )
                exam["offerings"].append(offering)
                offerings[offering.casefold()] = exam_id

        if labs - known_keys:
            raise ValueError(f"Unknown lab final keys: {', '.join(sorted(labs - known_keys))}")

        students, seen_students = {}, set()
        for row, r in records(book, "Students", "tblStudents"):
            student = text(r["StudentID"])
            picks = [
                text(r[f"Course{i}"]).casefold()
                for i in range(1, 8)
                if text(r[f"Course{i}"])
            ]
            if not student or student.casefold() in seen_students or len(set(picks)) != len(picks):
                raise ValueError(
                    f"Students row {row}: check the StudentID and repeated course choices."
                )
            seen_students.add(student.casefold())
            unknown = set(picks) - offerings.keys()
            if unknown:
                raise ValueError(
                    f"Students row {row}: unknown offering(s) {', '.join(sorted(unknown))}."
                )
            students[student] = tuple(sorted({offerings[p] for p in picks if offerings[p]}))
            for exam_id in students[student]:
                exams[exam_id]["students"].append(student)

        rooms, seen_rooms = {}, set()
        for row, r in records(book, "Rooms", "tblRooms"):
            room_id = text(r["RoomID"])
            kind = text(r["RoomType"])
            seats = float(r["UsableExamSeats"])
            if (
                not room_id or room_id.casefold() in seen_rooms or kind not in {HALL, LAB}
                or not seats.is_integer() or seats < 0
            ):
                raise ValueError(
                    f"Rooms row {row}: check RoomID, RoomType, and UsableExamSeats."
                )
            seen_rooms.add(room_id.casefold())
            rooms[room_id] = (text(r["RoomName"]), kind, int(seats))

        return settings, exams, students, rooms
    finally:
        book.close()


def build_slots(settings, weeks):
    first = day(settings["PeriodStartDate"])
    opening, closing = map(clock, (settings["WindowStart"], settings["WindowEnd"]))
    duration = int(settings["DefaultExamMinutes"])
    gap = int(settings["TurnaroundMinutes"])
    days_per_week = int(settings["ExamDaysPerWeek"])

    if weeks not in (1, 2) or not 1 <= days_per_week <= 5 or opening >= closing or duration < 1 or gap < 0:
        raise ValueError("Check the exam period, daily window, duration, and turnaround in Settings.")

    slots = {}
    for week in range(weeks):
        current = first + timedelta(weeks=week)
        for _ in range(days_per_week):
            while current.weekday() >= 5:
                current += timedelta(days=1)
            start = datetime.combine(current, opening)
            stop = datetime.combine(current, closing)
            session = 1
            while start + timedelta(minutes=duration) <= stop:
                end = start + timedelta(minutes=duration)
                slot_id = f"{current.isoformat()}@{start:%H:%M}"
                slots[slot_id] = (current, start.time(), end.time(), session)
                start += timedelta(minutes=duration + gap)
                session += 1
            current += timedelta(days=1)
    return slots


def solve(settings, exams, students, rooms, slots):
    if not exams:
        return "Optimal", {}, "There are no finals.", None
    if not slots:
        return "Infeasible", {}, "The daily window has no full-length sessions.", None

    cap = int(settings["MaxExamsPerStudentPerDay"])
    alpha = float(settings["SameDayWeight"])
    beta = float(settings["BackToBackWeight"])
    if cap < 1 or any(not isfinite(w) or w < 0 for w in (alpha, beta)):
        raise ValueError("The daily cap must be positive and penalty weights finite and nonnegative.")

    headcount = {e: len(exam["students"]) for e, exam in exams.items()}
    capacities = {
        kind: sorted(seats for _, room_type, seats in rooms.values() if room_type == kind)
        for kind in (HALL, LAB)
    }
    for e, exam in exams.items():
        if not any(seats >= headcount[e] for seats in capacities[exam["type"]]):
            return "Infeasible", {}, f"{e} needs {headcount[e]} seats in a single {exam['type']}.", None

    cohorts = Counter(tuple(sorted(es)) for es in students.values() if es)
    professors, dates = defaultdict(list), defaultdict(list)
    for e, exam in exams.items():
        professors[exam["professor"].casefold()].append(e)
    for t, (exam_date, _, _, _) in slots.items():
        dates[exam_date].append(t)

    model, x = assignment_model(exams, slots, headcount, capacities)
    seed = initial_schedule(exams, slots, dates, capacities, cap, alpha, beta)
    if seed:
        for (e, t), variable in x.items():
            variable.setInitialValue(int(seed[e] == t))

    # Prevent student and professor clashes.
    conflict_groups = set(cohorts) | {tuple(sorted(es)) for es in professors.values()}
    for es in conflict_groups:
        if len(es) > 1:
            for t in slots:
                model += pl.lpSum(x[e, t] for e in es) <= 1

    cost = []
    for group_number, (es, multiplicity) in enumerate(cohorts.items()):
        if len(es) < 2:
            continue
        attendance = {t: pl.lpSum(x[e, t] for e in es) for t in slots}
        initial = Counter(seed[e] for e in es) if seed else Counter()
        for d, times in enumerate(dates.values()):
            daily = pl.lpSum(attendance[t] for t in times)
            model += daily <= cap
            limit = min(cap, len(es), len(times))

            # Count pairs when a student has more than one exam that day.
            if alpha and limit > 1:
                pairs = pl.LpVariable(f"pairs_{group_number}_{d}", lowBound=0)
                for k in range(1, limit):
                    model += pairs >= k * daily - k * (k + 1) // 2
                if seed:
                    count = sum(initial[t] for t in times)
                    pairs.setInitialValue(count * (count - 1) // 2)
                cost.append(multiplicity * alpha * pairs)

            # Charge the extra penalty for consecutive sessions on the same day.
            if beta and limit > 1:
                for j, (a, b) in enumerate(zip(times, times[1:])):
                    adjacent = pl.LpVariable(f"adj_{group_number}_{d}_{j}", lowBound=0)
                    model += adjacent >= attendance[a] + attendance[b] - 1
                    if seed:
                        adjacent.setInitialValue(int(initial[a] + initial[b] == 2))
                    cost.append(multiplicity * beta * adjacent)

    model += pl.lpSum(cost)

    status, chosen, message = run_cbc(model, x, bool(seed))
    return status, assign_rooms(chosen, exams, rooms, headcount), message, "CBC"


def assignment_model(exams, slots, headcount, capacities):
    model = pl.LpProblem("ExamTimetabling", pl.LpMinimize)
    x = {
        (e, t): pl.LpVariable(f"x_{i}_{j}", cat="Binary")
        for i, e in enumerate(exams)
        for j, t in enumerate(slots)
    }
    for e in exams:
        model += pl.lpSum(x[e, t] for t in slots) == 1

    # At every capacity threshold, the exams that need those rooms must fit.
    for kind in (HALL, LAB):
        sized = {e: headcount[e] for e in exams if exams[e]["type"] == kind}
        seen_room_counts = set()
        for threshold in sorted(set(sized.values())):
            group = [e for e, n in sized.items() if n >= threshold]
            room_count = sum(seats >= threshold for seats in capacities[kind])
            if room_count not in seen_room_counts and len(group) > room_count:
                for t in slots:
                    model += pl.lpSum(x[e, t] for e in group) <= room_count
            seen_room_counts.add(room_count)
    return model, x


def initial_schedule(exams, slots, dates, capacities, cap, alpha, beta):
    """Suggest a feasible start. CBC can change every assignment."""
    deadline = perf_counter() + min(2.0, TIME_LIMIT / 10)
    rng = Random(0)
    exam_students = {e: set(exam["students"]) for e, exam in exams.items()}
    conflicts = {e: set() for e in exams}
    for first, second in combinations(exams, 2):
        same_professor = exams[first]["professor"].casefold() == exams[second]["professor"].casefold()
        if same_professor or exam_students[first] & exam_students[second]:
            conflicts[first].add(second)
            conflicts[second].add(first)
        if perf_counter() >= deadline:
            return {}
    ordered_rooms = {kind: sorted(values, reverse=True) for kind, values in capacities.items()}
    daily_slots = tuple(dates.values())
    best, best_score = {}, float("inf")
    while perf_counter() < deadline:
        groups, group_for = [], {}
        pending = set(exams)
        tie = {e: rng.random() for e in exams}
        while pending and perf_counter() < deadline:
            blocked = {e: {group_for[f] for f in conflicts[e] if f in group_for} for e in pending}
            e = max(pending, key=lambda item: (len(blocked[item]), len(conflicts[item]), tie[item]))
            kind = exams[e]["type"]
            index = len(groups)
            for j, group in enumerate(groups):
                if j in blocked[e]:
                    continue
                sizes = sorted(
                    [len(exam_students[e]), *(len(exam_students[f]) for f in group if exams[f]["type"] == kind)],
                    reverse=True,
                )
                seats = ordered_rooms[kind]
                if len(sizes) <= len(seats) and all(size <= room for size, room in zip(sizes, seats)):
                    index = j
                    break
            if index == len(groups):
                groups.append([])
            groups[index].append(e)
            group_for[e] = index
            pending.remove(e)
        if pending or len(groups) > len(slots):
            continue
        students = [set().union(*(exam_students[e] for e in group)) for group in groups]
        overlap = [[len(a & b) for b in students] for a in students]
        for _ in range(20):
            if perf_counter() >= deadline:
                break
            order = list(range(len(groups)))
            rng.shuffle(order)
            if _ % 3 == 0:
                order.sort(key=lambda j: -(sum(overlap[j]) - overlap[j][j]))
            per_day = [Counter() for _ in daily_slots]
            occupied, placement = {}, {}
            score = 0.0
            for group in order:
                choices = []
                for d, times in enumerate(daily_slots):
                    if any(per_day[d][s] >= cap for s in students[group]):
                        continue
                    same_day = sum(per_day[d][s] for s in students[group])
                    for t in times:
                        session = slots[t][3]
                        if (d, session) in occupied:
                            continue
                        adjacent = sum(overlap[group][occupied[d, k]]
                                       for k in (session - 1, session + 1) if (d, k) in occupied)
                        penalty = alpha * same_day + beta * adjacent
                        choices.append((penalty, (session - 1) % 2, rng.random(), d, session, t))
                if not choices:
                    break
                penalty, _, _, d, session, t = min(choices)
                score += penalty
                occupied[d, session] = group
                per_day[d].update(students[group])
                placement.update((e, t) for e in groups[group])
            if len(placement) == len(exams) and score < best_score:
                best, best_score = placement, score
    return best


def selected_slots(x):
    return {e: t for (e, t), var in x.items() if var.value() is not None and var.value() > 0.5}


def run_cbc(model, x, warm_start):
    options = dict(msg=True, timeLimit=TIME_LIMIT, gapRel=0, warmStart=warm_start, keepFiles=True)
    backend = pl.COIN_CMD(**options)
    if not backend.available():
        backend = pl.PULP_CBC_CMD(**options)
    if not backend.available():
        return "Solver error", {}, "CBC is unavailable in this Python environment."
    model.name = f"exam_{uuid4().hex}"
    try:
        model.solve(backend)
        if model.sol_status == pl.LpSolutionOptimal:
            return "Optimal", selected_slots(x), ""
        if model.sol_status == pl.LpSolutionIntegerFeasible:
            return "Feasible", selected_slots(x), "CBC returned a timetable before proving optimality."
        if model.status == pl.LpStatusInfeasible:
            return "Infeasible", {}, "CBC proved that the constraints cannot all be satisfied."
        return "No timetable", {}, "CBC stopped before finding a feasible timetable."
    except (OSError, pl.PulpSolverError) as exc:
        return "Solver error", {}, str(exc)
    finally:
        for suffix in ("lp", "mps", "mst", "sol"):
            Path(f"{model.name}-pulp.{suffix}").unlink(missing_ok=True)


def assign_rooms(chosen, exams, rooms, headcount):
    # Match the largest exams first to the smallest adequate free rooms.
    assignments = {}
    for t in dict.fromkeys(chosen.values()):
        free = sorted(rooms, key=lambda r: rooms[r][2])
        slot_exams = sorted(
            (e for e in chosen if chosen[e] == t), key=headcount.get, reverse=True
        )
        for e in slot_exams:
            room = next(
                r for r in free
                if rooms[r][1] == exams[e]["type"] and rooms[r][2] >= headcount[e]
            )
            assignments[e] = (t, room)
            free.remove(room)
    return assignments


def report(settings, exams, students, rooms, slots, assigned):
    ratio = int(settings["StudentsPerProctor"])
    if ratio < 1:
        raise ValueError("StudentsPerProctor must be positive.")

    timeline, student_rows, room_rows, utilization, problems = [], [], [], [], []
    room_use, slot_use, professors = defaultdict(list), defaultdict(list), defaultdict(list)

    for e, (t, r) in assigned.items():
        d, start, end, _ = slots[t]
        exam = exams[e]
        name, kind, seats = rooms[r]
        count = len(exam["students"])
        proctors = (count + ratio - 1) // ratio
        row = dict(
            ExamID=e, ExamName=exam["name"], CourseOfferingIDs=", ".join(exam["offerings"]),
            ExamDate=d.isoformat(), StartTime=start.strftime("%H:%M"), EndTime=end.strftime("%H:%M"),
            SlotID=t, RoomID=r, RoomName=name, RoomType=kind, Students=count, RoomSeats=seats,
            ProfessorID=exam["professor"], ProctorsRequired=proctors,
            AdditionalProctors=max(0, proctors - 1),
        )
        timeline.append(row)
        room_use[t, r].append(row)
        slot_use[t].append(row)
        professors[t, exam["professor"].casefold()].append(e)
        if kind != exam["type"] or count > seats:
            problems.append(f"{e}: incompatible room or insufficient seats in {r}.")

    missing = sorted(exams.keys() - assigned.keys())
    problems.extend(f"{e}: no assignment." for e in missing)
    professor_clashes = sum(len(es) * (len(es) - 1) // 2 for es in professors.values())
    problems.extend(
        f"{t}: professor {p} has {', '.join(es)}."
        for (t, p), es in professors.items() if len(es) > 1
    )

    same_day = back_to_back = student_clashes = two_exam_days = over_limit = 0
    for student, es in students.items():
        by_day = defaultdict(list)
        for e in es:
            if e in assigned:
                t, r = assigned[e]
                d, start, end, session = slots[t]
                by_day[d].append((session, e))
                student_rows.append(dict(
                    StudentID=student, ExamID=e, ExamName=exams[e]["name"],
                    ExamDate=d.isoformat(), StartTime=start.strftime("%H:%M"),
                    EndTime=end.strftime("%H:%M"), RoomID=r, RoomName=rooms[r][0],
                ))

        for d, entries in by_day.items():
            two_exam_days += len(entries) == 2
            if len(entries) > int(settings["MaxExamsPerStudentPerDay"]):
                over_limit += 1
                problems.append(f"{student}: exceeds the daily cap on {d}.")
            for (a, e), (b, f) in combinations(entries, 2):
                same_day += 1
                back_to_back += abs(a - b) == 1
                if a == b:
                    student_clashes += 1
                    problems.append(f"{student}: {e} and {f} clash on {d}.")

    room_conflicts = sum(len(rows) * (len(rows) - 1) // 2 for rows in room_use.values())
    for (t, r), rows in room_use.items():
        if len(rows) > 1:
            problems.append(f"{r}: multiple exams in {t}.")

    for t in slots:
        for r, (name, kind, seats) in rooms.items():
            rows = room_use.get((t, r), ())
            used = sum(row["Students"] for row in rows)
            room_rows.append(dict(
                SlotID=t, RoomID=r, RoomName=name, RoomType=kind,
                ExamIDs=", ".join(row["ExamID"] for row in rows),
                SeatsUsed=used, SeatsAvailable=seats,
                SeatUtilization=used / seats if seats else None,
            ))

        rows = slot_use.get(t, ())
        seats = sum(r[2] for r in rooms.values())
        labs = sum(r[2] for r in rooms.values() if r[1] == LAB)
        used = sum(row["Students"] for row in rows)
        lab_used = sum(row["Students"] for row in rows if row["RoomType"] == LAB)
        utilization.append(dict(
            SlotID=t, Exams=len(rows), SeatsUsed=used, SeatsAvailable=seats,
            SeatUtilization=used / seats if seats else None, LabSeatsUsed=lab_used,
            LabSeatsAvailable=labs, LabUtilization=lab_used / labs if labs else None,
            ProctorsRequired=sum(row["ProctorsRequired"] for row in rows),
            AdditionalProctors=sum(row["AdditionalProctors"] for row in rows),
        ))

    score = (
        float(settings["SameDayWeight"]) * same_day
        + float(settings["BackToBackWeight"]) * back_to_back
    )
    kpis = dict(
        ExamsScheduled=len(assigned), Finals=len(exams), MissingExams=len(missing),
        StudentClashes=student_clashes, ProfessorClashes=professor_clashes,
        RoomConflicts=room_conflicts, DailyLimitViolations=over_limit,
        TwoExamStudentDays=two_exam_days, SameDayPairs=same_day, BackToBackPairs=back_to_back,
        PenaltyScore=None if missing else score,
        PeakProctors=max((r["ProctorsRequired"] for r in utilization), default=0),
    )

    timeline.sort(key=lambda r: (r["ExamDate"], r["StartTime"], r["RoomID"]))
    student_rows.sort(key=lambda r: (r["StudentID"], r["ExamDate"], r["StartTime"]))
    calendar = defaultdict(dict)
    for t, (d, start, end, _) in slots.items():
        calendar[f"{start:%H:%M}–{end:%H:%M}"][d.isoformat()] = "\n".join(
            f'{r["ExamID"]} {r["ExamName"]} ({r["RoomID"]})' for r in slot_use.get(t, ())
        )

    return kpis, problems, {
        "Exam Timetable": pd.DataFrame(timeline),
        "Calendar": pd.DataFrame.from_dict(calendar, orient="index").rename_axis("Session").reset_index(),
        "Student Schedules": pd.DataFrame(student_rows),
        "Room & Lab Use": pd.DataFrame(room_rows),
        "Slot Utilization": pd.DataFrame(utilization),
    }


def main():
    if INPUT.resolve() == OUTPUT.resolve():
        raise ValueError("Choose an output filename different from the input files.")

    settings, exams, students, rooms = read_inputs(INPUT)
    weeks = int(settings["WeeksToUse"])
    slots = build_slots(settings, weeks)
    status, assigned, note, solver = solve(settings, exams, students, rooms, slots)
    counts, issues, sheets = report(settings, exams, students, rooms, slots, assigned)

    if status in {"Optimal", "Feasible"} and issues:
        raise RuntimeError(f"Solver output failed validation: {issues[0]}")
    if status not in {"Optimal", "Feasible"}:
        for metric in counts.keys() - {"ExamsScheduled", "Finals", "MissingExams"}:
            counts[metric] = None
        sheets = {name: frame.iloc[:0] for name, frame in sheets.items()}

    summary = dict(Weeks=weeks, Solver=solver, Status=status, Slots=len(slots), **counts)
    print(f"{solver or 'Model'} result: {status}.")
    if note:
        print(note)

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        pd.DataFrame([summary]).T.rename(columns={0: "Value"}).rename_axis("Metric").to_excel(
            writer, sheet_name="Summary"
        )
        for name, frame in sheets.items():
            for page, start in enumerate(range(0, max(1, len(frame)), 1_048_575), 1):
                tab = name if page == 1 else f"{name} {page}"
                frame.iloc[start:start + 1_048_575].to_excel(writer, sheet_name=tab, index=False)
        for worksheet in writer.sheets.values():
            for row in worksheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        cell.data_type = "s"
            worksheet.freeze_panes = "A2"
            for cell in worksheet[1]:
                column = get_column_letter(cell.column)
                worksheet.column_dimensions[column].width = max(14, min(48, len(text(cell.value)) + 3))
        calendar = writer.sheets["Calendar"]
        writer.sheets["Summary"].column_dimensions["A"].width = 25
        for column in range(2, calendar.max_column + 1):
            calendar.column_dimensions[get_column_letter(column)].width = 36
        for row in calendar.iter_rows(min_row=2, min_col=2):
            lines = max((text(cell.value).count("\n") + 1 for cell in row), default=1)
            if row:
                calendar.row_dimensions[row[0].row].height = min(409, max(19, 20 * lines))
                calendar.cell(row[0].row, 1).alignment = Alignment(vertical="top")
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    print(f"Saved {OUTPUT.resolve()}")


if __name__ == "__main__":
    main()

