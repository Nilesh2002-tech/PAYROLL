"""Attendance rules, salary calculation and salary-slip PDFs."""
import calendar
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
from fpdf import FPDF


@dataclass
class Config:
    late_allowed: int = 4
    early_allowed: int = 2
    basis: str = "calendar"        # calendar | 30 | working
    deduct_absent: bool = True


def tmin(v):
    """Keka 'h.mm' value (10.1 == 10:10) -> minutes since midnight."""
    h, _, m = str(v).strip().partition(".")
    return int(h) * 60 + int((m or "0").ljust(2, "0")[:2])


def hm(m):
    return f"{m // 60}:{m % 60:02d}"


def pdate(v):
    if hasattr(v, "date") and not isinstance(v, str):
        return v.date() if isinstance(v, datetime) else v
    try:
        return datetime.strptime(str(v).strip(), "%d %b %Y").date()
    except ValueError:
        return None


def _table(raw):
    for i, row in raw.iterrows():
        if any("employee number" in str(c).lower() for c in row):
            t = raw.iloc[i + 1:].copy()
            t.columns = [str(c).strip() for c in row]
            return t.reset_index(drop=True)


def read_workbook(f):
    try:  # calamine tolerates the odd styles in Keka exports; openpyxl may not
        sheets = pd.read_excel(f, sheet_name=None, header=None, engine="calamine")
    except Exception:
        f.seek(0)
        sheets = pd.read_excel(f, sheet_name=None, header=None)
    tabs = {n: _table(d) for n, d in sheets.items()}
    tabs = {n: t for n, t in tabs.items() if t is not None}
    punch = next((t for n, t in tabs.items() if "punch" in n.lower()), None)
    if punch is None:
        punch = next((t for t in tabs.values() if any("in time" in c.lower() for c in t.columns)), None)
    sal = next((t for n, t in tabs.items() if "salary" in n.lower()), None)
    if punch is None or sal is None:
        raise ValueError('Workbook needs a punch sheet and a "Salary Data" sheet.')
    return punch, sal


def _col(df, key):
    return next(c for c in df.columns if key in c.lower())


def compute(punch, sal, cfg: Config):
    cd, ci, co = _col(punch, "date"), _col(punch, "number"), _col(punch, "in time")
    cout = _col(punch, "out time")
    log, seen = {}, set()
    for _, r in punch.iterrows():
        d = pdate(r[cd])
        if d is None or pd.isna(r[co]) or pd.isna(r[cout]):
            continue
        log.setdefault(str(r[ci]).strip(), {})[d] = (tmin(r[co]), tmin(r[cout]))
        seen.add(d)
    first = min(seen)
    y, m = first.year, first.month
    n = calendar.monthrange(y, m)[1]
    days = []
    for d in range(1, n + 1):
        dt = datetime(y, m, d).date()
        wd = dt.weekday()  # Mon=0 .. Sun=6
        off = wd == 6 or (wd == 5 and ((d - 1) // 7 + 1) in (2, 4))
        days.append((dt, wd, (not off) and dt in seen))  # nobody punched => holiday
    work_days = sum(w for *_, w in days)
    den = {"calendar": n, "30": 30, "working": work_days}[cfg.basis]

    sid, snm, ssal = _col(sal, "number"), _col(sal, "name"), _col(sal, "salary")
    rows, details = [], {}
    for _, e in sal.dropna(subset=[sid]).iterrows():
        eid, p = str(e[sid]).strip(), log.get(str(e[sid]).strip(), {})
        L = E = H = A = pr = 0
        ot, det = 0.0, []
        for dt, wd, work in days:
            x = p.get(dt)
            if x is None:
                if work:
                    A += 1
                    det.append({"Date": dt, "In": "", "Out": "", "Hours": "", "Flags": "Absent"})
                continue
            i, o = x
            u, f, ob = o - i, [], 0.0
            pr += 1
            if not work:
                ob = (u // 210) * 0.5
                f.append("Off-day work")
            else:
                half = i > 660 or u <= 240
                flex = 540 < i <= 600
                late = not half and (i > 600 or (flex and u < 540))
                early = not half and ((i <= 540 and o < 1080) or (i > 600 and o < 960))
                shift_end = max(1080, i + 540) if i <= 600 else 1080
                H, L, E = H + half, L + late, E + early
                f += [t for t, c in (("Half day", half), ("Late", late), ("Early leave", early)) if c]
                if o > shift_end:
                    ob = ((o - shift_end) // 210) * 0.5
            if ob:
                ot += ob
                f.append(f"OT {ob}d")
            det.append({"Date": dt, "In": hm(i), "Out": hm(o), "Hours": hm(u), "Flags": ", ".join(f)})
        xl, xe = max(0, L - cfg.late_allowed), max(0, E - cfg.early_allowed)
        s = float(e[ssal])
        rate = s / den
        ded = [("Absent days (LOP)", A if cfg.deduct_absent else 0, 1.0), ("Half days", H, .5),
               ("Excess late marks", xl, .5), ("Excess early leaving", xe, .5)]
        ded = [(t, q, f, q * f * rate) for t, q, f in ded]
        total_ded, ot_amt = sum(d[3] for d in ded), ot * rate
        rows.append(dict(ID=eid, Employee=str(e[snm]).strip(), Salary=s, Present=pr, Absent=A, Late=L,
                         ExcessLate=xl, HalfDays=H, Early=E, ExcessEarly=xe, OTDays=ot, Rate=rate,
                         Deductions=total_ded, Overtime=ot_amt, Net=s - total_ded + ot_amt, _ded=ded))
        details[eid] = pd.DataFrame(det)
    return dict(period=datetime(y, m, 1).strftime("%B %Y"), work_days=work_days, basis_days=den,
                rows=rows, details=details)


def _inr(v):
    return f"Rs. {round(v):,}"


def slip_pdf(rows, period, work_days, company) -> bytes:
    pdf = FPDF()
    for e in rows:
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 17)
        pdf.cell(130, 9, company)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 9, "SALARY SLIP", align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, f"Pay period: {period}", new_x="LMARGIN", new_y="NEXT")
        pdf.line(10, pdf.get_y() + 1, 200, pdf.get_y() + 1)
        pdf.ln(5)
        for k, v in [("Employee Name", e["Employee"]), ("Employee ID", e["ID"]), ("Pay Period", period),
                     ("Working days / Present / Absent", f"{work_days} / {e['Present']} / {e['Absent']}"),
                     ("Late marks / Early leaves / Half days", f"{e['Late']} / {e['Early']} / {e['HalfDays']}")]:
            pdf.cell(95, 6.5, k)
            pdf.cell(0, 6.5, str(v), new_x="LMARGIN", new_y="NEXT")
        pdf.ln(4)

        def head(t):
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_fill_color(235, 238, 243)
            pdf.cell(0, 7, " " + t, fill=True, new_x="LMARGIN", new_y="NEXT")

        def line(t, v, bold=False):
            pdf.set_font("Helvetica", "B" if bold else "", 10)
            pdf.cell(130, 6.5, t)
            pdf.cell(0, 6.5, v, align="R", new_x="LMARGIN", new_y="NEXT")

        head("EARNINGS")
        line(f"Monthly salary (per-day rate {_inr(e['Rate'])})", _inr(e["Salary"]))
        line(f"Overtime ({e['OTDays']:g} days)", _inr(e["Overtime"]))
        pdf.ln(3)
        head("DEDUCTIONS")
        for t, q, f, a in e["_ded"]:
            line(f"{t} ({q} x {f:g} day)", _inr(a))
        line("Total deductions", _inr(e["Deductions"]), True)
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(130, 9, "NET PAYABLE SALARY")
        pdf.cell(0, 9, _inr(e["Net"]), align="R", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(0, 6, "Computer-generated salary slip. Net = salary - deductions + overtime.")
    return bytes(pdf.output())