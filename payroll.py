"""Attendance rules, salary calculation and salary-slip PDFs."""
import calendar
from pathlib import Path
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


LOGO = Path(__file__).with_name("logo.png")
BRAND = (219, 107, 39)      # orange from the logo
BRAND_LIGHT = (253, 243, 233)
GREY = (90, 98, 108)

_ONES = ("Zero One Two Three Four Five Six Seven Eight Nine Ten Eleven Twelve Thirteen Fourteen Fifteen "
         "Sixteen Seventeen Eighteen Nineteen").split()
_TENS = "_ _ Twenty Thirty Forty Fifty Sixty Seventy Eighty Ninety".split()


def _below100(n):
    return _ONES[n] if n < 20 else _TENS[n // 10] + (f" {_ONES[n % 10]}" if n % 10 else "")


def rupees_in_words(v):
    n = max(0, int(round(v)))
    if n == 0:
        return "Rupees Zero Only"
    parts = []
    for div, name in ((10_000_000, "Crore"), (100_000, "Lakh"), (1000, "Thousand"), (100, "Hundred")):
        q, n = divmod(n, div)
        if q:
            parts.append(f"{_below100(q)} {name}")
    if n:
        parts.append(_below100(n))
    return "Rupees " + " ".join(parts) + " Only"


def _inr(v):
    return f"{round(v):,}"


def slip_pdf(rows, period, work_days, company) -> bytes:
    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(False)
    L, W = 12, 186          # left margin, usable width
    for e in rows:
        pdf.add_page()
        # ---- header band -------------------------------------------------
        pdf.set_fill_color(*BRAND)
        pdf.rect(0, 0, 210, 4, "F")
        if LOGO.exists():
            pdf.image(str(LOGO), x=L, y=12, h=20)
        pdf.set_xy(L + 26, 13)
        pdf.set_font("Helvetica", "B", 18)
        pdf.set_text_color(30, 36, 48)
        pdf.cell(0, 9, company, new_x="LMARGIN", new_y="NEXT")
        pdf.set_x(L + 26)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*GREY)
        pdf.cell(0, 6, f"Salary Slip for the month of {period}")
        pdf.set_draw_color(*BRAND)
        pdf.set_line_width(0.6)
        pdf.line(L, 36, L + W, 36)

        # ---- employee details -------------------------------------------
        y = 43
        pdf.set_draw_color(215, 219, 224)
        pdf.set_line_width(0.2)
        pdf.set_fill_color(*BRAND_LIGHT)
        pdf.rect(L, y, W, 8, "DF")
        pdf.set_xy(L + 3, y + 1.5)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(*BRAND)
        pdf.cell(0, 5, "EMPLOYEE DETAILS")
        y += 8
        cells = [("Employee Name", e["Employee"]), ("Employee ID", e["ID"]),
                 ("Pay Period", period), ("Per-day Rate", "Rs. " + _inr(e["Rate"])),
                 ("Working Days", str(work_days)), ("Days Present", str(e["Present"])),
                 ("Absent Days", str(e["Absent"])), ("Half Days", str(e["HalfDays"])),
                 ("Late Marks", str(e["Late"])), ("Early Leavings", str(e["Early"]))]
        rh, half = 9, W / 2
        for i in range(0, len(cells), 2):
            for j in range(2):
                k, v = cells[i + j]
                x = L + j * half
                pdf.rect(x, y, half, rh)
                pdf.set_xy(x + 3, y + 1)
                pdf.set_font("Helvetica", "", 7.5)
                pdf.set_text_color(*GREY)
                pdf.cell(half - 6, 3.5, k.upper())
                pdf.set_xy(x + 3, y + 4.3)
                pdf.set_font("Helvetica", "B", 10)
                pdf.set_text_color(30, 36, 48)
                pdf.cell(half - 6, 4, v)
            y += rh

        # ---- earnings / deductions table ---------------------------------
        y += 8
        cw = [58, 35, 58, 35]
        pdf.set_fill_color(*BRAND)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_xy(L, y)
        for w, t, a in zip(cw, ["EARNINGS", "AMOUNT (Rs.)", "DEDUCTIONS", "AMOUNT (Rs.)"], "LRLR"):
            pdf.cell(w, 8, " " + t + " " if a == "L" else t + "  ", border=1, align=a, fill=True)
        y += 8
        earn = [("Monthly Salary", e["Salary"]), (f"Overtime ({e['OTDays']:g} day)", e["Overtime"])]
        ded = [(f"{t} ({q} x {f:g})" if q else t, a) for t, q, f, a in e["_ded"]]
        pdf.set_text_color(30, 36, 48)
        for i in range(max(len(earn), len(ded))):
            pdf.set_xy(L, y)
            pdf.set_font("Helvetica", "", 9.5)
            le, ld = (earn[i] if i < len(earn) else ("", None)), ded[i]
            pdf.cell(cw[0], 8, " " + le[0], border=1)
            pdf.cell(cw[1], 8, _inr(le[1]) + "  " if le[1] is not None else "", border=1, align="R")
            pdf.cell(cw[2], 8, " " + ld[0], border=1)
            pdf.cell(cw[3], 8, _inr(ld[1]) + "  ", border=1, align="R")
            y += 8
        gross = e["Salary"] + e["Overtime"]
        pdf.set_fill_color(*BRAND_LIGHT)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_xy(L, y)
        for w, t, a in zip(cw, [" Gross Earnings", _inr(gross) + "  ", " Total Deductions", _inr(e["Deductions"]) + "  "], "LRLR"):
            pdf.cell(w, 9, t, border=1, align=a, fill=True)
        y += 9

        # ---- net pay box --------------------------------------------------
        y += 8
        pdf.set_fill_color(*BRAND)
        pdf.rect(L, y, W, 16, "F")
        pdf.set_xy(L + 5, y + 4.5)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(90, 7, "NET PAYABLE SALARY")
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(W - 100, 7, "Rs. " + _inr(e["Net"]), align="R")
        y += 16
        pdf.set_draw_color(215, 219, 224)
        pdf.rect(L, y, W, 9)
        pdf.set_xy(L + 4, y + 2)
        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(*GREY)
        pdf.cell(W - 8, 5, "In words: " + rupees_in_words(e["Net"]))

        # ---- notes & signatures -------------------------------------------
        y += 17
        pdf.set_xy(L, y)
        pdf.set_font("Helvetica", "", 8)
        pdf.multi_cell(W, 4.2, "Net salary = Monthly salary - Deductions + Overtime. Per-day rate is "
                       "applied to absent days, half days (0.5), late marks and early leavings beyond the "
                       "monthly allowance (0.5 each), and overtime blocks of 3.5 hours (0.5 day each).")
        y = 258
        pdf.set_draw_color(120, 126, 134)
        for x, t in ((L, "Employer Signature"), (L + W - 60, "Employee Signature")):
            pdf.line(x, y, x + 60, y)
            pdf.set_xy(x, y + 1)
            pdf.set_font("Helvetica", "", 8.5)
            pdf.cell(60, 5, t, align="C")
        pdf.set_fill_color(*BRAND)
        pdf.rect(0, 291, 210, 6, "F")
        pdf.set_xy(0, 292)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(210, 4, "This is a computer-generated salary slip and does not require a signature.", align="C")
    return bytes(pdf.output())