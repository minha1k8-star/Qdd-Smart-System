"""
Independent Python re-implementation of the QDD Smart System calculation
pipeline, built directly from the real Excel formulas documented in
docs/04_Algorithm_Specification.md (XU_LY_LENH, DOAN_CONG_SUAT, DIEN_TICH,
TINH_TOAN). Used to validate the documented algorithm against real
operational data + the independently hand-calculated reference workbook.
"""
import csv
import datetime
import glob
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from load_xlsx import load_patched

REPO_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))
BASE = os.path.join(REPO_ROOT, "test-data")

RAMP_RATE = 3.5          # CAI_DAT!B7, MW/phut
QDD_V_COEF = 0.9188       # CAI_DAT!B8
TOLERANCE = 0.03          # CAI_DAT!B9 (+-3%)
EPS = 0.000001

# ---------- CSV (Qdc / Qmp) ----------

def read_kwhgiao(path):
    """Return list of 48 floats from the KwhGiao row of a 6001/6303 CSV."""
    with open(path, newline="", encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 2 and row[1].strip().upper() == "KWHGIAO":
                vals = [float(x) for x in row[2:50]]
                if len(vals) != 48:
                    raise ValueError(f"{path}: expected 48 values, got {len(vals)}")
                return vals
    raise ValueError(f"{path}: KwhGiao row not found")


def find_csv(day, meter):
    pattern_day = f"{day:02d}"
    hits = glob.glob(os.path.join(BASE, "**", f"{pattern_day}07{meter}.CSV"), recursive=True)
    hits += glob.glob(os.path.join(BASE, "**", f"{pattern_day}07{meter}.csv"), recursive=True)
    hits = sorted(set(hits))
    if not hits:
        raise FileNotFoundError(f"No CSV for day={day} meter={meter}")
    return hits[0]


# ---------- Command list (LENH_GOC equivalent) ----------

COMMAND_FILES = [
    "DanhSachLenhKetThuc-01.xlsx",
    "DanhSachLenhKetThuc-02.xlsx",
    "DanhSachLenhKetThuc-03.xlsx",
    "DanhSachLenhKetThuc-06.xlsx",
    "DanhSachLenhKetThuc-07.xlsx",
    "DanhSachLenhKetThuc-10.xlsx",
    "DanhSachLenhKetThuc-20260722_132212931.xlsx",   # day 19
    "DanhSachLenhKetThuc-20260722_183728481KT.xlsx",  # days 16,17,18
]


def load_all_commands():
    """Load every LENH_GOC-shaped row from all DanhSachLenhKetThuc files."""
    rows = []
    seen = set()
    for fn in COMMAND_FILES:
        hits = glob.glob(os.path.join(BASE, "**", fn), recursive=True)
        if not hits:
            continue
        path = hits[0]
        if path in seen:
            continue
        seen.add(path)
        wb = load_patched(path)
        ws = wb["All"] if "All" in wb.sheetnames else wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=4, values_only=True):
            if r[0] is None:
                continue
            rows.append(r)
    return rows


def effective_commands(all_rows, target_date, unit):
    """Apply R01-R03 filter + power selection, return list of dicts sorted by BDTH."""
    out = []
    for r in all_rows:
        # columns per LENH_GOC: A ID,B NhaMay,C ToMay,D NoiDung,E CSraLenh,
        # F CShoanThanh,G BDTH,H hoanThanh,...,P Hoanthanh(bool/1),Q Dunglenh,...,Y NguonLenh
        to_may = (r[2] or "").strip().upper()[:2]
        bdth = r[6]
        hoanthanh = r[15]
        dung = r[16]
        cs_ra_lenh = r[4]
        cs_hoan_thanh = r[5]
        nguon = (r[24] or "").strip().upper()
        if bdth is None:
            continue
        if not isinstance(bdth, (datetime.datetime, datetime.date)):
            continue
        bdth_date = bdth.date() if isinstance(bdth, datetime.datetime) else bdth
        if bdth_date != target_date:
            continue
        if to_may != unit.upper()[:2]:
            continue
        # Q3 formula: VALUE(P) else TRUE/"TRUE" -> 1
        completed = 1 if (hoanthanh in (1, True, "TRUE", "True") or hoanthanh == 1) else 0
        if completed != 1:
            continue
        valid = False
        if nguon == "SO" and isinstance(cs_hoan_thanh, (int, float)) and cs_hoan_thanh > 0:
            valid = True
        if nguon == "MO" and isinstance(cs_ra_lenh, (int, float)) and cs_ra_lenh > 0:
            valid = True
        if not valid:
            continue
        stopped_early = dung in (True, "TRUE", "True")
        if nguon == "SO":
            p_hieu_luc = cs_hoan_thanh
        elif nguon == "MO" and stopped_early and isinstance(cs_hoan_thanh, (int, float)) and cs_hoan_thanh > 0:
            p_hieu_luc = cs_hoan_thanh
        else:
            p_hieu_luc = cs_ra_lenh
        seconds = (bdth - datetime.datetime.combine(bdth_date, datetime.time())).total_seconds()
        out.append({
            "id": r[0], "bdth": bdth, "seconds": seconds, "p": p_hieu_luc,
            "nguon": nguon, "raw": r,
        })
    out.sort(key=lambda x: x["seconds"])
    return out


# ---------- Ramp engine (XU_LY_LENH) ----------

def build_ramp_rows(cmds, p0):
    """Reproduce XU_LY_LENH B..K columns exactly per the documented formulas."""
    rows = []
    prev_F = p0
    prev_B = None
    prev_I = None
    prev_D = None
    for i, c in enumerate(cmds):
        B = c["seconds"]
        D = c["p"]
        if i == 0:
            F = p0
        else:
            if B >= prev_I:
                F = prev_D
            else:
                denom = max(prev_I - prev_B, EPS)
                F = prev_F + (prev_D - prev_F) * (B - prev_B) / denom
        if abs(D - F) < EPS:
            H = 0.0
        else:
            H = abs(D - F) / RAMP_RATE * 60.0
        I = B + H
        rows.append({"B": B, "D": D, "F": F, "H": H, "I": I})
        prev_F, prev_B, prev_I, prev_D = F, B, I, D
    return rows


# ---------- Segments (DOAN_CONG_SUAT) ----------

def build_segments(ramp_rows, p0):
    segs = []
    first_b = ramp_rows[0]["B"] if ramp_rows else 86400.0
    segs.append((0.0, min(first_b, 86400.0), p0, p0))
    for i, rr in enumerate(ramp_rows):
        next_b = ramp_rows[i + 1]["B"] if i + 1 < len(ramp_rows) else 86400.0
        ramp_end = min(rr["I"], next_b, 86400.0)
        if ramp_end > rr["B"]:
            segs.append((rr["B"], ramp_end, rr["F"], rr["D"]))
        hold_end = min(next_b, 86400.0)
        if hold_end > ramp_end:
            segs.append((ramp_end, hold_end, rr["D"], rr["D"]))
    return segs


# ---------- Area integration (DIEN_TICH -> TINH_TOAN Qdd) ----------

def cycle_qdd(segs, cyc_start, cyc_end):
    total = 0.0
    for (s, e, p_s, p_e) in segs:
        if e <= cyc_start or s >= cyc_end or e == s:
            continue
        ov_start = max(cyc_start, s)
        ov_end = min(cyc_end, e)
        if ov_end <= ov_start:
            continue
        def p_at(t):
            return p_s + (p_e - p_s) * (t - s) / (e - s)
        area = (p_at(ov_start) + p_at(ov_end)) / 2.0 * (ov_end - ov_start)
        total += area
    return total / 1800.0


def compute_day(cmds, p0):
    ramp_rows = build_ramp_rows(cmds, p0)
    segs = build_segments(ramp_rows, p0)
    qdd = []
    for i in range(48):
        cs, ce = i * 1800.0, (i + 1) * 1800.0
        qdd.append(cycle_qdd(segs, cs, ce))
    return qdd


def infer_p0(cmds, manual_qdd_period1):
    """Safe P0 inference: if first command starts at/after 00:30, period-1
    Qdd from the manual reference equals P0 exactly (pure HOLD segment)."""
    if not cmds or cmds[0]["seconds"] >= 1800 - EPS:
        return manual_qdd_period1, True
    return None, False
