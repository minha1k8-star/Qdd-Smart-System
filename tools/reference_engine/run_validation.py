import datetime
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from load_xlsx import load_patched
import qdd_engine as eng

BASE = eng.BASE
REF_FILE = os.path.join(BASE, "Kiểm tra Qdu ngày 200726.xlsx")

DAYS = [1, 2, 3, 6, 7, 10, 16, 17, 18, 19]
UNITS = ["S1", "S2"]


def ref_sheet_name(day):
    return f"{day:02d}.07.2026"


def load_ref(wb, day):
    name = ref_sheet_name(day)
    if name not in wb.sheetnames:
        cands = [n for n in wb.sheetnames if n.strip() == name]
        if not cands:
            return None
        name = cands[0]
    ws = wb[name]
    data = {"S1": [], "S2": []}
    for row in range(4, 52):
        b = ws.cell(row=row, column=2).value   # B: Qdd S1
        c = ws.cell(row=row, column=3).value   # C: Qdd_V S1
        f = ws.cell(row=row, column=6).value   # F: Qdu S1
        l = ws.cell(row=row, column=12).value  # L: Qdd S2
        m = ws.cell(row=row, column=13).value  # M: Qdd_V S2
        p = ws.cell(row=row, column=16).value  # P: Qdu S2
        data["S1"].append((b, c, f))
        data["S2"].append((l, m, p))
    return data


def main():
    all_rows = eng.load_all_commands()
    wb = load_patched(REF_FILE)

    for day in DAYS:
        the_date = datetime.date(2026, 7, day)
        ref = load_ref(wb, day)
        if ref is None:
            print(f"Day {day:02d}: NO REFERENCE SHEET FOUND")
            continue
        csv6001 = eng.read_kwhgiao(eng.find_csv(day, "6001"))
        csv6303 = eng.read_kwhgiao(eng.find_csv(day, "6303"))

        for unit in UNITS:
            ref_u = ref[unit]
            ref_qdd1 = ref_u[0][0]
            if not isinstance(ref_qdd1, (int, float)):
                continue  # no ref data for this unit that day
            cmds = eng.effective_commands(all_rows, the_date, unit)
            if not cmds and unit == "S2":
                # no commands captured for S2 on days without a command file
                continue
            p0, safe = eng.infer_p0(cmds, ref_qdd1)
            if not safe:
                print(f"Day {day:02d} {unit}: UNSAFE P0 inference (first cmd within period 1) - SKIPPED, needs manual P0")
                continue
            qdd = eng.compute_day(cmds, p0)

            diffs = []
            for i in range(48):
                ref_val = ref_u[i][0]
                if not isinstance(ref_val, (int, float)):
                    continue
                diffs.append(abs(qdd[i] - ref_val))
            if not diffs:
                print(f"Day {day:02d} {unit}: no comparable periods")
                continue
            max_diff = max(diffs)
            n_over_01 = sum(1 for d in diffs if d > 0.1)
            print(f"Day {day:02d} {unit}: n_cmds={len(cmds):2d} P0={p0:7.2f} "
                  f"max|dQdd|={max_diff:8.4f} MW  periods>0.1MW={n_over_01}/{len(diffs)}")
            if max_diff > 0.1:
                scored = [(abs(qdd[i] - ref_u[i][0]), i) for i in range(48)
                          if isinstance(ref_u[i][0], (int, float))]
                worst = [i for _, i in sorted(scored, reverse=True)[:5]]
                for i in worst:
                    rv = ref_u[i][0]
                    print(f"    chu ky {i+1:02d} [{i*30//60:02d}:{i*30%60:02d}-...]: "
                          f"python={qdd[i]:.4f}  ref={rv:.4f}  diff={qdd[i]-rv:+.4f}")


if __name__ == "__main__":
    main()
