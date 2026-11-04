"""Port fidel a Python de l'Office Script de M365 (03_Office_Script_Parser.ts) per provar compatibilitat."""
import datetime, re
import openpyxl
SHIFT_ROWS = {"8-16": (4, 9), "16-20": (11, 16), "20-08": (18, 23)}


def parse_cell(s):
    s = s.strip()
    if not s or s.startswith("(") or s.startswith("Dissabte") or s.startswith("Diumenge"):
        return None
    if s == "FESTIU" or s.endswith(":"):
        return None
    cleaned = re.sub(r"^FESTIU\s*[—-]\s*Torn\s+(dia|nit)\s+\d+-\d+:?\s*", "", s, flags=re.I)
    pats = [(r"^Grup\s*1\s*:?\s*(.+)$", "G1", re.I), (r"^Grup\s*2\s*:?\s*(.+)$", "G2", re.I),
            (r"^N\s+i\s+B\s*:?\s*(.+)$", "NB", re.I), (r"^N\s*:?\s*(.+)$", "N", 0), (r"^B\s*[:.]?\s*(.+)$", "B", 0)]
    for rx, role, fl in pats:
        m = re.match(rx, cleaned, fl)
        if m and m.group(1):
            name = re.sub(r"\s*\([^)]*\)\s*$", "", m.group(1).strip()).strip()
            name = re.sub(r"^(Dra?\.?|Dr\.?)\s*", "", name, flags=re.I).strip()
            if name and not name.startswith("["):
                return role, name
    return None


def parse(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for sn in [s for s in wb.sheetnames if "setmana" in s.lower()]:
        ws = wb[sn]
        for col in range(7):
            raw = ws.cell(row=3, column=col + 2).value
            if not isinstance(raw, (datetime.date, datetime.datetime)):
                continue
            d = raw.date() if isinstance(raw, datetime.datetime) else raw
            for lbl, (a, b) in SHIFT_ROWS.items():
                for r in range(a, b + 1):
                    v = ws.cell(row=r, column=col + 2).value
                    if isinstance(v, str):
                        p = parse_cell(v)
                        if p:
                            out.append((d, lbl, p[0], p[1]))
    return out, (out[0][0].month if out else None)
