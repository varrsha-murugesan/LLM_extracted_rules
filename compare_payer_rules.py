#!/usr/bin/env python3
"""
Compare reimbursement rules from three payer rule workbooks by CPT/HCPCS code.

Usage:
    python compare_payer_rules.py --uhc UHC.xlsx --aetna Aetna.xlsx --humana Humana.xlsx [-o Payer_CPT_Rule_Comparison.xlsx]

Each input workbook must contain:
  - a rules sheet as its FIRST sheet, with columns: Policy ID, Rule ID, Rule Statement, LOB, SQL_Query
  - a 'CPT Code Rules' sheet, with columns: CPT/HCPCS Code, Policy ID, Rule ID, Rule Category

Output sheets: Summary, Coverage Matrix, Rule Types by Code, Side-by-Side Rules,
Payer-Specific Edits, Query Reuse Candidates, Category Map.
Requires: openpyxl (pip install openpyxl)
"""
import argparse, collections, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter as L

# ----------------------------------------------------------------- settings you may edit
# Source rule category -> common category. Add a line here if a workbook uses a new category.
CMAP = {
    'Frequency Limitation': 'Frequency',
    'Units / Quantity Limit': 'Units / Time', 'Time-Based / Unit Requirement': 'Units / Time',
    'Diagnosis Coding Requirement': 'Diagnosis / Indication', 'Diagnosis / Indication Requirement': 'Diagnosis / Indication',
    'Coverage Requirement': 'Coverage / Non-Covered', 'Non-Covered Service': 'Coverage / Non-Covered',
    'Modifier Requirement': 'Modifier', 'Modifier Exception': 'Modifier',
    'Code Relationship / Bundling': 'Bundling / Same-Day', 'Duplicate / Same-Day Billing': 'Bundling / Same-Day', 'Multiple Procedure': 'Bundling / Same-Day',
    'Payment Adjustment': 'Payment / Reimbursement', 'Reimbursement Limitation': 'Payment / Reimbursement',
    'Age Restriction': 'Age', 'Place of Service': 'Place of Service', 'Place of Service Restriction': 'Place of Service',
    'Provider Scope': 'Provider', 'Provider Eligibility': 'Provider', 'Provider Eligibility / Code Eligibility': 'Provider',
    'Code Level Selection': 'Level of Service', 'Level of Service': 'Level of Service', 'Global Surgery': 'Global Surgery',
    'Claim Submission': 'Claim Submission / Documentation', 'Documentation / Medical Necessity': 'Claim Submission / Documentation',
    'Code List': 'Code List / Scope', 'Code Identification': 'Code List / Scope', 'Applicability': 'Code List / Scope',
    'Definition': 'Code List / Scope', 'Cross-Policy Reference': 'Code List / Scope', 'Exception': 'Code List / Scope',
    'Rental / Purchase Limitation': 'Payment / Reimbursement', 'Replacement / Repair': 'Payment / Reimbursement',
    'Prior Authorization': 'Claim Submission / Documentation', 'NDC Requirement': 'Claim Submission / Documentation',
}
CATS = ['Frequency', 'Units / Time', 'Age', 'Diagnosis / Indication', 'Coverage / Non-Covered', 'Modifier', 'Place of Service',
        'Level of Service', 'Bundling / Same-Day', 'Global Surgery', 'Payment / Reimbursement', 'Provider',
        'Claim Submission / Documentation', 'Code List / Scope', 'Other']
LIMIT_CATS = CATS[:9]          # categories shown in Side-by-Side Rules
MAX_RULES_PER_CELL = 5         # rules listed per payer cell before "+ n more"
MAX_STATEMENT_CHARS = 400      # rule statements longer than this are cut with "..."

# ----------------------------------------------------------------- load
def load(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    if 'CPT Code Rules' not in wb.sheetnames:
        sys.exit(f"{path}: no 'CPT Code Rules' sheet")
    def rows(ws):
        it = ws.iter_rows(values_only=True); hdr = [str(h).strip() if h else '' for h in next(it)]
        return [dict(zip(hdr, r)) for r in it]
    cr = rows(wb['CPT Code Rules'])
    rules = {}
    for d in rows(wb[wb.sheetnames[0]]):
        if d.get('Policy ID') is None: continue
        rules[(d['Policy ID'], d['Rule ID'])] = d
    wb.close()
    return cr, rules

def sql_status(rule):
    s = str((rule or {}).get('SQL_Query') or '').strip()
    if s.startswith('--') or s.upper().startswith(('SELECT', 'WITH')): return 'SQL'
    if s.startswith('Not Feasible'): return 'Not Feasible'
    if s.startswith('Manual Review'): return 'Manual Review'
    return 'Unknown'

# ----------------------------------------------------------------- compare
def compare(data):
    P = list(data)
    unmapped = collections.Counter()
    idx = collections.defaultdict(lambda: {p: collections.defaultdict(list) for p in P})
    for p in P:
        for r in data[p]['cr']:
            code = str(r.get('CPT/HCPCS Code') or '').strip()
            if not code: continue
            src = r.get('Rule Category') or ''
            cat = CMAP.get(src)
            if cat is None: unmapped[(p, src)] += 1; cat = 'Other'
            key = (r.get('Policy ID'), r.get('Rule ID'))
            lst = idx[code][p][cat]
            if key not in lst: lst.append(key)
    has = lambda c, p: any(idx[c][p].values())
    shared = sorted([c for c in idx if sum(has(c, p) for p in P) >= 2], key=lambda c: (c[0].isalpha(), c))
    all3 = {c for c in shared if all(has(c, p) for p in P)}
    ctype = lambda c: 'HCPCS' if c[0].isalpha() else 'CPT'
    yes = lambda c: 'Yes' if c in all3 else 'No'

    def keys(c, p): return [k for lst in idx[c][p].values() for k in lst]
    def uniq(seq): return list(dict.fromkeys(seq))
    def fmt(p, ks):
        out = []
        for k in ks[:MAX_RULES_PER_CELL]:
            r = data[p]['rules'].get(k, {})
            s = str(r.get('Rule Statement') or '')
            if len(s) > MAX_STATEMENT_CHARS: s = s[:MAX_STATEMENT_CHARS] + ' ...'
            out.append(f"{k[0]} {k[1]} [{sql_status(r)}]: {s}")
        if len(ks) > MAX_RULES_PER_CELL: out.append(f"+ {len(ks) - MAX_RULES_PER_CELL} more rule(s) in this category")
        return "\n".join(out)

    V = dict(P=P, COV=[], TYP=[], SBS=[], SPEC=[], REUSE=[], unmapped=unmapped)
    for c in shared:
        cov = [p for p in P if has(c, p)]
        V['COV'].append([c, ctype(c), yes(c), len(cov), ', '.join(cov)] + [len(set(keys(c, p))) for p in P]
                        + ['; '.join(uniq(k[0] for k in keys(c, p))) for p in P]
                        + ['; '.join(sorted({str(data[p]['rules'].get(k, {}).get('LOB') or '') for k in keys(c, p)} - {''})) for p in P])
        V['TYP'].append([c, ctype(c), yes(c)] + [', '.join(p for p in P if idx[c][p].get(cat)) for cat in CATS])
        for cat in CATS:
            hv = [p for p in P if idx[c][p].get(cat)]
            if cat in LIMIT_CATS and len(hv) >= 2:
                V['SBS'].append([c, ctype(c), yes(c), cat, ', '.join(hv), ', '.join(p for p in cov if p not in hv) or 'None']
                                + [fmt(p, idx[c][p][cat]) if p in hv else '' for p in P])
            sq = {p: [k for k in idx[c][p].get(cat, []) if sql_status(data[p]['rules'].get(k)) == 'SQL'] for p in hv}
            sqp = [p for p in hv if sq[p]]
            if len(sqp) >= 2:
                V['REUSE'].append([c, ctype(c), cat, ', '.join(sqp)] + [fmt(p, sq[p]) if p in sqp else '' for p in P])
            if len(hv) == 1 and len(cov) >= 2 and cat != 'Code List / Scope':
                p = hv[0]
                V['SPEC'].append([c, ctype(c), yes(c), cat, p, ', '.join(q for q in cov if q != p), len(idx[c][p][cat]), fmt(p, idx[c][p][cat])])
    srcmap = collections.Counter((p, r.get('Rule Category') or '') for p in P for r in data[p]['cr'])
    V['MAP'] = sorted([[CMAP.get(k, 'Other'), p, k, n] for (p, k), n in srcmap.items()], key=lambda x: (CATS.index(x[0]), P.index(x[1]), x[2]))
    V['SRC'] = {
        'Rules (rule sheet rows)': {p: len(data[p]['rules']) for p in P},
        'Code-rule associations': {p: len(data[p]['cr']) for p in P},
        'Unique codes with rules': {p: sum(1 for c in idx if has(c, p)) for p in P},
        'Codes covered by this payer only': {p: sum(1 for c in idx if has(c, p) and sum(has(c, q) for q in P) == 1) for p in P},
    }
    return V

# ----------------------------------------------------------------- write
def write(V, files, out):
    P = V['P']
    HDR = PatternFill("solid", fgColor="1F3864"); HF = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    BF = Font(name="Arial", size=9); TF = Font(name="Arial", bold=True, size=14, color="1F3864"); SF = Font(name="Arial", bold=True, size=11, color="1F3864")
    wb = openpyxl.Workbook(); summ = wb.active; summ.title = "Summary"
    def sheet(name, cols, rows, widths, tname):
        ws = wb.create_sheet(name); ws.append(cols)
        for c in ws[1]: c.fill, c.font, c.alignment = HDR, HF, Alignment(wrap_text=True, vertical="center")
        for r in rows: ws.append(r)
        for row in ws.iter_rows(min_row=2):
            for c in row: c.font = BF; c.alignment = Alignment(wrap_text=True, vertical="top")
        for i, w in enumerate(widths, 1): ws.column_dimensions[L(i)].width = w
        ws.freeze_panes = "B2"
        if rows:
            t = Table(displayName=tname, ref=f"A1:{L(len(cols))}{len(rows) + 1}")
            t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True); ws.add_table(t)
    sheet("Coverage Matrix", ["CPT/HCPCS Code", "Code Type", "In All Three", "Number of Payers", "Payers"] + [f"{p} Rules" for p in P]
          + [f"{p} Policies" for p in P] + [f"{p} LOB" for p in P], V['COV'], [12, 9, 10, 10, 20, 9, 9, 9, 34, 34, 34, 22, 22, 22], "CoverageMatrix")
    sheet("Rule Types by Code", ["CPT/HCPCS Code", "Code Type", "In All Three"] + CATS, V['TYP'], [12, 9, 10] + [14] * len(CATS), "RuleTypesByCode")
    sheet("Side-by-Side Rules", ["CPT/HCPCS Code", "Code Type", "In All Three", "Common Category", "Payers With Rules In Category",
          "Covering Payers Without This Category"] + [f"{p} Rules" for p in P], V['SBS'], [12, 9, 10, 18, 18, 18, 60, 60, 60], "SideBySide")
    sheet("Payer-Specific Edits", ["CPT/HCPCS Code", "Code Type", "In All Three", "Common Category", "Only Payer With This Rule Type",
          "Other Payers Covering The Code", "Number of Rules", "Rules"], V['SPEC'], [12, 9, 10, 20, 14, 18, 10, 80], "PayerSpecific")
    sheet("Query Reuse Candidates", ["CPT/HCPCS Code", "Code Type", "Common Category", "Payers With SQL In Category"] + [f"{p} SQL Rules" for p in P],
          V['REUSE'], [12, 9, 18, 18, 60, 60, 60], "QueryReuse")
    sheet("Category Map", ["Common Category", "Payer", "Source Rule Category", "Code-Rule Rows"], V['MAP'], [30, 10, 40, 14], "CategoryMap")

    ws = summ
    for col, w in zip("ABCDEFGH", [34, 14, 14, 14, 16, 16, 16, 16]): ws.column_dimensions[col].width = w
    ws["A1"] = f"Payer CPT/HCPCS Rule Comparison: {', '.join(P)}"; ws["A1"].font = TF
    ws["A2"] = ("Compares the rules in the payer workbooks by CPT/HCPCS code; only codes with rules from at least two payers are compared. "
                "Check each payer's line of business (LOB columns) before treating a difference as a payer policy difference.")
    ws.merge_cells("A2:H2"); ws["A2"].alignment = Alignment(wrap_text=True, vertical="top"); ws["A2"].font = BF; ws.row_dimensions[2].height = 32
    r = 4; ws.cell(r, 1, "Source workbooks").font = SF; r += 1
    for c, h in enumerate(["Measure"] + P, 1): x = ws.cell(r, c, h); x.fill, x.font = HDR, HF
    for lab, d in V['SRC'].items():
        r += 1; ws.cell(r, 1, lab).font = BF
        for i, p in enumerate(P, 2): x = ws.cell(r, i, d[p]); x.font = Font(name="Arial", size=9, color="0000FF"); x.number_format = '#,##0'
    r += 1; ws.cell(r, 1, "Blue figures are counted from: " + "; ".join(f"{p} = {os.path.basename(f)}" for p, f in zip(P, files))).font = Font(name="Arial", size=8, italic=True)
    r += 2; ws.cell(r, 1, "Shared codes").font = SF; r += 1
    for c, h in enumerate(["Measure", "Count"], 1): x = ws.cell(r, c, h); x.fill, x.font = HDR, HF
    CM = "'Coverage Matrix'"
    rows = [("Codes covered by 2 or more payers", f"=COUNTA({CM}!A:A)-1"), ("Codes covered by all three payers", f'=COUNTIF({CM}!C:C,"Yes")')]
    for i in range(len(P)):
        for j in range(i + 1, len(P)):
            rows.append((f"{P[i]} and {P[j]} only", f'=COUNTIF({CM}!E:E,"{P[i]}, {P[j]}")'))
    rows += [("Side-by-side comparison rows", "=COUNTA('Side-by-Side Rules'!A:A)-1"), ("Payer-specific edit rows", "=COUNTA('Payer-Specific Edits'!A:A)-1"),
             ("Query reuse candidate rows", "=COUNTA('Query Reuse Candidates'!A:A)-1")]
    for lab, fm in rows:
        r += 1; ws.cell(r, 1, lab).font = BF; ws.cell(r, 2, fm).font = BF
    r += 2; ws.cell(r, 1, "Shared codes by common rule category").font = SF; r += 1
    hd = ["Common Category"] + [f"Shared codes with {p} rules" for p in P] + ["Codes where all payers have this category", "Side-by-side rows", "Payer-specific rows", "Query reuse rows"]
    for c, h in enumerate(hd, 1): x = ws.cell(r, c, h); x.fill, x.font, x.alignment = HDR, HF, Alignment(wrap_text=True)
    n = len(V['TYP']) + 1
    for i, cat in enumerate(CATS):
        r += 1; col = L(4 + i); rng = f"'Rule Types by Code'!{col}2:{col}{max(n, 2)}"
        ws.cell(r, 1, cat).font = BF
        for j, p in enumerate(P, 2): ws.cell(r, j, f'=COUNTIF({rng},"*{p}*")').font = BF
        ws.cell(r, 2 + len(P), f'=COUNTIF({rng},"{", ".join(P)}")').font = BF
        ws.cell(r, 3 + len(P), f"=COUNTIF('Side-by-Side Rules'!D:D,A{r})").font = BF
        ws.cell(r, 4 + len(P), f"=COUNTIF('Payer-Specific Edits'!D:D,A{r})").font = BF
        ws.cell(r, 5 + len(P), f"=COUNTIF('Query Reuse Candidates'!C:C,A{r})").font = BF
    wb.save(out)

def main():
    ap = argparse.ArgumentParser(description="Compare payer reimbursement rule workbooks by CPT/HCPCS code.")
    ap.add_argument("--uhc", required=True); ap.add_argument("--aetna", required=True); ap.add_argument("--humana", required=True)
    ap.add_argument("-o", "--output", default="Payer_CPT_Rule_Comparison.xlsx")
    a = ap.parse_args()
    files = {'UHC': a.uhc, 'Aetna': a.aetna, 'Humana': a.humana}
    data = {}
    for p, f in files.items():
        print(f"Reading {p}: {f}")
        cr, rules = load(f); data[p] = dict(cr=cr, rules=rules)
    V = compare(data)
    write(V, list(files.values()), a.output)
    print(f"Shared codes: {len(V['COV'])} | side-by-side rows: {len(V['SBS'])} | payer-specific rows: {len(V['SPEC'])} | query reuse rows: {len(V['REUSE'])}")
    if V['unmapped']:
        print("Rule categories not in CMAP (counted as 'Other'):")
        for (p, k), n in V['unmapped'].items(): print(f"  {p}: {k} ({n} rows)")
    print(f"Saved {a.output}  (formulas in the Summary sheet calculate when the file is opened in Excel)")

if __name__ == "__main__":
    main()
