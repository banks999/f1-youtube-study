"""Cross-check the Analysis sheet (and its helper columns on Clean) against independent pandas calculations.
Run after build_workbook.py and the LibreOffice recalculation:  python verify_analysis.py"""
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook

wb = load_workbook("f1_youtube_study.xlsx", data_only=True)
clean = pd.read_csv("data/clean_videos.csv", parse_dates=["pub_date"])
cal = pd.read_csv("data/f1_calendar.csv", parse_dates=["weekend_start", "race_date"])
fails = []


def check(name, ok):
    print(("PASS " if ok else "FAIL ") + name)
    if not ok:
        fails.append(name)


def close(a, b, tol=1e-9):
    return abs(float(a) - float(b)) <= tol * max(1.0, abs(float(b)))


# ---- derived columns, independently
clean["month"] = clean.pub_date.dt.year * 100 + clean.pub_date.dt.month
clean["rel"] = clean.views / clean.groupby(["channel_label", "format", "month"]).views.transform("median")
rd = cal.race_date.values.astype("datetime64[D]")
d = clean.pub_date.values.astype("datetime64[D]")
off = (d[:, None] - rd[None, :]).astype(int)
near = np.abs(off).argmin(axis=1)          # ties go to the earlier race, as in the sheet
clean["offset"] = off[np.arange(len(clean)), near]
clean["near_race"] = cal.race_name.values[near]


def is_hl(r):
    tl = r.title.lower()
    return int(r.group == "f1_official" and r.format == "Standard" and " highlights | 202" in tl
               and tl.startswith(("race", "sprint", "qualifying", "fp", "practice"))
               and "f2" not in tl and "f3" not in tl and "academy" not in tl)


clean["hl"] = clean.apply(is_hl, axis=1)

# ---- helper columns on Clean
ws = wb["Clean"]
rows = list(ws.iter_rows(min_row=2, min_col=19, max_col=23, values_only=True))
h = pd.DataFrame(rows, columns=["race_idx", "near_idx", "offset", "near_race", "hl"])
check("Clean: days_from_race_day", (h.offset.astype(int).values == clean.offset.values).all())
check("Clean: nearest_race", (h.near_race.values == clean.near_race.values).all())
check("Clean: highlight_package flag", (h.hl.astype(int).values == clean.hl.values).all())

# ---- Analysis sheet, located by section titles
a = wb["Analysis"]
def row_of(prefix, start=1):
    for r in range(start, a.max_row + 1):
        v = a.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    raise KeyError(prefix)

def cell(r, c):
    return a.cell(row=r, column=c).value

f1 = clean[clean.group == "f1_official"]
f1s = f1[f1.format == "Standard"]

# A. highlight kinds
rA = row_of("A. ")
for i, k in enumerate(["FP1", "FP2", "FP3", "Qualifying", "Sprint Qualifying", "Sprint", "Race"]):
    r = rA + 2 + i
    sub = f1s[(f1s.hl == 1) & f1s.title.str.lower().str.startswith((k + " highlights").lower())]
    check(f"A {k}: n", cell(r, 2) == len(sub))
    check(f"A {k}: median views", close(cell(r, 3), sub.views.median()))
r = row_of("Highlight packages (videos)")
hl = f1s[f1s.hl == 1]
check("A packages count", cell(r, 2) == len(hl))
check("A share of videos", close(cell(r + 1, 2), len(hl) / len(f1)))
check("A share of all views", close(cell(r + 2, 2), hl.views.sum() / f1.views.sum()))
check("A share of Standard views", close(cell(r + 3, 2), hl.views.sum() / f1s.views.sum()))
check("A median views packages", close(cell(r + 4, 2), hl.views.median()))
check("A median views other Standard", close(cell(r + 5, 2), f1s[f1s.hl == 0].views.median()))
check("A median views all Standard", close(cell(r + 6, 2), f1s.views.median()))
r = row_of("All Standard videos")
for j, w in enumerate(["Race weekend", "Post-race", "Other"]):
    sub = f1s[f1s.race_window == w]; sub2 = sub[sub.hl == 0]
    check(f"A rel all {w}", close(cell(r, 2 + j), sub.rel.median()) and cell(r, 5 + j) == len(sub))
    check(f"A rel without packages {w}", close(cell(r + 1, 2 + j), sub2.rel.median()) and cell(r + 1, 5 + j) == len(sub2))
rR = row_of("Race Highlights views, race by race")
races = cal[(cal.race_date >= clean.pub_date.min()) & (cal.race_date <= clean.pub_date.max())]
check("A race list length", len(races) == 23)
for i, rc in enumerate(races.itertuples()):
    sub = hl[(hl.title.str.startswith("Race Highlights")) & (hl.near_race == rc.race_name)]
    check(f"A Race Highlights views {rc.race_name}", cell(rR + 2 + i, 1) == rc.race_name and cell(rR + 2 + i, 3) == sub.views.sum())

# B. day by day
rB = row_of("B. ")
gk = ["f1_official", "team", "independent"]
nch = [clean[clean.group == g].channel_label.nunique() for g in gk]
ws_, we_ = pd.Timestamp("2025-11-17"), pd.Timestamp("2026-11-09")
for i, k in enumerate(range(-3, 4)):
    r = rB + 2 + i
    nr = ((cal.race_date >= ws_ - pd.Timedelta(days=k)) & (cal.race_date <= we_ - pd.Timedelta(days=k))).sum()
    for g, key in enumerate(gk):
        sub = clean[(clean.group == key) & (clean.offset == k)]
        subs = sub[sub.format == "Standard"]
        check(f"B uploads {key} day {k:+d}", close(cell(r, 2 + g), len(sub) / (nch[g] * nr)))
        check(f"B median rel {key} day {k:+d}", close(cell(r, 5 + g), subs.rel.median()) and cell(r, 8 + g) == len(subs))

# C. race by race
rC = row_of("C. ")
for i, rc in enumerate(races.itertuples()):
    r = rC + 2 + i
    for g, (key, exw) in enumerate([("f1_official", 0), ("team", 0), ("independent", 0), ("independent", 1)]):
        sub = clean[(clean.near_race == rc.race_name) & (clean.group == key) & (clean.format == "Standard") & (clean.race_window == "Race weekend")]
        if exw:
            sub = sub[sub.channel_label != "Peter Windsor"]
        got = cell(r, 3 + g)
        ok = cell(r, 7 + g) == len(sub) and ((got in (None, "")) if len(sub) < 3 else close(got, sub.rel.median()))
        check(f"C {rc.race_name} col{g}", ok)
r = row_of("Races where the median beat 1.00")
exp = []
for key, exw in [("f1_official", 0), ("team", 0), ("independent", 0), ("independent", 1)]:
    n = 0
    for rc in races.itertuples():
        sub = clean[(clean.near_race == rc.race_name) & (clean.group == key) & (clean.format == "Standard") & (clean.race_window == "Race weekend")]
        if exw:
            sub = sub[sub.channel_label != "Peter Windsor"]
        if len(sub) >= 3 and sub.rel.median() > 1:
            n += 1
    exp.append(n)
check("C races above 1.00 (22/15/18/13)", [cell(r, 3 + g) for g in range(4)] == exp)

# D. independents with / without Windsor
rD = row_of("D. ")
ind = clean[(clean.group == "independent") & (clean.format == "Standard")]
for i, exw in enumerate([0, 1]):
    sub0 = ind[ind.channel_label != "Peter Windsor"] if exw else ind
    for j, w in enumerate(["Race weekend", "Post-race", "Other"]):
        sub = sub0[sub0.race_window == w]
        check(f"D {'ex-Windsor' if exw else 'all'} {w}", close(cell(rD + 2 + i, 2 + j), sub.rel.median()) and cell(rD + 2 + i, 5 + j) == len(sub))

# E. Windsor
rE = row_of("E. ")
pw = clean[(clean.channel_label == "Peter Windsor") & (clean.format == "Standard")]
an = pw.title.str.contains("analysis", case=False)
for j, w in enumerate(["Race weekend", "Post-race", "Other"]):
    r = rE + 2 + j
    s1, s2 = pw[an & (pw.race_window == w)], pw[~an & (pw.race_window == w)]
    check(f"E analysis {w}", cell(r, 2) == len(s1) and close(cell(r, 3), s1.views.median()))
    check(f"E other {w}", cell(r, 4) == len(s2) and close(cell(r, 5), s2.views.median()))
check("E typical Standard video", close(cell(rE + 5, 3), pw.views.median()))

# F. concentration
rF = row_of("F. ")
r = rF + 2
while cell(r, 1) and not str(cell(r, 1)).startswith("Relative views"):
    ch = cell(r, 1)
    sub = clean[clean.channel_label == ch]
    n = len(sub); k = max(1, int(np.floor(n * 0.1 + 0.5)))
    thr = sub.views.sort_values(ascending=False).iloc[k - 1]
    check(f"F {ch}: top-10% share", close(cell(r, 5), sub[sub.views >= thr].views.sum() / sub.views.sum()))
    check(f"F {ch}: Shorts share of views", close(cell(r, 6), sub[sub.format == "Short"].views.sum() / sub.views.sum()))
    top = sub.loc[sub.rel.idxmax()]
    check(f"F {ch}: biggest hit", close(cell(r, 7), top.rel) and cell(r, 8) == top.title)
    r += 1

# G. post-race kinds
rG = row_of("G. ")
pr = f1s[f1s.race_window == "Post-race"]
t = pr.title
chs = t.str.contains("Chasing The Dream", case=False)
pod = t.str.contains("F1 Nation", case=False) | t.str.contains("Podcast", case=False)
pal = t.str.contains("Palmer", case=False)
masks = [chs, ~chs & pod, ~chs & ~pod & pal, ~chs & ~pod & ~pal]
for i, m in enumerate(masks):
    check(f"G kind {i}", cell(rG + 2 + i, 2) == int(m.sum()) and close(cell(rG + 2 + i, 3), pr[m].rel.median()))
check("G total", cell(rG + 6, 2) == len(pr))
allw = f1s.title
chs2 = allw.str.contains("Chasing The Dream", case=False)
pod2 = allw.str.contains("F1 Nation", case=False) | allw.str.contains("Podcast", case=False)
pal2 = allw.str.contains("Palmer", case=False)
masks2 = [chs2, ~chs2 & pod2, ~chs2 & ~pod2 & pal2, ~chs2 & ~pod2 & ~pal2]
for i, m in enumerate(masks2):
    check(f"G kind {i} all windows", cell(rG + 2 + i, 4) == int(m.sum()) and close(cell(rG + 2 + i, 5), f1s[m].rel.median()))

# H. reach per subscriber
rH = row_of("H. ")
r = rH + 2
nH = 0
while cell(r, 1) and not str(cell(r, 1)).startswith("Subscribers are as of"):
    ch = cell(r, 1)
    sub = clean[clean.channel_label == ch]
    check(f"H {ch}: subscribers", cell(r, 3) == sub.channel_subs.iloc[0])
    check(f"H {ch}: videos", cell(r, 4) == len(sub))
    check(f"H {ch}: median views", close(cell(r, 5), sub.views.median()))
    check(f"H {ch}: median views / subscribers", close(cell(r, 6), sub.views.median() / sub.channel_subs.iloc[0]))
    check(f"H {ch}: small-sample flag", (cell(r, 7) in (None, "")) == (len(sub) >= 30))
    nH += 1
    r += 1
check("H channel count", nH == 14)

# I. team channels and their own drivers
rI = row_of("I. ")
tt = clean.title.str.lower()
def has(names):
    m = pd.Series(False, index=clean.index)
    for nm in names:
        m = m | tt.str.contains(nm.lower(), regex=False)
    return m
any_w = pd.Series(False, index=clean.index)
all_names = []
for i in range(4):
    rr = rI + 2 + i
    ch, d1, d2 = cell(rr, 1), cell(rr, 2), cell(rr, 3)
    all_names += [d1, d2]
    isch = clean.channel_label == ch
    m = has([d1, d2])
    w, o = clean[isch & m], clean[isch & ~m]
    check(f"I {ch}: naming a driver", cell(rr, 4) == len(w) and close(cell(rr, 5), w.rel.median()))
    check(f"I {ch}: other videos", cell(rr, 6) == len(o) and close(cell(rr, 7), o.rel.median()))
    any_w = any_w | (isch & m)
rp = rI + 6
isteam = clean.group == "team"
tw, to = clean[any_w], clean[isteam & ~any_w]
check("I all four teams: naming a driver", cell(rp, 4) == len(tw) and close(cell(rp, 5), tw.rel.median()))
check("I all four teams: other videos", cell(rp, 6) == len(to) and close(cell(rp, 7), to.rel.median()))
m8 = has(all_names)
nw, no = clean[~isteam & m8], clean[~isteam & ~m8]
check("I contrast: naming a driver", cell(rp + 1, 4) == len(nw) and close(cell(rp + 1, 5), nw.rel.median()))
check("I contrast: other videos", cell(rp + 1, 6) == len(no) and close(cell(rp + 1, 7), no.rel.median()))

print("\nFAILURES:", fails if fails else "none")
sys.exit(1 if fails else 0)
