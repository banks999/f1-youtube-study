"""Cross-check the recalculated workbook against independent pandas calculations."""
import sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook

wb = load_workbook("f1_youtube_study.xlsx", data_only=True)
clean = pd.read_csv("data/clean_videos.csv", parse_dates=["pub_date"])
ws = wb["Clean"]
rows = list(ws.iter_rows(min_row=2, max_col=18, values_only=True))
x = pd.DataFrame(rows, columns=["channel","group","subs","vid","title","pub","dur","views","likes","comments","age","format","eng","vps","window","month","base","rel"])
fails = []
def check(name, ok):
    print(("PASS " if ok else "FAIL ") + name)
    if not ok: fails.append(name)

check("row count", len(x) == len(clean))
check("video ids line up", (x.vid.values == clean.video_id.values).all())
check("age_days", (x.age.values == clean.age_days.values).all())
check("format", (x.format.values == clean.format.values).all())
check("race_window (Excel formulas vs Python rules)", (x.window.values == clean.race_window.values).all())
check("engagement_rate", np.allclose(x.eng.astype(float), clean.engagement_rate))
check("views_per_sub", np.allclose(x.vps.astype(float), clean.views_per_sub))
clean["month"] = clean.pub_date.dt.year*100 + clean.pub_date.dt.month
clean["base"] = clean.groupby(["channel_label","format","month"]).views.transform("median")
clean["rel"] = clean.views / clean.base
check("baseline median", np.allclose(x.base.astype(float), clean.base))
check("relative_views", np.allclose(x.rel.astype(float), clean.rel))

s = wb["Summary"]
def grid(r1, r2, c1, c2):
    return [[s.cell(row=r, column=c).value for c in range(c1, c2+1)] for r in range(r1, r2+1)]

# T1
import glob
channels = pd.concat([pd.read_csv(f) for f in sorted(glob.glob("data/channels_resolved*.csv"))])["label"].tolist()
channels = [c for c in channels if c in set(clean.channel_label)]
for i, ch in enumerate(channels):
    r = 5 + i
    d = clean[clean.channel_label == ch]; st = d[d.format == "Standard"]; sh = d[d.format == "Short"]
    exp = [d.group.iloc[0], d.channel_subs.iloc[0], len(d), (d.format=="Short").mean(), len(d)/(( (pd.Timestamp("2026-11-09")-pd.Timestamp("2025-11-17")).days+1)/7),
           st.views.median() if len(st) else "n/a", sh.views.median() if len(sh) else "n/a"]
    got = [s.cell(row=r, column=c).value for c in range(2, 9)]
    ok = all((abs(float(g)-float(e)) < 1e-6*max(1,abs(float(e)))) if not isinstance(e,str) or e=="n/a" and False else g == e for g, e in zip(got, exp)) if False else True
    ok = True
    for g, e in zip(got, exp):
        if isinstance(e, str): ok &= (g == e)
        else: ok &= abs(float(g) - float(e)) < 1e-6 * max(1, abs(float(e)))
    check(f"T1 {ch}", ok)
    eng = st.engagement_rate.median() if len(st) else "n/a"
    jv = s.cell(row=r, column=10).value
    check(f"T1 {ch} median engagement", (jv == "n/a") if isinstance(eng, str) else abs(float(jv)-eng) < 1e-9)

# T2/T3 located by labels
def find(text):
    for r in range(1, 120):
        v = s.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith(text): return r
t2 = find("2. When"); t3 = find("3. Do race"); t4 = find("4. The same"); t5 = find("5. Why")
days = pd.date_range("2025-11-17", "2026-11-09")
cal = pd.read_csv("data/f1_calendar.csv", parse_dates=["weekend_start","race_date"])
def win(d):
    for c in cal.itertuples():
        if c.weekend_start <= d <= c.race_date: return "Race weekend"
    for c in cal.itertuples():
        if c.race_date < d <= c.race_date + pd.Timedelta(days=3): return "Post-race"
    return "Other"
dw = pd.Series([win(d) for d in days]); ndays = dw.value_counts()
W = ["Race weekend","Post-race","Other"]
for i, (g, lab) in enumerate([("f1_official","F1"),("team","Teams"),("independent","Indep")]):
    r = t2 + 2 + i
    nch = clean[clean.group == g].channel_label.nunique()
    for j, w in enumerate(W):
        exp = ((clean.group == g) & (clean.race_window == w)).sum() / ndays[w] / nch
        check(f"T2 {lab} {w}", abs(float(s.cell(row=r, column=2+j).value) - exp) < 1e-9)
r = t2 + 5
for j, w in enumerate(W):
    check(f"T2 share of days {w}", abs(float(s.cell(row=r, column=2+j).value) - ndays[w]/len(days)) < 1e-9)

for k in range(6):
    r = t3 + 2 + k
    lab, fmt = s.cell(row=r, column=1).value, s.cell(row=r, column=2).value
    g = {"F1 official":"f1_official","Teams":"team","Independents":"independent"}[lab]
    for j, w in enumerate(W):
        sub = clean[(clean.group==g)&(clean.format==fmt)&(clean.race_window==w)]
        got, n = s.cell(row=r, column=3+j).value, s.cell(row=r, column=6+j).value
        ok = (n == len(sub)) and ((got == "n/a") if len(sub)==0 else abs(float(got) - sub.rel.median()) < 1e-9)
        check(f"T3 {lab} {fmt} {w} (n={len(sub)})", ok)

r = t4 + 2
while s.cell(row=r, column=1).value:
    ch, fmt = s.cell(row=r, column=1).value, s.cell(row=r, column=2).value
    for j, w in enumerate(W):
        sub = clean[(clean.channel_label==ch)&(clean.format==fmt)&(clean.race_window==w)]
        got, n = s.cell(row=r, column=3+j).value, s.cell(row=r, column=6+j).value
        ok = (n == len(sub)) and ((got == "n/a") if len(sub)==0 else abs(float(got) - sub.rel.median()) < 1e-9)
        check(f"T4 {ch} {fmt} {w} (n={len(sub)})", ok)
    r += 1

for j, w in enumerate(W):
    sub = clean[(clean.format=="Standard")&(clean.race_window==w)]
    check(f"T5 age {w}", abs(float(s.cell(row=t5+2, column=3+j).value) - sub.age_days.median()) < 1e-9)

print("\nFAILURES:", fails if fails else "none")
sys.exit(1 if fails else 0)
