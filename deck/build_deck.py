import json, os
from decimal import Decimal, ROUND_HALF_UP
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ART = os.path.join(HERE, "out")  # slides, deck.css and deck.json are written here
os.makedirs(f"{ART}/slides", exist_ok=True)
wb = openpyxl.load_workbook(os.path.join(HERE, "..", "f1_youtube_study.xlsx"), data_only=True)
S, A = wb["Summary"], wb["Analysis"]

def row_of(ws, prefix):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    raise KeyError(prefix)
def cv(ws, r, c): return ws.cell(row=r, column=c).value

# ---- number formatting: half-up on the decimal value, like the workbook's display
def rnd(x, d): return Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
def pct(x, d=0): return f"{rnd(Decimal(repr(float(x))) * 100, d)}%"
def mult(x, d=2): return f"{rnd(x, d)}×"
def num(x, d=2): return f"{rnd(x, d)}"
def k(v):
    if v >= 1e6: return f"{rnd(v / 1e6, 1)}M"
    if v >= 1e4: return f"{rnd(v / 1e3, 0)}K"
    return f"{rnd(v / 1e3, 1)}K"

# ---- data from the workbook
score = [dict(name=cv(S, r, 1), group=cv(S, r, 2), subs=cv(S, r, 3), kept=cv(S, r, 4), shorts=cv(S, r, 5)) for r in range(5, 19)]
t2 = row_of(S, "2. When")
cad = {cv(S, t2 + 2 + i, 1): [cv(S, t2 + 2 + i, c) for c in (2, 3, 4, 5)] for i in range(3)}
days_share = [cv(S, t2 + 5, c) for c in (2, 3, 4)]
t3 = row_of(S, "3. Do race")
team_std = [cv(S, t3 + 3, c) for c in (3, 4, 5)]
assert cv(S, t3 + 3, 1) == "Teams" and cv(S, t3 + 3, 2) == "Standard"
t4 = row_of(S, "4. The same")
chan = {}
r = t4 + 2
while cv(S, r, 1):
    if cv(S, r, 2) == "Standard": chan[cv(S, r, 1)] = (cv(S, r, 3), cv(S, r, 6))
    r += 1
rA = row_of(A, "A. ")
kinds = {cv(A, rA + 2 + i, 1): (cv(A, rA + 2 + i, 2), cv(A, rA + 2 + i, 3)) for i in range(7)}
rw = row_of(A, "Highlight packages (videos)")
hl_n, hl_sh_videos, hl_sh_views, hl_sh_std = [cv(A, rw + i, 2) for i in range(4)]
hl_med, other_med, all_std_med = [cv(A, rw + 4 + i, 2) for i in range(3)]
rr_ = row_of(A, "All Standard videos")
f1_all = [cv(A, rr_, 2 + j) for j in range(3)]
f1_nohl = [cv(A, rr_ + 1, 2 + j) for j in range(3)]
rB = row_of(A, "B. ")
daysB = [[cv(A, rB + 2 + i, c) for c in range(1, 11)] for i in range(7)]
ordinary = [cv(A, rB + 9, c) for c in (2, 3, 4)]
rC = row_of(A, "Races where the median beat 1.00")
beat = [cv(A, rC, 3 + g) for g in range(4)]
rD = row_of(A, "D. ")
ind_all = [cv(A, rD + 2, 2 + j) for j in range(3)]; ind_all_n = [cv(A, rD + 2, 5 + j) for j in range(3)]
ind_exw = [cv(A, rD + 3, 2 + j) for j in range(3)]; ind_exw_n = [cv(A, rD + 3, 5 + j) for j in range(3)]
rE = row_of(A, "E. ")
an_rw_n, an_rw_med = cv(A, rE + 2, 2), cv(A, rE + 2, 3)
an_post_n, an_post_med = cv(A, rE + 3, 2), cv(A, rE + 3, 3)
pw_typ = cv(A, rE + 5, 3)
rF = row_of(A, "F. ")
conc = {}
r = rF + 2
while cv(A, r, 1) and not str(cv(A, r, 1)).startswith("Relative views"):
    conc[cv(A, r, 1)] = dict(top10=cv(A, r, 5), shorts=cv(A, r, 6), hit=cv(A, r, 7), title=cv(A, r, 8)); r += 1
rG = row_of(A, "G. ")
gk = [(cv(A, rG + 2 + i, 1), cv(A, rG + 2 + i, 2), cv(A, rG + 2 + i, 3), cv(A, rG + 2 + i, 4), cv(A, rG + 2 + i, 5)) for i in range(4)]
post_total = cv(A, rG + 6, 2)
rH = row_of(A, "H. ")
reach = {}
r = rH + 2
while cv(A, r, 1) and not str(cv(A, r, 1)).startswith("Subscribers are as of"):
    reach[cv(A, r, 1)] = dict(subs=cv(A, r, 3), n=cv(A, r, 4), med=cv(A, r, 5), ratio=cv(A, r, 6)); r += 1
rI = row_of(A, "I. ")
own = [dict(team=cv(A, rI + 2 + i, 1), d1=cv(A, rI + 2 + i, 2), d2=cv(A, rI + 2 + i, 3), n_w=cv(A, rI + 2 + i, 4),
            m_w=cv(A, rI + 2 + i, 5), n_o=cv(A, rI + 2 + i, 6), m_o=cv(A, rI + 2 + i, 7)) for i in range(4)]
own_all = dict(n_w=cv(A, rI + 6, 4), m_w=cv(A, rI + 6, 5), n_o=cv(A, rI + 6, 6), m_o=cv(A, rI + 6, 7))
own_ctr = dict(n_w=cv(A, rI + 7, 4), m_w=cv(A, rI + 7, 5), n_o=cv(A, rI + 7, 6), m_o=cv(A, rI + 7, 7))
tot_videos = sum(c["kept"] for c in score)

RW, PR, OT = "#F50537", "#0B0C0E", "#79808A"



# ---------------------------------------------------------------- building blocks
def eb(t): return f'<div class="eb"><div class="slash"></div><p class="eyebrow">{t}</p></div>'
def foot(t): return f'<div class="rule"></div><p class="src">{t}</p>'
def deco(): return ('<div class="dec" style="left:1648px;top:128px;width:40px;height:96px;background:#F50537"></div>'
                    '<div class="dec" style="left:1712px;top:128px;width:40px;height:96px;background:#AEB4BB"></div>')
def page(id_, ebt, title, body, src, dark=False, big=False):
    h = f'<h2{" class=\"big\"" if big else ""}>{title}</h2>'
    return (f'<section id="{id_}"{" class=\"dark\"" if dark else ""}>\n{deco()}\n{eb(ebt)}\n{h}\n{body}\n{foot(src)}\n</section>\n')
def legend(items, extra=""):
    o = '<div class="legend">'
    for c, l in items:
        sw = '<div class="lsw"></div>' if c == "line" else f'<div class="sw" style="background:{c}"></div>'
        o += f'<div class="it">{sw}<p>{l}</p></div>'
    return o + extra + "</div>"
def mega(num, cap, cls="mg2", white=False):
    return (f'<div class="col"><p class="mega {cls}{" mw" if white else ""}">{num}</p><p class="mcap">{cap}</p></div>')
slides = {}
WH = "#EDEFF1"

# ---------------------------------------------------------------- 1 cover
slides["cover"] = f"""<section id="cover" class="dark">
<div class="slash" style="position:absolute;left:128px;top:136px;width:64px;height:16px"></div>
<p class="eyebrow" style="position:absolute;left:216px;top:124px;width:900px;height:36px">A CREATOR-DATA STUDY</p>
<h1 class="ctitle" style="position:absolute;left:128px;top:210px;width:1400px;height:170px">F1 on YouTube</h1>
<p class="cq" style="position:absolute;left:128px;top:410px;width:1300px;height:130px">Does the race calendar change what F1’s channels upload, and what gets watched?</p>
<div style="position:absolute;left:1560px;top:136px;width:40px;height:360px;background:#F50537"></div>
<div style="position:absolute;left:1624px;top:136px;width:40px;height:360px;background:#AEB4BB"></div>
<div style="position:absolute;left:1688px;top:136px;width:40px;height:360px;background:#2A2D33"></div>
<div class="crule" style="position:absolute;left:128px;top:650px;width:1664px;height:2px"></div>
<div class="stats" style="position:absolute;left:128px;top:676px;width:1664px;height:180px">
<div class="stat"><p class="snum">{len(score)}</p><p class="scap">channels: the sport’s, four teams, nine independents</p></div>
<div class="stat"><p class="snum">{tot_videos:,}</p><p class="scap">videos analysed</p></div>
<div class="stat"><p class="snum">23</p><p class="scap">race weekends, Nov 2025 to Nov 2026</p></div>
</div>
<p class="byline" style="position:absolute;left:128px;top:890px;width:1664px;height:40px">Jugal Shah · SDBI, Mumbai</p>
</section>
"""

# ---------------------------------------------------------------- 2 statement
lift = [cad["F1 official"][3], cad["Teams"][3], cad["Independents"][3]]
slides["statement"] = f"""<section id="statement" class="dark">
{deco()}
{eb("THE SHORT VERSION")}
<h2 class="stmt">Race weekends change posting.<br>Only a few videos get the views.</h2>
<div class="cols3">
<div class="big3"><div class="b3rule"></div><p class="b3n">{mult(lift[0], 1)}</p><p class="b3c">more uploads a day from the sport’s channel on race weekends</p></div>
<div class="big3"><div class="b3rule"></div><p class="b3n">{pct(hl_sh_views, 1)}</p><p class="b3c">of its views come from just {hl_n} videos</p></div>
<div class="big3"><div class="b3rule"></div><p class="b3n">{mult(ind_exw[0])}</p><p class="b3c">independents’ race-weekend lift once Peter Windsor is out</p></div>
</div>
{foot("Source: f1_youtube_study.xlsx. Method, limits and every figure are in the report.")}
</section>
"""

# ---------------------------------------------------------------- 3 rhythm
def host_days(vals, ordv, dec):
    base, mxh, pitch, bw, W = 250, 190, 76, 56, 528
    scale = mxh / max(vals)
    cols = [OT, RW, RW, RW, PR, PR, PR]
    labs = ["−3", "−2", "−1", "0", "+1", "+2", "+3"]
    p = [f'<div class="ref" style="top:{round(base - ordv * scale)}px;width:{W}px"></div>',
         f'<div class="base" style="top:{base}px;width:{W}px"></div>']
    for i, (v, c) in enumerate(zip(vals, cols)):
        h = round(v * scale); l = i * pitch
        p.append(f'<div class="bar" style="left:{l + 8}px;top:{base - h}px;width:{bw}px;height:{h}px;background:{c}"></div>')
        p.append(f'<p class="val" style="left:{l}px;top:{base - h - 36}px;width:72px">{num(v, dec)}</p>')
        p.append(f'<p class="xl" style="left:{l}px;top:{base + 8}px;width:72px">{labs[i]}</p>')
    return f'<div class="host" style="width:{W}px;height:296px">' + "\n".join(p) + "</div>"
nm = ["Sport’s channel", "Teams", "Independents"]
panels = "".join(
    f'<div class="panel"><p class="ph1">{nm[g]}</p><p class="mega mg3">{mult(lift[g], 1)}</p>'
    f'<p class="pcap">more uploads a day on race weekends</p>'
    f'{host_days([daysB[i][1 + g] for i in range(7)], ordinary[g], (1, 1, 2)[g])}</div>' for g in range(3))
slides["rhythm"] = page("rhythm", "UPLOAD RHYTHM", "Everyone posts more around race day",
    legend([(RW, "Race weekend"), (PR, "Post-race"), (OT, "Other days"), ("line", "Ordinary-day level")])
    + f'<div class="hrow3">{panels}</div>',
    "Source: workbook, Analysis B. Bars are uploads per channel per day; axis is days from race day, 0 = race day.")

# ---------------------------------------------------------------- 4 highlights (dark)
order = ["Race", "Qualifying", "Sprint", "Sprint Qualifying", "FP2", "FP1", "FP3"]
lab_map = {"Race": "Race highlights", "Qualifying": "Qualifying highlights", "Sprint": "Sprint highlights",
           "Sprint Qualifying": "Sprint qualifying", "FP2": "Practice 2", "FP1": "Practice 1", "FP3": "Practice 3"}
mx = max(kinds[kk][1] for kk in order)
x0, maxw, pitch = 300, 420, 58
def hrow(i, lab, v, col):
    w = max(4, round(v / mx * maxw))
    return (f'<p class="cn" style="left:0px;top:{i*pitch+6}px;width:290px">{lab}</p>'
            f'<div class="bar" style="left:{x0}px;top:{i*pitch+8}px;width:{w}px;height:32px;background:{col}"></div>'
            f'<p class="cv" style="left:{x0+w+14}px;top:{i*pitch+6}px;width:100px">{k(v)}</p>')
rowsA = "\n".join(hrow(i, lab_map[kk], kinds[kk][1], RW) for i, kk in enumerate(order))
hostH = (f'<div class="host" style="width:840px;height:{pitch*7}px">{rowsA}</div>'
         f'<div style="width:840px;height:2px;background:#2A2D33;margin:6px 0"></div>'
         f'<div class="host" style="width:840px;height:{pitch}px">{hrow(0, "All other videos", other_med, OT)}</div>')
slides["highlights"] = page("highlights", "THE SPORT’S CHANNEL", "Highlight videos carry the weekend",
    f'<div class="cols" style="gap:64px"><div style="width:760px;flex-shrink:0">{mega(pct(hl_sh_views, 1), f"of its views come from {hl_n} highlight videos, only {pct(hl_sh_videos, 1)} of its uploads", "mg1")}</div>'
    f'<div style="width:840px;flex-shrink:0"><p class="ch">Median views per video, Standard videos</p>{hostH}</div></div>',
    "Source: workbook, Analysis A. Highlight videos are race-weekend session highlights.", dark=True)

# ---------------------------------------------------------------- 5 post-race (light)
rowsP = [("F2 “Chasing The Dream”", gk[0][2]), ("Podcasts", gk[1][2]), ("Jolyon Palmer analysis", gk[2][2]), ("Everything else", gk[3][2])]
sc, x0, pitch = 340, 330, 96
p = [f'<div class="vref" style="left:{x0+sc}px;top:0px;height:{pitch*4}px"></div>']
for i, (lab, v) in enumerate(rowsP):
    w = max(4, round(v * sc)); col = RW if v >= 1 else OT
    p.append(f'<p class="cn" style="left:0px;top:{i*pitch+26}px;width:320px">{lab}</p>')
    p.append(f'<div class="bar" style="left:{x0}px;top:{i*pitch+18}px;width:{w}px;height:52px;background:{col}"></div>')
    p.append(f'<p class="cv" style="left:{x0+max(w, sc)+18}px;top:{i*pitch+26}px;width:130px">{mult(v)}</p>')
p.append(f'<p class="cl" style="left:{x0+sc-50}px;top:{pitch*4+6}px;width:100px">1.00</p>')
hostP = f'<div class="host" style="width:1000px;height:{pitch*4+40}px">' + "\n".join(p) + "</div>"
three_n = sum(g_[1] for g_ in gk[:3])
slides["postrace"] = page("postrace", "THE SPORT’S CHANNEL, AFTER THE RACE", "After a race, the dip is what gets posted",
    f'<div class="cols" style="gap:64px"><div style="width:1000px;flex-shrink:0"><p class="ch">Median relative views, Standard videos posted in the 3 days after a race</p>{hostP}'
    f'<p class="cap">{three_n} of the {post_total} post-race videos are the three kinds above.</p></div>'
    f'<div class="col" style="width:600px;flex-shrink:0">{mega(mult(f1_all[1]), "all post-race Standard videos", "mg2", True)}<div style="height:30px"></div>{mega(mult(gk[3][2]), "everything except the first three kinds", "mg2")}</div></div>',
    "Source: workbook, Analysis G. Kinds are identified by words in the title.")

# ---------------------------------------------------------------- 6 independents (dark)
def g(nm_): return chan[nm_]
rowsI = [("Peter Windsor", g("Peter Windsor")[0], RW), ("TacticalRab", g("F1 News TacticalRab")[0], OT),
         ("Kym Illman", g("Kym Illman")[0], OT), ("Cameron Cc", g("Cameron Cc")[0], OT)]
sc, x0, pitch = 150, 250, 100
p = [f'<div class="vref" style="left:{x0+sc}px;top:0px;height:{pitch*4}px"></div>']
for i, (lab, v, col) in enumerate(rowsI):
    w = max(4, round(v * sc))
    p.append(f'<p class="cn" style="left:0px;top:{i*pitch+30}px;width:240px">{lab}</p>')
    p.append(f'<div class="bar" style="left:{x0}px;top:{i*pitch+22}px;width:{w}px;height:52px;background:{col}"></div>')
    p.append(f'<p class="cv" style="left:{x0+max(w, sc)+18}px;top:{i*pitch+30}px;width:120px">{mult(v)}</p>')
p.append(f'<p class="cl" style="left:{x0+sc-50}px;top:{pitch*4+6}px;width:100px">1.00</p>')
hostI = f'<div class="host" style="width:900px;height:{pitch*4+40}px">' + "\n".join(p) + "</div>"
slides["independents"] = page("independents", "INDEPENDENT CREATORS", "Among independents, the lift is one channel",
    f'<div class="cols" style="gap:64px"><div class="col" style="width:700px;flex-shrink:0">{mega(mult(ind_all[0]), "all independents", "mg2", True)}<div style="height:24px"></div>{mega(mult(ind_exw[0]), "without Peter Windsor", "mg2")}</div>'
    f'<div style="width:900px;flex-shrink:0"><p class="ch">Race-weekend relative views, Standard videos</p>{hostI}</div></div>',
    "Source: workbook, Summary table 4 and Analysis D. Race weekend = first practice day to race day.", dark=True)

# ---------------------------------------------------------------- 7 teams (light): biggest hits
tm = [("Red Bull", "Red Bull Racing", "“please don’t try this at home…”"), ("Williams", "Williams", "Carlos Sainz takes on Ansu Fati"),
      ("Mercedes", "Mercedes-AMG F1", "“age is just experience”"), ("McLaren", "McLaren", "“When Adele met the Papaya team”")]
mxh_, base_, cw, cgap = 300, 400, 232, 24
mxv = max(conc[key]["hit"] for _, key, _ in tm)
p = [f'<div class="base" style="top:{base_}px;width:1000px"></div>']
for i, (lab, key, hit) in enumerate(tm):
    v = conc[key]["hit"]; h = max(6, round(v / mxv * mxh_)); l = i * (cw + cgap)
    p.append(f'<div class="bar" style="left:{l}px;top:{base_ - h}px;width:{cw}px;height:{h}px;background:{PR}"></div>')
    p.append(f'<p class="bn" style="left:{l}px;top:{base_ - h - 70}px;width:{cw}px;height:64px">{str(rnd(v, 0))}×</p>')
    p.append(f'<p class="cn" style="left:{l}px;top:{base_ + 12}px;width:{cw}px;height:36px"><b>{lab}</b></p>')
    p.append(f'<p class="sub" style="left:{l}px;top:{base_ + 50}px;width:{cw}px;height:64px">{hit}</p>')
t10 = [conc[key]["top10"] for _, key, _ in tm]
top10_rng = f"{rnd(min(t10) * 100, 0)}–{pct(max(t10))}"
hostT = f'<div class="host" style="width:1000px;height:520px">' + "\n".join(p) + "</div>"
slides["teams"] = page("teams", "THE TEAMS", "Team views come from hits, not the calendar",
    f'<div class="cols" style="gap:64px"><div style="width:1000px;flex-shrink:0"><p class="ch">Each team’s biggest hit, as a multiple of its normal video</p>{hostT}</div>'
    f'<div class="col" style="width:600px;flex-shrink:0;margin-top:40px">{mega(top10_rng, "of a team’s views come from its top 10% of videos", "mg3")}<div style="height:44px"></div>{mega(mult(team_std[0]), f"team videos on race weekends, against {mult(team_std[2])} on ordinary days", "mg3", True)}</div></div>',
    "Source: workbook, Analysis F and Summary table 3. Biggest hit counts Shorts and Standard videos.")

# ---------------------------------------------------------------- 8 drivers (dark)
short = {"Red Bull Racing": "Red Bull", "Mercedes-AMG F1": "Mercedes", "McLaren": "McLaren", "Williams": "Williams"}
sc, x0, pitch = 220, 300, 104
def drv_host(items):
    p = [f'<div class="vref" style="left:{x0+sc}px;top:0px;height:{pitch*len(items)}px"></div>']
    for i, d in enumerate(items):
        w1 = max(4, round(d["m_w"] * sc)); w2 = max(4, round(d["m_o"] * sc)); t0 = i * pitch
        p.append(f'<p class="cn" style="left:0px;top:{t0+10}px;width:290px">{short[d["team"]]}</p>')
        p.append(f'<p class="sub" style="left:0px;top:{t0+44}px;width:290px">{d["d1"]}, {d["d2"]}</p>')
        p.append(f'<div class="bar" style="left:{x0}px;top:{t0+8}px;width:{w1}px;height:38px;background:{RW}"></div>')
        p.append(f'<p class="cv" style="left:{x0+max(w1, sc)+16}px;top:{t0+8}px;width:140px">{mult(d["m_w"])}</p>')
        p.append(f'<div class="bar" style="left:{x0}px;top:{t0+52}px;width:{w2}px;height:38px;background:{OT}"></div>')
        p.append(f'<p class="cv" style="left:{x0+max(w2, sc)+16}px;top:{t0+52}px;width:140px">{mult(d["m_o"])}</p>')
    return f'<div class="host" style="width:960px;height:{pitch*len(items)}px">' + "\n".join(p) + "</div>"
hostD = drv_host(own[:2]) + drv_host(own[2:])
slides["drivers"] = page("drivers", "THE TEAMS", "On team channels, driver names pull",
    f'<div class="cols" style="gap:64px"><div class="col" style="width:640px;flex-shrink:0">{mega(mult(own_all["m_w"]), "team videos naming one of their own drivers", "mg2")}<div style="height:24px"></div>{mega(mult(own_ctr["m_w"]), "the same drivers named on other channels", "mg2", True)}</div>'
    f'<div style="width:960px;flex-shrink:0">{legend([(RW, "Names a team driver"), (OT, "Other videos")])}{hostD}</div></div>',
    "Source: workbook, Analysis I. Relative views: 1.00× is a normal video for the channel.", dark=True)

# ---------------------------------------------------------------- 9 reach (light)
names = {"Formula 1": "Formula 1", "Mercedes-AMG F1": "Mercedes", "Red Bull Racing": "Red Bull", "F1 News TacticalRab": "TacticalRab"}
indep = {"Formula Addict", "F1 News TacticalRab", "Cameron Cc", "Tommo", "Kym Illman", "Peter Windsor"}
chs = sorted([c for c in reach if reach[c]["n"] >= 30], key=lambda c: -reach[c]["ratio"])
def rfmt(x): return pct(x, 1 if x < 0.1 else 0)
sc, x0, pitch = 1000, 320, 48
def reach_host(sub, off):
    p = [f'<div class="vref" style="left:{x0+sc}px;top:0px;height:{pitch*len(sub)}px"></div>']
    for i, c in enumerate(sub):
        r_ = reach[c]["ratio"]; w = max(4, round(r_ * sc)); col = RW if c in indep else PR
        p.append(f'<p class="cn" style="left:0px;top:{i*pitch+6}px;width:300px">{names.get(c, c)}</p>')
        p.append(f'<div class="bar" style="left:{x0}px;top:{i*pitch+8}px;width:{w}px;height:30px;background:{col}"></div>')
        p.append(f'<p class="cv" style="left:{x0+w+14}px;top:{i*pitch+6}px;width:110px">{rfmt(r_)}</p>')
    if off == 1:
        p.append(f'<p class="sub" style="left:{x0+sc+14}px;top:{pitch*2+6}px;width:300px">views = subscribers</p>')
    return f'<div class="host" style="width:1664px;height:{pitch*len(sub)}px">' + "\n".join(p) + "</div>"
hostR = reach_host(chs[:6], 0) + reach_host(chs[6:], 1)
slides["reach"] = page("reach", "REACH", "Four independents get far more views per subscriber",
    legend([(RW, "Independents"), (PR, "Sport’s channel and teams")], '<p class="th1" style="margin-left:auto">Median views per video, as a share of subscribers</p>') + hostR,
    "Source: workbook, Analysis H. Three channels with under 30 videos are left out; see the report.")

order_ = ["cover", "statement", "rhythm", "highlights", "postrace", "independents", "teams", "drivers", "reach"]
for id_ in order_:
    open(f"{ART}/slides/{id_}.html", "w", encoding="utf-8").write(slides[id_])
open(f"{ART}/deck.css", "w", encoding="utf-8").write(open(os.path.join(HERE, "deck.css"), encoding="utf-8").read())
deck = {"v": 5, "createdOnFiles": {"v": 1, "at": "2026-11-16T13:10:00Z"}, "lists": "css",
        "format": {"v": 1, "shared": ["project/deck.css"]},
        "title": "F1 on YouTube: teams, the sport’s channel and independent creators",
        "cover": "cover", "order": order_,
        "sections": {"s1": {"description": "The question and the short version", "start": "cover"},
                     "s2": {"description": "What the numbers show", "start": "rhythm"}},
        "faces": {"archivo": {"family": "Archivo", "href": "https://fonts.googleapis.com/css2?family=Archivo:wght@400..800&display=swap"},
                  "inter": {"family": "Inter", "href": "https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap"}},
        "designSystems": []}
json.dump(deck, open(f"{ART}/deck.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("built", order_)
