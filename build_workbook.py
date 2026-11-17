#!/usr/bin/env python3
"""
build_workbook.py - build the Excel workbook from the cleaned data.

Reads   data/clean_videos.csv, data/raw_videos*.csv, data/f1_calendar.csv, data/cleaning_log.csv,
        data/channels_resolved.csv
Writes  f1_youtube_study.xlsx

Everything on the Summary sheet is an Excel formula over the Clean sheet, so the numbers can be
traced cell by cell. Run LibreOffice recalculation afterwards so cached values are stored:
    python <skill>/scripts/recalc.py f1_youtube_study.xlsx
"""
import argparse
import datetime as dt
import glob

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_SUB = Font(name=FONT, size=11, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", start_color="1F3864")
FILL_SOFT = PatternFill("solid", start_color="F2F2F2")

GROUPS = [("f1_official", "F1 official"), ("team", "Teams"), ("independent", "Independents")]
WINDOWS = ["Race weekend", "Post-race", "Other"]
FORMATS = ["Standard", "Short"]


def style_range(ws, cells, font=None, fill=None, align=None, fmt=None):
    for row in ws[cells] if isinstance(cells, str) else cells:
        for c in row if isinstance(row, tuple) else (row,):
            if font:
                c.font = font
            if fill:
                c.fill = fill
            if align:
                c.alignment = align
            if fmt:
                c.number_format = fmt


def header(ws, row, labels, col=1):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=text)
        c.font, c.fill = F_HEAD, FILL_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_df(ws, df, number_formats=None):
    header(ws, 1, list(df.columns))
    number_formats = number_formats or {}
    for r, row in enumerate(df.itertuples(index=False), start=2):
        for c, v in enumerate(row, start=1):
            if isinstance(v, float) and pd.isna(v):
                v = None
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = F_BASE
            fmt = number_formats.get(df.columns[c - 1])
            if fmt:
                cell.number_format = fmt


def add_analysis(a, clean, cal, last, cal_last, t1_first, t1_last, t2_first, channels, C):
    """Analysis sheet: deeper cuts behind the findings. Every number is a formula over the Clean sheet."""
    CAL_RACE = f"Calendar!$F$2:$F${cal_last}"
    A_, B_, E_, H_, L_, O_, R_ = (C("A"), C("B"), C("E"), C("H"), C("L"), C("O"), C("R"))
    U_, V_, W_ = C("U"), C("V"), C("W")

    def put(r, c, v, fmt=None, font=F_BASE):
        cell = a.cell(row=r, column=c, value=v)
        cell.font = font
        if fmt:
            cell.number_format = fmt
        return cell

    def arr(r, c, f, fmt=None):
        ref = f"{get_column_letter(c)}{r}"
        a[ref] = ArrayFormula(ref, f)
        a[ref].font = F_BASE
        if fmt:
            a[ref].number_format = fmt

    def title(r, text):
        put(r, 1, text, font=F_SUB)

    def note(r, text):
        put(r, 1, text, font=F_NOTE)

    put(1, 1, "Analysis", font=F_TITLE)
    note(2, "Deeper cuts behind the findings. Every number is a formula over the Clean sheet (helper columns S-W). "
            "Relative views and windows are defined as on the Summary sheet.")

    # ---------------------------------------------------------------- A. highlight packages
    r = 4
    title(r, "A. The sport's channel: highlight packages (practice, qualifying, sprint and race highlights of the 2025 and 2026 seasons)")
    header(a, r + 1, ["Highlight kind", "Videos", "Median views"])
    kinds = ["FP1", "FP2", "FP3", "Qualifying", "Sprint Qualifying", "Sprint", "Race"]
    kind_first = r + 2
    for i, k in enumerate(kinds):
        rr = kind_first + i
        put(rr, 1, k)
        put(rr, 2, f'=COUNTIFS({W_},1,{E_},A{rr}&" Highlights*")', "#,##0")
        arr(rr, 3, f'=MEDIAN(IF(({W_}=1)*(LEFT({E_},LEN(A{rr}&" Highlights"))=A{rr}&" Highlights"),{H_}))', "#,##0")
    r = kind_first + len(kinds) + 1
    header(a, r, ["How much of the channel they carry", "Value"])
    wt = r + 1
    rows = [
        ("Highlight packages (videos)", f"=SUM({W_})", "#,##0"),
        ("Share of the channel's videos", f'=SUM({W_})/COUNTIFS({B_},"f1_official")', "0.0%"),
        ("Share of the channel's views (Shorts included)", f'=SUMIFS({H_},{W_},1)/SUMIFS({H_},{B_},"f1_official")', "0.0%"),
        ("Share of the channel's Standard-video views", f'=SUMIFS({H_},{W_},1)/SUMIFS({H_},{B_},"f1_official",{L_},"Standard")', "0.0%"),
    ]
    for i, (lab, f, fmt) in enumerate(rows):
        put(wt + i, 1, lab)
        put(wt + i, 2, f, fmt)
    k = wt + len(rows)
    put(k, 1, "Median views: highlight packages")
    arr(k, 2, f"=MEDIAN(IF({W_}=1,{H_}))", "#,##0")
    put(k + 1, 1, "Median views: every other Standard video on the channel")
    arr(k + 1, 2, f'=MEDIAN(IF(({B_}="f1_official")*({L_}="Standard")*({W_}=0),{H_}))', "#,##0")
    put(k + 2, 1, "Median views: all Standard videos on the channel")
    arr(k + 2, 2, f'=MEDIAN(IF(({B_}="f1_official")*({L_}="Standard"),{H_}))', "#,##0")
    r = k + 4
    header(a, r, ["Sport's channel, Standard videos: median relative views", "Race weekend", "Post-race", "Other",
                  "n: Race weekend", "n: Post-race", "n: Other"])
    a.row_dimensions[r].height = 30
    wins = ["Race weekend", "Post-race", "Other"]
    put(r + 1, 1, "All Standard videos")
    put(r + 2, 1, "Without the highlight packages")
    for j, w in enumerate(wins):
        col = 2 + j
        hdr = f"{get_column_letter(col)}${r}"
        arr(r + 1, col, f'=MEDIAN(IF(({B_}="f1_official")*({L_}="Standard")*({O_}={hdr}),{R_}))', "0.00")
        arr(r + 2, col, f'=MEDIAN(IF(({B_}="f1_official")*({L_}="Standard")*({W_}=0)*({O_}={hdr}),{R_}))', "0.00")
        put(r + 1, 5 + j, f'=COUNTIFS({B_},"f1_official",{L_},"Standard",{O_},{hdr})', "#,##0")
        put(r + 2, 5 + j, f'=COUNTIFS({B_},"f1_official",{L_},"Standard",{W_},0,{O_},{hdr})', "#,##0")
    note(r + 3, "Without the packages, race-weekend uploads are about as popular as the channel's other videos: the lift comes from the packages.")
    wt_rows = dict(kind_first=kind_first, weight=wt, med=k, win=r)

    # race highlights per race
    r = r + 5
    title(r, "Race Highlights views, race by race (views at the snapshot date)")
    header(a, r + 1, ["Race", "Race day", "Race Highlights views"])
    lo, hi = clean["pub_date"].min(), clean["pub_date"].max()
    races = cal[(cal["race_date"] >= lo) & (cal["race_date"] <= hi)].reset_index()
    rh_first = r + 2
    for i, rc in enumerate(races.itertuples()):
        rr = rh_first + i
        put(rr, 1, f"=Calendar!C{rc.index + 2}")
        put(rr, 2, f"=Calendar!F{rc.index + 2}", "yyyy-mm-dd")
        put(rr, 3, f'=SUMIFS({H_},{W_},1,{E_},"Race Highlights*",{V_},A{rr})', "#,##0")
    rh_last = rh_first + len(races) - 1
    put(rh_last + 1, 1, "Smallest / median / largest", font=F_BOLD)
    put(rh_last + 1, 2, f"=MIN(C{rh_first}:C{rh_last})", "#,##0")
    put(rh_last + 1, 3, f"=MEDIAN(C{rh_first}:C{rh_last})", "#,##0")
    put(rh_last + 1, 4, f"=MAX(C{rh_first}:C{rh_last})", "#,##0")
    note(rh_last + 2, "The latest races have had the least time to collect views.")

    # ---------------------------------------------------------------- B. day by day
    r = rh_last + 5
    title(r, "B. Day by day around race day (0 = race day, -2 = Friday of a normal weekend, +1 = the Monday after)")
    header(a, r + 1, ["Days from race day", "Uploads/channel/day: Sport's channel", "Uploads/channel/day: Teams",
                      "Uploads/channel/day: Independents", "Median relative views (Standard): Sport's channel",
                      "Median relative views (Standard): Teams", "Median relative views (Standard): Independents",
                      "n: Sport's channel", "n: Teams", "n: Independents"])
    a.row_dimensions[r + 1].height = 58
    day_first = r + 2
    gkeys = ["f1_official", "team", "independent"]
    for i, kk in enumerate(range(-3, 4)):
        rr = day_first + i
        put(rr, 1, kk, "+0;-0;0", F_INPUT)
        for g, gk in enumerate(gkeys):
            n_races = f'COUNTIFS({CAL_RACE},">="&(Notes!$B$5-$A{rr}),{CAL_RACE},"<="&(Notes!$B$7-$A{rr}))'
            put(rr, 2 + g, f'=COUNTIFS({B_},"{gk}",{U_},$A{rr})/(Summary!$F${t2_first + g}*{n_races})', "0.00")
            arr(rr, 5 + g, f'=MEDIAN(IF(({B_}="{gk}")*({L_}="Standard")*({U_}=$A{rr}),{R_}))', "0.00")
            put(rr, 8 + g, f'=COUNTIFS({B_},"{gk}",{L_},"Standard",{U_},$A{rr})', "#,##0")
    day_last = day_first + 6
    put(day_last + 1, 1, "Ordinary days (Summary table 2)", font=F_BOLD)
    for g in range(3):
        put(day_last + 1, 2 + g, f"=Summary!$D${t2_first + g}", "0.00")
    note(day_last + 2, "Uploads are divided by the channels in the group and by the number of races whose day falls inside the data window.")

    # ---------------------------------------------------------------- C. race by race
    r = day_last + 5
    title(r, "C. Race by race: median relative views of race-weekend Standard videos (blank where a race has fewer than 3 videos)")
    header(a, r + 1, ["Race", "Race day", "Sport's channel", "Teams", "Independents", "Independents without Peter Windsor",
                      "n: Sport's channel", "n: Teams", "n: Independents", "n: Independents without Windsor"])
    a.row_dimensions[r + 1].height = 44
    rc_first = r + 2
    for i, rc in enumerate(races.itertuples()):
        rr = rc_first + i
        put(rr, 1, f"=Calendar!C{rc.index + 2}")
        put(rr, 2, f"=Calendar!F{rc.index + 2}", "yyyy-mm-dd")
        specs = [("f1_official", False), ("team", False), ("independent", False), ("independent", True)]
        for g, (gk, exw) in enumerate(specs):
            extra_cnt = f',{A_},"<>Peter Windsor"' if exw else ""
            extra_arr = f'*({A_}<>"Peter Windsor")' if exw else ""
            put(rr, 7 + g, f'=COUNTIFS({V_},$A{rr},{B_},"{gk}",{L_},"Standard",{O_},"Race weekend"{extra_cnt})', "#,##0")
            cnt = f"{get_column_letter(7 + g)}{rr}"
            arr(rr, 3 + g, (f'=IF({cnt}<3,"",MEDIAN(IF(({V_}=$A{rr})*({B_}="{gk}")*({L_}="Standard")*'
                            f'({O_}="Race weekend"){extra_arr},{R_})))'), "0.00")
    rc_last = rc_first + len(races) - 1
    put(rc_last + 1, 1, "Races with a median", font=F_BOLD)
    put(rc_last + 2, 1, "Races where the median beat 1.00 (a normal video)", font=F_BOLD)
    for g in range(4):
        col = get_column_letter(3 + g)
        put(rc_last + 1, 3 + g, f"=COUNT({col}{rc_first}:{col}{rc_last})", "0")
        put(rc_last + 2, 3 + g, f'=COUNTIF({col}{rc_first}:{col}{rc_last},">1")', "0")

    # ---------------------------------------------------------------- D. independents without Windsor
    r = rc_last + 5
    title(r, "D. Independents, Standard videos: with and without Peter Windsor (median relative views)")
    header(a, r + 1, ["", "Race weekend", "Post-race", "Other", "n: Race weekend", "n: Post-race", "n: Other"])
    for i, (lab, exw) in enumerate([("All independents", False), ("Without Peter Windsor", True)]):
        rr = r + 2 + i
        put(rr, 1, lab)
        extra_cnt = f',{A_},"<>Peter Windsor"' if exw else ""
        extra_arr = f'*({A_}<>"Peter Windsor")' if exw else ""
        for j, w in enumerate(wins):
            hdr = f"{get_column_letter(2 + j)}${r + 1}"
            arr(rr, 2 + j, f'=MEDIAN(IF(({B_}="independent")*({L_}="Standard")*({O_}={hdr}){extra_arr},{R_}))', "0.00")
            put(rr, 5 + j, f'=COUNTIFS({B_},"independent",{L_},"Standard",{O_},{hdr}{extra_cnt})', "#,##0")
    dwin = r + 1

    # ---------------------------------------------------------------- E. Windsor's race analyses
    r = r + 6
    title(r, "E. Peter Windsor: videos with 'analysis' in the title vs his other Standard videos (median views)")
    header(a, r + 1, ["Window", "'Analysis' videos (n)", "'Analysis' median views", "Other Standard videos (n)",
                      "Other Standard median views"])
    a.row_dimensions[r + 1].height = 30
    for j, w in enumerate(wins):
        rr = r + 2 + j
        put(rr, 1, w)
        flag = f'ISNUMBER(SEARCH("analysis",{E_}))'
        base = f'({A_}="Peter Windsor")*({L_}="Standard")*({O_}=$A{rr})'
        put(rr, 2, f'=SUMPRODUCT({base}*{flag})', "#,##0")
        arr(rr, 3, f'=IF(B{rr}=0,"n/a",MEDIAN(IF({base}*{flag},{H_})))', "#,##0")
        put(rr, 4, f'=SUMPRODUCT({base}*(1-{flag}))', "#,##0")
        arr(rr, 5, f'=IF(D{rr}=0,"n/a",MEDIAN(IF({base}*(1-{flag}),{H_})))', "#,##0")
    put(r + 5, 1, "His typical Standard video (all windows)")
    arr(r + 5, 3, f'=MEDIAN(IF(({A_}="Peter Windsor")*({L_}="Standard"),{H_}))', "#,##0")

    # ---------------------------------------------------------------- F. concentration and Shorts
    r = r + 8
    title(r, "F. How concentrated are views? Top 10% of each channel's videos, Shorts share of views, and the biggest hit")
    header(a, r + 1, ["Channel", "Videos", "Top 10% (videos)", "Views of the 10th-percentile video",
                      "Share of views from the top 10%", "Share of views from Shorts", "Biggest hit: relative views",
                      "Biggest hit: title"])
    a.row_dimensions[r + 1].height = 44
    cf = r + 2
    for i, ch in enumerate(channels):
        rr = cf + i
        put(rr, 1, ch)
        put(rr, 2, f"=COUNTIFS({A_},A{rr})", "#,##0")
        put(rr, 3, f"=MAX(1,ROUND(B{rr}*0.1,0))", "0")
        arr(rr, 4, f"=LARGE(IF({A_}=A{rr},{H_}),C{rr})", "#,##0")
        put(rr, 5, f'=SUMIFS({H_},{A_},A{rr},{H_},">="&D{rr})/SUMIFS({H_},{A_},A{rr})', "0%")
        put(rr, 6, f'=SUMIFS({H_},{A_},A{rr},{L_},"Short")/SUMIFS({H_},{A_},A{rr})', "0%")
        arr(rr, 7, f"=MAX(IF({A_}=A{rr},{R_}))", "0.0")
        arr(rr, 8, f"=INDEX({E_},MATCH(1,({A_}=A{rr})*({R_}=G{rr}),0))")
    cl = cf + len(channels) - 1
    note(cl + 1, "Relative views of 83 = 83 times that channel's normal video in the same format and month. Biggest hits count Shorts and Standard videos.")

    # ---------------------------------------------------------------- G. post-race fade
    r = cl + 4
    title(r, "G. Sport's channel, Standard videos in the 3 days after a race: what is posted and how it does")
    note(r + 7, "The first three kinds score well below 1.00 in every window (right-hand columns), not just after races, so the post-race dip is about what gets posted then.")
    header(a, r + 1, ["Kind of video (by title)", "Videos (n)", "Median relative views",
                      "Same kind, all windows: videos (n)", "Same kind, all windows: median relative views"])
    a.row_dimensions[r + 1].height = 44
    chasing = f'ISNUMBER(SEARCH("Chasing The Dream",{E_}))'
    podcast = f'((ISNUMBER(SEARCH("F1 Nation",{E_}))+ISNUMBER(SEARCH("Podcast",{E_})))>0)'
    palmer = f'ISNUMBER(SEARCH("Palmer",{E_}))'
    base = f'({B_}="f1_official")*({L_}="Standard")*({O_}="Post-race")'
    kinds2 = [
        ("'Chasing The Dream' (F2 behind the scenes)", f"{chasing}"),
        ("F1 Nation and other podcasts", f"(1-{chasing})*{podcast}"),
        ("Jolyon Palmer analysis", f"(1-{chasing})*(1-{podcast})*{palmer}"),
        ("Everything else", f"(1-{chasing})*(1-{podcast})*(1-{palmer})"),
    ]
    for i, (lab, flag) in enumerate(kinds2):
        rr = r + 2 + i
        put(rr, 1, lab)
        put(rr, 2, f"=SUMPRODUCT({base}*{flag})", "#,##0")
        arr(rr, 3, f'=IF(B{rr}=0,"n/a",MEDIAN(IF({base}*{flag},{R_})))', "0.00")
        base_all = f'({B_}="f1_official")*({L_}="Standard")'
        put(rr, 4, f"=SUMPRODUCT({base_all}*{flag})", "#,##0")
        arr(rr, 5, f'=IF(D{rr}=0,"n/a",MEDIAN(IF({base_all}*{flag},{R_})))', "0.00")
    put(r + 6, 1, "All post-race Standard videos", font=F_BOLD)
    put(r + 6, 2, f"=SUM(B{r + 2}:B{r + 5})", "#,##0")

    # ---------------------------------------------------------------- H. reach per subscriber
    r = r + 10
    title(r, "H. Reach: what a typical video gets, compared with the channel's subscriber count")
    header(a, r + 1, ["Channel", "Group", "Subscribers", "Videos", "Median views (all videos)",
                      "Median views / subscribers", "Note"])
    a.row_dimensions[r + 1].height = 44
    hf = r + 2
    for i, ch in enumerate(channels):
        rr = hf + i
        put(rr, 1, ch)
        put(rr, 2, f"=INDEX({B_},MATCH(A{rr},{A_},0))")
        put(rr, 3, f"=INDEX({C("C")},MATCH(A{rr},{A_},0))", "#,##0")
        put(rr, 4, f"=COUNTIFS({A_},A{rr})", "#,##0")
        arr(rr, 5, f"=MEDIAN(IF({A_}=A{rr},{H_}))", "#,##0")
        put(rr, 6, f"=E{rr}/C{rr}", "0.0%")
        put(rr, 7, f'=IF(D{rr}<30,"fewer than 30 videos: indicative only","")', font=F_NOTE)
    hl_ = hf + len(channels) - 1
    note(hl_ + 1, "Subscribers are as of the snapshot date. Above 100% is possible: a video can reach people who do not subscribe. "
                  "Both formats are pooled, so channels that post mostly Shorts are compared on their Shorts.")

    # ---------------------------------------------------------------- I. team channels and their own drivers
    r = hl_ + 4
    title(r, "I. Team channels: videos that name one of the team's own drivers, against the channel's other videos")
    header(a, r + 1, ["Team channel", "Driver 1 (typed in)", "Driver 2 (typed in)", "Videos naming a driver (n)",
                      "Median relative views: naming a driver", "Other videos (n)", "Median relative views: other videos"])
    a.row_dimensions[r + 1].height = 44
    teams_ = [("Red Bull Racing", "Verstappen", "Hadjar"), ("Mercedes-AMG F1", "Russell", "Antonelli"),
              ("McLaren", "Norris", "Piastri"), ("Williams", "Albon", "Sainz")]
    t_first = r + 2

    def named(rr):      # 1 when the title names either of the row's two drivers
        return (f"((ISNUMBER(SEARCH($B{rr},{E_}))+ISNUMBER(SEARCH($C{rr},{E_})))>0)")

    for i, (ch, d1, d2) in enumerate(teams_):
        rr = t_first + i
        put(rr, 1, ch)
        put(rr, 2, d1, font=F_INPUT)
        put(rr, 3, d2, font=F_INPUT)
        w_ = f"({A_}=$A{rr})*{named(rr)}"
        o_ = f"({A_}=$A{rr})*(1-{named(rr)})"
        put(rr, 4, f"=SUMPRODUCT({w_})", "#,##0")
        arr(rr, 5, f'=IF(D{rr}=0,"n/a",MEDIAN(IF({w_},{R_})))', "0.00")
        put(rr, 6, f"=SUMPRODUCT({o_})", "#,##0")
        arr(rr, 7, f'=IF(F{rr}=0,"n/a",MEDIAN(IF({o_},{R_})))', "0.00")
    t_last = t_first + len(teams_) - 1
    rp = t_last + 1
    put(rp, 1, "All four team channels", font=F_BOLD)
    w_all = "(" + "+".join(f"({A_}=$A{t_first + i})*{named(t_first + i)}" for i in range(4)) + ")"
    t_any = "(" + "+".join(f"({A_}=$A{t_first + i})" for i in range(4)) + ")"
    o_all = f"({t_any}*(1-{w_all}))"
    put(rp, 4, f"=SUMPRODUCT({w_all})", "#,##0")
    arr(rp, 5, f"=MEDIAN(IF({w_all},{R_}))", "0.00")
    put(rp, 6, f"=SUMPRODUCT({o_all})", "#,##0")
    arr(rp, 7, f"=MEDIAN(IF({o_all},{R_}))", "0.00")
    rc2 = rp + 1
    put(rc2, 1, "For contrast: the sport's channel and independents, naming any of those eight drivers", font=F_BOLD)
    any8 = "((" + "+".join(f"ISNUMBER(SEARCH($B${t_first + i},{E_}))+ISNUMBER(SEARCH($C${t_first + i},{E_}))" for i in range(4)) + ")>0)"
    nt_ = f"({B_}<>\"team\")"
    put(rc2, 4, f"=SUMPRODUCT({nt_}*{any8})", "#,##0")
    arr(rc2, 5, f"=MEDIAN(IF({nt_}*{any8},{R_}))", "0.00")
    put(rc2, 6, f"=SUMPRODUCT({nt_}*(1-{any8}))", "#,##0")
    arr(rc2, 7, f"=MEDIAN(IF({nt_}*(1-{any8}),{R_}))", "0.00")
    note(rc2 + 1, "Both formats pooled; relative views already adjust for each channel's normal video in the same format and month. "
                  "Drivers are the 2026 line-ups, typed in. A name is matched anywhere in the title, whatever the capitalisation.")

    for col, w in zip("ABCDEFGHIJ", [58, 18, 20, 20, 22, 22, 22, 22, 16, 16]):
        a.column_dimensions[col].width = w
    a.column_dimensions["H"].width = 60
    a.freeze_panes = "A3"
    return dict(kind_first=kind_first, weight=wt, med=k, win=wt_rows["win"], rh_first=rh_first, rh_last=rh_last,
                day_first=day_first, rc_first=rc_first, rc_last=rc_last, dwin=dwin, concentration_first=cf,
                concentration_last=cl)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="f1_youtube_study.xlsx")
    args = ap.parse_args()

    clean = pd.read_csv("data/clean_videos.csv", parse_dates=["pub_date"])
    raw = pd.concat([pd.read_csv(f) for f in sorted(glob.glob("data/raw_videos*.csv"))], ignore_index=True)
    cal = pd.read_csv("data/f1_calendar.csv", parse_dates=["weekend_start", "race_date"])
    log = pd.read_csv("data/cleaning_log.csv")
    resolved = pd.concat([pd.read_csv(f) for f in sorted(glob.glob("data/channels_resolved*.csv"))], ignore_index=True)

    resolved = resolved[resolved["label"].isin(clean["channel_label"].unique())].reset_index(drop=True)
    import os
    excluded = pd.read_csv("data/excluded_channels.csv") if os.path.exists("data/excluded_channels.csv") else None
    pull_date = pd.to_datetime(raw["pulled_at"]).max().tz_localize(None).normalize().date()
    window_start = dt.date(2025, 11, 17)
    min_age = 7
    n = len(clean)
    last = n + 1                       # last data row on Clean
    cal_last = len(cal) + 1            # last data row on Calendar

    wb = Workbook()
    ws_notes = wb.active
    ws_notes.title = "Notes"
    ws_sum = wb.create_sheet("Summary")
    ws_an = wb.create_sheet("Analysis")
    ws_clean = wb.create_sheet("Clean")
    ws_base = wb.create_sheet("Baselines")
    ws_days = wb.create_sheet("Days")
    ws_cal = wb.create_sheet("Calendar")
    ws_raw = wb.create_sheet("Raw")

    # ---------------------------------------------------------------- Calendar
    cal_out = cal.rename(columns={"weekend_start": "weekend_start (first practice day)", "race_date": "race_date"})
    write_df(ws_cal, cal_out, {"weekend_start (first practice day)": "yyyy-mm-dd", "race_date": "yyyy-mm-dd"})
    for r in range(2, cal_last + 1):          # hard-coded inputs -> blue
        for c in range(1, 8):
            ws_cal.cell(row=r, column=c).font = F_INPUT
    for col, w in zip("ABCDEFG", [8, 8, 32, 14, 30, 14, 16]):
        ws_cal.column_dimensions[col].width = w
    ws_cal.freeze_panes = "A2"
    ws_cal["I1"] = "Source: Jolpica F1 API (Ergast successor), api.jolpi.ca/ergast/f1/{2025,2026}.json, fetched 2026-11-16."
    ws_cal["I2"] = "weekend_start = date of first practice session. Race 16 of 2026 is named 'Bahrain Grand Prix in Malaysia' by the API; kept as published."
    ws_cal["I1"].font = ws_cal["I2"].font = F_NOTE

    # ------------------------------------------------------------------- Notes inputs
    ws_notes["A1"] = "F1 on YouTube: team channels, the sport's channel, and independent creators"
    ws_notes["A1"].font = F_TITLE
    ws_notes["A3"] = "Inputs (blue = typed in; change these and the workbook recalculates)"
    ws_notes["A3"].font = F_SUB
    inputs = [
        (4, "Stats snapshot (pull) date", pull_date, "yyyy-mm-dd", "From the pull timestamp in the raw data."),
        (5, "Window start (earliest upload kept)", window_start, "yyyy-mm-dd", "The --since date passed to yt_collect.py."),
        (6, "Minimum video age (days)", min_age, "0", "Younger videos are dropped: their view counts are still climbing."),
        (7, "Window end (latest upload kept)", "=B4-B6", "yyyy-mm-dd", "Pull date minus minimum age."),
        (8, "Short = duration up to (seconds)", 60, "0", "Proxy for YouTube Shorts; YouTube does not expose a Shorts flag in the API."),
        (9, "Post-race window (days after race day)", 3, "0", "Uploads Mon-Wed after a race are tagged 'Post-race'."),
        (10, "Weeks in window", "=(B7-B5+1)/7", "0.0", "Used for videos-per-week."),
    ]
    for r, label, val, fmt, note in inputs:
        ws_notes.cell(row=r, column=1, value=label).font = F_BASE
        c = ws_notes.cell(row=r, column=2, value=val)
        c.number_format = fmt
        c.font = F_BASE if isinstance(val, str) and val.startswith("=") else F_INPUT
        ws_notes.cell(row=r, column=3, value=note).font = F_NOTE

    # ------------------------------------------------------------------- Clean
    clean_cols = ["channel", "group", "subscribers", "video_id", "title", "published (UTC date)",
                  "duration_sec", "views", "likes", "comments", "age_days", "format",
                  "engagement_rate", "views_per_sub", "race_window", "month",
                  "same_month_median_views", "relative_views",
                  "race_idx", "nearest_race_idx", "days_from_race_day", "nearest_race", "highlight_package"]
    header(ws_clean, 1, clean_cols)
    cal_rng_start = f"Calendar!$E$2:$E${cal_last}"
    cal_rng_race = f"Calendar!$F$2:$F${cal_last}"
    clean["month"] = clean["pub_date"].dt.year * 100 + clean["pub_date"].dt.month
    base_keys = (clean[["channel_label", "format", "month"]].drop_duplicates()
                 .sort_values(["channel_label", "format", "month"]).values.tolist())
    n_base = len(base_keys)
    base_last = n_base + 1

    def window_formula(ref):
        return (f'=IF(COUNTIFS({cal_rng_start},"<="&{ref},{cal_rng_race},">="&{ref})>0,"Race weekend",'
                f'IF(COUNTIFS({cal_rng_race},"<"&{ref},{cal_rng_race},">="&{ref}-Notes!$B$9)>0,"Post-race","Other"))')

    for i, row in enumerate(clean.itertuples(index=False), start=2):
        vals = [row.channel_label, row.group, int(row.channel_subs), row.video_id, row.title,
                row.pub_date.date(), int(row.duration_sec), int(row.views), int(row.likes), int(row.comments)]
        for c, v in enumerate(vals, start=1):
            cell = ws_clean.cell(row=i, column=c, value=v)
            cell.font = F_BASE
        ws_clean.cell(row=i, column=6).number_format = "yyyy-mm-dd"
        ws_clean.cell(row=i, column=11, value=f"=Notes!$B$4-F{i}")
        ws_clean.cell(row=i, column=12, value=f'=IF(G{i}<=Notes!$B$8,"Short","Standard")')
        ws_clean.cell(row=i, column=13, value=f"=(I{i}+J{i})/H{i}")
        ws_clean.cell(row=i, column=14, value=f"=H{i}/C{i}")
        ws_clean.cell(row=i, column=15, value=window_formula(f"F{i}"))
        ws_clean.cell(row=i, column=16, value=f"=YEAR(F{i})*100+MONTH(F{i})")
        ws_clean.cell(row=i, column=17,
                      value=f'=INDEX(Baselines!$F$2:$F${base_last},MATCH(A{i}&"|"&L{i}&"|"&P{i},Baselines!$A$2:$A${base_last},0))')
        ws_clean.cell(row=i, column=18, value=f"=H{i}/Q{i}")
        # helper columns for the Analysis sheet
        ws_clean.cell(row=i, column=19, value=f"=MATCH(F{i},{cal_rng_race},1)")
        ws_clean.cell(row=i, column=20,
                      value=(f"=IF(ABS(F{i}-INDEX({cal_rng_race},S{i}))<=ABS(F{i}-INDEX({cal_rng_race},S{i}+1)),"
                             f"S{i},S{i}+1)"))
        ws_clean.cell(row=i, column=21, value=f"=F{i}-INDEX({cal_rng_race},T{i})")
        ws_clean.cell(row=i, column=22, value=f"=INDEX(Calendar!$C$2:$C${cal_last},T{i})")
        ws_clean.cell(row=i, column=23,
                      value=(f'=IF(AND(B{i}="f1_official",L{i}="Standard",ISNUMBER(SEARCH(" Highlights | 202",E{i})),'
                             f'OR(LEFT(E{i},4)="Race",LEFT(E{i},6)="Sprint",LEFT(E{i},10)="Qualifying",'
                             f'LEFT(E{i},2)="FP",LEFT(E{i},8)="Practice"),'
                             f'ISERROR(SEARCH("F2",E{i})),ISERROR(SEARCH("F3",E{i})),ISERROR(SEARCH("Academy",E{i}))),1,0)'))
        for c in range(11, 24):
            ws_clean.cell(row=i, column=c).font = F_BASE
        ws_clean.cell(row=i, column=13).number_format = "0.00%"
        ws_clean.cell(row=i, column=14).number_format = "0.0000"
        ws_clean.cell(row=i, column=17).number_format = "#,##0"
        ws_clean.cell(row=i, column=18).number_format = "0.00"
        for c in (3, 8, 9, 10):
            ws_clean.cell(row=i, column=c).number_format = "#,##0"
    for col, w in zip("ABCDEFGHIJKLMNOPQRSTUVW", [18, 12, 12, 13, 50, 14, 11, 12, 10, 10, 9, 10, 12, 12, 14, 9, 16, 12,
                                                  9, 12, 12, 30, 12]):
        ws_clean.column_dimensions[col].width = w
    ws_clean.freeze_panes = "B2"
    ws_clean.auto_filter.ref = f"A1:W{last}"

    def C(col):
        return f"Clean!${col}$2:${col}${last}"

    # ------------------------------------------------------------------- Baselines
    header(ws_base, 1, ["key", "channel", "format", "month", "videos", "median views"])
    for r, (ch, fmt, mo) in enumerate(base_keys, start=2):
        ws_base.cell(row=r, column=1, value=f'=B{r}&"|"&C{r}&"|"&D{r}')
        ws_base.cell(row=r, column=2, value=ch)
        ws_base.cell(row=r, column=3, value=fmt)
        ws_base.cell(row=r, column=4, value=int(mo))
        ws_base.cell(row=r, column=5, value=f"=COUNTIFS({C('A')},B{r},{C('L')},C{r},{C('P')},D{r})")
        ws_base[f"F{r}"] = ArrayFormula(
            f"F{r}", f"=MEDIAN(IF(({C('A')}=B{r})*({C('L')}=C{r})*({C('P')}=D{r}),{C('H')}))")
        ws_base.cell(row=r, column=6).number_format = "#,##0"
        for c in range(1, 7):
            ws_base.cell(row=r, column=c).font = F_BASE
    for col, w in zip("ABCDEF", [34, 18, 12, 10, 10, 16]):
        ws_base.column_dimensions[col].width = w
    ws_base["H1"] = ("Median views for each channel, format and upload month (e.g. 202603 = March 2026). "
                     "Each video's 'relative_views' on the Clean sheet is its views divided by this number, "
                     "so it is compared only with videos of similar age.")
    ws_base["H1"].font = F_NOTE

    # ------------------------------------------------------------------- Days
    header(ws_days, 1, ["date", "race_window"])
    n_days = (pull_date - dt.timedelta(days=min_age) - window_start).days + 1
    for r in range(2, n_days + 2):
        ws_days.cell(row=r, column=1, value="=Notes!$B$5" if r == 2 else f"=A{r-1}+1").number_format = "yyyy-mm-dd"
        ws_days.cell(row=r, column=2, value=window_formula(f"A{r}"))
        ws_days.cell(row=r, column=1).font = ws_days.cell(row=r, column=2).font = F_BASE
    days_last = n_days + 1
    ws_days.column_dimensions["A"].width = 14
    ws_days.column_dimensions["B"].width = 16
    ws_days["D1"] = "One row per calendar day in the window, tagged with the same rule as the Clean sheet. Used for 'share of calendar days'."
    ws_days["D1"].font = F_NOTE

    # ------------------------------------------------------------------- Raw
    raw_out = raw[["channel_label", "group", "channel_subs", "video_id", "title", "published_at",
                   "duration_sec", "views", "likes", "comments", "was_live", "pulled_at"]]
    write_df(ws_raw, raw_out)
    for col, w in zip("ABCDEFGHIJKL", [18, 12, 12, 13, 50, 22, 11, 12, 10, 10, 9, 22]):
        ws_raw.column_dimensions[col].width = w
    ws_raw.freeze_panes = "A2"

    # ------------------------------------------------------------------- Summary
    s = ws_sum
    s["A1"] = "Summary"
    s["A1"].font = F_TITLE
    s["A2"] = ("All figures are formulas over the Clean sheet. Medians, not averages, because views are heavily skewed. "
               "Treat any cell based on fewer than 30 videos as indicative only.")
    s["A2"].font = F_NOTE

    # --- T1 channel scorecard
    s["A3"] = "1. Channel scorecard"
    s["A3"].font = F_SUB
    header(s, 4, ["Channel", "Group", "Subscribers", "Videos kept", "Shorts share", "Videos per week",
                  "Median views (Standard)", "Median views (Short)", "Median views ÷ subscribers (Standard)",
                  "Median engagement rate (Standard)", "Standard videos (n)"])
    s.row_dimensions[4].height = 42
    channels = list(resolved["label"])
    t1_first, t1_last = 5, 5 + len(channels) - 1
    for r, ch in enumerate(channels, start=5):
        s.cell(row=r, column=1, value=ch)
        s.cell(row=r, column=2, value=f"=INDEX({C('B')},MATCH(A{r},{C('A')},0))")
        s.cell(row=r, column=3, value=f"=INDEX({C('C')},MATCH(A{r},{C('A')},0))")
        s.cell(row=r, column=4, value=f"=COUNTIFS({C('A')},A{r})")
        s.cell(row=r, column=5, value=f'=COUNTIFS({C("A")},A{r},{C("L")},"Short")/D{r}')
        s.cell(row=r, column=6, value=f"=D{r}/Notes!$B$10")
        s[f"G{r}"] = ArrayFormula(
            f"G{r}", f'=IF(K{r}=0,"n/a",MEDIAN(IF(({C("A")}=A{r})*({C("L")}="Standard"),{C("H")})))')
        s[f"H{r}"] = ArrayFormula(
            f"H{r}", f'=IF(COUNTIFS({C("A")},A{r},{C("L")},"Short")=0,"n/a",MEDIAN(IF(({C("A")}=A{r})*({C("L")}="Short"),{C("H")})))')
        s.cell(row=r, column=9, value=f'=IFERROR(G{r}/C{r},"n/a")')
        s[f"J{r}"] = ArrayFormula(
            f"J{r}", f'=IF(K{r}=0,"n/a",MEDIAN(IF(({C("A")}=A{r})*({C("L")}="Standard"),{C("M")})))')
        s.cell(row=r, column=11, value=f'=COUNTIFS({C("A")},A{r},{C("L")},"Standard")')
        for c, fmt in [(3, "#,##0"), (4, "#,##0"), (5, "0%"), (6, "0.0"), (7, "#,##0"), (8, "#,##0"),
                       (9, "0.000"), (10, "0.0%"), (11, "#,##0")]:
            s.cell(row=r, column=c).number_format = fmt
        for c in range(1, 12):
            s.cell(row=r, column=c).font = F_BASE
        s.cell(row=r, column=1).font = F_INPUT

    # --- T2 cadence
    t2 = t1_last + 3
    s.cell(row=t2, column=1, value="2. When do they upload? Uploads per channel per day, by race window").font = F_SUB
    header(s, t2 + 1, ["Group"] + WINDOWS + ["Race-weekend lift vs Other", "Channels in group"])
    s.row_dimensions[t2 + 1].height = 30
    t2_first = t2 + 2
    for i, (gkey, glabel) in enumerate(GROUPS):
        r = t2_first + i
        s.cell(row=r, column=1, value=glabel)
        s.cell(row=r, column=6, value=f'=COUNTIFS($B${t1_first}:$B${t1_last},"{gkey}")')
        for j, w in enumerate(WINDOWS):
            col = get_column_letter(2 + j)
            s.cell(row=r, column=2 + j,
                   value=(f'=IF($F{r}=0,"n/a",COUNTIFS({C("B")},"{gkey}",{C("O")},{col}${t2 + 1})'
                          f'/COUNTIF(Days!$B$2:$B${days_last},{col}${t2 + 1})/$F{r})'))
            s.cell(row=r, column=2 + j).number_format = "0.00"
        s.cell(row=r, column=5, value=f'=IFERROR(B{r}/D{r},"n/a")').number_format = '0.0"x"'
        for c in range(1, 7):
            s.cell(row=r, column=c).font = F_BASE
    t2_days = t2_first + len(GROUPS)
    s.cell(row=t2_days, column=1, value="Share of calendar days").font = F_BOLD
    for j, w in enumerate(WINDOWS):
        col = get_column_letter(2 + j)
        c = s.cell(row=t2_days, column=2 + j,
                   value=f"=COUNTIF(Days!$B$2:$B${days_last},{col}${t2 + 1})/COUNTA(Days!$B$2:$B${days_last})")
        c.number_format = "0%"
        c.font = F_BASE
    s.cell(row=t2_days + 1, column=1,
           value="A lift above 1.0x means the group posts more per day on race weekends than on ordinary days.").font = F_NOTE

    # --- T3 relative views by group
    t3 = t2_days + 4
    s.cell(row=t3, column=1,
           value="3. Do race-weekend uploads get more views? Median relative views by group "
                 "(1.00 = that channel's normal video in the same format)").font = F_SUB
    header(s, t3 + 1, ["Group", "Format"] + WINDOWS + [f"n: {w}" for w in WINDOWS])
    s.row_dimensions[t3 + 1].height = 30
    row = t3 + 2
    t3_std_first = row
    for fmt in FORMATS:
        for gkey, glabel in GROUPS:
            s.cell(row=row, column=1, value=glabel)
            s.cell(row=row, column=2, value=fmt)
            for j, w in enumerate(WINDOWS):
                hdr = f"{get_column_letter(3 + j)}${t3 + 1}"
                ncell = f"{get_column_letter(6 + j)}{row}"
                s.cell(row=row, column=6 + j,
                       value=f'=COUNTIFS({C("B")},"{gkey}",{C("L")},$B{row},{C("O")},{hdr})')
                s[f"{get_column_letter(3 + j)}{row}"] = ArrayFormula(
                    f"{get_column_letter(3 + j)}{row}",
                    (f'=IF({ncell}=0,"n/a",MEDIAN(IF(({C("B")}="{gkey}")*({C("L")}=$B{row})*({C("O")}={hdr}),{C("R")})))'))
                s.cell(row=row, column=3 + j).number_format = "0.00"
            for c in range(1, 9):
                s.cell(row=row, column=c).font = F_BASE
            row += 1
    t3_std_last = t3_std_first + len(GROUPS) - 1
    s.cell(row=row, column=1,
           value="Relative views = a video's views / the median views of the same channel's videos in the same format "
                 "uploaded in the same calendar month. Grey = fewer than 10 videos behind the number. "
                 "Groups pool every video, so channels that post a lot count for more; the channel table below shows who drives each group.").font = F_NOTE

    # --- T4 per channel
    t4 = row + 3
    s.cell(row=t4, column=1, value="4. The same, channel by channel").font = F_SUB
    header(s, t4 + 1, ["Channel", "Format"] + WINDOWS + [f"n: {w}" for w in WINDOWS])
    s.row_dimensions[t4 + 1].height = 30
    row = t4 + 2
    for fmt in FORMATS:
        for ch in channels:
            if not ((clean["channel_label"] == ch) & (clean["format"] == fmt)).any():
                continue
            s.cell(row=row, column=1, value=ch)
            s.cell(row=row, column=2, value=fmt)
            for j, w in enumerate(WINDOWS):
                hdr = f"{get_column_letter(3 + j)}${t4 + 1}"
                ncell = f"{get_column_letter(6 + j)}{row}"
                s.cell(row=row, column=6 + j,
                       value=f'=COUNTIFS({C("A")},$A{row},{C("L")},$B{row},{C("O")},{hdr})')
                s[f"{get_column_letter(3 + j)}{row}"] = ArrayFormula(
                    f"{get_column_letter(3 + j)}{row}",
                    (f'=IF({ncell}=0,"n/a",MEDIAN(IF(({C("A")}=$A{row})*({C("L")}=$B{row})*({C("O")}={hdr}),{C("R")})))'))
                s.cell(row=row, column=3 + j).number_format = "0.00"
            for c in range(1, 9):
                s.cell(row=row, column=c).font = F_BASE
            row += 1

    from openpyxl.formatting.rule import FormulaRule
    grey = Font(name=FONT, size=10, color="A6A6A6", italic=True)
    for first, last_row in [(t3 + 2, t3 + 1 + 2 * len(GROUPS)), (t4 + 2, row - 1)]:
        for j in range(3):
            med, ncol = get_column_letter(3 + j), get_column_letter(6 + j)
            s.conditional_formatting.add(
                f"{med}{first}:{med}{last_row}",
                FormulaRule(formula=[f"{ncol}{first}<10"], font=grey))

    # --- T5 age check
    t5 = row + 2
    s.cell(row=t5, column=1, value="5. Why the same-month comparison: median video age in days, by window (Standard videos)").font = F_SUB
    header(s, t5 + 1, ["", ""] + WINDOWS)
    s.cell(row=t5 + 2, column=1, value="All channels").font = F_BASE
    for j, w in enumerate(WINDOWS):
        hdr = f"{get_column_letter(3 + j)}${t5 + 1}"
        cell = f"{get_column_letter(3 + j)}{t5 + 2}"
        s[cell] = ArrayFormula(cell, f'=MEDIAN(IF(({C("L")}="Standard")*({C("O")}={hdr}),{C("K")}))')
        s[cell].number_format = "0"
        s[cell].font = F_BASE
    s.cell(row=t5 + 3, column=1,
           value="These differ because there are no races from December to February, so 'Other' holds more old videos. That is why relative views compare each video only with the same channel's videos from the same month.").font = F_NOTE

    for col, w in zip("ABCDEFGHIJK", [26, 14, 14, 13, 14, 14, 16, 16, 18, 18, 14]):
        s.column_dimensions[col].width = w
    s.freeze_panes = "A3"

    # ------------------------------------------------------------------- charts
    def bar(title, ytitle, anchor, data_ref, cats_ref, log=False):
        ch = BarChart()
        ch.type, ch.style = "col", 10
        ch.title, ch.y_axis.title = title, ytitle
        ch.add_data(data_ref, titles_from_data=True)
        ch.set_categories(cats_ref)
        ch.height, ch.width = 8.5, 17
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        if log:
            ch.y_axis.scaling.logBase = 10
        s.add_chart(ch, anchor)

    bar("Uploads per channel per day, by race window", "uploads / channel / day", "M3",
        Reference(s, min_col=2, max_col=4, min_row=t2 + 1, max_row=t2_first + len(GROUPS) - 1),
        Reference(s, min_col=1, min_row=t2_first, max_row=t2_first + len(GROUPS) - 1))
    bar("Median relative views (Standard videos), by race window", "1.00 = channel's usual video", "M21",
        Reference(s, min_col=3, max_col=5, min_row=t3 + 1, max_row=t3_std_last),
        Reference(s, min_col=1, min_row=t3_std_first, max_row=t3_std_last))
    bar("Median views per subscriber (Standard videos), log scale", "views ÷ subscribers", "M39",
        Reference(s, min_col=9, min_row=4, max_row=t1_last),
        Reference(s, min_col=1, min_row=t1_first, max_row=t1_last), log=True)

    # ------------------------------------------------------------------- Notes text
    r = 13
    ws_notes.cell(row=r, column=1, value="Question").font = F_SUB
    ws_notes.cell(row=r + 1, column=1,
                  value="Do Formula 1's team channels, the sport's own channel, and independent creators behave differently "
                        "on YouTube, and does the race calendar change what they post and how it performs?").font = F_BASE
    r += 3
    ws_notes.cell(row=r, column=1, value="Where the data comes from").font = F_SUB
    for t in [
        "YouTube Data API v3, public statistics only (views, likes, comments, duration, publish time), pulled with yt_collect.py.",
        "Race dates: Jolpica F1 API (Ergast successor). Both are on the Calendar and Raw sheets as pulled.",
        "View counts are a snapshot taken on the pull date; they keep growing afterwards.",
    ]:
        r += 1
        ws_notes.cell(row=r, column=1, value=t).font = F_BASE
    r += 2
    ws_notes.cell(row=r, column=1, value="Cleaning steps (rows removed at each rule)").font = F_SUB
    for step, rows in log.itertuples(index=False):
        r += 1
        ws_notes.cell(row=r, column=1, value=step).font = F_BASE
        c = ws_notes.cell(row=r, column=2, value=int(rows))
        c.font, c.number_format = F_INPUT, "#,##0"
    if excluded is not None and len(excluded):
        r += 2
        ws_notes.cell(row=r, column=1, value="Channels pulled but left out of the analysis (and why)").font = F_SUB
        for lab, why in excluded.itertuples(index=False):
            r += 1
            ws_notes.cell(row=r, column=1, value=lab).font = F_BASE
            ws_notes.cell(row=r, column=3, value=why).font = F_BASE
    r += 2
    ws_notes.cell(row=r, column=1, value="Definitions").font = F_SUB
    for t in [
        "Short: duration up to 60 seconds. Standard: anything longer. (The API has no Shorts flag, so this is a proxy.)",
        "Race weekend: uploaded between the first practice day and race day, inclusive. Post-race: the 3 days after race day. Other: the rest.",
        "Dates are UTC. A video uploaded late on a Sunday in Asia can fall on the next UTC day.",
        "Relative views: a video's views divided by the median views of the same channel's videos in the same format and the same upload month. Older videos have had longer to collect views, so comparing within a month keeps the comparison fair.",
        "Engagement rate: (likes + comments) / views.",
    ]:
        r += 1
        ws_notes.cell(row=r, column=1, value=t).font = F_BASE
    r += 2
    ws_notes.cell(row=r, column=1, value="Limits to keep in mind").font = F_SUB
    for t in [
        "Views are a snapshot. Within a month, videos still differ in age by up to a few weeks, which matters most for the newest months.",
        "Team channels mix race content with promotion, partner and brand videos. This study does not separate them.",
        "Independent creators post rarely, so independent figures rest on few videos. Check the n columns.",
        "Observational data: it shows what happened, not why.",
    ]:
        r += 1
        ws_notes.cell(row=r, column=1, value=t).font = F_BASE
    r += 2
    ws_notes.cell(row=r, column=1, value="Where the findings are").font = F_SUB
    r += 1
    ws_notes.cell(row=r, column=1,
                  value="Summary = group and channel tables. Analysis = the deeper cuts behind the findings (highlight packages, "
                        "day-by-day rhythm, per-race checks, concentration). Every number in both is a formula over Clean.").font = F_BASE
    ws_notes.column_dimensions["A"].width = 52
    ws_notes.column_dimensions["B"].width = 14
    ws_notes.column_dimensions["C"].width = 80

    add_analysis(ws_an, clean, cal, last, cal_last, t1_first, t1_last, t2_first, channels, C)
    wb.save(args.out)
    print(f"Saved {args.out}: Clean rows {n}, Raw rows {len(raw_out)}, Days {n_days}")


if __name__ == "__main__":
    main()
