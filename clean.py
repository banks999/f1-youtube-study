#!/usr/bin/env python3
"""
clean.py - turn the raw YouTube pull into the analysis table.

Reads   data/raw_videos*.csv   (one or more pulls from yt_collect.py, merged)
        data/f1_calendar.csv
Writes  data/clean_videos.csv  (one row per kept video, with the derived columns below)
        data/cleaning_log.csv  (how many rows each rule removed, in the order applied)

Rules, applied in this order
  0. Out of scope    channels listed in data/excluded_channels.csv (with the reason) are left out.
  1. Duplicates      same video_id in more than one pull -> keep the most recent pull.
  2. Live streams    YouTube reports live-stream details -> drop. Streams (press conferences,
                     full sessions) collect views differently from normal uploads.
  3. Too recent      published < MIN_AGE_DAYS days before the pull -> drop. Their view counts
                     are still climbing and would look artificially weak.
  4. Bad duration    missing or zero length -> drop.

Derived columns (the Excel workbook recomputes these with formulas; this script is the cross-check)
  format           'Short' if duration <= 60 s, else 'Standard'
  age_days         days between publish date (UTC) and the pull date
  race_window      'Race weekend'  publish date from the weekend's first practice day to race day
                   'Post-race'     the 3 days after race day
                   'Other'         everything else
  race_name        the Grand Prix the window belongs to (blank for 'Other')
  sprint_weekend   1 if that race weekend had a sprint
  engagement_rate  (likes + comments) / views
  views_per_sub    views / channel subscribers at pull time
"""
import argparse
import glob
import os

import pandas as pd

SHORT_MAX_SEC = 60
POST_RACE_DAYS = 3


def assign_race_window(pub_dates, calendar):
    """Vectorised-enough: 5k videos x 47 races. Returns (window, race_name, sprint) lists."""
    windows, names, sprints = [], [], []
    cal = list(calendar.itertuples(index=False))
    for d in pub_dates:
        window, name, sprint = "Other", "", 0
        for r in cal:
            if r.weekend_start <= d <= r.race_date:
                window, name, sprint = "Race weekend", r.race_name, int(r.sprint_weekend)
                break
        else:
            for r in cal:
                if r.race_date < d <= r.race_date + pd.Timedelta(days=POST_RACE_DAYS):
                    window, name, sprint = "Post-race", r.race_name, int(r.sprint_weekend)
                    break
        windows.append(window)
        names.append(name)
        sprints.append(sprint)
    return windows, names, sprints


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--raw-glob", default="data/raw_videos*.csv")
    ap.add_argument("--calendar", default="data/f1_calendar.csv")
    ap.add_argument("--out", default="data/clean_videos.csv")
    ap.add_argument("--log", default="data/cleaning_log.csv")
    ap.add_argument("--excluded", default="data/excluded_channels.csv")
    ap.add_argument("--min-age-days", type=int, default=7)
    args = ap.parse_args()

    files = sorted(glob.glob(args.raw_glob))
    if not files:
        raise SystemExit(f"No files match {args.raw_glob}")
    raw = pd.concat([pd.read_csv(f) for f in files], ignore_index=True)
    log = [("raw rows (all pulls merged)", len(raw))]

    # 0. channels deliberately left out of scope (listed with reasons in data/excluded_channels.csv)
    if os.path.exists(args.excluded):
        excluded = pd.read_csv(args.excluded)
        raw = raw[~raw["channel_label"].isin(excluded["label"])]
        log.append(("after leaving out channels in excluded_channels.csv", len(raw)))

    # 1. duplicates: keep the latest pull of each video
    raw = raw.sort_values("pulled_at").drop_duplicates("video_id", keep="last")
    log.append(("after removing duplicate video_ids", len(raw)))

    # 2. live streams
    raw = raw[raw["was_live"] == 0]
    log.append(("after removing live streams", len(raw)))

    # 3. too recent
    raw["published_at"] = pd.to_datetime(raw["published_at"], utc=True)
    raw["pulled_at"] = pd.to_datetime(raw["pulled_at"], utc=True)
    raw["pub_date"] = raw["published_at"].dt.tz_localize(None).dt.normalize()
    pull_date = raw["pulled_at"].dt.tz_localize(None).dt.normalize()
    raw["age_days"] = (pull_date - raw["pub_date"]).dt.days
    raw = raw[raw["age_days"] >= args.min_age_days]
    log.append((f"after removing videos younger than {args.min_age_days} days", len(raw)))

    # 4. bad duration
    raw = raw[raw["duration_sec"].fillna(0) > 0]
    log.append(("after removing missing/zero duration", len(raw)))

    # derived columns
    raw["format"] = raw["duration_sec"].apply(lambda s: "Short" if s <= SHORT_MAX_SEC else "Standard")
    cal = pd.read_csv(args.calendar, parse_dates=["weekend_start", "race_date"])
    raw["race_window"], raw["race_name"], raw["sprint_weekend"] = assign_race_window(raw["pub_date"], cal)
    raw["engagement_rate"] = (raw["likes"] + raw["comments"]) / raw["views"]
    raw["views_per_sub"] = raw["views"] / raw["channel_subs"]

    cols = ["channel_label", "group", "channel_subs", "video_id", "title", "published_at", "pub_date",
            "age_days", "duration_sec", "format", "views", "likes", "comments", "race_window",
            "race_name", "sprint_weekend", "engagement_rate", "views_per_sub"]
    out = raw[cols].sort_values(["channel_label", "published_at"]).reset_index(drop=True)
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    out.to_csv(args.out, index=False)
    pd.DataFrame(log, columns=["step", "rows"]).to_csv(args.log, index=False)

    print(pd.DataFrame(log, columns=["step", "rows"]).to_string(index=False))
    print(f"\nWrote {len(out)} rows -> {args.out}")


if __name__ == "__main__":
    main()
