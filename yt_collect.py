#!/usr/bin/env python3
"""
yt_collect.py - pull public upload metadata and stats for a list of YouTube channels.

Uses only the YouTube Data API v3 (official, free quota) and the Python standard library.
Each API call used here costs 1 quota unit; the default daily quota is 10,000 units, and a
full run for ~8 channels over 12 months uses well under 300.

Setup
    1. Create an API key (Google Cloud Console -> APIs & Services -> YouTube Data API v3).
    2. export YOUTUBE_API_KEY="your-key"        (never commit the key)

Usage
    python yt_collect.py --dry-run                      # resolve channels only, no video pull
    python yt_collect.py --since 2025-11-17             # full pull, writes data/raw_videos.csv
    python yt_collect.py --channels data/channels_extra.csv --out data/raw_videos_extra.csv --since 2025-11-17
                                                        # a second batch; clean.py merges every raw_videos*.csv

channels.csv takes an @handle or a channel id (UC...) in the first column.

Steps per channel
    channels.list(forHandle)        -> channel id, title, subscribers, uploads-playlist id
    playlistItems.list (uploads)    -> every upload id + publish time, newest first, stop at --since
    videos.list (50 ids per call)   -> title, duration, views, likes, comments, live-stream flag
"""
import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

API = "https://www.googleapis.com/youtube/v3"

RAW_COLUMNS = [
    "channel_handle", "channel_label", "group", "channel_id", "channel_title", "channel_subs",
    "video_id", "title", "published_at", "duration_sec", "views", "likes", "comments",
    "was_live", "category_id", "pulled_at",
]


class QuotaError(RuntimeError):
    pass


def api_get(endpoint, params, key, retries=3):
    """GET one API endpoint, return parsed JSON. The key is never printed."""
    query = dict(params)
    query["key"] = key
    url = f"{API}/{endpoint}?{urllib.parse.urlencode(query)}"
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, timeout=30) as resp:
                return json.load(resp)
        except urllib.error.HTTPError as err:
            body = err.read().decode("utf-8", "replace")
            if err.code == 403 and "quota" in body.lower():
                raise QuotaError("YouTube API daily quota exhausted; try again tomorrow.")
            if err.code in (500, 502, 503, 504) and attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            raise RuntimeError(f"HTTP {err.code} from {endpoint}: {body[:300]}")
        except urllib.error.URLError:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
                continue
            raise
    raise RuntimeError(f"Failed after {retries} attempts: {endpoint}")


_DURATION = re.compile(r"^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$")


def parse_duration(iso):
    """ISO-8601 duration ('PT1H2M3S', 'P0D' for upcoming lives) -> seconds."""
    m = _DURATION.match(iso or "")
    if not m:
        return None
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    return d * 86400 + h * 3600 + mi * 60 + s


def to_int(value):
    """Counts come back as strings and are absent when hidden/disabled -> None."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def is_channel_id(text):
    """YouTube channel ids are 24 characters starting with 'UC'."""
    return len(text) == 24 and text.startswith("UC")


def resolve_channel(handle, key):
    """`handle` is either an @handle or a channel id (UC...)."""
    lookup = {"id": handle} if is_channel_id(handle) else {"forHandle": handle.lstrip("@")}
    data = api_get("channels", {"part": "snippet,contentDetails,statistics", **lookup}, key)
    items = data.get("items", [])
    if not items:
        return None
    ch = items[0]
    stats = ch.get("statistics", {})
    return {
        "channel_id": ch["id"],
        "channel_title": ch["snippet"]["title"],
        "channel_subs": None if stats.get("hiddenSubscriberCount") else to_int(stats.get("subscriberCount")),
        "uploads_playlist": ch["contentDetails"]["relatedPlaylists"]["uploads"],
    }


def list_upload_ids(playlist_id, since, key):
    """Return [(video_id, published_at)] newest-first, down to the --since date."""
    found, token = [], None
    while True:
        params = {"part": "contentDetails", "playlistId": playlist_id, "maxResults": 50}
        if token:
            params["pageToken"] = token
        data = api_get("playlistItems", params, key)
        page_all_old = True
        for item in data.get("items", []):
            cd = item["contentDetails"]
            published = cd.get("videoPublishedAt")
            if not published:  # private / deleted entry
                continue
            if published[:10] >= since:
                page_all_old = False
                found.append((cd["videoId"], published))
        token = data.get("nextPageToken")
        if not token or page_all_old:
            break
    return found


def fetch_video_rows(ids, key):
    rows = {}
    for i in range(0, len(ids), 50):
        batch = ids[i:i + 50]
        data = api_get("videos", {"part": "snippet,contentDetails,statistics,liveStreamingDetails",
                                  "id": ",".join(batch), "maxResults": 50}, key)
        for v in data.get("items", []):
            sn, st = v["snippet"], v.get("statistics", {})
            rows[v["id"]] = {
                "video_id": v["id"],
                "title": sn["title"],
                "published_at": sn["publishedAt"],
                "duration_sec": parse_duration(v["contentDetails"].get("duration")),
                "views": to_int(st.get("viewCount")),
                "likes": to_int(st.get("likeCount")),
                "comments": to_int(st.get("commentCount")),
                "was_live": int("liveStreamingDetails" in v),
                "category_id": sn.get("categoryId"),
            }
    return rows


def read_channels(path):
    with open(path, newline="", encoding="utf-8") as f:
        return [r for r in csv.DictReader(f) if r.get("handle", "").strip()]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--channels", default="data/channels.csv")
    ap.add_argument("--since", default=(dt.date.today() - dt.timedelta(days=365)).isoformat(),
                    help="earliest publish date to keep, YYYY-MM-DD (default: 365 days ago)")
    ap.add_argument("--out", default="data/raw_videos.csv")
    ap.add_argument("--resolved-out", default=None,
                    help="default: derived from --out (raw_videos_x.csv -> channels_resolved_x.csv)")
    ap.add_argument("--dry-run", action="store_true", help="resolve channels only; skip the video pull")
    args = ap.parse_args()

    # Windows consoles default to a legacy codepage; channel titles can contain any character.
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    key = os.environ.get("YOUTUBE_API_KEY")
    if not key:
        sys.exit("Set YOUTUBE_API_KEY first (see the docstring at the top of this file).")

    if args.resolved_out is None:
        folder, name = os.path.split(args.out)
        args.resolved_out = os.path.join(folder, name.replace("raw_videos", "channels_resolved", 1))

    channels = read_channels(args.channels)
    pulled_at = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    resolved, all_rows, problems = [], [], []

    for ch in channels:
        handle = ch["handle"].strip()
        info = resolve_channel(handle, key)
        if info is None:
            problems.append(handle)
            print(f"  ! {handle}: no channel found for this handle - fix it in {args.channels}")
            continue
        resolved.append({"handle": handle, "label": ch["label"], "group": ch["group"],
                         "channel_id": info["channel_id"], "channel_title": info["channel_title"],
                         "channel_subs": info["channel_subs"]})
        print(f"  {handle:<18} -> {info['channel_title']} ({info['channel_subs']} subs)")
        if args.dry_run:
            continue

        uploads = list_upload_ids(info["uploads_playlist"], args.since, key)
        stats = fetch_video_rows([vid for vid, _ in uploads], key)
        for vid, _ in uploads:
            if vid not in stats:
                continue
            all_rows.append({
                "channel_handle": handle, "channel_label": ch["label"], "group": ch["group"],
                "channel_id": info["channel_id"], "channel_title": info["channel_title"],
                "channel_subs": info["channel_subs"], **stats[vid], "pulled_at": pulled_at,
            })
        print(f"      {len([r for r in all_rows if r['channel_handle'] == handle])} videos since {args.since}")

    os.makedirs(os.path.dirname(args.resolved_out) or ".", exist_ok=True)
    with open(args.resolved_out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["handle", "label", "group", "channel_id", "channel_title", "channel_subs"])
        w.writeheader()
        w.writerows(resolved)

    if not args.dry_run:
        with open(args.out, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=RAW_COLUMNS)
            w.writeheader()
            w.writerows(all_rows)
        print(f"\nWrote {len(all_rows)} rows to {args.out} (stats snapshot taken {pulled_at})")

    if problems:
        print(f"\nUnresolved handles: {', '.join(problems)}")


if __name__ == "__main__":
    main()
