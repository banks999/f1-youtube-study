import re, sys, glob, os
import os
ART = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
CSSH = {"val": 34, "xl": 34, "cn": 36, "cv": 36, "cl": 30, "sub": 30, "base": 2, "ref": 2, "hair": 1}
def px(style, key):
    m = re.search(rf"(?<![-\w]){key}:(-?\d+(?:\.\d+)?)px", style)
    return float(m.group(1)) if m else None
bad = 0
for f in sorted(glob.glob(f"{ART}/slides/*.html")):
    html = open(f, encoding="utf-8").read()
    name = os.path.basename(f)
    # hosts: split on '<div class="host"' and take until the matching close by counting divs
    for m in re.finditer(r'<div class="host" style="([^"]*)">', html):
        W, H = px(m.group(1), "width"), px(m.group(1), "height")
        i = m.end(); depth = 1; j = i
        while depth and j < len(html):
            o = html.find("<div", j); c = html.find("</div>", j)
            if c == -1: break
            if o != -1 and o < c: depth += 1; j = o + 4
            else: depth -= 1; j = c + 6
        body = html[i:j - 6]
        kids = re.findall(r'<(?:div|p) class="([\w ]+)" style="([^"]*)"', body)
        if len(kids) > 24: print(name, "HOST HAS", len(kids), "CHILDREN"); bad += 1
        for cls, st in kids:
            l, t, w, h = px(st, "left"), px(st, "top"), px(st, "width"), px(st, "height")
            c0 = cls.split()[0]
            if h is None: h = CSSH.get(c0)
            if w is None and c0 in ("base", "ref"): w = W
            if w is None: w = 2 if c0 == "vref" else 0
            if l is None: l = 0
            if h is None: h = 0
            if l + w > W + 0.5 or t + h > H + 0.5 or l < 0 or t < 0:
                print(f"{name}: {cls} box l={l} t={t} w={w} h={h} outside host {W}x{H}"); bad += 1
    # slide-level pinned
    if 'id="cover"' in html:
        for st in re.findall(r'style="(position:absolute[^"]*)"', html):
            l, t, w, h = (px(st, k) for k in ("left", "top", "width", "height"))
            if l + w > 1792 or t + h > 952: print(name, "cover pinned outside", st); bad += 1
print("geometry problems:", bad)
