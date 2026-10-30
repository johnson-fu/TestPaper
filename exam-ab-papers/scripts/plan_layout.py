#!/usr/bin/env python3
"""規劃題目卷換欄位置與答案卷作答框高度 → layout.json

1. 以 probe（空白卷、自然排版）量出每題高度。
2. 用 DP 選出換欄題號，讓每欄填滿比例盡量平均（計算題接在最後一欄）。
3. 每欄填滿比例上限 max_fill（預設 0.93）；放不下就把字級降 0.5pt 重試。
4. 依各欄剩餘空間加大題間距（同一份卷採同一間距，避免各欄疏密不一）。
5. 量答案卷最下方框線，把計算題作答框高度調到頁尾前約 22pt。

用法：python3 plan_layout.py [--pt 14] [--ls 1.2] [--max-fill 0.93]   （在考試工作資料夾內）
"""
import argparse, json, os, subprocess, sys
from PIL import Image
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdfmeasure as pm

HERE = os.path.dirname(os.path.abspath(__file__))
GEN = os.path.join(HERE, "gen_paper.js")
cfg = json.load(open("exam.config.json", encoding="utf-8"))
ap = argparse.ArgumentParser()
ap.add_argument("--pt", type=float, default=cfg.get("pt", 14))
ap.add_argument("--ls", type=float, default=cfg.get("ls", 1.2))
ap.add_argument("--max-fill", type=float, default=0.93)
a = ap.parse_args()
PAGES = cfg.get("paper_pages", 2); K = 2 * PAGES
os.makedirs("probe", exist_ok=True)
n_mc = int(subprocess.run(["node", "-e", 'console.log(require(require("path").resolve(process.argv[1])).mc.length)',
                           cfg.get("data", "exam_data.js")], capture_output=True, text=True, check=True).stdout)
G0 = 100


def gen(kind, ver, out):
    r = subprocess.run(["node", GEN, kind, ver, out, "probe"], capture_output=True, text=True)
    if r.returncode: sys.exit(r.stderr)


def plan(ver, layout):
    json.dump(layout, open("layout.json", "w"), indent=1)
    gen("paper", ver, f"probe/{ver}.docx")
    m = pm.measure(pm.render(f"probe/{ver}.docx"), n_mc)
    h, calc = m["h"], m["calc"]
    head = m["first_q"] - m["top1"]
    top_body = 40
    cap = [pm.BOTTOM - m["top1"] - head, pm.BOTTOM - m["top1"] - 6] + [pm.BOTTOM - top_body] * 2
    cap = cap[:K] if PAGES == 2 else cap[:2]
    pre = [0]
    for x in h: pre.append(pre[-1] + x)
    seg = lambda i, j: pre[j] - pre[i]          # 題 i..j-1（0 起算）
    best = None
    # DP：dp[k][j] = 前 j 題放進前 k 欄的最小「最大填滿比例」
    INF = float("inf")
    dp = [[INF] * (n_mc + 1) for _ in range(K + 1)]; arg = [[0] * (n_mc + 1) for _ in range(K + 1)]
    dp[0][0] = 0
    for k in range(1, K + 1):
        for j in range(n_mc + 1):
            for i in range(j + 1):
                if i == j and not (k == K and j == n_mc): continue   # 只有最後一欄可以只放計算題
                extra = calc if (k == K and j == n_mc) else 0
                r = max(dp[k - 1][i], (seg(i, j) + extra) / cap[k - 1])
                if r < dp[k][j]: dp[k][j], arg[k][j] = r, i
    ratio = dp[K][n_mc]
    cuts, j = [], n_mc
    for k in range(K, 0, -1):
        i = arg[k][j]; cuts.append((i, j)); j = i
    cuts.reverse()
    cols = [seg(i, j) + (calc if k == K - 1 else 0) for k, (i, j) in enumerate(cuts)]
    breaks = [i + 1 for (i, j) in cuts[1:] if i < n_mc] + ([n_mc + 1] if cuts[-1][0] == n_mc else [])
    # 依最滿的那一欄決定統一題間距
    gap_extra = min(((a.max_fill * c - u) / max(j - i - 1, 1)) for (i, j), c, u in zip(cuts, cap, cols) if j > i)
    gap = int(G0 + max(0, min(gap_extra * 20, 160)))
    return ratio, breaks, gap, [round(c / k, 3) for c, k in zip(cols, cap)]


layout = {}
pt = a.pt
while True:
    ok = True
    for ver in "AB":
        layout[ver] = {"pt": pt, "ls": a.ls, "gaps": [G0] * K}
    for ver in "AB":
        ratio, breaks, gap, fills = plan(ver, layout)
        print(f"{ver} 卷：字級 {pt}pt，換欄於第 {breaks} 題前，各欄填滿 {fills}，題間距 {gap} twips")
        if ratio > a.max_fill: ok = False; break
        layout[ver]["gaps"] = [G0] * K
        layout[ver].update({"breaks": breaks, "gaps": [gap] * K})
    if ok:   # 以最終版面實際轉檔確認頁數與欄底
        json.dump(layout, open("layout.json", "w"), indent=1)
        for ver in "AB":
            gen("paper", ver, f"probe/{ver}_final.docx")
            pdf = pm.render(f"probe/{ver}_final.docx")
            low = max(w[4] for w in pm.words(pdf) if w[4] < pm.BOTTOM + 5)
            if pm.pages(pdf) != PAGES or low > pm.BOTTOM - 10:
                print(f"{ver} 卷實際轉檔為 {pm.pages(pdf)} 頁（欄底 {low:.0f}pt），需降字級"); ok = False
    if ok: break
    pt -= 0.5
    if pt < 10: sys.exit("字級降到 10pt 仍放不下，請減少題目或改為 3 頁以上")
    print(f"放不下 {PAGES} 頁，字級降為 {pt}pt 重試")

# ---------- 答案卷作答框高度 ----------
M = len(json.loads(subprocess.run(["node", "-e", 'console.log(JSON.stringify(require(require("path").resolve(process.argv[1])).calc||[]))',
                                   cfg.get("data", "exam_data.js")], capture_output=True, text=True).stdout))
if M:
    w = cfg.get("sheet", {}).get("calc_weights", [1] * M)
    base = [int(12950 * x / sum(w)) for x in w]
    for ver in "AB": layout[ver]["calc_heights"] = base
    json.dump(layout, open("layout.json", "w"), indent=1)
    gen("sheet", "A", "probe/sheet.docx")
    pdf = pm.render("probe/sheet.docx")
    subprocess.run(["pdftoppm", "-gray", "-r", "100", "-singlefile", pdf, "probe/sheet"], check=True)
    im = Image.open("probe/sheet.pgm"); W, Hh = im.size; px = im.load()
    rows = [y for y in range(Hh) if sum(1 for x in range(W // 4, 3 * W // 4) if px[x, y] < 100) > W * 0.45]
    page_h = 20636 / 20; bottom = page_h - 794 / 20
    last = rows[-1] / Hh * page_h
    slack_tw = int((bottom - 22 - last) * 20)
    pages_now = pm.pages(pdf)
    if pages_now > 1: slack_tw = -3000
    new = [max(1500, int(b + slack_tw * x / sum(w))) for b, x in zip(base, w)]
    for ver in "AB": layout[ver]["calc_heights"] = new
    print(f"答案卷作答框高度（twips）：{new}")
json.dump(layout, open("layout.json", "w"), indent=1)
print("→ layout.json")
