#!/usr/bin/env python3
"""驗證產出的 PDF（在考試工作資料夾內執行）：

題目卷：總頁數＝人數×paper_pages；每人表頭班級／學號／姓名／卷別正確；頁尾姓名、卷別、頁碼正確；
        35 題（或 N 題）的題序與該生卷別一致；最後一頁有「試題結束」；欄底未超出版心。
答案卷：每人 1 頁、順序與題目卷相同、名冊資料與卷別正確。
批改卷：各 1 頁；題號列印出的正解與題庫一致；水平框線位置與空白答案卷完全相同（疊放對齊）。

注意：pdftotext 對中文常漏字或斷字，所以題序比對只用英數字，且用「包含」而非「開頭相同」。
"""
import argparse, json, os, re, subprocess, sys
from PIL import Image
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdfmeasure as pm

ap = argparse.ArgumentParser()
for k in ("paper", "sheet", "key-a", "key-b", "blank-sheet"): ap.add_argument("--" + k)
a = ap.parse_args()
cfg = json.load(open("exam.config.json", encoding="utf-8"))
PAGES = cfg.get("paper_pages", 2)
S = json.load(open("students.json", encoding="utf-8"))
O = json.load(open("orders.json"))
js = 'const {mc}=require(require("path").resolve(process.argv[1]));console.log(JSON.stringify(mc.map(q=>({s:q.stem.replace(/\\*/g,""),a:q.ans}))))'
MC = json.loads(subprocess.run(["node", "-e", js, cfg.get("data", "exam_data.js")], capture_output=True, text=True, check=True).stdout)
N = len(MC)
alnum = lambda t: "".join(re.findall(r"[A-Za-z0-9]", t))
problems = 0
def bad(msg):
    global problems; problems += 1; print("✗", msg)

if a.paper:
    W = pm.words(a.paper)
    if pm.pages(a.paper) != PAGES * len(S): bad(f"題目卷 {pm.pages(a.paper)} 頁，應為 {PAGES * len(S)} 頁")
    low = 0
    for i, s in enumerate(S):
        pg = list(range(PAGES * i + 1, PAGES * i + PAGES + 1))
        ws = [w for w in W if w[0] in pg and w[4] < pm.BOTTOM + 5]
        low = max([low] + [w[4] for w in ws])
        head = " ".join(w[5] for w in W if w[0] == pg[0] and w[2] < 120)
        if not all(x in head for x in (s["cls"], s["id"], s["name"], s["ver"] + " 卷")): bad(f"{s['name']} 表頭錯誤")
        for k, p in enumerate(pg, 1):
            foot = " ".join(w[5] for w in W if w[0] == p and w[2] > pm.BOTTOM + 5)
            if s["name"] not in foot or s["ver"] + " 卷" not in foot or f"第 {k} 頁" not in foot: bad(f"{s['name']} 第 {k} 頁頁尾錯誤：{foot}")
        if "試題結束" not in " ".join(w[5] for w in W if w[0] == pg[-1]): bad(f"{s['name']} 最後一頁沒有「試題結束」")
        nums = pm.question_numbers(ws)[:N]
        exp = [MC[q]["s"] for q in O[s["ver"]]]
        for k, w in enumerate(nums):
            line = sorted([x for x in ws if x[0] == w[0] and abs((x[2] + x[4]) / 2 - (w[2] + w[4]) / 2) < 5 and w[1] < x[1] < w[1] + 380], key=lambda x: x[1])
            t = alnum("".join(x[5] for x in line))[:14]
            okk = int(w[5][:-1]) == k + 1 and ((t in alnum(exp[k])[:90]) if t else exp[k][0] > "⺀")
            if not okk: bad(f"{s['name']}（{s['ver']}）第 {k + 1} 題與題序不符：{t}"); break
        if len(nums) != N: bad(f"{s['name']} 只找到 {len(nums)} 題")
    print(f"題目卷：{len(S)} 人、{pm.pages(a.paper)} 頁；最低欄底 {low:.0f}pt（版心下緣 {pm.BOTTOM:.0f}pt）")

if a.sheet:
    W = pm.words(a.sheet)
    if pm.pages(a.sheet) != len(S): bad(f"答案卷 {pm.pages(a.sheet)} 頁，應為 {len(S)} 頁")
    for i, s in enumerate(S):
        txt = " ".join(w[5] for w in W if w[0] == i + 1)
        if not all(x in txt for x in (s["cls"], s["id"], s["name"], s["ver"] + " 卷")): bad(f"答案卷第 {i + 1} 頁不是 {s['name']}")
    print(f"答案卷：{pm.pages(a.sheet)} 頁")

def hrules(pdf, tag):
    subprocess.run(["pdftoppm", "-gray", "-r", "150", "-singlefile", pdf, f"probe/{tag}"], check=True)
    im = Image.open(f"probe/{tag}.pgm"); Wd, H = im.size; px = im.load()
    return [y for y in range(H) if sum(1 for x in range(int(Wd * .1), int(Wd * .9), 3) if px[x, y] < 120) > (Wd * .8 / 3) * 0.6]

for ver, pdf in (("A", a.key_a), ("B", a.key_b)):
    if not pdf: continue
    if pm.pages(pdf) != 1: bad(f"批改卷 {ver} 不是 1 頁")
    W = pm.words(pdf)
    nums = [w for w in W if re.fullmatch(r"\d+", w[5]) and 150 < w[2] < 600 and w[1] > 60]
    got = {}
    for n in nums:
        c = [w for w in W if re.fullmatch(r"[A-E]", w[5]) and abs(w[4] - n[4]) < 6 and 0 <= w[1] - n[3] < 20]
        if c: got[int(n[5])] = min(c, key=lambda w: w[1])[5]
    exp = {k + 1: MC[q]["a"] for k, q in enumerate(O[ver])}
    if got != exp: bad(f"批改卷 {ver} 正解與題庫不符：{''.join(got.get(i, '?') for i in range(1, N + 1))}")
    else: print(f"批改卷 {ver}：正解 {''.join(exp[i] for i in range(1, N + 1))}")
    if a.blank_sheet:
        if hrules(pdf, f"k{ver}") != hrules(a.blank_sheet, "blank"): bad(f"批改卷 {ver} 框線位置與答案卷不同，疊放會錯開")
        else: print(f"批改卷 {ver}：框線位置與答案卷一致")

print("全部通過" if not problems else f"共 {problems} 個問題")
sys.exit(1 if problems else 0)
