#!/usr/bin/env python3
"""產生 A、B 卷選擇題題序 → orders.json

A 卷＝題庫原順序；B 卷在 exam.config.json 的 shuffle.groups 範圍內各自打亂，並滿足：
  * 每一題在 B 卷的位置都和 A 卷不同
  * A 卷相鄰的兩題，在 B 卷不會以相同次序相鄰（keep_together 指定的除外）
  * keep_together 內的題目在 B 卷保持相連、次序不變（例如後一題沿用前一題的定義）
  * 同一個答案最多連續 max_same_answer_run 題
  * 有附圖的題目彼此間隔 > 2 題（盡量做到，做不到就放寬）

shuffle 設定（題號皆為 A 卷題號、從 1 起算，範圍含頭尾）：
  {"seed": 11530, "groups": [[1, 16], [17, 35]], "keep_together": [[31, 32]], "max_same_answer_run": 3}
未設定 groups 時整份一起打亂。

用法：python3 make_order.py   （在考試工作資料夾內）
"""
import json, random, subprocess, sys, os

cfg = json.load(open("exam.config.json", encoding="utf-8")) if os.path.exists("exam.config.json") else {}
data = cfg.get("data", "exam_data.js")
js = 'const {mc}=require(require("path").resolve(process.argv[1]));console.log(JSON.stringify(mc.map(q=>({a:q.ans,f:q.fig||null}))))'
mc = json.loads(subprocess.run(["node", "-e", js, data], capture_output=True, text=True, check=True).stdout)
N = len(mc); ans = [q["a"] for q in mc]; figs = [i for i, q in enumerate(mc) if q["f"]]
sh = cfg.get("shuffle", {})
groups = [(a - 1, b - 1) for a, b in sh.get("groups", [[1, N]])]
keep = [[x - 1 for x in blk] for blk in sh.get("keep_together", [])]
maxrun = sh.get("max_same_answer_run", 3)
covered = sorted(i for a, b in groups for i in range(a, b + 1))
if covered != list(range(N)): sys.exit("shuffle.groups 必須剛好涵蓋全部題目且不重疊")
kept_pairs = {(blk[i], blk[i + 1]) for blk in keep for i in range(len(blk) - 1)}

def ok(order, fig_gap):
    if any(order[i] == i for i in range(N)): return False
    adjA = {(i, i + 1) for i in range(N - 1)} - kept_pairs
    if any((order[i], order[i + 1]) in adjA for i in range(N - 1)): return False
    run = 1
    for i in range(1, N):
        run = run + 1 if ans[order[i]] == ans[order[i - 1]] else 1
        if run > maxrun: return False
    pos = [order.index(f) for f in figs]
    return all(abs(p - q) > fig_gap for p in pos for q in pos if p != q)

rng = random.Random(sh.get("seed", 1))
for fig_gap in (2, 1, 0):
    for tries in range(200000):
        order = []
        for a, b in groups:
            units, inblk = [], set()
            for blk in keep:
                if a <= blk[0] <= b: units.append(blk); inblk |= set(blk)
            units += [[i] for i in range(a, b + 1) if i not in inblk]
            rng.shuffle(units)
            order += [x for u in units for x in u]
        if ok(order, fig_gap): break
    else:
        continue
    break
else:
    sys.exit("找不到符合條件的 B 卷題序：請放寬 max_same_answer_run 或調整 groups")

json.dump({"A": list(range(N)), "B": order}, open("orders.json", "w"))
print("B 卷題序（A 卷題號）：", [o + 1 for o in order])
print("A 卷答案：", "".join(ans))
print("B 卷答案：", "".join(ans[o] for o in order))
