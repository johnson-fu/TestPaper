#!/usr/bin/env python3
"""產生全部考卷並轉 PDF、自動驗證。

輸出（out/）：
  <exam_name>_題目卷_AB混合.docx   A3，每人 paper_pages 頁，依 students.json 順序（A、B 混合）
  <exam_name>_答案卷_AB混合.docx   B4，每人 1 頁，順序同上
  <exam_name>_批改卷_A卷.docx / _B卷.docx   B4，可疊在答案卷上
  <exam_name>_參考答案.docx         A4，教師用
  另可加 --split 產生 A、B 分開的題目卷／答案卷。

用法：python3 build_all.py [--split] [--no-pdf]   （在考試工作資料夾內，需先跑 prep_roster / make_order / plan_layout）
"""
import argparse, json, os, subprocess, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pdfmeasure as pm

HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument("--split", action="store_true", help="另外產生 A、B 分開的檔案")
ap.add_argument("--no-pdf", action="store_true")
ap.add_argument("--out", default="out")
a = ap.parse_args()
cfg = json.load(open("exam.config.json", encoding="utf-8"))
name = (cfg.get("exam_name") or cfg.get("title", "考試").rstrip("卷")).replace(" ", "_")
os.makedirs(a.out, exist_ok=True)

def node(*args):
    r = subprocess.run(["node", *args], capture_output=True, text=True)
    if r.returncode: sys.exit(r.stderr)
    print(r.stdout.strip())

jobs = [("paper", "ALL", f"{name}_題目卷_AB混合"), ("sheet", "ALL", f"{name}_答案卷_AB混合"),
        ("key", "A", f"{name}_批改卷_A卷"), ("key", "B", f"{name}_批改卷_B卷")]
if a.split:
    jobs += [("paper", v, f"{name}_題目卷_{v}卷") for v in "AB"] + [("sheet", v, f"{name}_答案卷_{v}卷") for v in "AB"]
files = {}
for kind, ver, base in jobs:
    out = os.path.join(a.out, base + ".docx")
    node(os.path.join(HERE, "gen_paper.js"), kind, ver, out)
    files[(kind, ver)] = out
key = os.path.join(a.out, f"{name}_參考答案.docx")
node(os.path.join(HERE, "build_key.js"), key)

if not a.no_pdf:
    for f in list(files.values()) + [key]:
        pm.render(f)
    # 對齊檢查需要一份空白答案卷
    node(os.path.join(HERE, "gen_paper.js"), "sheet", "A", os.path.join("probe", "sheet_blank.docx"), "probe")
    pm.render(os.path.join("probe", "sheet_blank.docx"))
    r = subprocess.run([sys.executable, os.path.join(HERE, "verify_outputs.py"),
                        "--paper", files[("paper", "ALL")][:-5] + ".pdf",
                        "--sheet", files[("sheet", "ALL")][:-5] + ".pdf",
                        "--key-a", files[("key", "A")][:-5] + ".pdf",
                        "--key-b", files[("key", "B")][:-5] + ".pdf",
                        "--blank-sheet", os.path.join("probe", "sheet_blank.pdf")])
    sys.exit(r.returncode)
