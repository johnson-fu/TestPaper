#!/usr/bin/env python3
"""驗證學生卷與解答卷 PDF：單頁、紙張尺寸、填空區格線位置完全一致。

用法：
    python3 verify_sheets.py config.json student.pdf key.pdf
需要：pdfinfo、pdftoppm（poppler-utils）、Pillow、numpy
"""
import json
import os
import re
import subprocess
import sys
import tempfile

import numpy as np
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_answer_sheets as M  # noqa: E402

DPI = 150
PX = DPI / 25.4  # 每 mm 像素數


def pages_and_size(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    n = int(re.search(r"Pages:\s+(\d+)", out).group(1))
    w, h = map(float, re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", out).groups())
    return n, w * 25.4 / 72, h * 25.4 / 72


def render(pdf, tmp, tag):
    base = os.path.join(tmp, tag)
    subprocess.run(["pdftoppm", "-r", str(DPI), "-gray", "-f", "1", "-l", "1", pdf, base], check=True)
    f = [x for x in os.listdir(tmp) if x.startswith(tag)][0]
    return np.array(Image.open(os.path.join(tmp, f)))


def edges(v):
    d = v < 100
    return [i for i in range(1, len(d)) if d[i] and not d[i - 1]]


def main():
    cfg_path, stu, key = sys.argv[1:4]
    with open(cfg_path, encoding="utf-8") as f:
        user = json.load(f)
    cfg = json.loads(json.dumps(M.DEFAULTS))
    for k, v in user.items():
        if isinstance(v, dict) and isinstance(cfg.get(k), dict):
            cfg[k].update(v)
        else:
            cfg[k] = v
    pa = cfg["paper"]
    W, H, NCOL, ROWS, ROW_H, _ = M.geometry(cfg)
    num_w = cfg["fill_in"].get("num_width_mm", 34)
    ok = True

    for name, pdf in (("學生卷", stu), ("解答卷", key)):
        n, w, h = pages_and_size(pdf)
        good = n == 1 and abs(w - pa["width_mm"]) < 1 and abs(h - pa["height_mm"]) < 1
        ok &= good
        print(f"{'✔' if good else '✘'} {name}：{n} 頁，{w:.0f}×{h:.0f} mm")

    with tempfile.TemporaryDirectory() as tmp:
        a, b = render(stu, tmp, "s"), render(key, tmp, "k")
        top = pa["margin_tb_mm"] + M.TITLE_H + M.GAP_TITLE + M.ID_ROW_H + M.GAP_SECTION + M.HEADING_H
        bot = top + ROWS * ROW_H
        y0, y1 = int((top - 1) * PX), int((bot + 1) * PX)
        # 橫線：在每一欄答案格中間往下掃（兩份卷子的答案格都是空白）
        for k in range(NCOL):
            x = int((pa["margin_lr_mm"] + k * W / NCOL + num_w + 5) * PX)
            ea, eb = edges(a[y0:y1, x]), edges(b[y0:y1, x])
            same = ea == eb and len(ea) == ROWS + 1
            ok &= same
            print(f"{'✔' if same else '✘'} 第 {k + 1} 欄橫線 {len(ea)} 條" + ("" if same else f"（解答卷 {len(eb)} 條）"))
        # 直線：在第一列頂端下方 1 mm 橫向掃描（避開文字）
        y = int((top + 1) * PX)
        ea, eb = edges(a[y, :]), edges(b[y, :])
        same = ea == eb
        ok &= same
        print(f"{'✔' if same else '✘'} 直線位置 {len(ea)} 條" + ("" if same else f"：{ea} vs {eb}"))
        # 最下緣留白
        rows = np.where((a < 200).any(axis=1))[0]
        bottom_gap = pa["height_mm"] - rows[-1] / PX
        print(f"  學生卷內容距頁底 {bottom_gap:.1f} mm")

    print("全部通過" if ok else "有項目未通過")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
