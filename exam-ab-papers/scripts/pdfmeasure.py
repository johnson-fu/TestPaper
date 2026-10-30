"""PDF 轉檔與量測工具（plan_layout.py、verify_outputs.py 共用）"""
import glob, html, os, re, shutil, subprocess

A3_H = 1190.55                     # pt
BOTTOM = A3_H - (794 + 120) / 20   # 題目卷本文下緣（pt）
COL_SPLIT = 421                    # 左右欄分界 x（pt）


def soffice_cmd():
    """優先用 Claude docx skill 的 soffice 包裝（沙盒環境），否則用系統的 soffice / libreoffice。"""
    for p in glob.glob("/root/.claude/skills/**/docx/scripts/office/soffice.py", recursive=True) + \
             glob.glob("/mnt/skills/public/docx/scripts/office/soffice.py"):
        return ["python3", p]
    for exe in ("soffice", "libreoffice", r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if shutil.which(exe) or os.path.exists(exe):
            return [exe]
    raise SystemExit("找不到 LibreOffice（soffice）")


def render(docx):
    outdir = os.path.dirname(os.path.abspath(docx))
    subprocess.run(soffice_cmd() + ["--headless", "--convert-to", "pdf", "--outdir", outdir, docx],
                   capture_output=True, check=False)
    pdf = os.path.splitext(docx)[0] + ".pdf"
    if not os.path.exists(pdf):
        raise SystemExit(f"轉 PDF 失敗：{docx}")
    return pdf


def pages(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    return int(re.search(r"Pages:\s+(\d+)", out).group(1))


def words(pdf):
    """回傳 (page, xMin, yMin, xMax, yMax, text) 清單。"""
    x = subprocess.run(["pdftotext", "-bbox", pdf, "-"], capture_output=True, text=True).stdout
    out = []
    for pi, p in enumerate(x.split("<page ")[1:], 1):
        for a, b, c, d, w in re.findall(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)<', p):
            out.append((pi, float(a), float(b), float(c), float(d), html.unescape(w)))
    return out


def col_of(w):
    return (w[0], "L" if w[1] < COL_SPLIT else "R")


def question_numbers(ws):
    """題號（「12.」）且位於欄的左緣，依閱讀順序排列。"""
    nums = [w for w in ws if re.fullmatch(r"\d+\.", w[5]) and (abs(w[1] - 40) < 6 or abs(w[1] - 431) < 6)]
    return sorted(nums, key=lambda w: (w[0], w[1] > COL_SPLIT, w[2]))


def measure(pdf, n_mc):
    """自然排版的試卷 → 每題高度（含題間距）、第一欄起點、計算題區塊高度。"""
    ws = [w for w in words(pdf) if w[4] < BOTTOM - 8]
    last = {}
    for w in ws:
        last[col_of(w)] = max(last.get(col_of(w), 0), w[4])
    nums = question_numbers(ws)
    h = []
    for i, w in enumerate(nums[:n_mc]):
        nxt = nums[i + 1] if i + 1 < len(nums) else None
        h.append((nxt[2] if nxt and col_of(nxt) == col_of(w) else last[col_of(w)] + 4) - w[2])
    head1 = next(w for w in ws if w[5].startswith("一、"))
    calc_head = next((w for w in ws if w[5].startswith("二、")), None)
    end = next(w for w in ws if "試題結束" in w[5])
    start = calc_head[2] if calc_head else end[2]
    return {"h": h, "top1": head1[2], "first_q": nums[0][2], "calc": end[4] - start + 6}
