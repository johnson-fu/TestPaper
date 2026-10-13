#!/usr/bin/env python3
"""把題目卷 .docx 等比例縮小字級，轉成不超過指定頁數的 PDF（用 LibreOffice）。

會由大到小嘗試字級，挑出「頁數符合、沒有缺字、底線空格沒被拆開」的最大字級。
原始 .docx 不會被修改。

用法：
    python3 fit_question_pdf.py 題目.docx 輸出.pdf --pages 2 [--keep-from 二、] [--min-pt 11]

處理的 LibreOffice 問題：
  1. 雙欄 + 連續分節 + 「段落不分頁(keepLines)」時，欄位平衡會把多行段落截成只剩第一行。
     → 在 --keep-from 標記（預設「二、」）之前的段落移除 keepLines；之後（問答題）保留。
  2. 中文段落裡的底線 ____ 被當成中文字，任兩個底線之間都能斷行，空格被拆成兩段。
     → 在每段底線前插入 1pt 白色英文字母 x，讓底線改以英文單字處理，不再被拆開。
        （印出看不到；但從 PDF 複製文字時會多出 x）
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter

ANCHOR = ('<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
          '<w:color w:val="FFFFFF"/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:t>x</w:t></w:r>')


def body_size(xml):
    """最常見的字級（半點）視為內文字級"""
    c = Counter(int(v) for v in re.findall(r'<w:sz w:val="(\d+)"/>', xml) if int(v) > 4)
    return c.most_common(1)[0][0]


def split_run(m):
    run = m.group(0)
    if "__" not in run:
        return run
    rpr = re.search(r"<w:rPr>(.*?)</w:rPr>", run, re.S)
    rpr = rpr.group(1) if rpr else ""
    t = re.search(r"<w:t[^>]*>([^<]*)</w:t>", run)
    if not t:
        return run
    res = ""
    for p in re.split(r"(_{2,})", t.group(1)):
        if not p:
            continue
        if p.startswith("__"):
            res += ANCHOR
        res += f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{p}</w:t></w:r>'
    return res


def transform(src, out, factor, keep_from):
    def fix(xml, name):
        xml = re.sub(r'<w:(sz|szCs) w:val="(\d+)"/>',
                     lambda m: m.group(0) if int(m.group(2)) <= 4
                     else f'<w:{m.group(1)} w:val="{round(int(m.group(2)) * factor)}"/>', xml)
        if name == "word/document.xml":
            i = xml.find(keep_from) if keep_from else -1
            if i < 0:
                i = len(xml)
            xml = xml[:i].replace("<w:keepLines/>", "") + xml[i:]
            xml = re.sub(r"<w:r(?: [^>]*)?>(?:(?!</w:r>).)*?</w:r>", split_run, xml, flags=re.S)
        return xml

    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename in ("word/document.xml", "word/numbering.xml"):
            data = fix(data.decode("utf8"), it.filename).encode("utf8")
        zout.writestr(it, data)
    zout.close()


def pdfinfo(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    n = int(re.search(r"Pages:\s+(\d+)", out).group(1))
    w, h = map(float, re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", out).groups())
    return n, w, h


def pdftext(pdf, page=None, x=None, w=None, h=None):
    cmd = ["pdftotext"]
    if page:
        cmd += ["-f", str(page), "-l", str(page)]
    if x is not None:
        cmd += ["-x", str(int(x)), "-y", "0", "-W", str(int(w) + 1), "-H", str(int(h) + 1)]
    return subprocess.run(cmd + [pdf, "-"], capture_output=True, text=True).stdout


def docx_text(path):
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf8")
    paras = re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
    return ["".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)) for p in paras]


def check(src_docx, pdf, ncols):
    paras = docx_text(src_docx)
    norm = lambda s: re.sub(r"[\s_⁠x]", "", s)
    src_chars = Counter(norm("".join(paras)))
    got_chars = Counter(norm(pdftext(pdf)))
    missing = sum((src_chars - got_chars).values())
    n, w, h = pdfinfo(pdf)
    src_blanks = Counter(len(b) for p in paras for b in re.findall(r"_{2,}", p))
    got_blanks = Counter()
    cw = w / ncols
    for pg in range(1, n + 1):
        for c in range(ncols):
            got_blanks += Counter(len(b) for b in re.findall(r"_+", pdftext(pdf, pg, c * cw, cw, h)))
    split = sum((src_blanks - got_blanks).values())
    return n, missing, split


def to_pdf(docx, outdir):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
    return os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("pdf")
    ap.add_argument("--pages", type=int, default=2)
    ap.add_argument("--keep-from", default="二、", help="此文字之後的段落保留 keepLines")
    ap.add_argument("--min-pt", type=float, default=11)
    ap.add_argument("--max-pt", type=float, default=None, help="最大嘗試字級（預設為原字級）")
    ap.add_argument("--columns", type=int, default=2, help="版面欄數（用於檢查底線是否被拆開）")
    a = ap.parse_args()

    xml = zipfile.ZipFile(a.docx).read("word/document.xml").decode("utf8")
    base = body_size(xml) / 2
    print(f"原內文字級 {base:g}pt")
    with tempfile.TemporaryDirectory() as tmp:
        pt = min(base, a.max_pt) if a.max_pt else base
        while pt >= a.min_pt:
            d = os.path.join(tmp, f"t{pt:g}.docx")
            transform(a.docx, d, pt / base, a.keep_from)
            pdf = to_pdf(d, tmp)
            n, missing, split = check(a.docx, pdf, a.columns)
            print(f"  {pt:g}pt → {n} 頁，缺字 {missing}，被拆開的空格 {split}")
            if n <= a.pages and missing == 0 and split == 0:
                shutil.copy(pdf, a.pdf)
                print(f"✔ 採用 {pt:g}pt，已輸出 {a.pdf}")
                return
            pt -= 0.5
    sys.exit("✘ 找不到符合條件的字級，請調低 --min-pt 或增加頁數")


if __name__ == "__main__":
    main()
