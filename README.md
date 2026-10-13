# TestPaper

考試卷相關的 Claude skills。每個資料夾是一個 skill（`SKILL.md` + `scripts/` + `examples/`）。

| Skill | 用途 |
|---|---|
| [exam-answer-sheet](exam-answer-sheet/SKILL.md) | 依試題產生單面表格式學生答案卷與可疊合比對的解答卷；題目卷縮頁轉 PDF；作答時間估計 |

## 快速開始

```bash
pip install python-docx pillow numpy      # 另需 LibreOffice 與 poppler-utils
cd exam-answer-sheet
python3 scripts/make_answer_sheets.py examples/115-cs-intro-midterm.json out --pdf
python3 scripts/verify_sheets.py examples/115-cs-intro-midterm.json \
  "out/115學年度計算機概論期中考答案卷.pdf" "out/115學年度計算機概論期中考答案卷_解答.pdf"
```
