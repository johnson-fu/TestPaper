# TestPaper

考試卷相關的 Claude skills。每個資料夾是一個 skill（`SKILL.md` + `scripts/` + `examples/`）。

| Skill | 用途 |
|---|---|
| [exam-answer-sheet](exam-answer-sheet/SKILL.md) | 依試題產生單面表格式學生答案卷與可疊合比對的解答卷；題目卷縮頁轉 PDF；作答時間估計 |
| [exam-ab-papers](exam-ab-papers/SKILL.md) | 依題庫、A/B 座位表與名冊產生個人化 A/B 卷：A3 雙欄題目卷＋B4 答案卷（印班級學號姓名、依座位排序）、可疊合的批改卷、教師參考答案 |

## 快速開始

### exam-answer-sheet

```bash
pip install python-docx pillow numpy      # 另需 LibreOffice 與 poppler-utils
cd exam-answer-sheet
python3 scripts/make_answer_sheets.py examples/115-cs-intro-midterm.json out --pdf
python3 scripts/verify_sheets.py examples/115-cs-intro-midterm.json \
  "out/115學年度計算機概論期中考答案卷.pdf" "out/115學年度計算機概論期中考答案卷_解答.pdf"
```

### exam-ab-papers

```bash
npm install -g docx                         # Node.js 套件
pip install python-docx openpyxl pillow     # 另需 LibreOffice 與 poppler-utils
cp -r exam-ab-papers/examples/demo /tmp/demo && cd /tmp/demo   # 在獨立工作資料夾執行
S=<repo>/exam-ab-papers/scripts
python3 $S/prep_roster.py seats.docx roster.xlsx --order column
python3 $S/make_order.py
python3 $S/plan_layout.py
python3 $S/build_all.py            # 產出在 out/，最後一行應為「全部通過」
```

> 本 repo 為公開：真實考題、答案、學生名冊不要 commit。
