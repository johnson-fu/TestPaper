---
name: exam-ab-papers
description: 依題庫（.docx）、標示 A/B 卷的座位表（.docx）與學生名冊（.xlsx）產生個人化 A/B 卷——A3 雙欄正反兩面題目卷與 JIS B4 答案卷都印上班級、學號、姓名、卷別，A/B 卷選擇題題序不同，全部依座位由上而下、由左至右排成一份；另產生可疊在答案卷上的 B4 批改卷（正解印在題號列）與教師用參考答案。當使用者要出 A/B 卷、依座位表發考卷、在考卷上印學生資料、或要批改套版時使用。
---

# A/B 卷個人化考卷產生器

一次產出整套期中／期末考用卷，**每位學生一份**，依座位順序排好直接發放：

| 檔案（`out/`） | 紙張 | 內容 |
|---|---|---|
| `<考試>_題目卷_AB混合.docx` | A3 直式雙欄，每人 2 頁（正反兩面） | 表頭印班級／學號／姓名／卷別；選擇題依卷別排序；計算題；頁尾印卷別、學號姓名、頁碼 |
| `<考試>_答案卷_AB混合.docx` | JIS B4（257×364 mm），每人 1 頁 | 名冊欄、得分欄、選擇題答案格（每列 18 題）、計算題作答框 |
| `<考試>_批改卷_A卷.docx`、`_B卷.docx` | JIS B4 | 與答案卷**幾何完全相同**；題號列印「小字題號＋紅色正解」、答案列留白，疊在答案卷上逐題對照 |
| `<考試>_參考答案.docx` | A4 | A、B 卷答案速查、A↔B 題號對照與解析、計算題詳解 |

加 `--split` 另外產生 A、B 分開的題目卷與答案卷。

## 需要的工具

Node.js＋`docx`（npm）、Python 3＋`python-docx`、`openpyxl`、`Pillow`、LibreOffice（`soffice`）、poppler（`pdfinfo`／`pdftotext`／`pdftoppm`）。

```
scripts/prep_roster.py    座位表 .docx ＋ 名冊 .xlsx → students.json（含排序）
scripts/make_order.py     A 卷＝原順序；B 卷在指定範圍內打亂 → orders.json
scripts/plan_layout.py    量測每題高度、DP 選換欄點、必要時自動降字級、調整答案卷作答框 → layout.json
scripts/gen_paper.js      paper｜sheet｜key 產生 .docx（被 build_all.py 呼叫）
scripts/build_key.js      教師用參考答案
scripts/build_all.py      產生全部檔案、轉 PDF、自動驗證
scripts/verify_outputs.py 驗證頁數、每人資料與卷別、題序、批改卷正解與框線對齊
scripts/pdfmeasure.py     轉檔與量測共用函式
examples/demo/            可直接執行的範例（原創題目、虛構名單）
```

## 工作流程

所有指令都在**考試工作資料夾**內執行（不要放在這個 skill 資料夾裡）。這個 repo 是公開的：**真實考題、答案、學生名冊、教科書附圖都不可以 commit。**

### 1. 題庫 → `exam_data.js`

格式見 `examples/demo/exam_data.js`：

```js
const mc = [{ stem, opts: [4 個選項], ans: "A"~"D", exp: "解析", fig?: "圖片鍵", code?: [程式行], layout?: 4|2|1 }];
const calc = [{ stem, defs?: [補充定義行], sol: [詳解段落] }];
module.exports = { mc, calc };
```

- 讀題庫：`Read` 或 `pandoc -t plain`；附圖在 docx 的 `word/media/`（`unzip` 取出成 PNG，放 `img/`），並確認每張圖對應哪一題。
- **逐題檢查答案**。題庫常有缺答案的題目：自己解，能窮舉的用 Python 驗算（排列組合、真值表、迴圈計數），在交付時告訴老師補了哪幾題。
- 斜體寫 `*n*`；數學符號直接用 Unicode（`∧ ∨ ¬ → ↔ ∀ ∃ ≤ ≥ ⋅ − ₁ ² ⋯`）。乘號用 `⋅`（U+22C5），**不要用 `·`**；撇號用 `'`，**不要用 `’`**——後兩者在 Word 會被當成中文字、顯示成全形。
- 後一題沿用前一題定義的（例如「承上題」、同一個 h(x, y)），記下題號放進 `keep_together`。
- 原文明顯不通順可以微調措辭，但要在交付時說明。

### 2. `exam.config.json`

| 欄位 | 說明 |
|---|---|
| `title` / `exam_name` | 題目卷標題（例「115年離散數學 期中考卷」）／答案卷、批改卷、檔名用（例「115年離散數學 期中考」） |
| `mc_points` / `calc_points` | 每題配分；注意事項與各大題標題會自動算總分 |
| `paper_pages` | 題目卷每人頁數：2（正反兩面，預設）或 1 |
| `pt` / `ls` | 起始字級（預設 14）與行距倍數（預設 1.2）；放不下時 `plan_layout.py` 自動降 0.5pt |
| `figures` | `{"鍵": {"file": "img/x.png", "cm_w": 10}}` 置於題幹下方；用 `cm_h` 則放在題幹右側（高瘦的圖） |
| `shuffle` | `seed`、`groups`（A 卷題號範圍，各自打亂，例如計數題、邏輯題分開）、`keep_together`、`max_same_answer_run` |
| `sheet` | `per_row`（答案格每列題數，預設 18）、`calc_weights`（各計算題作答框高度比例） |
| `fonts` | 預設 Times New Roman／標楷體／Cambria Math／Consolas |

### 3. 座位表與名冊 → `students.json`

```bash
python3 SKILL_DIR/scripts/prep_roster.py 座位表.docx 名冊.xlsx --order column
```

- 座位表：Word 表格，第一列為行別、第一欄為列號，每格「姓名 (A)」或「姓名（B）」，空位寫 X。
- 名冊：第一列要有「班級」「學號」「姓名」。
- `--order column`（預設）＝**由上而下、由左至右**（先排完第 1 行的每一列，再換第 2 行）；`row`＝由左至右、由上而下。老師說「由上而下、由左至右」就是 `column`。
- 程式會檢查：座位表的人都在名冊裡、沒有同名；名冊有但座位表沒有的人只提醒、不產生考卷。

### 4. B 卷題序 → `orders.json`

```bash
python3 SKILL_DIR/scripts/make_order.py
```

只調換選擇題順序，選項不變、計算題不變。保證：每題位置都與 A 卷不同、A 卷相鄰的兩題在 B 卷不再相鄰、`keep_together` 保持相連、同一答案連續不超過 `max_same_answer_run`、附圖題彼此錯開。

### 5. 版面 → `layout.json`

```bash
python3 SKILL_DIR/scripts/plan_layout.py
```

輸出各卷換欄題號、各欄填滿比例、題間距、答案卷作答框高度。填滿上限 0.93（保留給 Word 與 LibreOffice 的排版差異）；放不下會自動降字級，並以最終版面實際轉檔確認頁數。

### 6. 產生並驗證

```bash
python3 SKILL_DIR/scripts/build_all.py [--split]
```

最後一行要是「全部通過」。接著**一定要看圖**：`pdftoppm -jpeg -r 60 -f 1 -l 2 out/…題目卷_AB混合.pdf v`，檢查附圖、程式碼、長選項換行、批改卷的紅字。

### 7. 交付時告訴老師

- 檔案清單與份數／頁數、學生順序（列出姓名＋卷別）、B 卷答案、補上或修改的題目。
- 列印：題目卷 A3 **雙面、長邊翻頁**；答案卷與批改卷 B4，**一律選「實際大小／100%」**，否則批改卷疊上去會錯位。可以把批改卷印在透明片上直接疊，或印在紙上沿答案列剪開，對齊答案卷的題號列。
- 預覽是在沒有標楷體的環境轉的；若老師用 Word 開啟後頁數不對，回報後降字級或調整作答框高度重產。

## 實作細節與陷阱（修改程式前必讀）

**版面**
- 本文段落用固定行高（`EXACT`），所以頁數可預測。但**圖片段落必須改回 `AUTO`**，否則圖只剩一行高；**標題段落**也要 `AUTO`，否則 22pt 字繼承本文行高會被切頭。
- 不要讓 Word 自然分欄：整題 `keepNext／keepLines` 加上大附圖，常在欄底留下一大塊空白並把後面全部推出頁面。改由 `plan_layout.py` 量測後插入 `ColumnBreak` 段落（固定 1pt 行高）。
- 題目卷每人兩個 section：單欄表頭＋連續分節雙欄本文。第二人起用 `ODD_PAGE`，萬一有人多出一頁會插空白頁，而不是讓後面所有人的正反面錯位；頁碼每人從 1 起算，「共 N 頁」直接寫數字（`NUMPAGES` 會算整份文件）。
- 選項依估計寬度自動 4／2／1 個一列；附圖 `cm_h` 用無框表格放右側。
- `C(n, k)` 逗號後自動改成不斷行空白，避免拆成兩行。

**批改卷**
- 與答案卷共用同一函式，只改「格子裡的文字」：標題、大題說明、名冊欄文字可換，任何列高、欄寬、段落行高都不能變；大題說明必須維持單行。
- 正解放在題號列（老師說的「答案放在標題上」），答案列留白，疊上去不會蓋住學生的字。
- `verify_outputs.py` 把兩份轉成 150 dpi 圖逐列比對水平框線；直線在答案列比對。

**驗證**
- `pdftotext` 抽中文常漏字（例如「依」「為」、「q 是 p」整段），題序比對只用英數字、用「包含」比對，否則會整批誤報。
- 每人資料檢查表頭（班級／學號／姓名／卷別）與每頁頁尾（姓名、卷別、第 k 頁）。

**環境**
- 沙盒的 LibreOffice 沒有標楷體與 Cambria Math，預覽字型不同但寬度相近；各欄保留約 7% 以上空間。
- 台灣說的 B4 是 JIS 257×364 mm（`14570 × 20636` DXA）；A3 為 `16838 × 23811` DXA。
- 與 `exam-answer-sheet` 的差別：那個 skill 產生不記名的單份答案卷（Python／python-docx）；本 skill 產生依座位個人化、A/B 卷混合的整套用卷（docx-js）。
