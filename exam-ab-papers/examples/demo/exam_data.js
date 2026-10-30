// 範例題庫（示範用的原創題目，非實際考題）
// mc：選擇題  { stem, opts[4], ans, exp, fig?, code?, layout? }
//   stem/opts 中 *x* 為斜體；數學符號直接用 Unicode（∧ ∨ ¬ → ↔ ∀ ∃ ≤ ≥ ⋅ − ₁ ² ⋯），程式會自動改用數學字型
//   fig：對應 exam.config.json 的 figures 鍵；code：等寬字程式碼行；layout：強制選項 4／2／1 個一列
// calc：計算或問答題 { stem, defs?, sol[] }（sol 只出現在參考答案）
const mc = [
  { stem: "How many bit strings of length 8 contain exactly three 1s?",
    opts: ["56", "28", "70", "8"], ans: "A", exp: "C(8, 3) = 56。" },
  { stem: "In how many ways can 5 different books be arranged on a shelf?",
    opts: ["25", "60", "120", "720"], ans: "C", exp: "5! = 120。" },
  { stem: "How many nonnegative integer solutions does x₁ + x₂ + x₃ = 5 have?",
    opts: ["10", "15", "21", "35"], ans: "C", exp: "C(5 + 3 − 1, 5) = C(7, 5) = 21。" },
  { stem: "How many edges does the graph shown in the figure have?",
    opts: ["4", "5", "6", "8"], ans: "B", fig: "graph", exp: "四邊加一條對角線，共 5 條邊。" },
  { stem: "What is the coefficient of x³ in the expansion of (1 + x)⁵?",
    opts: ["5", "10", "15", "20"], ans: "B", exp: "C(5, 3) = 10。" },
  { stem: "Which of the following is logically equivalent to ¬(p ∧ q)?",
    opts: ["¬p ∧ ¬q", "¬p ∨ ¬q", "p ∨ q", "p ∧ ¬q"], ans: "B", exp: "DeMorgan's Law。" },
  { stem: "下列何者為 tautology？",
    opts: ["p ∨ ¬p", "p ∧ ¬p", "p → ¬p", "¬p → p"], ans: "A", exp: "p ∨ ¬p 恆為真。" },
  { stem: "Which of the following is the negation of ∀x p(x)?",
    opts: ["∀x ¬p(x)", "∃x p(x)", "∃x ¬p(x)", "¬∃x p(x)"], ans: "C", exp: "¬∀x p(x) ⇔ ∃x ¬p(x)。" },
  { stem: "令 p：今天下雨，q：我帶傘。「若今天下雨，則我帶傘」可表示為下列何者？",
    opts: ["p ∧ q", "p → q", "q → p", "p ∨ q"], ans: "B", exp: "「若 p 則 q」為 p → q。" },
  { stem: "承上題，該敘述的 contrapositive 為下列何者？",
    opts: ["q → p", "¬p → ¬q", "¬q → ¬p", "¬p ∨ q"], ans: "C", exp: "p → q 的 contrapositive 為 ¬q → ¬p。" },
];

const calc = [
  { stem: "Prove that for every integer *n*, *n*³ − *n* is divisible by 3.",
    sol: ["n³ − n = (n − 1)n(n + 1) 為三個連續整數之積，其中必有一個是 3 的倍數，故可被 3 整除。"] },
  { stem: "How many 4-letter strings can be formed from the letters A, B, C, D, E, F if repetition is not allowed?",
    sol: ["P(6, 4) = 6⋅5⋅4⋅3 = 360。"] },
];

module.exports = { mc, calc };
