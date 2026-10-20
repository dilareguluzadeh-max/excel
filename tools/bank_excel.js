/* Excel & Google Sheets sual bankı: 42 müstəqil praktiki tapşırıq (8 modul).
 * Mənbə: Excel_GoogleSheets_Exam_Questions_Bank.docx
 *
 * Hər tapşırıq:
 *   n, mod, diff        — bank nömrəsi, modul (1–8), çətinlik (1–5)
 *   title, skill        — başlıq və yoxlanan bacarıq (sənəddən)
 *   setup               — tələbəyə hazır verilən məlumat ({cells, fmt})
 *   steps               — tapşırığın şərti (addımlar)
 *   checks              — avtomatik yoxlama meyarları (grader.js), hər birinin çəkisi w
 *   sol                 — etalon həll (verify.js 100% olduğunu yoxlayır)
 */
"use strict";
const SE = require("../sheetengine.js");

const MODULES = {
    1: "Basic Formulas, Absolute & Mixed Cell Referencing",
    2: "Data Cleaning, Formatting & Text Functions",
    3: "Logical & Lookup Functions",
    4: "Conditional Aggregations",
    5: "Descriptive Statistics & Exploratory Data Analysis",
    6: "Data Validation & Conditional Formatting",
    7: "PivotTables, Custom Grouping & Calculated Fields",
    8: "Data Visualization, Trendlines & Scenario Analysis",
};

// ---- köməkçilər --------------------------------------------------------
function put(topLeft, matrix) {
    const p = SE.parseAddr(topLeft), out = {};
    matrix.forEach((row, i) => row.forEach((v, j) => {
        if (v === null || v === undefined || v === "") return;
        out[SE.addr(p.c + j, p.r + i)] = String(v);
    }));
    return out;
}
function down(col, r1, r2, f) {            // düsturu r1-dən r2-yə qədər aşağı doldurur
    const out = {};
    for (let r = r1; r <= r2; r++) out[col + r] = SE.shiftFormula(f, r - r1, 0);
    return out;
}
function block(range, f) {                  // iki ölçülü doldurma (f — yuxarı-sol xananın düsturu)
    const g = SE.parseRange(range), out = {};
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out[SE.addr(c, r)] = SE.shiftFormula(f, r - g.r1, c - g.c1);
    return out;
}
const pct = (d) => ({ t: "pct", d });
const fmtRange = (range, f) => Object.fromEntries(SE.cellsIn(range).map((a) => [a, f]));
const TASKS = [];
const task = (o) => TASKS.push(o);

// ======================== MODUL I ========================================
task({
    n: 1, mod: 1, diff: 1,
    title: "Tax Amount Calculation Using Absolute Cell References",
    skill: "Distinguishing absolute ($) and relative cell references; copying formulas with AutoFill / fill handle.",
    setup: {
        cells: { ...put("A1", [["Product", "Quantity", "Unit Price ($)", "Sales ($)", "Tax (10%)"], ["Laptop", 10, 1200], ["Smartphone", 25, 800], ["Tablet", 15, 500], ["Monitor", 12, 300]]), F1: "Tax Rate", G1: "0.1" },
        fmt: { G1: pct(0) },
    },
    steps: [
        "The table is entered in range `A1:E5` and the Tax Rate constant is in cell `G1` (10%).",
        "In cell `D2`, calculate Sales Amount using a relative reference (Quantity × Unit Price) and drag the fill handle down to `D5`.",
        "In cell `E2`, calculate Tax Amount using an absolute reference to `G1`, so that the reference does not move when the formula is copied down to `E5`.",
    ],
    checks: [
        { t: "vs", cells: { D2: 12000, D3: 20000, D4: 7500, D5: 3600 }, w: 2, l: "D2:D5 — Sales Amount (formulas)" },
        { t: "vs", cells: { E2: 1200, E3: 2000, E4: 750, E5: 360 }, w: 2, l: "E2:E5 — Tax Amount (formulas)" },
        { t: "re", cells: ["E2"], re: "\\$G\\$1", l: "E2 — absolute reference to the tax rate ($G$1)" },
        { t: "re", cells: ["E5"], re: "\\$G\\$1", l: "E5 — the reference to the tax rate did not move when copied down" },
    ],
    sol: { cells: { ...down("D", 2, 5, "=B2*C2"), ...down("E", 2, 5, "=D2*$G$1") } },
});

task({
    n: 2, mod: 1, diff: 2,
    title: "Dynamic Growth Projection with Mixed Cell References",
    skill: "Using mixed cell references ($A2 and B$1) to fill a two-dimensional table with one formula.",
    setup: { cells: put("A1", [["Base Revenue ($)", 0.05, 0.10, 0.15], [100000], [150000], [80000]]) },
    steps: [
        "The base revenues are in `A2:A4` (rows: Region North, Region South, Region East) and the growth rates are in the header row `B1:D1`.",
        "Write a single formula in cell `B2` that fixes the column of the base revenue (`$A2`) and the row of the growth rate (`B$1`): Projected Revenue = Base × (1 + Growth Rate).",
        "Copy cell `B2` across the whole range `B2:D4` in a single action (fill handle, or copy and paste).",
    ],
    checks: [
        { t: "vs", cells: { B2: 105000, C2: 110000, D2: 115000, B3: 157500, C3: 165000, D3: 172500, B4: 84000, C4: 88000, D4: 92000 }, w: 3, l: "B2:D4 — all nine projected values" },
        { t: "re", cells: ["B2"], re: "(\\$A2.*B\\$1|B\\$1.*\\$A2)", l: "B2 — mixed references ($A2 and B$1)" },
        { t: "re", cells: ["D4"], re: "(\\$A4.*D\\$1|D\\$1.*\\$A4)", l: "D4 — the mixed references stayed anchored after copying" },
    ],
    sol: { cells: block("B2:D4", "=$A2*(1+B$1)") },
});

task({
    n: 3, mod: 1, diff: 2,
    title: "Profit Margin Percentage & Overall Weighted Margin",
    skill: "Formula construction, percentage formatting and order of operations.",
    setup: { cells: put("A1", [["Item", "Revenue ($)", "Cost ($)", "Profit ($)", "Profit Margin (%)"], ["Alpha", 50000, 32000], ["Beta", 80000, 60000], ["Gamma", 120000, 78000], ["Overall"]]) },
    steps: [
        "The data is entered in `A1:C4` and the label \"Overall\" is in `A5`. Calculate Profit in column `D` (Revenue − Cost).",
        "Calculate Profit Margin in column `E` (Profit / Revenue) and format `E2:E4` as Percentage with 2 decimal places.",
        "In cell `E5`, calculate the overall weighted profit margin for the entire business as total profit divided by total revenue.",
    ],
    checks: [
        { t: "vs", cells: { D2: 18000, D3: 20000, D4: 42000 }, w: 1, l: "D2:D4 — Profit (formulas)" },
        { t: "vs", cells: { E2: 0.36, E3: 0.25, E4: 0.35 }, w: 2, l: "E2:E4 — Profit Margin (formulas)" },
        { t: "fmt", r: "E2:E4", type: "pct", d: 2, w: 1, l: "E2:E4 — Percentage format with 2 decimal places" },
        { t: "v", c: "E5", v: 0.32, w: 1, l: "E5 — overall margin value" },
        { t: "dyn", sets: { B2: "100000" }, cells: { E5: 130000 / 300000 }, tol: 1e-9, w: 2, l: "E5 — weighted margin (total profit ÷ total revenue), not the average of margins" },
    ],
    sol: { cells: { ...down("D", 2, 4, "=B2-C2"), ...down("E", 2, 4, "=D2/B2"), E5: "=SUM(D2:D4)/SUM(B2:B4)" }, fmt: fmtRange("E2:E4", pct(2)) },
});

task({
    n: 4, mod: 1, diff: 2,
    title: "Compound Interest Future Value Modeling",
    skill: "Constructing exponent formulas (^) for financial projections and rounding results.",
    setup: { cells: put("A1", [["Principal ($)", "Annual Rate", "Years (t)", "Compounding Periods per Year (n)", "Future Value ($)"], [10000, 0.06, 5, 12], [25000, 0.045, 10, 4], [50000, 0.08, 3, 365]]) },
    steps: [
        "The investment parameters are entered in `A2:D4` (headers in row 1).",
        "In cell `E2`, enter the compound interest formula FV = P × (1 + r/n)^(n×t) and round it to 2 decimal places.",
        "Copy the formula down to row 4.",
    ],
    checks: [
        { t: "vs", cells: { E2: 13488.5, E3: 39109.42, E4: 63560.79 }, tol: 0.0051, w: 3, l: "E2:E4 — Future Value (formulas)" },
        { t: "fn", cells: ["E2"], all: ["ROUND"], w: 1, l: "E2 — result is rounded with ROUND" },
        { t: "vs", cells: { E2: 13488.5, E3: 39109.42, E4: 63560.79 }, tol: 1e-9, w: 1, l: "E2:E4 — rounded to exactly 2 decimals" },
    ],
    sol: { cells: down("E", 2, 4, "=ROUND(A2*(1+B2/D2)^(C2*D2),2)") },
});

task({
    n: 5, mod: 1, diff: 1,
    title: "Discount Calculation with Parentheses and Order of Operations",
    skill: "Building arithmetic formulas with parentheses and the correct order of operations.",
    setup: { cells: put("A1", [["Customer Type", "Gross Order ($)", "Discount Rate", "Net Payable ($)"], ["Retail", 1200, 0.05], ["Wholesale", 5000, 0.15], ["VIP", 8500, 0.20]]) },
    steps: [
        "The data is entered in range `A1:C4`.",
        "In cell `D2`, construct a formula calculating Net Payable Amount: Net = Gross × (1 − Discount Rate).",
        "Fill the formula down to `D4` and calculate the total of column `D` in cell `D5`.",
    ],
    checks: [
        { t: "vs", cells: { D2: 1140, D3: 4250, D4: 6800 }, w: 3, l: "D2:D4 — Net Payable (formulas)" },
        { t: "v", c: "D5", v: 12190, w: 1, l: "D5 — total net revenue" },
        { t: "fn", cells: ["D5"], all: ["SUM"], w: 1, l: "D5 — uses SUM" },
        { t: "dyn", sets: { C2: "0.5" }, cells: { D2: 600, D5: 11650 }, w: 1, l: "D2 / D5 — recalculate when the discount rate changes" },
    ],
    sol: { cells: { ...down("D", 2, 4, "=B2*(1-C2)"), D5: "=SUM(D2:D4)" } },
});

// ======================== MODUL II =======================================
task({
    n: 6, mod: 2, diff: 1,
    title: "Text Cleaning with TRIM and PROPER Functions",
    skill: "Removing irregular spacing and standardizing capitalization in raw imported text.",
    setup: { cells: { A1: "Raw Customer Name (Messy)", B1: "Cleaned Name", A2: "  jOhn    dOe  ", A3: "aLiCe   sMiTh   ", A4: " roBeRt jOhNsOn " } },
    steps: [
        "The raw text is entered in `A2:A4` with irregular leading, trailing and internal spaces.",
        "In cell `B2`, combine `TRIM` (removes excess spaces) and `PROPER` (capitalizes the first letter of each word).",
        "Copy the formula down to `B4`.",
    ],
    checks: [
        { t: "vs", cells: { B2: "John Doe", B3: "Alice Smith", B4: "Robert Johnson" }, w: 3, l: "B2:B4 — cleaned names (formulas)" },
        { t: "fn", cells: ["B2"], all: ["TRIM", "PROPER"], w: 1, l: "B2 — uses TRIM and PROPER" },
    ],
    sol: { cells: down("B", 2, 4, "=PROPER(TRIM(A2))") },
});

task({
    n: 7, mod: 2, diff: 2,
    title: "Splitting Full Names with LEFT, MID and SEARCH",
    skill: "Splitting a combined text column into First Name and Last Name using formulas that work in every spreadsheet tool.",
    setup: { cells: put("A1", [["Full Name", "First Name", "Last Name"], ["David Miller"], ["Elena Rostova"], ["Marcus Vance"]]) },
    steps: [
        "The full names are entered in range `A2:A4`.",
        "In cell `B2`, extract the First Name using `LEFT` together with `SEARCH` (position of the space).",
        "In cell `C2`, extract the Last Name using `MID` together with `SEARCH` and `LEN`.",
        "Copy both formulas down to row 4. (`TEXTSPLIT` / `SPLIT` are not available on every tool, so use `LEFT` / `MID` / `SEARCH`.)",
    ],
    checks: [
        { t: "vs", cells: { B2: "David", B3: "Elena", B4: "Marcus" }, w: 2, l: "B2:B4 — First Names (formulas)" },
        { t: "vs", cells: { C2: "Miller", C3: "Rostova", C4: "Vance" }, w: 2, l: "C2:C4 — Last Names (formulas)" },
        { t: "fn", cells: ["B2"], all: ["LEFT"], any: ["SEARCH", "FIND"], w: 1, l: "B2 — uses LEFT with SEARCH" },
        { t: "fn", cells: ["C2"], all: ["MID"], any: ["SEARCH", "FIND"], w: 1, l: "C2 — uses MID with SEARCH" },
    ],
    sol: { cells: { ...down("B", 2, 4, '=LEFT(A2,SEARCH(" ",A2)-1)'), ...down("C", 2, 4, '=MID(A2,SEARCH(" ",A2)+1,LEN(A2))') } },
});

task({
    n: 8, mod: 2, diff: 3,
    title: "Converting Text-Formatted Numbers with VALUE and SUBSTITUTE",
    skill: "Diagnosing silent errors (numbers stored as text) and converting them to numeric data.",
    setup: { cells: put("A1", [["Code", "Sales String (Text)", "Numeric Value"], [101, "'1,250.50"], [102, "'3,400.00"], [103, "'850.75"]]) },
    steps: [
        "The codes are in `A2:A4`. The sales values in `B2:B4` are stored as TEXT (the apostrophe only forces text entry and is not part of the value).",
        "In cell `C2`, remove the thousands separator with `SUBSTITUTE` and convert the result into a true number with `VALUE`. Copy down to `C4`.",
        "In cell `C5`, calculate `=SUM(C2:C4)`. In cell `B5`, calculate `=SUM(B2:B4)` and compare the two results.",
    ],
    checks: [
        { t: "vs", cells: { C2: 1250.5, C3: 3400, C4: 850.75 }, w: 3, l: "C2:C4 — numeric values (formulas)" },
        { t: "fn", cells: ["C2"], all: ["VALUE", "SUBSTITUTE"], w: 1, l: "C2 — uses SUBSTITUTE and VALUE" },
        { t: "v", c: "C5", v: 5501.25, w: 1, l: "C5 — sum of the converted values" },
        { t: "v", c: "B5", v: 0, w: 1, l: "B5 — SUM over the text range" },
    ],
    sol: { cells: { ...down("C", 2, 4, '=VALUE(SUBSTITUTE(B2,",",""))'), C5: "=SUM(C2:C4)", B5: "=SUM(B2:B4)" } },
});

task({
    n: 9, mod: 2, diff: 3,
    title: "Extracting Substrings with MID and FIND for Product Codes",
    skill: "Parsing structured SKU identifiers to extract category codes.",
    setup: { cells: put("A1", [["SKU Code", "Category Identifier"], ["PROD-CAT102-2024"], ["PROD-ELEC05-2024"], ["PROD-FURN88-2024"]]) },
    steps: [
        "The SKU codes are entered in `A2:A4`. The category code is the text between the first and the second hyphen.",
        "In cell `B2`, use `MID` together with `FIND` to locate the hyphens and isolate the category substring.",
        "Copy the formula down to `B4`.",
    ],
    checks: [
        { t: "vs", cells: { B2: "CAT102", B3: "ELEC05", B4: "FURN88" }, w: 3, l: "B2:B4 — category identifiers (formulas)" },
        { t: "fn", cells: ["B2"], all: ["MID"], w: 1, l: "B2 — uses MID" },
    ],
    sol: { cells: down("B", 2, 4, '=MID(A2,FIND("-",A2)+1,FIND("-",A2,FIND("-",A2)+1)-FIND("-",A2)-1)') },
});

task({
    n: 10, mod: 2, diff: 2,
    title: "Handling Division-by-Zero Errors with IFERROR",
    skill: "Auditing #DIV/0! errors and handling them cleanly.",
    setup: { cells: put("A1", [["Region", "Total Sales ($)", "Units Sold", "Avg Price per Unit ($)"], ["North", 50000, 250], ["South", 30000, 0], ["East", 45000, 180]]) },
    steps: [
        "The dataset is entered in range `A1:C4` (Units Sold for 'South' is 0).",
        "In cell `D2`, calculate Avg Price (Sales / Units) and copy down. Notice the `#DIV/0!` error in `D3`.",
        "Wrap the calculation in `IFERROR` so that it displays `0` (or the text 'No Sales') instead of the error.",
        "In cell `D5`, calculate `=SUM(D2:D4)` to confirm the column now totals without error.",
    ],
    checks: [
        { t: "v", c: "D2", v: 200, w: 1, l: "D2 — average price (North)" },
        { t: "v", c: "D3", v: [0, "No Sales"], ci: true, w: 2, l: "D3 — the error is handled (0 or 'No Sales')" },
        { t: "v", c: "D4", v: 250, w: 1, l: "D4 — average price (East)" },
        { t: "fn", cells: ["D2", "D3", "D4"], all: ["IFERROR"], w: 1, l: "D2:D4 — use IFERROR" },
        { t: "v", c: "D5", v: 450, w: 1, l: "D5 — the column totals without error" },
    ],
    sol: { cells: { ...down("D", 2, 4, "=IFERROR(B2/C2,0)"), D5: "=SUM(D2:D4)" } },
});

// ======================== MODUL III ======================================
task({
    n: 11, mod: 3, diff: 2,
    title: "Multi-Tiered Performance Grading with Nested IF",
    skill: "Building hierarchical conditional logic.",
    setup: { cells: put("A1", [["Student Name", "Exam Score (0-100)", "Grade Category"], ["Alexander", 92], ["Beatrix", 78], ["Charles", 64], ["Diana", 45]]) },
    steps: [
        "The data is entered in range `A1:B5`.",
        "In cell `C2`, write a Nested IF assigning grades: Score ≥ 90 → 'A'; ≥ 75 → 'B'; ≥ 60 → 'C'; otherwise 'F'.",
        "Fill the formula down to `C5`.",
    ],
    checks: [
        { t: "vs", cells: { C2: "A", C3: "B", C4: "C", C5: "F" }, w: 3, l: "C2:C5 — grades (formulas)" },
        { t: "dyn", sets: { B2: "89", B3: "60", B4: "90", B5: "59" }, cells: { C2: "B", C3: "C", C4: "A", C5: "F" }, w: 2, l: "C2:C5 — boundary values (90, 75, 60) are handled correctly" },
        { t: "fn", cells: ["C2"], all: ["IF"], w: 1, l: "C2 — uses IF" },
    ],
    sol: { cells: down("C", 2, 5, '=IF(B2>=90,"A",IF(B2>=75,"B",IF(B2>=60,"C","F")))') },
});

task({
    n: 12, mod: 3, diff: 2,
    title: "Sales Bonus Calculation with Combined IF and AND Logic",
    skill: "Evaluating several simultaneous criteria with the logical AND function.",
    setup: { cells: put("A1", [["Sales Rep", "Total Sales ($)", "New Clients", "Bonus ($)"], ["John", 120000, 6], ["Sara", 95000, 8], ["Mike", 150000, 3]]) },
    steps: [
        "The data is entered in range `A1:C4` and the header 'Bonus ($)' is in `D1`.",
        "A Sales Rep gets a bonus of $5,000 ONLY IF Sales ≥ $100,000 AND New Clients ≥ 5. Otherwise the bonus is 0.",
        "In cell `D2`, write the formula combining `IF` and `AND`, then fill it down to `D4`. In `D5`, total the bonuses.",
    ],
    checks: [
        { t: "vs", cells: { D2: 5000, D3: 0, D4: 0 }, w: 3, l: "D2:D4 — bonuses (formulas)" },
        { t: "v", c: "D5", v: 5000, w: 1, l: "D5 — total bonus paid" },
        { t: "fn", cells: ["D2"], all: ["IF", "AND"], w: 1, l: "D2 — uses IF with AND" },
        { t: "dyn", sets: { B3: "100000", C3: "5", C4: "5" }, cells: { D3: 5000, D4: 5000 }, w: 2, l: "D3:D4 — both conditions work when the inputs change" },
    ],
    sol: { cells: { ...down("D", 2, 4, "=IF(AND(B2>=100000,C2>=5),5000,0)"), D5: "=SUM(D2:D4)" } },
});

task({
    n: 13, mod: 3, diff: 2,
    title: "Exact Match Lookup with VLOOKUP",
    skill: "Retrieving record attributes from a lookup table.",
    setup: { cells: { ...put("A1", [["Employee ID", "Name", "Department", "Salary ($)"], ["E101", "Alice", "HR", 55000], ["E102", "Bob", "IT", 72000], ["E103", "Carol", "Finance", 68000]]), F2: "Lookup key", G2: "E102", F3: "Salary" } },
    steps: [
        "The master table is entered in `A1:D4` and the lookup key 'E102' is in cell `G2`.",
        "In cell `G3`, write a `VLOOKUP` formula that fetches the salary (4th column) with an exact match (`FALSE`).",
        "Change `G2` to 'E103' and check that `G3` updates dynamically.",
    ],
    checks: [
        { t: "v", c: "G3", v: 72000, w: 2, l: "G3 — salary of E102" },
        { t: "fn", cells: ["G3"], all: ["VLOOKUP"], w: 1, l: "G3 — uses VLOOKUP" },
        { t: "dyn", sets: { G2: "E103" }, cells: { G3: 68000 }, w: 1, l: "G3 — updates when the key changes" },
        { t: "dyn", sets: { G2: "E999" }, cells: { G3: { err: "#N/A" } }, w: 1, l: "G3 — exact match (a missing key gives #N/A)" },
    ],
    sol: { cells: { G3: "=VLOOKUP(G2,$A$1:$D$4,4,FALSE)" } },
});

task({
    n: 14, mod: 3, diff: 3,
    title: "Left Lookup with INDEX and MATCH",
    skill: "Overcoming the VLOOKUP limitation by looking up values to the left of the key column.",
    setup: { cells: { ...put("A1", [["Department", "Employee ID", "Name"], ["Logistics", "E501", "Frank"], ["Marketing", "E502", "Grace"], ["IT", "E503", "Henry"]]), E2: "Query ID", F2: "E502", E3: "Department" } },
    steps: [
        "The data table is entered in `A1:C4`. Note that the Employee ID is in column `B`, while the Department is in column `A` (to the left of the ID). The query ID 'E502' is in `F2`.",
        "In cell `F3`, write an `INDEX` and `MATCH` formula that returns the Department for the ID in `F2`.",
    ],
    checks: [
        { t: "v", c: "F3", v: "Marketing", w: 2, l: "F3 — department of E502" },
        { t: "fn", cells: ["F3"], all: ["INDEX", "MATCH"], w: 1, l: "F3 — uses INDEX with MATCH" },
        { t: "dyn", sets: { F2: "E503" }, cells: { F3: "IT" }, w: 1, l: "F3 — updates when the ID changes" },
        { t: "dyn", sets: { F2: "E501" }, cells: { F3: "Logistics" }, w: 1, l: "F3 — works for every ID" },
    ],
    sol: { cells: { F3: "=INDEX(A2:A4,MATCH(F2,B2:B4,0))" } },
});

task({
    n: 15, mod: 3, diff: 3,
    title: "Flexible Lookup with XLOOKUP and Missing-Value Handling",
    skill: "Using XLOOKUP with its built-in 'not found' argument.",
    setup: { cells: { ...put("A1", [["Product Code", "Product Name", "Price ($)"], ["P-01", "Keyboard", 45], ["P-02", "Mouse", 25], ["P-03", "Webcam", 85]]), E2: "Query Code", F2: "P-05", E3: "Price" } },
    steps: [
        "The product list is entered in `A1:C4` and the query code 'P-05' (which does not exist) is in cell `F2`.",
        "In cell `F3`, write an `XLOOKUP` formula that retrieves the Price and uses the `if_not_found` argument to display 'Product Not Found'. (On an older Excel version an `IFERROR` + `INDEX` + `MATCH` formula is accepted.)",
        "Change `F2` to 'P-02' and check that `F3` displays 25.",
    ],
    checks: [
        { t: "v", c: "F3", v: "Product Not Found", w: 2, l: "F3 — missing code is handled with the text 'Product Not Found'" },
        { t: "dyn", sets: { F2: "P-02" }, cells: { F3: 25 }, w: 1, l: "F3 — price of P-02" },
        { t: "dyn", sets: { F2: "P-03" }, cells: { F3: 85 }, w: 1, l: "F3 — price of P-03" },
        { t: "fn", cells: ["F3"], any: ["XLOOKUP", "INDEX"], w: 1, l: "F3 — uses XLOOKUP (or INDEX / MATCH)" },
    ],
    sol: { cells: { F3: '=XLOOKUP(F2,A2:A4,C2:C4,"Product Not Found")' } },
});

task({
    n: 16, mod: 3, diff: 3,
    title: "Two-Way Matrix Lookup using INDEX with Double MATCH",
    skill: "Finding the value at the intersection of a dynamic row and column header.",
    setup: { cells: { ...put("A1", [["Region", "Q1 Sales", "Q2 Sales", "Q3 Sales"], ["North", 15000, 18000, 21000], ["South", 12000, 14000, 16000], ["East", 20000, 22000, 25000]]), F1: "Region", G1: "South", F2: "Quarter", G2: "Q2 Sales", F3: "Sales" } },
    steps: [
        "The quarterly regional sales matrix is entered in range `A1:D4`. The Region 'South' is in `G1` and the Quarter 'Q2 Sales' is in `G2`.",
        "In cell `G3`, write an `INDEX` formula with two `MATCH` functions (row match and column match) to extract the intersection.",
    ],
    checks: [
        { t: "v", c: "G3", v: 14000, w: 2, l: "G3 — South / Q2 Sales" },
        { t: "fn", cells: ["G3"], all: ["INDEX", "MATCH"], w: 1, l: "G3 — uses INDEX with MATCH" },
        { t: "re", cells: ["G3"], re: "MATCH\\(.*MATCH\\(", w: 1, l: "G3 — two MATCH functions (row and column)" },
        { t: "dyn", sets: { G1: "East", G2: "Q3 Sales" }, cells: { G3: 25000 }, w: 1, l: "G3 — East / Q3 Sales" },
        { t: "dyn", sets: { G1: "North", G2: "Q1 Sales" }, cells: { G3: 15000 }, w: 1, l: "G3 — North / Q1 Sales" },
    ],
    sol: { cells: { G3: "=INDEX(B2:D4,MATCH(G1,A2:A4,0),MATCH(G2,B1:D1,0))" } },
});

// ======================== MODUL IV =======================================
task({
    n: 17, mod: 4, diff: 1,
    title: "Conditional Totaling with SUMIF",
    skill: "Summing values based on a single category criterion.",
    setup: { cells: { ...put("A1", [["Department", "Expense Category", "Amount ($)"], ["IT", "Hardware", 4500], ["HR", "Training", 2200], ["IT", "Software", 3100], ["Sales", "Travel", 1800], ["IT", "Hardware", 1200]]), D2: "IT total" } },
    steps: [
        "The expense dataset is entered in `A1:C6`.",
        "In cell `E2`, write a `SUMIF` formula that calculates the total expenses of the 'IT' department only.",
        "Verify the result manually.",
    ],
    checks: [
        { t: "v", c: "E2", v: 8800, w: 2, l: "E2 — total IT expenses" },
        { t: "fn", cells: ["E2"], all: ["SUMIF"], w: 1, l: "E2 — uses SUMIF" },
        { t: "dyn", sets: { C2: "5000" }, cells: { E2: 9300 }, w: 1, l: "E2 — updates when the amounts change" },
    ],
    sol: { cells: { E2: '=SUMIF(A2:A6,"IT",C2:C6)' } },
});

task({
    n: 18, mod: 4, diff: 1,
    title: "Transaction Counting with COUNTIF",
    skill: "Counting records that satisfy a text or numeric condition.",
    setup: { cells: { ...put("A1", [["Order ID", "Customer", "Order Value ($)"], ["ORD-101", "Acme", 450], ["ORD-102", "Beta", 1200], ["ORD-103", "Acme", 850], ["ORD-104", "Delta", 2100], ["ORD-105", "Acme", 150]]), D2: "Orders by Acme", D3: "Orders > $500" } },
    steps: [
        "The order records are entered in `A1:C6`.",
        "In cell `E2`, count the total orders placed by customer 'Acme'.",
        "In cell `E3`, count the total orders with an Order Value strictly greater than $500.",
    ],
    checks: [
        { t: "v", c: "E2", v: 3, w: 2, l: "E2 — orders placed by Acme" },
        { t: "v", c: "E3", v: 3, w: 2, l: "E3 — orders greater than $500" },
        { t: "fn", cells: ["E2", "E3"], all: ["COUNTIF"], w: 1, l: "E2:E3 — use COUNTIF" },
        { t: "dyn", sets: { C2: "600", B3: "Acme" }, cells: { E2: 4, E3: 4 }, w: 1, l: "E2:E3 — update when the data changes" },
    ],
    sol: { cells: { E2: '=COUNTIF(B2:B6,"Acme")', E3: '=COUNTIF(C2:C6,">500")' } },
});

task({
    n: 19, mod: 4, diff: 2,
    title: "Multi-Criteria Summation using SUMIFS",
    skill: "Summing data across several simultaneous category and year conditions.",
    setup: { cells: { ...put("A1", [["Region", "Category", "Year", "Revenue ($)"], ["West", "Electronics", 2023, 50000], ["West", "Furniture", 2023, 30000], ["West", "Electronics", 2024, 65000], ["East", "Electronics", 2023, 40000]]), E2: "West / Electronics / 2023" } },
    steps: [
        "The dataset is entered in `A1:D5`.",
        "In cell `F2`, calculate the total Revenue for Region = 'West' AND Category = 'Electronics' AND Year = 2023.",
        "Remember the argument order: the sum range comes first in `SUMIFS`.",
    ],
    checks: [
        { t: "v", c: "F2", v: 50000, w: 2, l: "F2 — West / Electronics / 2023 revenue" },
        { t: "fn", cells: ["F2"], all: ["SUMIFS"], w: 1, l: "F2 — uses SUMIFS" },
        { t: "dyn", sets: { C4: "2023" }, cells: { F2: 115000 }, w: 2, l: "F2 — all three criteria are applied dynamically" },
    ],
    sol: { cells: { F2: '=SUMIFS(D2:D5,A2:A5,"West",B2:B5,"Electronics",C2:C5,2023)' } },
});

task({
    n: 20, mod: 4, diff: 2,
    title: "Selective Averaging with AVERAGEIF",
    skill: "Calculating conditional averages.",
    setup: { cells: { ...put("A1", [["Store", "Customer Rating (1-5)", "Sales ($)"], ["Store A", 4.5, 12000], ["Store B", 3.2, 8000], ["Store A", 4.8, 15000], ["Store A", 2.1, 4000]]), D2: "Avg sales, Store A", D3: "Avg sales, rating >= 4" } },
    steps: [
        "The store dataset is entered in `A1:C5`.",
        "In cell `E2`, write an `AVERAGEIF` formula to calculate the average Sales of 'Store A'.",
        "In cell `E3`, write an `AVERAGEIF` formula to calculate the average Sales of records with Rating ≥ 4.0.",
    ],
    checks: [
        { t: "v", c: "E2", v: 10333.333333, tol: 0.006, w: 2, l: "E2 — average sales of Store A" },
        { t: "v", c: "E3", v: 13500, w: 2, l: "E3 — average sales with rating ≥ 4" },
        { t: "fn", cells: ["E2", "E3"], all: ["AVERAGEIF"], w: 1, l: "E2:E3 — use AVERAGEIF" },
    ],
    sol: { cells: { E2: '=AVERAGEIF(A2:A5,"Store A",C2:C5)', E3: '=AVERAGEIF(B2:B5,">=4",C2:C5)' } },
});

task({
    n: 21, mod: 4, diff: 2,
    title: "Multi-Criteria Counting within a Numeric Range using COUNTIFS",
    skill: "Counting records bounded by a numeric range and a category.",
    setup: { cells: { ...put("A1", [["Transaction ID", "Sales Rep", "Deal Size ($)"], ["T1", "Sarah", 15000], ["T2", "Sarah", 45000], ["T3", "Sarah", 80000], ["T4", "Bob", 30000]]), D2: "Sarah, $20k–$100k" } },
    steps: [
        "The dataset is entered in `A1:C5`.",
        "In cell `E2`, count the transactions handled by 'Sarah' where the Deal Size is between $20,000 and $100,000 inclusive.",
    ],
    checks: [
        { t: "v", c: "E2", v: 2, w: 2, l: "E2 — number of matching transactions" },
        { t: "fn", cells: ["E2"], all: ["COUNTIFS"], w: 1, l: "E2 — uses COUNTIFS" },
        { t: "dyn", sets: { C2: "20000", C3: "100000" }, cells: { E2: 3 }, w: 1, l: "E2 — both range boundaries are inclusive" },
        { t: "dyn", sets: { C2: "19999", C3: "100001" }, cells: { E2: 1 }, w: 1, l: "E2 — values outside the range are excluded" },
    ],
    sol: { cells: { E2: '=COUNTIFS(B2:B5,"Sarah",C2:C5,">=20000",C2:C5,"<=100000")' } },
});

// ======================== MODUL V ========================================
task({
    n: 22, mod: 5, diff: 2,
    title: "Central Tendency: Mean, Median and Mode Comparison",
    skill: "Computing measures of central tendency and recognizing skewness.",
    setup: { cells: { ...put("A1", [["Employee ID", "Monthly Salary ($)"], [101, 3200], [102, 3400], [103, 3500], [104, 3500], [105, 3600], [106, 18000]]), D2: "Mean", D3: "Median", D4: "Mode" } },
    steps: [
        "The salary dataset is entered in range `A1:B7` (note the executive outlier in row 7).",
        "In cells `E2`, `E3` and `E4`, calculate the Mean (`AVERAGE`), Median (`MEDIAN`) and Mode (`MODE.SNGL`).",
        "Compare Mean and Median to identify positive skewness caused by the outlier.",
    ],
    checks: [
        { t: "v", c: "E2", v: 5866.666667, tol: 0.006, w: 2, l: "E2 — Mean" },
        { t: "v", c: "E3", v: 3500, w: 2, l: "E3 — Median" },
        { t: "v", c: "E4", v: 3500, w: 2, l: "E4 — Mode" },
        { t: "fn", cells: ["E2"], any: ["AVERAGE"], w: 1, l: "E2 — uses AVERAGE" },
        { t: "fn", cells: ["E3"], any: ["MEDIAN"], w: 1, l: "E3 — uses MEDIAN" },
        { t: "fn", cells: ["E4"], any: ["MODE.SNGL", "MODE"], w: 1, l: "E4 — uses MODE.SNGL" },
    ],
    sol: { cells: { E2: "=AVERAGE(B2:B7)", E3: "=MEDIAN(B2:B7)", E4: "=MODE.SNGL(B2:B7)" } },
});

task({
    n: 23, mod: 5, diff: 2,
    title: "Dispersion Metrics: Sample Standard Deviation",
    skill: "Measuring risk and data dispersion.",
    setup: { cells: { ...put("A1", [["Month", "Stock Price A ($)", "Stock Price B ($)"], ["Jan", 100, 100], ["Feb", 102, 85], ["Mar", 98, 115], ["Apr", 101, 70], ["May", 99, 130]]), A7: "StDev (sample)" } },
    steps: [
        "The stock price movements are entered in `A1:C6`.",
        "In cells `B7` and `C7`, calculate the Sample Standard Deviation (`STDEV.S`) of both stocks.",
        "Identify which stock exhibits higher volatility.",
    ],
    checks: [
        { t: "v", c: "B7", v: 1.581139, tol: 0.006, w: 2, l: "B7 — StDev of Stock A" },
        { t: "v", c: "C7", v: 23.72, tol: 0.006, w: 2, l: "C7 — StDev of Stock B" },
        { t: "fn", cells: ["B7", "C7"], any: ["STDEV.S", "STDEV"], w: 1, l: "B7:C7 — use STDEV.S" },
        { t: "fn", cells: ["B7", "C7"], none: ["STDEV.P", "STDEVP"], w: 1, l: "B7:C7 — the SAMPLE (not population) standard deviation" },
    ],
    sol: { cells: { B7: "=STDEV.S(B2:B6)", C7: "=STDEV.S(C2:C6)" } },
});

task({
    n: 24, mod: 5, diff: 3,
    title: "Quartiles, Interquartile Range (IQR) and Outlier Fence",
    skill: "Calculating positional measures for outlier identification.",
    setup: { cells: { ...put("A1", [["Response Time (Minutes)"], [12], [15], [18], [22], [25], [30], [35], [42], [55], [88]]), B2: "Q1", B3: "Q3", B4: "IQR", B5: "Upper fence" } },
    steps: [
        "The ordered response times are entered in range `A2:A11` (header in `A1`).",
        "In cell `C2`, calculate the First Quartile Q1 `=QUARTILE(A2:A11,1)`. In cell `C3`, calculate the Third Quartile Q3 `=QUARTILE(A2:A11,3)`.",
        "In cell `C4`, compute the Interquartile Range (IQR = Q3 − Q1).",
        "In cell `C5`, compute the upper outlier fence (Q3 + 1.5 × IQR) and decide whether the maximum value is an outlier.",
    ],
    checks: [
        { t: "vs", cells: { C2: 19, C3: 40.25 }, w: 2, l: "C2:C3 — Q1 and Q3" },
        { t: "v", c: "C4", v: 21.25, w: 2, l: "C4 — IQR" },
        { t: "v", c: "C5", v: 72.125, w: 2, l: "C5 — upper outlier fence" },
        { t: "fn", cells: ["C2", "C3"], any: ["QUARTILE", "QUARTILE.INC", "PERCENTILE", "PERCENTILE.INC"], w: 1, l: "C2:C3 — use QUARTILE" },
        { t: "dyn", sets: { A9: "50" }, cells: { C3: 46.25 }, w: 1, l: "C3 — updates when the data changes" },
    ],
    sol: { cells: { C2: "=QUARTILE(A2:A11,1)", C3: "=QUARTILE(A2:A11,3)", C4: "=C3-C2", C5: "=C3+1.5*C4" } },
});

task({
    n: 25, mod: 5, diff: 1,
    title: "Relative Variability using the Coefficient of Variation (CV)",
    skill: "Normalizing standard deviation against the mean.",
    setup: { cells: put("A1", [["Process", "Mean Time (sec)", "Standard Deviation (sec)", "CV"], ["Process A", 120, 12], ["Process B", 15, 3]]) },
    steps: [
        "The parameters are entered in range `A1:C3` and the header 'CV' is in `D1`.",
        "In cell `D2`, calculate the Coefficient of Variation (StDev / Mean) and format it as a Percentage. Copy down to `D3`.",
        "Compare the CV of Process A and Process B to determine relative consistency.",
    ],
    checks: [
        { t: "vs", cells: { D2: 0.1, D3: 0.2 }, w: 3, l: "D2:D3 — CV values (formulas)" },
        { t: "fmt", r: "D2:D3", type: "pct", w: 1, l: "D2:D3 — Percentage format" },
    ],
    sol: { cells: down("D", 2, 3, "=C2/B2"), fmt: fmtRange("D2:D3", pct(1)) },
});

task({
    n: 26, mod: 5, diff: 2,
    title: "Descriptive Statistics Summary Table built with Functions",
    skill: "Producing a complete statistical summary with worksheet functions (the same result as the Excel Analysis ToolPak, but available in every tool).",
    setup: { cells: { ...put("A1", [["Exam Score"], [65], [72], [85], [88], [90], [92], [95], [58], [77], [84]]), ...put("C1", [["Mean"], ["Standard Error"], ["Median"], ["Standard Deviation"], ["Minimum"], ["Maximum"], ["Count"]]) } },
    steps: [
        "The 10 exam scores are in range `A2:A11` and the labels Mean, Standard Error, Median, Standard Deviation, Minimum, Maximum, Count are in `C1:C7`.",
        "In `D1:D7`, write the matching formulas (Standard Error = Standard Deviation / `SQRT`(Count)).",
        "Verify the generated summary table.",
    ],
    checks: [
        { t: "vs", cells: { D1: 80.6, D3: 84.5 }, tol: 0.006, w: 2, l: "D1, D3 — Mean and Median" },
        { t: "vs", cells: { D2: 3.8764, D4: 12.2589 }, tol: 0.006, w: 2, l: "D2, D4 — Standard Error and Standard Deviation" },
        { t: "vs", cells: { D5: 58, D6: 95, D7: 10 }, w: 2, l: "D5:D7 — Minimum, Maximum and Count" },
        { t: "fn", cells: ["D2"], all: ["SQRT"], any: ["STDEV.S", "STDEV"], w: 1, l: "D2 — Standard Error uses STDEV.S and SQRT" },
    ],
    sol: { cells: { D1: "=AVERAGE(A2:A11)", D2: "=STDEV.S(A2:A11)/SQRT(COUNT(A2:A11))", D3: "=MEDIAN(A2:A11)", D4: "=STDEV.S(A2:A11)", D5: "=MIN(A2:A11)", D6: "=MAX(A2:A11)", D7: "=COUNT(A2:A11)" } },
});

// ======================== MODUL VI =======================================
task({
    n: 27, mod: 6, diff: 3,
    title: "Data Validation: Numeric Limits with Error Alert",
    skill: "Restricting data entry to a valid numeric interval.",
    setup: { cells: put("A1", [["Student Name", "Score (Allowed: 0 to 100)"], ["John"], ["Alice"], ["Bob"]]) },
    steps: [
        "The data is entered in `A1:B4`. Select range `B2:B10`.",
        "Open the **Data Validation** dialog (toolbar) and allow only whole numbers between 0 and 100.",
        "Configure the error message 'Please enter a score between 0 and 100.'",
        "Test the rule by entering 150 in cell `B2` (it must be rejected).",
    ],
    checks: [
        { t: "dv", r: "B2:B10", type: "whole", min: 0, max: 100, op: "between", w: 3, l: "B2:B10 — whole number between 0 and 100" },
        { t: "dvmsg", r: "B2:B10", msg: "Please enter a score between 0 and 100.", w: 1, l: "B2:B10 — error message text" },
        { t: "dvrej", c: "B2", rej: ["150", "-1", "7.5", "abc"], acc: ["0", "55", "100"], w: 2, l: "B2 — invalid entries are rejected, valid ones accepted" },
    ],
    sol: { dv: [{ range: "B2:B10", type: "whole", op: "between", min: 0, max: 100, title: "Invalid Score", msg: "Please enter a score between 0 and 100.", style: "stop" }] },
});

task({
    n: 28, mod: 6, diff: 2,
    title: "Dropdown List Creation using Data Validation",
    skill: "Standardizing categorical input with a controlled dropdown list.",
    setup: { cells: { ...put("A1", [["Ticket ID", "Department Dropdown"], ["TCK-01"], ["TCK-02"], ["TCK-03"], ["TCK-04"]]), ...put("G1", [["IT"], ["HR"], ["Finance"], ["Sales"]]) } },
    steps: [
        "The allowed list is entered in range `G1:G4` = 'IT', 'HR', 'Finance', 'Sales'.",
        "Select range `B2:B5`.",
        "Apply data validation of type **List** that takes its items from the range `$G$1:$G$4`.",
        "Verify that a dropdown arrow appears in cell `B2` and that only the 4 allowed items can be chosen.",
    ],
    checks: [
        { t: "dv", r: "B2:B5", type: "list", w: 1, l: "B2:B5 — validation of type List" },
        { t: "dvsrc", r: "B2:B5", src: "G1:G4", w: 2, l: "B2:B5 — list items come from the range $G$1:$G$4" },
        { t: "dvitems", r: "B2:B5", items: ["IT", "HR", "Finance", "Sales"], w: 1, l: "B2:B5 — exactly the 4 allowed items" },
        { t: "dvrej", c: "B3", rej: ["Marketing", "xyz"], acc: ["HR", "Sales"], w: 1, l: "B3 — other items are rejected" },
    ],
    sol: { dv: [{ range: "B2:B5", type: "list", source: "$G$1:$G$4", msg: "", style: "stop" }] },
});

task({
    n: 29, mod: 6, diff: 2,
    title: "Conditional Formatting based on Value Thresholds",
    skill: "Highlighting critical metrics dynamically.",
    setup: { cells: put("A1", [["Project", "Budget ($)", "Actual Spend ($)", "Variance ($)"], ["Project A", 100000, 115000], ["Project B", 80000, 72000], ["Project C", 50000, 58000]]) },
    steps: [
        "The data is entered in `A1:C4`. In column `D`, calculate Variance = Budget − Actual Spend.",
        "Select range `D2:D4`.",
        "Create a **Conditional Formatting** rule that formats cells with a value less than 0.",
        "Choose a light red fill with dark red text.",
    ],
    checks: [
        { t: "vs", cells: { D2: -15000, D3: 8000, D4: -8000 }, w: 2, l: "D2:D4 — Variance (formulas)" },
        { t: "cf", r: "D2:D4", hl: ["D2", "D4"], w: 3, l: "D2:D4 — exactly the negative variances are highlighted" },
        { t: "cf", r: "D2:D4", hl: ["D2", "D4"], fills: ["lightred"], w: 1, l: "D2:D4 — light red fill with dark red text" },
    ],
    sol: { cells: down("D", 2, 4, "=B2-C2"), cf: [{ range: "D2:D4", kind: "cell", op: "lt", v1: "0", fill: "lightred" }] },
});

task({
    n: 30, mod: 6, diff: 3,
    title: "Formula-Based Conditional Formatting for Duplicate Detection",
    skill: "Using a logical formula inside conditional formatting to flag non-unique entries.",
    setup: { cells: put("A1", [["Invoice ID", "Amount ($)"], ["INV-1001", 500], ["INV-1002", 750], ["INV-1001", 500], ["INV-1003", 1200]]) },
    steps: [
        "The data is entered in `A1:B5`. Select range `A2:A5`.",
        "Create a **Conditional Formatting** rule that uses a custom formula.",
        "Enter the formula `=COUNTIF($A$2:$A$5,A2)>1`.",
        "Set the fill color to yellow.",
    ],
    checks: [
        { t: "cf", r: "A2:A5", hl: ["A2", "A4"], w: 3, l: "A2:A5 — exactly the duplicate invoice IDs are highlighted" },
        { t: "cf", r: "A2:A5", hl: ["A2", "A4"], kind: "formula", w: 2, l: "A2:A5 — rule type: custom formula" },
        { t: "cf", r: "A2:A5", hl: ["A2", "A4"], fills: ["yellow", "lightyellow"], w: 1, l: "A2:A5 — yellow fill" },
    ],
    sol: { cf: [{ range: "A2:A5", kind: "formula", formula: "=COUNTIF($A$2:$A$5,A2)>1", fill: "yellow" }] },
});

task({
    n: 31, mod: 6, diff: 3,
    title: "Highlighting Outliers with Z-Score Threshold Formatting",
    skill: "Automating a visual statistical audit rule.",
    setup: { cells: put("A1", [["Machine ID", "Temp (°C)", "Mean", "StDev", "Z-Score"], ["M1", 45, 48, 4], ["M2", 47, 48, 4], ["M3", 62, 48, 4]]) },
    steps: [
        "The data is entered in `A1:D4`. In column `E`, calculate the Z-Score = (Temp − Mean) / StDev.",
        "Apply **Conditional Formatting** to range `E2:E4` with the custom formula `=ABS(E2)>2` and a red fill.",
        "Observe which machine is flagged as a statistical anomaly.",
    ],
    checks: [
        { t: "vs", cells: { E2: -0.75, E3: -0.25, E4: 3.5 }, w: 2, l: "E2:E4 — Z-Scores (formulas)" },
        { t: "cf", r: "E2:E4", hl: ["E4"], w: 3, l: "E2:E4 — only the outlier (|Z| > 2) is highlighted" },
        { t: "cf", r: "E2:E4", hl: ["E4"], kind: "formula", fills: ["red", "lightred"], w: 1, l: "E2:E4 — custom-formula rule with a red fill" },
    ],
    sol: { cells: down("E", 2, 4, "=(B2-C2)/D2"), cf: [{ range: "E2:E4", kind: "formula", formula: "=ABS(E2)>2", fill: "red" }] },
});

// ======================== MODUL VII ======================================
task({
    n: 32, mod: 7, diff: 3,
    title: "Building a Multi-Dimensional PivotTable",
    skill: "Aggregating transactional data interactively.",
    setup: { cells: put("A1", [["Region", "Sales Rep", "Category", "Sales ($)"], ["North", "John", "Tech", 12000], ["North", "John", "Office", 4000], ["South", "Sara", "Tech", 15000], ["South", "Sara", "Tech", 8000]]) },
    steps: [
        "The dataset is entered in `A1:D5`.",
        "Insert a **PivotTable** (toolbar) from the range `A1:D5`.",
        "Place 'Region' in Rows, 'Category' in Columns and 'Sales ($)' in Values (Sum of Sales).",
        "Format the value cells as Currency ($).",
    ],
    checks: [
        { t: "pivot", spec: { src: "A1:D5", rows: "Region", cols: "Category", vals: [{ f: "Sales ($)", agg: "sum" }] }, w: 3, l: "PivotTable: Rows = Region, Columns = Category, Values = Sum of Sales ($)" },
        { t: "pivotv", spec: { rows: "Region", cols: "Category", vals: [{ f: "Sales ($)", agg: "sum" }] }, expect: [{ r: "North", c: "Tech", v: 12000 }, { r: "North", c: "Office", v: 4000 }, { r: "South", c: "Tech", v: 23000 }, { grand: 1, v: 39000 }], w: 1, l: "PivotTable output values" },
        { t: "pivotfmt", spec: { rows: "Region", cols: "Category", vals: [{ f: "Sales ($)", agg: "sum" }] }, numFmt: "currency", w: 1, l: "PivotTable values formatted as Currency" },
    ],
    sol: { pivots: [{ id: "P1", src: "A1:D5", rows: "Region", cols: "Category", values: [{ field: "Sales ($)", agg: "sum", show: "value" }], filters: {}, numFmt: "currency" }] },
});

task({
    n: 33, mod: 7, diff: 4,
    title: "Grouping Numeric Values into Custom Intervals in a Pivot Table",
    skill: "Grouping a numeric dimension into discrete bins. A helper column is used so the same method works in every tool.",
    setup: { cells: put("A1", [["Discount Rate", "Profit ($)", "Discount Band"], [0, 15000], [0.05, 12000], [0.12, 8000], [0.25, -2000], [0.4, -8000]]) },
    steps: [
        "The data is entered in `A1:B6` and the header 'Discount Band' is in `C1`.",
        "In cell `C2`, calculate the lower bound of the 10-point band that the discount belongs to (for example 0.12 → 0.1), and copy it down to `C6`.",
        "Insert a **PivotTable** from `A1:C6` with 'Discount Band' in Rows and Sum of 'Profit ($)' in Values.",
        "Identify the discount band at which total profit turns negative.",
    ],
    checks: [
        { t: "vs", cells: { C2: 0, C3: 0, C4: 0.1, C5: 0.2, C6: 0.4 }, w: 2, l: "C2:C6 — discount bands (formulas)" },
        { t: "fn", cells: ["C2"], any: ["ROUNDDOWN", "INT", "FLOOR", "TRUNC"], w: 1, l: "C2 — rounds the discount down to its band" },
        { t: "pivot", spec: { src: "A1:C6", rows: "Discount Band", vals: [{ f: "Profit ($)", agg: "sum" }] }, w: 2, l: "PivotTable: Rows = Discount Band, Values = Sum of Profit ($)" },
        { t: "pivotv", spec: { rows: "Discount Band", vals: [{ f: "Profit ($)", agg: "sum" }] }, expect: [{ r: 0, v: 27000 }, { r: 0.1, v: 8000 }, { r: 0.2, v: -2000 }, { r: 0.4, v: -8000 }], w: 1, l: "PivotTable output values per band" },
    ],
    sol: { cells: down("C", 2, 6, "=ROUNDDOWN(A2*10,0)/10"), pivots: [{ id: "P1", src: "A1:C6", rows: "Discount Band", cols: "", values: [{ field: "Profit ($)", agg: "sum", show: "value" }], filters: {}, numFmt: "general" }] },
});

task({
    n: 34, mod: 7, diff: 3,
    title: "Adding a Commission Measure to a Pivot Table",
    skill: "Extending pivot analysis with a derived measure.",
    setup: { cells: put("A1", [["Sales Rep", "Total Revenue ($)", "Total Cost ($)", "Commission ($)"], ["Alice", 200000, 140000], ["Bob", 150000, 90000]]) },
    steps: [
        "The data is entered in `A1:C3` and the header 'Commission ($)' is in `D1`.",
        "In cell `D2`, calculate Commission = (Revenue − Cost) × 10% and copy it down to `D3`.",
        "Insert a **PivotTable** from `A1:D3` with 'Sales Rep' in Rows and Sum of 'Commission ($)' in Values.",
    ],
    checks: [
        { t: "vs", cells: { D2: 6000, D3: 6000 }, w: 2, l: "D2:D3 — Commission (formulas)" },
        { t: "dyn", sets: { B2: "300000" }, cells: { D2: 16000 }, w: 1, l: "D2 — updates when revenue changes" },
        { t: "pivot", spec: { src: "A1:D3", rows: "Sales Rep", vals: [{ f: "Commission ($)", agg: "sum" }] }, w: 2, l: "PivotTable: Rows = Sales Rep, Values = Sum of Commission ($)" },
        { t: "pivotv", spec: { rows: "Sales Rep", vals: [{ f: "Commission ($)", agg: "sum" }] }, expect: [{ r: "Alice", v: 6000 }, { r: "Bob", v: 6000 }, { grand: 1, v: 12000 }], w: 1, l: "PivotTable output values" },
    ],
    sol: { cells: down("D", 2, 3, "=(B2-C2)*0.1"), pivots: [{ id: "P1", src: "A1:D3", rows: "Sales Rep", cols: "", values: [{ field: "Commission ($)", agg: "sum", show: "value" }], filters: {}, numFmt: "general" }] },
});

task({
    n: 35, mod: 7, diff: 3,
    title: "Interactive Filtering with Slicers",
    skill: "Implementing interactive filter controls.",
    setup: { cells: put("A1", [["Date", "Region", "Sales ($)"], ["2024-01-15", "North", 5000], ["2024-02-10", "South", 8000], ["2024-03-05", "North", 12000]]) },
    steps: [
        "The data is entered in `A1:C4`. Insert a **PivotTable** showing Sum of 'Sales ($)' by Region.",
        "Add a **Slicer** for the 'Region' field (button inside the PivotTable panel).",
        "Click 'North' in the slicer.",
        "Observe the interactive filtering of the PivotTable.",
    ],
    checks: [
        { t: "pivot", spec: { src: "A1:C4", rows: "Region", vals: [{ f: "Sales ($)", agg: "sum" }] }, w: 2, l: "PivotTable: Rows = Region, Values = Sum of Sales ($)" },
        { t: "pivot", spec: { src: "A1:C4", rows: "Region", vals: [{ f: "Sales ($)", agg: "sum" }], filters: { Region: ["North"] } }, w: 2, l: "Slicer on Region with only 'North' selected" },
        { t: "pivotv", spec: { rows: "Region", vals: [{ f: "Sales ($)", agg: "sum" }], filters: { Region: ["North"] } }, expect: [{ grand: 1, v: 17000 }], w: 1, l: "PivotTable shows only the North total" },
    ],
    sol: { pivots: [{ id: "P1", src: "A1:C4", rows: "Region", cols: "", values: [{ field: "Sales ($)", agg: "sum", show: "value" }], filters: { Region: ["North"] }, slicers: ["Region"], numFmt: "general" }] },
});

task({
    n: 36, mod: 7, diff: 3,
    title: "Showing Percent of Grand Total in a Pivot Table",
    skill: "Transforming raw sums into relative proportions.",
    setup: { cells: put("A1", [["Category", "Revenue ($)"], ["Electronics", 120000], ["Furniture", 50000], ["Clothing", 30000]]) },
    steps: [
        "The data is entered in `A1:B4`. Insert a **PivotTable** with Category in Rows and Revenue in Values.",
        "Add 'Revenue ($)' to Values a second time.",
        "Change the second value column so that it is shown as '% of Grand Total'.",
        "Verify the relative percentage distribution.",
    ],
    checks: [
        { t: "pivot", spec: { src: "A1:B4", rows: "Category", vals: [{ f: "Revenue ($)", agg: "sum", show: "value" }] }, w: 2, l: "PivotTable: Rows = Category, Values = Sum of Revenue ($)" },
        { t: "pivot", spec: { src: "A1:B4", rows: "Category", vals: [{ f: "Revenue ($)", agg: "sum" }, { f: "Revenue ($)", agg: "sum", show: "pct_grand" }] }, w: 2, l: "Second value column is shown as % of Grand Total" },
        { t: "pivotv", spec: { rows: "Category", vals: [{ f: "Revenue ($)", agg: "sum" }, { f: "Revenue ($)", agg: "sum", show: "pct_grand" }] }, expect: [{ r: "Electronics", vi: 1, v: 0.6 }, { r: "Furniture", vi: 1, v: 0.25 }, { r: "Clothing", vi: 1, v: 0.15 }, { grand: 1, vi: 0, v: 200000 }], w: 1, l: "PivotTable output values" },
    ],
    sol: { pivots: [{ id: "P1", src: "A1:B4", rows: "Category", cols: "", values: [{ field: "Revenue ($)", agg: "sum", show: "value" }, { field: "Revenue ($)", agg: "sum", show: "pct_grand" }], filters: {}, numFmt: "general" }] },
});

// ======================== MODUL VIII =====================================
task({
    n: 37, mod: 8, diff: 4,
    title: "Combo Chart with a Secondary Axis for Dual Metrics",
    skill: "Combining different scales (Revenue $ vs Profit Margin %) in a single chart.",
    setup: { cells: put("A1", [["Month", "Revenue ($)", "Profit Margin (%)"], ["Jan", 100000, 0.12], ["Feb", 125000, 0.15], ["Mar", 150000, 0.18], ["Apr", 140000, 0.14]]), fmt: fmtRange("C2:C5", pct(0)) },
    steps: [
        "The data is entered in `A1:C5` and column `C` is formatted as Percentage. Select `A1:C5`.",
        "Insert a **Chart** (toolbar) of type Combo: Revenue as columns on the primary (left) axis.",
        "Show Profit Margin as a line on the secondary (right) axis.",
        "Add the chart title 'Monthly Revenue & Margin Trend'.",
    ],
    checks: [
        { t: "chart", spec: { type: "combo", range: "A1:C5" }, w: 1, l: "Combo chart built from A1:C5" },
        { t: "chart", spec: { type: "combo", range: "A1:C5", series: [{ col: "B", kind: "column", axis: "primary" }] }, w: 1, l: "Revenue — columns on the primary axis" },
        { t: "chart", spec: { type: "combo", range: "A1:C5", series: [{ col: "C", kind: "line", axis: "secondary" }] }, w: 2, l: "Profit Margin — line on the secondary axis" },
        { t: "chart", spec: { type: "combo", range: "A1:C5", title: "Monthly Revenue & Margin Trend" }, w: 1, l: "Chart title" },
    ],
    sol: { charts: [{ id: "C1", type: "combo", range: "A1:C5", title: "Monthly Revenue & Margin Trend", series: [{ col: "B", kind: "column", axis: "primary" }, { col: "C", kind: "line", axis: "secondary" }], trendline: null }] },
});

task({
    n: 38, mod: 8, diff: 4,
    title: "Linear Trendline and R² on a Scatter Plot",
    skill: "Evaluating correlation and goodness of fit.",
    setup: { cells: { ...put("A1", [["Marketing Spend ($)", "Sales Units"], [1000, 120], [2000, 210], [3000, 310], [4000, 390], [5000, 510]]), E2: "← Slope", E3: "← Intercept", E4: "← R-squared" } },
    steps: [
        "The data is entered in `A1:B6`. Select `A1:B6`.",
        "Insert a **Chart** of type Scatter (X, Y).",
        "Add a linear trendline and display the equation and the R-squared value on the chart.",
        "In cells `D2`, `D3` and `D4` verify the chart with `SLOPE`, `INTERCEPT` and `RSQ`. Evaluate the strength of the model.",
    ],
    checks: [
        { t: "chart", spec: { type: "scatter", range: "A1:B6" }, w: 1, l: "Scatter chart built from A1:B6" },
        { t: "chart", spec: { type: "scatter", range: "A1:B6", trend: {} }, w: 1, l: "Linear trendline added" },
        { t: "chart", spec: { type: "scatter", range: "A1:B6", trend: { eq: true, r2: true } }, w: 1, l: "Trendline shows the equation and the R-squared value" },
        { t: "vs", cells: { D2: 0.096, D3: 20 }, tol: 1e-6, w: 2, l: "D2:D3 — SLOPE and INTERCEPT" },
        { t: "v", c: "D4", v: 0.9965, tol: 0.00006, w: 1, l: "D4 — RSQ" },
        { t: "fn", cells: ["D2"], all: ["SLOPE"], w: 0.5, l: "D2 — uses SLOPE" },
        { t: "fn", cells: ["D3"], all: ["INTERCEPT"], w: 0.5, l: "D3 — uses INTERCEPT" },
    ],
    sol: { cells: { D2: "=SLOPE(B2:B6,A2:A6)", D3: "=INTERCEPT(B2:B6,A2:A6)", D4: "=RSQ(B2:B6,A2:A6)" }, charts: [{ id: "C1", type: "scatter", range: "A1:B6", title: "", series: [{ col: "B", kind: "column", axis: "primary" }], trendline: { type: "linear", eq: true, r2: true } }] },
});

task({
    n: 39, mod: 8, diff: 2,
    title: "Linear Forecasting using FORECAST",
    skill: "Predicting a future value from historical trend data.",
    setup: { cells: { ...put("A1", [["Month (x)", "Sales ($) (y)"], [1, 12000], [2, 13500], [3, 15000], [4, 16800], [5, 18200]]), D1: "Target Month (x)", D2: 6, E1: "Forecast ($)" } },
    steps: [
        "The historical 5-month sales are entered in `A1:B6` and the target month x = 6 is in cell `D2`.",
        "In cell `E2`, write a `FORECAST` formula to calculate the projected sales of month 6.",
    ],
    checks: [
        { t: "v", c: "E2", v: 19810, tol: 0.006, w: 3, l: "E2 — projected sales of month 6" },
        { t: "fn", cells: ["E2"], any: ["FORECAST", "FORECAST.LINEAR"], w: 1, l: "E2 — uses FORECAST" },
        { t: "dyn", sets: { D2: "7" }, cells: { E2: 21380 }, tol: 0.006, w: 1, l: "E2 — updates when the target month changes" },
    ],
    sol: { cells: { E2: "=FORECAST(D2,B2:B6,A2:A6)" } },
});

task({
    n: 40, mod: 8, diff: 3,
    title: "Target Profit Analysis (Goal Seek and Algebraic Solution)",
    skill: "Reverse-engineering an input variable to reach a financial target.",
    setup: { cells: { ...put("A1", [["Parameter", "Value"], ["Unit Price ($)", 50], ["Units Sold", 1000], ["Fixed Costs ($)", 20000], ["Variable Cost per Unit ($)", 20]]), A6: "Net Profit ($)", B6: "=B2*B3-B4-B5*B3", A8: "Target Profit", B8: 25000, A9: "Required Units" } },
    steps: [
        "The profit model is in `A1:B6` (`B6` holds the formula `=B2*B3-B4-B5*B3`). Current Net Profit = $10,000.",
        "Business goal: achieve a Net Profit of exactly $25,000. The target is in `B8` (label 'Target Profit' in `A8`).",
        "In cell `B9` (label 'Required Units' in `A9`), solve algebraically: Units = (Target + Fixed Costs) / (Unit Price − Variable Cost per Unit).",
        "Type the result into `B3` and check that `B6` shows 25,000.",
    ],
    checks: [
        { t: "v", c: "B9", v: 1500, w: 3, l: "B9 — required units (formula)" },
        { t: "re", cells: ["B9"], re: "B8", w: 1, l: "B9 — refers to the target in B8" },
        { t: "dyn", sets: { B8: "30000" }, cells: { B9: 50000 / 30 }, tol: 1e-6, w: 1, l: "B9 — updates when the target changes" },
        { t: "v", c: "B3", v: 1500, f: false, const: true, w: 1, l: "B3 — Units Sold set to the required units" },
        { t: "v", c: "B6", v: 25000, f: false, const: true, w: 1, l: "B6 — Net Profit reaches the target" },
    ],
    sol: { cells: { B9: "=(B8+B4)/(B2-B5)", B3: "1500" } },
});

task({
    n: 41, mod: 8, diff: 4,
    title: "Two-Variable Loan Payment Matrix using Mixed References",
    skill: "Evaluating a multi-scenario matrix across interest rates and loan terms with one formula (replaces the Excel-only Data Table feature).",
    setup: { cells: { ...put("A1", [["Loan ($)", "Rate", "Months", "Monthly Payment ($)"], [100000, 0.06, 360]]), D2: "=PMT(B2/12,C2,-A2)", ...put("F1", [[0.05, 0.06, 0.07]]), ...put("E2", [[180], [240], [360]]) } },
    steps: [
        "The base data is entered in `A1:D2` (`D2` = `=PMT(B2/12,C2,-A2)`, base payment = $599.55).",
        "The matrix headers are entered: rates 5%, 6%, 7% in `F1:H1` and terms 180, 240, 360 months in `E2:E4`.",
        "In cell `F2`, write ONE payment formula that uses mixed references so it can be copied across `F2:H4`. Use the loan amount from `$A$2`.",
        "Copy `F2` across the whole range `F2:H4`.",
    ],
    checks: [
        { t: "vs", cells: { F2: 790.79, G2: 843.86, H2: 898.83, F3: 659.96, G3: 716.43, H3: 775.3, F4: 536.82, G4: 599.55, H4: 665.3 }, tol: 0.006, w: 4, l: "F2:H4 — all nine payments" },
        { t: "re", cells: ["F2"], re: "F\\$1", w: 1, l: "F2 — rate reference with the row fixed (F$1)" },
        { t: "re", cells: ["F2"], re: "\\$E2", w: 1, l: "F2 — term reference with the column fixed ($E2)" },
        { t: "re", cells: ["F2"], re: "\\$A\\$2", w: 1, l: "F2 — loan amount fixed ($A$2)" },
    ],
    sol: { cells: block("F2:H4", "=PMT(F$1/12,$E2,-$A$2)") },
});

task({
    n: 42, mod: 8, diff: 3,
    title: "Best / Base / Worst Case Scenario Model with a Selector",
    skill: "Modeling business outcomes under alternative assumptions (replaces the Excel-only Scenario Manager).",
    setup: { cells: { ...put("A1", [["Scenario", "Sales Units", "Price ($)", "Total Revenue ($)"], ["Best Case", 7000, 110], ["Base Case", 5000, 100], ["Worst Case", 3000, 85]]), F1: "Active Scenario", F2: "Units", F3: "Price", F4: "Revenue" } },
    steps: [
        "The scenario table is entered in `A1:D4`. Calculate Total Revenue in column `D` (Units × Price).",
        "In `G1` create a dropdown list (**Data Validation**, type List) from the range `A2:A4` and select 'Base Case'.",
        "In `G2` and `G3`, use `INDEX` and `MATCH` to pull the Units and Price of the scenario selected in `G1`.",
        "In `G4`, calculate the Revenue of the active scenario (`=G2*G3`). Switch between the three scenarios and observe the results.",
    ],
    checks: [
        { t: "vs", cells: { D2: 770000, D3: 500000, D4: 255000 }, w: 2, l: "D2:D4 — Total Revenue (formulas)" },
        { t: "dvsrc", r: "G1", src: "A2:A4", w: 2, l: "G1 — dropdown list from A2:A4" },
        { t: "fn", cells: ["G2", "G3"], all: ["INDEX", "MATCH"], w: 1, l: "G2:G3 — use INDEX with MATCH" },
        { t: "dyn", sets: { G1: "Worst Case" }, cells: { G2: 3000, G3: 85, G4: 255000 }, w: 2, l: "G2:G4 — Worst Case values" },
        { t: "dyn", sets: { G1: "Best Case" }, cells: { G2: 7000, G3: 110, G4: 770000 }, w: 1, l: "G2:G4 — Best Case values" },
    ],
    sol: { cells: { ...down("D", 2, 4, "=B2*C2"), G1: "Base Case", G2: "=INDEX(B2:B4,MATCH($G$1,A2:A4,0))", G3: "=INDEX(C2:C4,MATCH($G$1,A2:A4,0))", G4: "=G2*G3" }, dv: [{ range: "G1", type: "list", source: "$A$2:$A$4", msg: "", style: "stop" }] },
});

module.exports = { TASKS, MODULES };
