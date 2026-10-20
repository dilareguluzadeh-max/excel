"use strict";
const { TASKS } = require("./bank_excel.js");
const G = require("../grader.js");
const merge = (t) => {
  const s = t.setup || {}, o = t.sol || {};
  return { cells: { ...(s.cells||{}), ...(o.cells||{}) }, fmt: { ...(s.fmt||{}), ...(o.fmt||{}) },
    dv: o.dv||[], cf: o.cf||[], pivots: o.pivots||[], charts: o.charts||[] };
};
let bad = 0;
for (const t of TASKS) {
  const task = { id: "T"+t.n, pts: 17, checks: t.checks };
  const full = G.gradeTask(task, merge(t));
  const setup = G.gradeTask(task, { cells: (t.setup||{}).cells||{}, fmt: (t.setup||{}).fmt||{} });
  const okF = full.ok, okS = setup.pts === 0;
  if (!okF || !okS) { bad++; console.log("TASK", t.n, "full:", full.pts, "setup:", setup.pts, "FAILED:", full.results.filter(r=>!r.ok).map(r=>r.l).join(" | "), okS ? "" : "| setup-only scored: "+setup.results.filter(r=>r.ok).map(r=>r.l).join(" | ")); }
}
console.log(TASKS.length, "tasks;", bad, "problems");
