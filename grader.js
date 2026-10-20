/* ====================================================================
 * SDF 1 — Cədvəl tapşırıqlarının yoxlayıcısı (brauzer + Node-da eyni kod)
 *
 * gradeTask(task, state) → { pts, max, got, tot, ok, partial, msg, detail, results }
 *   task  : exam_data.js-dəki tapşırıq ({ pts, checks:[...] })
 *   state : tələbənin cədvəli { cells, fmt, dv, cf, pivots, charts }
 *
 * Hər tapşırıq bir neçə meyardan (check) ibarətdir; hər meyarın çəkisi (w) var.
 * Bal = round(pts × ödənən çəki / ümumi çəki)  →  0 ilə pts arasında qismən bal.
 * ==================================================================== */
(function (root, factory) {
    if (typeof module === "object" && module.exports) module.exports = factory(require("./sheetengine.js"));
    else root.Grader = factory(root.SheetEngine);
})(typeof self !== "undefined" ? self : this, function (SE) {
    "use strict";

    const clone = (o) => JSON.parse(JSON.stringify(o || {}));
    const norm = (s) => String(s ?? "").trim().toLowerCase();
    const strip = (f) => String(f || "").replace(/\s+/g, "").toUpperCase();

    function blankState() { return { cells: {}, fmt: {}, dv: [], cf: [], pivots: [], charts: [] }; }
    function normState(s) {
        const st = Object.assign(blankState(), clone(s));
        return st;
    }

    // dəyər müqayisəsi
    function valEq(actual, exp, tol, ci) {
        if (Array.isArray(exp)) return exp.some((e) => valEq(actual, e, tol, ci));
        if (exp && typeof exp === "object" && exp.err) return SE.isErr(actual) && actual.code === exp.err;
        if (SE.isErr(actual)) return false;
        if (typeof exp === "number") {
            if (typeof actual !== "number") return false;
            const t = tol !== undefined ? tol : Math.max(1e-9, Math.abs(exp) * 1e-9);
            return Math.abs(actual - exp) <= t;
        }
        if (typeof exp === "boolean") return actual === exp;
        if (typeof exp === "string") {
            if (typeof actual !== "string") return false;
            return ci ? norm(actual) === norm(exp) : actual.trim() === exp.trim();
        }
        return false;
    }
    const isFormula = (wb, a) => String(wb.raw(a)).trim()[0] === "=";

    function cellOk(wb, a, exp, ck) {
        if (ck.f !== false && !isFormula(wb, a) && !ck.const) return false;
        return valEq(wb.value(a), exp, ck.tol, ck.ci);
    }

    function applySets(state, sets) {
        const st = normState(state);
        for (const [a, v] of Object.entries(sets || {})) {
            if (v === "" || v === null) delete st.cells[a]; else st.cells[a] = String(v);
        }
        return new SE.Workbook(st);
    }

    function fnsIn(raw) {
        const out = new Set();
        const re = /([A-Z][A-Z0-9.]*)\(/g;
        let m;
        const up = String(raw || "").toUpperCase().replace(/"[^"]*"/g, '""');
        while ((m = re.exec(up))) out.add(m[1]);
        return out;
    }

    function findPivot(state, spec) {
        for (const p of state.pivots || []) {
            if (spec.src && !SE.rangeCovers(p.src, spec.src) && p.src.replace(/\$/g, "") !== spec.src) continue;
            if (spec.rows !== undefined && p.rows !== spec.rows) continue;
            if (spec.cols !== undefined && (p.cols || "") !== (spec.cols || "")) continue;
            if (spec.vals) {
                const pv = p.values || [];
                if (pv.length < spec.vals.length) continue;
                if (!spec.vals.every((v, i) => pv[i] && pv[i].field === v.f && (pv[i].agg || "sum") === (v.agg || "sum") && (v.show === undefined || (pv[i].show || "value") === v.show))) continue;
            }
            return p;
        }
        return null;
    }
    function filtersEq(p, want) {
        const f = p.filters || {};
        const keys = Object.keys(f).filter((k) => f[k] && f[k].length);
        const wk = Object.keys(want || {});
        if (keys.length !== wk.length) return false;
        return wk.every((k) => f[k] && f[k].length === want[k].length && want[k].every((x) => f[k].includes(x)));
    }

    function chartMatches(ch, spec) {
        if (spec.type && ch.type !== spec.type) return false;
        if (spec.range && ch.range.replace(/\$/g, "").toUpperCase() !== spec.range) return false;
        if (spec.title !== undefined && norm(ch.title) !== norm(spec.title)) return false;
        for (const s of spec.series || []) {
            const hit = (ch.series || []).find((x) => x.col === s.col);
            if (!hit) return false;
            if (s.kind && hit.kind !== s.kind) return false;
            if (s.axis && (hit.axis || "primary") !== s.axis) return false;
        }
        if (spec.trend) {
            if (!ch.trendline) return false;
            if (spec.trend.eq && !ch.trendline.eq) return false;
            if (spec.trend.r2 && !ch.trendline.r2) return false;
        }
        return true;
    }

    const CHECKS = {
        // bir xananın dəyəri (+ düstur olmalıdır)
        v(ck, wb) { return cellOk(wb, ck.c, ck.v, ck); },
        // bir neçə xana: hamısı düzgün olmalıdır
        vs(ck, wb) { return Object.entries(ck.cells).every(([a, e]) => cellOk(wb, a, e, ck)); },
        // düsturda funksiyalar iştirak etməlidir
        fn(ck, wb) {
            const cells = ck.cells || [ck.c];
            return cells.every((a) => {
                if (!isFormula(wb, a)) return false;
                const used = fnsIn(wb.raw(a));
                if (ck.all && !ck.all.every((f) => used.has(f))) return false;
                if (ck.any && !ck.any.some((f) => used.has(f))) return false;
                if (ck.none && ck.none.some((f) => used.has(f))) return false;
                return true;
            });
        },
        // düstur mətni nümunəyə uyğun olmalıdır
        re(ck, wb) {
            const cells = ck.cells || [ck.c];
            return cells.every((a) => isFormula(wb, a) && new RegExp(ck.re, "i").test(strip(wb.raw(a))));
        },
        // giriş xanalarını dəyişib nəticənin yeniləndiyini yoxlayır
        dyn(ck, wb, state) {
            const w2 = applySets(state, ck.sets);
            return Object.entries(ck.cells).every(([a, e]) => {
                if (ck.f !== false && !isFormula(w2, a)) return false;
                return valEq(w2.value(a), e, ck.tol, ck.ci);
            });
        },
        fmt(ck, wb, state) {
            return SE.cellsIn(ck.r).every((a) => {
                const f = (state.fmt || {})[a];
                return f && f.t === ck.type && (ck.d === undefined || (f.d ?? 2) === ck.d);
            });
        },
        noerr(ck, wb) { return SE.cellsIn(ck.r).every((a) => !SE.isErr(wb.value(a))); },
        // --- məlumat doğrulaması
        dv(ck, wb, state) {
            const rules = (state.dv || []).filter((r) => SE.rangeCovers(r.range, ck.r));
            return rules.some((r) => r.type === ck.type &&
                (ck.min === undefined || Number(r.min) === ck.min) && (ck.max === undefined || Number(r.max) === ck.max) &&
                (ck.op === undefined || (r.op || "between") === ck.op));
        },
        dvmsg(ck, wb, state) {
            const rules = (state.dv || []).filter((r) => SE.rangeCovers(r.range, ck.r));
            return rules.some((r) => norm(r.msg).replace(/\.$/, "") === norm(ck.msg).replace(/\.$/, ""));
        },
        dvrej(ck, wb) {
            return (ck.rej || []).every((x) => SE.dvViolation(wb, ck.c, x)) && (ck.acc || []).every((x) => !SE.dvViolation(wb, ck.c, x));
        },
        dvsrc(ck, wb, state) {   // siyahı mənbəyi diapazondur və gözlənilən diapazona işarə edir
            const rules = (state.dv || []).filter((r) => r.type === "list" && SE.rangeCovers(r.range, ck.r));
            return rules.some((r) => {
                const src = String(r.source || "").replace(/^=/, "").replace(/\$/g, "").trim().toUpperCase();
                return src === ck.src;
            });
        },
        dvitems(ck, wb, state) {
            const rules = (state.dv || []).filter((r) => r.type === "list" && SE.rangeCovers(r.range, ck.r));
            return rules.some((r) => {
                const items = SE.dvListItems(wb, r);
                return items.length === ck.items.length && ck.items.every((x) => items.includes(x));
            });
        },
        // --- şərti format
        cf(ck, wb, state) {
            const cells = SE.cellsIn(ck.r);
            const hl = cells.filter((a) => SE.cfStyleFor(wb, a));
            if (hl.length !== ck.hl.length || !ck.hl.every((a) => hl.includes(a))) return false;
            if (ck.fills) {
                if (!hl.every((a) => ck.fills.includes(SE.cfStyleFor(wb, a).id))) return false;
            }
            if (ck.kind && !(state.cf || []).some((r) => SE.rangeCovers(r.range, ck.r) && r.kind === ck.kind)) return false;
            return true;
        },
        // --- pivot
        pivot(ck, wb, state) {
            const p = findPivot(state, ck.spec);
            if (!p) return false;
            if (ck.spec.filters !== undefined && !filtersEq(p, ck.spec.filters)) return false;
            return true;
        },
        pivotv(ck, wb, state) {
            const p = findPivot(state, ck.spec);
            if (!p) return false;
            if (ck.spec.filters !== undefined && !filtersEq(p, ck.spec.filters)) return false;
            const res = SE.computePivot(wb, p);
            if (!res.ok) return false;
            for (const e of ck.expect) {
                const got = e.grand ? res.grand(e.vi || 0) : res.get(e.r, e.c, e.vi || 0);
                if (!valEq(got, e.v, e.tol !== undefined ? e.tol : 1e-6)) return false;
            }
            return true;
        },
        pivotfmt(ck, wb, state) {
            const p = findPivot(state, ck.spec);
            return !!p && (p.numFmt || "general") === ck.numFmt;
        },
        // --- qrafik
        chart(ck, wb, state) { return (state.charts || []).some((ch) => chartMatches(ch, ck.spec)); },
    };

    function gradeTask(task, stateIn) {
        const state = normState(stateIn);
        const wb = new SE.Workbook(state);
        const results = [];
        let got = 0, tot = 0;
        for (const ck of task.checks) {
            const w = ck.w === undefined ? 1 : ck.w;
            let ok = false;
            try { ok = !!CHECKS[ck.t](ck, wb, state); } catch (e) { ok = false; }
            wb.invalidate();
            results.push({ l: ck.l, ok, w });
            tot += w;
            if (ok) got += w;
        }
        const max = task.pts;
        const pts = tot ? Math.round(max * got / tot) : 0;
        const full = got === tot;
        const failed = results.filter((r) => !r.ok).map((r) => r.l);
        let msg, detail = "";
        if (full) msg = `✅ Task ${task.id}: CORRECT (+${max} points)`;
        else if (pts > 0) { msg = `🟡 Task ${task.id}: PARTIALLY CORRECT (+${pts} / ${max} points)`; detail = failed.join(" · "); }
        else { msg = `❌ Task ${task.id}: NOT CORRECT YET (0 / ${max} points)`; detail = failed.join(" · "); }
        return { id: task.id, pts: full ? max : pts, max, got, tot, ok: full, partial: !full && pts > 0, msg, detail, results };
    }

    function gradeVariant(variant, answers) {
        const out = {};
        let total = 0;
        for (const t of variant.tasks) {
            let st = null;
            const a = answers ? answers[t.id] : null;
            if (a) { try { st = typeof a === "string" ? JSON.parse(a) : a; } catch (e) { st = null; } }
            if (!st) st = t.setup || blankState();
            const r = gradeTask(t, st);
            out[t.id] = r;
            total += r.pts;
        }
        return { total, results: out };
    }

    return { gradeTask, gradeVariant, blankState, normState, findPivot, chartMatches };
});
