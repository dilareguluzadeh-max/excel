/* ====================================================================
 * SDF 1 — Cədvəl mühərriki (Excel / Google Sheets düsturları)
 * Brauzerdə (window.SheetEngine) və Node-da (require) eyni kod işləyir.
 * Asılılıq yoxdur. Düstur parseri, ~70 funksiya, doldurma (relativ istinad sürüşməsi),
 * məlumat doğrulaması, şərti format və pivot cədvəl hesablayıcısı.
 * ==================================================================== */
(function (root, factory) {
    if (typeof module === "object" && module.exports) module.exports = factory();
    else root.SheetEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
    "use strict";

    // ------------------------------------------------------------------
    // Xətalar
    // ------------------------------------------------------------------
    class XlError { constructor(code) { this.code = code; } toString() { return this.code; } }
    const E = {
        DIV0: new XlError("#DIV/0!"), VALUE: new XlError("#VALUE!"), REF: new XlError("#REF!"),
        NAME: new XlError("#NAME?"), NA: new XlError("#N/A"), NUM: new XlError("#NUM!"),
        CIRC: new XlError("#CIRC!"), PARSE: new XlError("#ERROR!"),
    };
    const isErr = (v) => v instanceof XlError;
    const MAX_ROWS = 200;

    // ------------------------------------------------------------------
    // Ünvanlar
    // ------------------------------------------------------------------
    function colToNum(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
    function numToCol(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
    function addr(c, r) { return numToCol(c) + r; }
    function parseAddr(a) {
        const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,5})$/.exec(String(a).trim());
        return m ? { c: colToNum(m[1]), r: +m[2] } : null;
    }
    function parseRange(s) {
        const parts = String(s).replace(/\$/g, "").split(":");
        const p1 = parseAddr(parts[0]), p2 = parseAddr(parts[1] || parts[0]);
        if (!p1 || !p2) return null;
        return { c1: Math.min(p1.c, p2.c), r1: Math.min(p1.r, p2.r), c2: Math.max(p1.c, p2.c), r2: Math.max(p1.r, p2.r) };
    }
    function cellsIn(rangeStr) {
        const g = parseRange(rangeStr); const out = [];
        if (!g) return out;
        for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push(addr(c, r));
        return out;
    }
    function rangeCovers(outer, inner) {
        const o = parseRange(outer), i = parseRange(inner);
        return !!(o && i && o.c1 <= i.c1 && o.r1 <= i.r1 && o.c2 >= i.c2 && o.r2 >= i.r2);
    }
    function inRange(rangeStr, a) {
        const g = parseRange(rangeStr), p = parseAddr(a);
        return !!(g && p && p.c >= g.c1 && p.c <= g.c2 && p.r >= g.r1 && p.r <= g.r2);
    }

    // ------------------------------------------------------------------
    // Tokenizer
    // ------------------------------------------------------------------
    const RE_REF = /^(?:(?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,5})(?::(\$?[A-Za-z]{1,3}\$?\d{1,5}))?(?![\w(])/;
    const RE_COLS = /^(?:(?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w(])/;
    const RE_IDENT = /^[A-Za-z_][A-Za-z0-9_.]*/;
    const RE_NUM = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/;

    function tokenize(src) {
        const toks = [];
        let i = 0;
        while (i < src.length) {
            const ch = src[i];
            if (/\s/.test(ch)) { i++; continue; }
            if (ch === '"') {
                let j = i + 1, s = "";
                for (;;) {
                    if (j >= src.length) throw E.PARSE;
                    if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
                    s += src[j++];
                }
                toks.push({ t: "str", v: s, s: i, e: j + 1 }); i = j + 1; continue;
            }
            const rest = src.slice(i);
            let m;
            if ((m = RE_NUM.exec(rest)) && !/^[A-Za-z_$]/.test(rest)) {
                toks.push({ t: "num", v: parseFloat(m[0]), s: i, e: i + m[0].length }); i += m[0].length; continue;
            }
            const id = RE_IDENT.exec(rest);
            if (id && rest[id[0].length] === "(") {
                toks.push({ t: "func", v: id[0].toUpperCase(), s: i, e: i + id[0].length }); i += id[0].length; continue;
            }
            if ((m = RE_REF.exec(rest))) {
                const a = parseRefText(m[1]), b = m[2] ? parseRefText(m[2]) : null;
                toks.push(b ? { t: "range", a, b, s: i, e: i + m[0].length } : { t: "ref", a, s: i, e: i + m[0].length });
                i += m[0].length; continue;
            }
            if ((m = RE_COLS.exec(rest))) {
                const a = parseRefText(m[1] + "1"), b = parseRefText(m[2] + MAX_ROWS);
                toks.push({ t: "range", a, b, s: i, e: i + m[0].length, whole: true }); i += m[0].length; continue;
            }
            if (id) {
                const up = id[0].toUpperCase();
                if (up === "TRUE" || up === "FALSE") toks.push({ t: "bool", v: up === "TRUE", s: i, e: i + id[0].length });
                else toks.push({ t: "name", v: id[0], s: i, e: i + id[0].length });
                i += id[0].length; continue;
            }
            const two = src.substr(i, 2);
            if (two === "<=" || two === ">=" || two === "<>") { toks.push({ t: "op", v: two, s: i, e: i + 2 }); i += 2; continue; }
            if ("+-*/^&=<>%".includes(ch)) { toks.push({ t: "op", v: ch, s: i, e: i + 1 }); i++; continue; }
            if (ch === "(") { toks.push({ t: "lp", s: i, e: i + 1 }); i++; continue; }
            if (ch === ")") { toks.push({ t: "rp", s: i, e: i + 1 }); i++; continue; }
            if (ch === "," || ch === ";") { toks.push({ t: "comma", s: i, e: i + 1 }); i++; continue; }
            throw E.PARSE;
        }
        return toks;
    }

    function parseRefText(t) {
        const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,5})$/.exec(t);
        return { ac: !!m[1], c: colToNum(m[2]), ar: !!m[3], r: +m[4] };
    }
    function refText(p) { return (p.ac ? "$" : "") + numToCol(p.c) + (p.ar ? "$" : "") + p.r; }

    // ------------------------------------------------------------------
    // Parser (rekursiv enmə)
    // ------------------------------------------------------------------
    function parseFormula(src) {
        const toks = tokenize(src);
        let p = 0;
        const peek = () => toks[p];
        const next = () => toks[p++];
        const isOp = (t, ...vs) => t && t.t === "op" && vs.includes(t.v);

        function cmp() {
            let l = concat();
            while (isOp(peek(), "=", "<>", "<", ">", "<=", ">=")) { const op = next().v; l = { t: "bin", op, l, r: concat() }; }
            return l;
        }
        function concat() {
            let l = add();
            while (isOp(peek(), "&")) { next(); l = { t: "bin", op: "&", l, r: add() }; }
            return l;
        }
        function add() {
            let l = mul();
            while (isOp(peek(), "+", "-")) { const op = next().v; l = { t: "bin", op, l, r: mul() }; }
            return l;
        }
        function mul() {
            let l = pow();
            while (isOp(peek(), "*", "/")) { const op = next().v; l = { t: "bin", op, l, r: pow() }; }
            return l;
        }
        function pow() {
            let l = unary();
            while (isOp(peek(), "^")) { next(); l = { t: "bin", op: "^", l, r: unary() }; }
            return l;
        }
        function unary() {
            if (isOp(peek(), "-", "+")) { const op = next().v; const e = unary(); return op === "-" ? { t: "un", e } : e; }
            return postfix();
        }
        function postfix() {
            let e = primary();
            while (isOp(peek(), "%")) { next(); e = { t: "pct", e }; }
            return e;
        }
        function primary() {
            const t = next();
            if (!t) throw E.PARSE;
            switch (t.t) {
                case "num": return { t: "num", v: t.v };
                case "str": return { t: "str", v: t.v };
                case "bool": return { t: "bool", v: t.v };
                case "ref": return { t: "ref", c: t.a.c, r: t.a.r };
                case "range": return { t: "range", c1: Math.min(t.a.c, t.b.c), r1: Math.min(t.a.r, t.b.r), c2: Math.max(t.a.c, t.b.c), r2: Math.max(t.a.r, t.b.r) };
                case "name": return { t: "name", v: t.v };
                case "lp": { const e = cmp(); const c = next(); if (!c || c.t !== "rp") throw E.PARSE; return e; }
                case "func": {
                    const lp = next();
                    if (!lp || lp.t !== "lp") throw E.PARSE;
                    const args = [];
                    if (peek() && peek().t === "rp") { next(); return { t: "call", name: t.v, args }; }
                    for (;;) {
                        if (peek() && (peek().t === "comma" || peek().t === "rp")) args.push({ t: "empty" });
                        else args.push(cmp());
                        const n = next();
                        if (!n) throw E.PARSE;
                        if (n.t === "rp") break;
                        if (n.t !== "comma") throw E.PARSE;
                    }
                    return { t: "call", name: t.v, args };
                }
                default: throw E.PARSE;
            }
        }
        const ast = cmp();
        if (p < toks.length) throw E.PARSE;
        return ast;
    }

    // Düstur mətnində relativ istinadları (dRow, dCol) qədər sürüşdürür (doldurma / kopyala-yapışdır)
    function shiftFormula(f, dRow, dCol) {
        if (typeof f !== "string" || f[0] !== "=") return f;
        let toks;
        try { toks = tokenize(f.slice(1)); } catch (e) { return f; }
        let out = "=", pos = 0;
        const body = f.slice(1);
        const sh = (p) => {
            const q = { ac: p.ac, ar: p.ar, c: p.ac ? p.c : p.c + dCol, r: p.ar ? p.r : p.r + dRow };
            return (q.c < 1 || q.r < 1) ? null : q;
        };
        for (const t of toks) {
            out += body.slice(pos, t.s);
            pos = t.e;
            const orig = body.slice(t.s, t.e);
            if (t.t === "ref" || t.t === "range") {
                const prefix = /^((?:'[^']+'|[A-Za-z_][\w.]*)!)/.exec(orig);
                const pre = prefix ? prefix[1] : "";
                const a = sh(t.a), b = t.b ? sh(t.b) : null;
                if (!a || (t.b && !b)) out += "#REF!";
                else out += pre + refText(a) + (b ? ":" + refText(b) : "");
            } else out += orig;
        }
        out += body.slice(pos);
        return out;
    }

    // ------------------------------------------------------------------
    // Dəyərlər: çevirmələr
    // ------------------------------------------------------------------
    function parseConst(raw) {
        if (raw === null || raw === undefined) return null;
        let s = String(raw);
        if (s[0] === "'") return s.slice(1);
        const t = s.trim();
        if (t === "") return null;
        if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) return parseFloat(t);
        if (/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(t)) return parseFloat(t) / 100;
        if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return parseFloat(t.replace(/,/g, ""));
        if (/^[+-]?\$\s?\d{1,3}(,\d{3})*(\.\d+)?$|^[+-]?\$\s?\d+(\.\d+)?$/.test(t)) return parseFloat(t.replace(/[$,\s]/g, ""));
        if (/^TRUE$/i.test(t)) return true;
        if (/^FALSE$/i.test(t)) return false;
        if (/^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!)$/.test(t)) return new XlError(t);
        return s;
    }

    function numToStr(n) {
        if (Number.isInteger(n)) return String(n);
        return String(+n.toPrecision(15));
    }
    function toStr(v) {
        if (v === null || v === undefined) return "";
        if (typeof v === "number") return numToStr(v);
        if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
        return String(v);
    }
    function toNum(v) {
        if (typeof v === "number") return v;
        if (v === null || v === undefined) return 0;
        if (typeof v === "boolean") return v ? 1 : 0;
        if (typeof v === "string") {
            const t = v.trim();
            const c = parseConst(t);
            if (typeof c === "number") return c;
            throw E.VALUE;
        }
        throw E.VALUE;
    }
    function toBool(v) {
        if (typeof v === "boolean") return v;
        if (typeof v === "number") return v !== 0;
        if (v === null || v === undefined) return false;
        if (typeof v === "string") {
            if (/^true$/i.test(v)) return true;
            if (/^false$/i.test(v)) return false;
        }
        throw E.VALUE;
    }
    const norm15 = (x) => (Number.isFinite(x) ? parseFloat(x.toPrecision(15)) : x);

    function typeRank(v) { return typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3; }
    function compare(a, b) {
        if (a === null || a === undefined) a = (typeof b === "string") ? "" : (typeof b === "boolean" ? false : 0);
        if (b === null || b === undefined) b = (typeof a === "string") ? "" : (typeof a === "boolean" ? false : 0);
        const ra = typeRank(a), rb = typeRank(b);
        if (ra !== rb) return ra < rb ? -1 : 1;
        if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
        if (ra === 0) { a = norm15(a); b = norm15(b); }
        return a < b ? -1 : a > b ? 1 : 0;
    }
    function looseEq(a, b) {
        if (a === null || a === undefined || b === null || b === undefined) return false;
        if (typeof a !== typeof b) return false;
        return compare(a, b) === 0;
    }

    // Kriteriya (SUMIF, COUNTIF ...)
    function wildcardRegex(p, anchored = true) {
        let re = "";
        for (let i = 0; i < p.length; i++) {
            const ch = p[i];
            if (ch === "~" && i + 1 < p.length) { re += p[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); continue; }
            if (ch === "*") re += ".*"; else if (ch === "?") re += ".";
            else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        }
        return new RegExp(anchored ? "^" + re + "$" : re, "i");
    }
    function matcher(crit) {
        if (isErr(crit)) throw crit;
        let op = "=", operand = crit;
        if (typeof crit === "string") {
            const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(crit);
            op = m[1] || "="; operand = m[2];
            const pc = parseConst(operand);
            if (typeof pc === "number" || typeof pc === "boolean") operand = pc;
        } else if (crit === null || crit === undefined) { operand = 0; }
        return (v) => {
            if (isErr(v)) return false;
            if (typeof operand === "number") {
                if (op === "=" || op === "<>") {
                    const eq = typeof v === "number" && norm15(v) === norm15(operand);
                    return op === "=" ? eq : !eq;
                }
                if (typeof v !== "number") return false;
                const a = norm15(v), b = norm15(operand);
                return op === "<" ? a < b : op === ">" ? a > b : op === "<=" ? a <= b : a >= b;
            }
            if (typeof operand === "boolean") {
                const eq = typeof v === "boolean" && v === operand;
                return op === "<>" ? !eq : eq;
            }
            // mətn
            if (op === "=" || op === "<>") {
                let eq;
                if (operand === "") eq = v === null || v === undefined || v === "";
                else eq = typeof v === "string" && wildcardRegex(operand).test(v);
                return op === "=" ? eq : !eq;
            }
            if (typeof v !== "string") return false;
            const c = compare(v, operand);
            return op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
        };
    }

    // ------------------------------------------------------------------
    // Diapazon obyekti
    // ------------------------------------------------------------------
    class Rng {
        constructor(wb, c1, r1, c2, r2, single) { this.wb = wb; this.c1 = c1; this.r1 = r1; this.c2 = c2; this.r2 = r2; this.single = !!single; }
        get rows() { return this.r2 - this.r1 + 1; }
        get cols() { return this.c2 - this.c1 + 1; }
        at(i, j) { return this.wb.value(addr(this.c1 + j, this.r1 + i)); }
        values() { const out = []; for (let i = 0; i < this.rows; i++) for (let j = 0; j < this.cols; j++) out.push(this.at(i, j)); return out; }
    }

    // ------------------------------------------------------------------
    // Funksiyalar
    // ------------------------------------------------------------------
    const FUNCS = Object.create(null);
    const LAZY = Object.create(null);
    const def = (names, fn) => { for (const n of names.split(" ")) FUNCS[n] = fn; };

    const S = (x) => {                          // skalyara çevir
        if (x instanceof Rng) {
            if (x.rows === 1 && x.cols === 1) { const v = x.at(0, 0); if (isErr(v)) throw v; return v; }
            throw E.VALUE;
        }
        if (isErr(x)) throw x;
        return x;
    };
    const N = (x) => toNum(S(x));
    const T = (x) => toStr(S(x));
    // Ədədləri toplayır: diapazonda yalnız ədədlər, birbaşa arqumentdə çevirmə
    function collectNums(args, opts = {}) {
        const out = [];
        for (const a of args) {
            if (a instanceof Rng) {
                for (const v of a.values()) {
                    if (isErr(v)) throw v;
                    if (typeof v === "number") out.push(v);
                    else if (opts.boolsAsNums && typeof v === "boolean") out.push(v ? 1 : 0);
                }
            } else {
                if (a && a.t === "empty") continue;
                if (isErr(a)) throw a;
                if (a === null || a === undefined) continue;
                if (typeof a === "string" && opts.skipText) continue;
                out.push(toNum(a));
            }
        }
        return out;
    }
    const has = (args, i) => args.length > i && args[i] !== null && args[i] !== undefined;
    const need = (args, min, max) => { if (args.length < min || args.length > (max ?? min)) throw E.VALUE; };
    const flat = (x) => (x instanceof Rng ? x.values() : [S(x)]);
    const sum = (a) => a.reduce((x, y) => x + y, 0);

    def("SUM", (a) => sum(collectNums(a)));
    def("AVERAGE", (a) => { const n = collectNums(a); if (!n.length) throw E.DIV0; return sum(n) / n.length; });
    def("MIN", (a) => { const n = collectNums(a); return n.length ? Math.min(...n) : 0; });
    def("MAX", (a) => { const n = collectNums(a); return n.length ? Math.max(...n) : 0; });
    def("COUNT", (a) => {
        let c = 0;
        for (const x of a) {
            if (x instanceof Rng) c += x.values().filter((v) => typeof v === "number").length;
            else if (!isErr(x) && x !== null && !(x && x.t === "empty")) { try { toNum(x); c++; } catch (e) { /* mətn */ } }
        }
        return c;
    });
    def("COUNTA", (a) => {
        let c = 0;
        for (const x of a) {
            if (x instanceof Rng) c += x.values().filter((v) => v !== null && v !== undefined && v !== "").length;
            else c++;
        }
        return c;
    });
    def("COUNTBLANK", (a) => { need(a, 1); return flat(a[0]).filter((v) => v === null || v === undefined || v === "").length; });
    def("MEDIAN", (a) => {
        const n = collectNums(a).sort((x, y) => x - y);
        if (!n.length) throw E.NUM;
        const m = n.length >> 1;
        return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
    });
    def("MODE.SNGL MODE", (a) => {
        const n = collectNums(a), cnt = new Map();
        for (const v of n) cnt.set(v, (cnt.get(v) || 0) + 1);
        let best = null, bc = 1;
        for (const [v, c] of cnt) if (c > bc) { best = v; bc = c; }
        if (best === null) throw E.NA;
        return best;
    });
    const variance = (n, sample) => {
        const k = n.length - (sample ? 1 : 0);
        if (k <= 0) throw E.DIV0;
        const m = sum(n) / n.length;
        return sum(n.map((x) => (x - m) ** 2)) / k;
    };
    def("VAR.S VAR", (a) => variance(collectNums(a), true));
    def("VAR.P VARP", (a) => variance(collectNums(a), false));
    def("STDEV.S STDEV", (a) => Math.sqrt(variance(collectNums(a), true)));
    def("STDEV.P STDEVP", (a) => Math.sqrt(variance(collectNums(a), false)));
    const percentile = (n, k) => {
        n = n.slice().sort((x, y) => x - y);
        if (!n.length || k < 0 || k > 1) throw E.NUM;
        const pos = k * (n.length - 1), lo = Math.floor(pos), hi = Math.ceil(pos);
        return n[lo] + (pos - lo) * (n[hi] - n[lo]);
    };
    def("PERCENTILE PERCENTILE.INC", (a) => { need(a, 2); return percentile(collectNums([a[0]]), N(a[1])); });
    def("QUARTILE QUARTILE.INC", (a) => {
        need(a, 2);
        const q = Math.trunc(N(a[1]));
        if (q < 0 || q > 4) throw E.NUM;
        return percentile(collectNums([a[0]]), q / 4);
    });
    def("LARGE", (a) => { need(a, 2); const n = collectNums([a[0]]).sort((x, y) => y - x); const k = N(a[1]); if (k < 1 || k > n.length) throw E.NUM; return n[k - 1]; });
    def("SMALL", (a) => { need(a, 2); const n = collectNums([a[0]]).sort((x, y) => x - y); const k = N(a[1]); if (k < 1 || k > n.length) throw E.NUM; return n[k - 1]; });

    function pairs(ys, xs) {
        const Y = flat(ys), X = flat(xs);
        if (Y.length !== X.length) throw E.NA;
        const out = [];
        for (let i = 0; i < Y.length; i++) if (typeof Y[i] === "number" && typeof X[i] === "number") out.push([X[i], Y[i]]);
        if (out.length < 2) throw E.DIV0;
        return out;
    }
    function regress(ys, xs) {
        const p = pairs(ys, xs), n = p.length;
        const mx = sum(p.map((q) => q[0])) / n, my = sum(p.map((q) => q[1])) / n;
        const sxx = sum(p.map((q) => (q[0] - mx) ** 2)), sxy = sum(p.map((q) => (q[0] - mx) * (q[1] - my))), syy = sum(p.map((q) => (q[1] - my) ** 2));
        if (sxx === 0) throw E.DIV0;
        return { slope: sxy / sxx, intercept: my - (sxy / sxx) * mx, r2: syy === 0 ? 1 : (sxy * sxy) / (sxx * syy), r: sxy / Math.sqrt(sxx * syy) };
    }
    def("SLOPE", (a) => { need(a, 2); return regress(a[0], a[1]).slope; });
    def("INTERCEPT", (a) => { need(a, 2); return regress(a[0], a[1]).intercept; });
    def("RSQ", (a) => { need(a, 2); return regress(a[0], a[1]).r2; });
    def("CORREL", (a) => { need(a, 2); return regress(a[0], a[1]).r; });
    def("FORECAST FORECAST.LINEAR", (a) => { need(a, 3); const r = regress(a[1], a[2]); return r.intercept + r.slope * N(a[0]); });

    // Riyaziyyat
    def("ABS", (a) => { need(a, 1); return Math.abs(N(a[0])); });
    def("SQRT", (a) => { need(a, 1); const x = N(a[0]); if (x < 0) throw E.NUM; return Math.sqrt(x); });
    def("POWER", (a) => { need(a, 2); return N(a[0]) ** N(a[1]); });
    def("EXP", (a) => { need(a, 1); return Math.exp(N(a[0])); });
    def("LN", (a) => { need(a, 1); const x = N(a[0]); if (x <= 0) throw E.NUM; return Math.log(x); });
    def("LOG10", (a) => { need(a, 1); const x = N(a[0]); if (x <= 0) throw E.NUM; return Math.log10(x); });
    def("PI", () => Math.PI);
    def("INT", (a) => { need(a, 1); return Math.floor(norm15(N(a[0]))); });
    def("MOD", (a) => { need(a, 2); const x = N(a[0]), y = N(a[1]); if (y === 0) throw E.DIV0; return x - y * Math.floor(x / y); });
    const roundTo = (x, d, mode) => {
        x = norm15(x);
        const f = 10 ** Math.abs(d), sgn = x < 0 ? -1 : 1, ax = Math.abs(x);
        const sc = d >= 0 ? norm15(ax * f) : norm15(ax / f);
        const r = mode === "down" ? Math.floor(sc) : mode === "up" ? Math.ceil(sc) : Math.floor(sc + 0.5);
        return sgn * (d >= 0 ? r / f : r * f);
    };
    def("ROUND", (a) => { need(a, 2); return roundTo(N(a[0]), Math.trunc(N(a[1])), "half"); });
    def("ROUNDDOWN", (a) => { need(a, 2); return roundTo(N(a[0]), Math.trunc(N(a[1])), "down"); });
    def("ROUNDUP", (a) => { need(a, 2); return roundTo(N(a[0]), Math.trunc(N(a[1])), "up"); });
    def("PMT", (a) => {
        need(a, 3, 5);
        const r = N(a[0]), n = N(a[1]), pv = N(a[2]), fv = has(a, 3) ? N(a[3]) : 0, ty = has(a, 4) ? N(a[4]) : 0;
        if (n === 0) throw E.NUM;
        if (r === 0) return -(pv + fv) / n;
        const g = (1 + r) ** n;
        return -(r * (pv * g + fv)) / ((1 + r * ty) * (g - 1));
    });
    def("FV", (a) => {
        need(a, 3, 5);
        const r = N(a[0]), n = N(a[1]), pmt = N(a[2]), pv = has(a, 3) ? N(a[3]) : 0, ty = has(a, 4) ? N(a[4]) : 0;
        if (r === 0) return -(pv + pmt * n);
        const g = (1 + r) ** n;
        return -(pv * g + pmt * (1 + r * ty) * (g - 1) / r);
    });
    def("SUMPRODUCT", (a) => {
        if (!a.length) throw E.VALUE;
        const arrs = a.map((x) => flat(x));
        const len = arrs[0].length;
        if (arrs.some((x) => x.length !== len)) throw E.VALUE;
        let t = 0;
        for (let i = 0; i < len; i++) { let p = 1; for (const arr of arrs) p *= typeof arr[i] === "number" ? arr[i] : 0; t += p; }
        return t;
    });

    // Mətn
    def("LEFT", (a) => { need(a, 1, 2); const n = has(a, 1) ? N(a[1]) : 1; if (n < 0) throw E.VALUE; return T(a[0]).slice(0, n); });
    def("RIGHT", (a) => { need(a, 1, 2); const n = has(a, 1) ? N(a[1]) : 1; if (n < 0) throw E.VALUE; const s = T(a[0]); return n === 0 ? "" : s.slice(-n); });
    def("MID", (a) => { need(a, 3); const s = T(a[0]), st = Math.trunc(N(a[1])), n = Math.trunc(N(a[2])); if (st < 1 || n < 0) throw E.VALUE; return s.substr(st - 1, n); });
    def("LEN", (a) => { need(a, 1); return T(a[0]).length; });
    def("UPPER", (a) => { need(a, 1); return T(a[0]).toUpperCase(); });
    def("LOWER", (a) => { need(a, 1); return T(a[0]).toLowerCase(); });
    def("PROPER", (a) => { need(a, 1); return T(a[0]).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, p, c) => p + c.toUpperCase()); });
    def("TRIM", (a) => { need(a, 1); return T(a[0]).replace(/ +/g, " ").replace(/^ | $/g, ""); });
    def("CLEAN", (a) => { need(a, 1); return T(a[0]).replace(/[\x00-\x1f]/g, ""); });
    def("CONCATENATE CONCAT", (a) => a.map((x) => flat(x).map(toStr).join("")).join(""));
    def("SUBSTITUTE", (a) => {
        need(a, 3, 4);
        const s = T(a[0]), o = T(a[1]), n = T(a[2]);
        if (o === "") return s;
        if (has(a, 3)) {
            const k = Math.trunc(N(a[3])); let idx = -1, cnt = 0, from = 0;
            while ((idx = s.indexOf(o, from)) !== -1) { cnt++; if (cnt === k) return s.slice(0, idx) + n + s.slice(idx + o.length); from = idx + o.length; }
            return s;
        }
        return s.split(o).join(n);
    });
    def("FIND", (a) => {
        need(a, 2, 3);
        const f = T(a[0]), w = T(a[1]), st = has(a, 2) ? Math.trunc(N(a[2])) : 1;
        if (st < 1 || st > w.length + 1) throw E.VALUE;
        const i = w.indexOf(f, st - 1);
        if (i < 0) throw E.VALUE;
        return i + 1;
    });
    def("SEARCH", (a) => {
        need(a, 2, 3);
        const f = T(a[0]), w = T(a[1]), st = has(a, 2) ? Math.trunc(N(a[2])) : 1;
        if (st < 1 || st > w.length + 1) throw E.VALUE;
        const m = wildcardRegex(f, false).exec(w.slice(st - 1));
        if (!m) throw E.VALUE;
        return st + m.index;
    });
    def("VALUE", (a) => {
        need(a, 1);
        const v = S(a[0]);
        if (typeof v === "number") return v;
        const c = parseConst(typeof v === "string" ? v.trim() : v);
        if (typeof c === "number") return c;
        throw E.VALUE;
    });
    def("TEXT", (a) => {
        need(a, 2);
        const v = S(a[0]), f = T(a[1]);
        if (typeof v !== "number") return toStr(v);
        const m = /^(\$)?#?,?#*0?(?:\.(0+))?(%)?$/.exec(f);
        if (!m) return numToStr(v);
        const d = m[2] ? m[2].length : 0;
        let x = m[3] ? v * 100 : v;
        let s = Math.abs(x).toFixed(d);
        if (f.includes(",")) s = s.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        return (x < 0 ? "-" : "") + (m[1] || "") + s + (m[3] ? "%" : "");
    });

    // Məntiq
    def("AND", (a) => { let r = true, any = false; for (const x of a) for (const v of flat(x)) { if (v === null || typeof v === "string" && x instanceof Rng) continue; any = true; if (!toBool(v)) r = false; } if (!any) throw E.VALUE; return r; });
    def("OR", (a) => { let r = false, any = false; for (const x of a) for (const v of flat(x)) { if (v === null || typeof v === "string" && x instanceof Rng) continue; any = true; if (toBool(v)) r = true; } if (!any) throw E.VALUE; return r; });
    def("NOT", (a) => { need(a, 1); return !toBool(S(a[0])); });
    def("TRUE", () => true);
    def("FALSE", () => false);
    def("ISNUMBER", (a) => { need(a, 1); try { return typeof S(a[0]) === "number"; } catch (e) { return false; } });
    def("ISTEXT", (a) => { need(a, 1); try { return typeof S(a[0]) === "string"; } catch (e) { return false; } });
    def("ISBLANK", (a) => { need(a, 1); try { const v = S(a[0]); return v === null || v === undefined; } catch (e) { return false; } });
    LAZY.IF = (args, ev) => {
        if (args.length < 2 || args.length > 3) throw E.VALUE;
        const c = toBool(ev(args[0]));
        if (c) return ev(args[1]);
        return args[2] ? ev(args[2]) : false;
    };
    LAZY.IFERROR = (args, ev) => {
        if (args.length !== 2) throw E.VALUE;
        try { const v = ev(args[0]); if (isErr(v)) return ev(args[1]); return v; } catch (e) { if (e instanceof XlError) return ev(args[1]); throw e; }
    };
    LAZY.IFNA = (args, ev) => {
        if (args.length !== 2) throw E.VALUE;
        try { const v = ev(args[0]); if (v === E.NA) return ev(args[1]); return v; } catch (e) { if (e === E.NA) return ev(args[1]); throw e; }
    };
    LAZY.ISERROR = (args, ev) => { try { return isErr(ev(args[0])); } catch (e) { if (e instanceof XlError) return true; throw e; } };
    LAZY.ISNA = (args, ev) => { try { return ev(args[0]) === E.NA; } catch (e) { if (e instanceof XlError) return e === E.NA; throw e; } };

    // Şərti toplama / sayma
    function critPairs(args, startIdx) {
        const out = [];
        for (let i = startIdx; i < args.length; i += 2) {
            if (!(args[i] instanceof Rng) || i + 1 >= args.length) throw E.VALUE;
            out.push({ vals: args[i].values(), test: matcher(S(args[i + 1])), n: args[i].values().length });
        }
        const len = out.length ? out[0].n : 0;
        if (out.some((p) => p.n !== len)) throw E.VALUE;
        return { pairs: out, len };
    }
    const hits = (cp) => { const idx = []; for (let i = 0; i < cp.len; i++) if (cp.pairs.every((p) => p.test(p.vals[i]))) idx.push(i); return idx; };
    def("COUNTIF", (a) => { need(a, 2); return hits(critPairs(a, 0)).length; });
    def("COUNTIFS", (a) => { if (a.length < 2 || a.length % 2) throw E.VALUE; return hits(critPairs(a, 0)).length; });
    def("SUMIF", (a) => {
        need(a, 2, 3);
        if (!(a[0] instanceof Rng)) throw E.VALUE;
        const cp = critPairs([a[0], a[1]], 0), sv = (a[2] instanceof Rng ? a[2] : a[0]).values();
        return sum(hits(cp).map((i) => (typeof sv[i] === "number" ? sv[i] : 0)));
    });
    def("SUMIFS", (a) => {
        if (a.length < 3 || a.length % 2 === 0 || !(a[0] instanceof Rng)) throw E.VALUE;
        const sv = a[0].values(), cp = critPairs(a, 1);
        if (cp.len !== sv.length) throw E.VALUE;
        return sum(hits(cp).map((i) => (typeof sv[i] === "number" ? sv[i] : 0)));
    });
    def("AVERAGEIF", (a) => {
        need(a, 2, 3);
        if (!(a[0] instanceof Rng)) throw E.VALUE;
        const cp = critPairs([a[0], a[1]], 0), sv = (a[2] instanceof Rng ? a[2] : a[0]).values();
        const v = hits(cp).map((i) => sv[i]).filter((x) => typeof x === "number");
        if (!v.length) throw E.DIV0;
        return sum(v) / v.length;
    });
    def("AVERAGEIFS", (a) => {
        if (a.length < 3 || a.length % 2 === 0 || !(a[0] instanceof Rng)) throw E.VALUE;
        const sv = a[0].values(), cp = critPairs(a, 1);
        const v = hits(cp).map((i) => sv[i]).filter((x) => typeof x === "number");
        if (!v.length) throw E.DIV0;
        return sum(v) / v.length;
    });

    // Axtarış
    def("VLOOKUP", (a) => {
        need(a, 3, 4);
        const key = S(a[0]);
        if (!(a[1] instanceof Rng)) throw E.VALUE;
        const t = a[1], ci = Math.trunc(N(a[2]));
        if (ci < 1) throw E.VALUE;
        if (ci > t.cols) throw E.REF;
        const exact = has(a, 3) ? !toBool(S(a[3])) : false;
        if (exact) {
            for (let i = 0; i < t.rows; i++) if (looseEq(typeof key === "string" ? t.at(i, 0) : t.at(i, 0), key) || (typeof key === "string" && typeof t.at(i, 0) === "string" && wildcardRegex(key).test(t.at(i, 0)))) { const v = t.at(i, ci - 1); if (isErr(v)) throw v; return v === null ? 0 : v; }
            throw E.NA;
        }
        let found = -1;
        for (let i = 0; i < t.rows; i++) { const v = t.at(i, 0); if (v === null || typeof v !== typeof key) continue; if (compare(v, key) <= 0) found = i; else break; }
        if (found < 0) throw E.NA;
        const v = t.at(found, ci - 1); if (isErr(v)) throw v; return v === null ? 0 : v;
    });
    def("HLOOKUP", (a) => {
        need(a, 3, 4);
        const key = S(a[0]);
        if (!(a[1] instanceof Rng)) throw E.VALUE;
        const t = a[1], ri = Math.trunc(N(a[2]));
        if (ri < 1) throw E.VALUE;
        if (ri > t.rows) throw E.REF;
        const exact = has(a, 3) ? !toBool(S(a[3])) : false;
        for (let j = 0; j < t.cols; j++) if (exact ? looseEq(t.at(0, j), key) : false) { const v = t.at(ri - 1, j); if (isErr(v)) throw v; return v === null ? 0 : v; }
        throw E.NA;
    });
    function matchIndex(key, arr, type) {
        const v = arr.values();
        if (type === 0) {
            for (let i = 0; i < v.length; i++) {
                if (looseEq(v[i], key) || (typeof key === "string" && typeof v[i] === "string" && wildcardRegex(key).test(v[i]))) return i + 1;
            }
            throw E.NA;
        }
        let found = -1;
        for (let i = 0; i < v.length; i++) {
            if (v[i] === null || typeof v[i] !== typeof key) continue;
            const c = compare(v[i], key);
            if (type > 0 ? c <= 0 : c >= 0) found = i; else break;
        }
        if (found < 0) throw E.NA;
        return found + 1;
    }
    def("MATCH", (a) => {
        need(a, 2, 3);
        if (!(a[1] instanceof Rng) || (a[1].rows > 1 && a[1].cols > 1)) throw E.NA;
        return matchIndex(S(a[0]), a[1], has(a, 2) ? Math.sign(N(a[2])) : 1);
    });
    def("INDEX", (a) => {
        need(a, 2, 3);
        const t = a[0];
        if (!(t instanceof Rng)) throw E.VALUE;
        let r = Math.trunc(N(a[1])), c = has(a, 2) ? Math.trunc(N(a[2])) : null;
        if (c === null) { if (t.rows === 1) { c = r; r = 1; } else c = 1; }
        if (r < 1 || c < 1 || r > t.rows || c > t.cols) throw E.REF;
        const v = t.at(r - 1, c - 1);
        if (isErr(v)) throw v;
        return v === null ? 0 : v;
    });
    def("XLOOKUP", (a) => {
        if (a.length < 3 || a.length > 6) throw E.VALUE;
        const key = S(a[0]);
        if (!(a[1] instanceof Rng) || !(a[2] instanceof Rng)) throw E.VALUE;
        const la = a[1].values(), ra = a[2].values();
        if (la.length !== ra.length) throw E.VALUE;
        const mode = has(a, 4) ? Math.trunc(N(a[4])) : 0;
        let idx = -1;
        if (mode === 0 || mode === 2) {
            for (let i = 0; i < la.length; i++) {
                if (looseEq(la[i], key) || (mode === 2 && typeof key === "string" && typeof la[i] === "string" && wildcardRegex(key).test(la[i]))) { idx = i; break; }
            }
        } else {
            let best = -1;
            for (let i = 0; i < la.length; i++) {
                if (la[i] === null || typeof la[i] !== typeof key) continue;
                const c = compare(la[i], key);
                if (c === 0) { best = i; break; }
                if (mode === -1 && c < 0 && (best < 0 || compare(la[i], la[best]) > 0)) best = i;
                if (mode === 1 && c > 0 && (best < 0 || compare(la[i], la[best]) < 0)) best = i;
            }
            idx = best;
        }
        if (idx < 0) {
            if (has(a, 3)) return S(a[3]);
            throw E.NA;
        }
        const v = ra[idx];
        if (isErr(v)) throw v;
        return v === null ? 0 : v;
    });

    // ------------------------------------------------------------------
    // İşçi kitabı
    // ------------------------------------------------------------------
    class Workbook {
        constructor(state) {
            this.state = state || { cells: {} };
            if (!this.state.cells) this.state.cells = {};
            this.cache = new Map(); this.stack = new Set(); this.asts = new Map();
        }
        get cells() { return this.state.cells; }
        raw(a) { const r = this.cells[a]; return r === undefined ? "" : r; }
        set(a, raw) {
            if (raw === "" || raw === null || raw === undefined) delete this.cells[a]; else this.cells[a] = String(raw);
            this.cache.clear();
        }
        invalidate() { this.cache.clear(); }
        parse(src) {
            let ast = this.asts.get(src);
            if (!ast) { ast = parseFormula(src); this.asts.set(src, ast); }
            return ast;
        }
        checkFormula(text) {
            try { parseFormula(text.replace(/^=/, "")); return null; } catch (e) { return "Invalid formula"; }
        }
        value(a) {
            if (this.cache.has(a)) return this.cache.get(a);
            const raw = this.cells[a];
            let v;
            if (raw === undefined || raw === null || raw === "") v = null;
            else if (raw[0] === "=") {
                if (this.stack.has(a)) return E.CIRC;
                this.stack.add(a);
                try { v = this.evalFormula(raw); if (v === null || v === undefined) v = 0; } finally { this.stack.delete(a); }
            } else v = parseConst(raw);
            this.cache.set(a, v);
            return v;
        }
        // "=..." düsturunu qiymətləndirir; xəta olarsa XlError qaytarır
        evalFormula(text) {
            try {
                const ast = this.parse(text.replace(/^=/, ""));
                let r = this.ev(ast, false);
                if (r instanceof Rng) r = S(r);
                if (isErr(r)) return r;
                return r === undefined ? null : r;
            } catch (e) {
                if (e instanceof XlError) return e;
                throw e;
            }
        }
        sc(x) { return S(x); }
        ev(n, wantRef) {
            switch (n.t) {
                case "num": case "str": case "bool": return n.v;
                case "empty": return null;
                case "ref": {
                    if (wantRef) return new Rng(this, n.c, n.r, n.c, n.r, true);
                    const v = this.value(addr(n.c, n.r)); if (isErr(v)) throw v; return v;
                }
                case "range": return new Rng(this, n.c1, n.r1, n.c2, n.r2);
                case "name": throw E.NAME;
                case "un": return -toNum(S(this.ev(n.e, false)));
                case "pct": return toNum(S(this.ev(n.e, false))) / 100;
                case "bin": {
                    const l = S(this.ev(n.l, false)), r = S(this.ev(n.r, false));
                    switch (n.op) {
                        case "+": return toNum(l) + toNum(r);
                        case "-": return toNum(l) - toNum(r);
                        case "*": return toNum(l) * toNum(r);
                        case "/": { const d = toNum(r); if (d === 0) throw E.DIV0; return toNum(l) / d; }
                        case "^": { const b = toNum(l), x = toNum(r); const v = b ** x; if (Number.isNaN(v)) throw E.NUM; return v; }
                        case "&": return toStr(l) + toStr(r);
                        case "=": return compare(l, r) === 0;
                        case "<>": return compare(l, r) !== 0;
                        case "<": return compare(l, r) < 0;
                        case ">": return compare(l, r) > 0;
                        case "<=": return compare(l, r) <= 0;
                        case ">=": return compare(l, r) >= 0;
                    }
                    throw E.PARSE;
                }
                case "call": {
                    if (LAZY[n.name]) return LAZY[n.name](n.args, (x) => { const v = this.ev(x, false); return v instanceof Rng ? S(v) : v; });
                    const fn = FUNCS[n.name];
                    if (!fn) throw E.NAME;
                    const args = n.args.map((x) => this.ev(x, true));
                    const r = fn(args);
                    if (typeof r === "number" && !Number.isFinite(r)) throw E.NUM;
                    return r;
                }
            }
            throw E.PARSE;
        }
    }

    // ------------------------------------------------------------------
    // Ədəd formatı
    // ------------------------------------------------------------------
    function groupDigits(s) { return s.replace(/\B(?=(\d{3})+(?!\d))/g, ","); }
    function formatNumber(v, fmt) {
        if (!fmt || fmt.t === "gen") {
            if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
            const s = String(+v.toPrecision(10));
            return s;
        }
        const d = fmt.d ?? 2;
        if (fmt.t === "pct") return (v * 100).toFixed(d) + "%";
        if (fmt.t === "cur") { const s = groupDigits(Math.abs(v).toFixed(d)); return (v < 0 ? "-" : "") + "$" + s; }
        if (fmt.t === "num") return (fmt.sep ? groupDigits(v.toFixed(d)) : v.toFixed(d));
        return String(v);
    }
    function formatValue(v, fmt) {
        if (v === null || v === undefined) return "";
        if (isErr(v)) return v.code;
        if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
        if (typeof v === "number") return formatNumber(v, fmt);
        return String(v);
    }

    // ------------------------------------------------------------------
    // Məlumat doğrulaması
    // ------------------------------------------------------------------
    function dvRulesFor(state, a) {
        return (state.dv || []).filter((r) => inRange(r.range, a));
    }
    function dvListItems(wb, rule) {
        let src = String(rule.source || "").trim();
        if (src[0] === "=") src = src.slice(1);
        if (/^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(src)) {
            return cellsIn(src).map((c) => wb.value(c)).filter((v) => v !== null && v !== "").map(toStr);
        }
        return src.split(",").map((s) => s.trim()).filter((s) => s !== "");
    }
    // Daxil edilən xam mətn qaydaya uyğundurmu? Qaytarır: null (uyğundur) və ya pozulan qayda
    function dvViolation(wb, a, raw) {
        const rules = dvRulesFor(wb.state, a);
        const s = String(raw ?? "");
        if (s.trim() === "" || s[0] === "=") return null;
        for (const rule of rules) {
            if (rule.type === "list") {
                const items = dvListItems(wb, rule);
                if (!items.some((it) => it.toLowerCase() === s.trim().toLowerCase())) return rule;
            } else if (rule.type === "whole" || rule.type === "decimal") {
                const v = parseConst(s);
                if (typeof v !== "number") return rule;
                if (rule.type === "whole" && !Number.isInteger(v)) return rule;
                const lo = Number(rule.min), hi = Number(rule.max), op = rule.op || "between";
                let ok = true;
                if (op === "between") ok = v >= lo && v <= hi;
                else if (op === "notbetween") ok = v < lo || v > hi;
                else if (op === "eq") ok = v === lo; else if (op === "neq") ok = v !== lo;
                else if (op === "gt") ok = v > lo; else if (op === "gte") ok = v >= lo;
                else if (op === "lt") ok = v < lo; else if (op === "lte") ok = v <= lo;
                if (!ok) return rule;
            }
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Şərti format
    // ------------------------------------------------------------------
    const CF_STYLES = {
        lightred: { name: "Light red fill with dark red text", fill: "#ffc7ce", font: "#9c0006", family: "red" },
        lightyellow: { name: "Yellow fill with dark yellow text", fill: "#ffeb9c", font: "#9c5700", family: "yellow" },
        lightgreen: { name: "Green fill with dark green text", fill: "#c6efce", font: "#006100", family: "green" },
        red: { name: "Red fill", fill: "#ff5050", font: "#ffffff", family: "red" },
        yellow: { name: "Yellow fill", fill: "#ffff00", font: "#000000", family: "yellow" },
        green: { name: "Green fill", fill: "#4fc36b", font: "#000000", family: "green" },
        orange: { name: "Orange fill", fill: "#ffb050", font: "#000000", family: "orange" },
        blue: { name: "Blue fill", fill: "#9cc2ff", font: "#000000", family: "blue" },
    };
    function cfRuleMatches(wb, rule, a) {
        const tl = parseRange(rule.range);
        const p = parseAddr(a);
        if (rule.kind === "formula") {
            let f = String(rule.formula || "").trim();
            if (f[0] !== "=") f = "=" + f;
            const shifted = shiftFormula(f, p.r - tl.r1, p.c - tl.c1);
            const v = wb.evalFormula(shifted);
            return v === true || (typeof v === "number" && v !== 0);
        }
        const v = wb.value(a);
        if (isErr(v)) return false;
        const lo = parseConst(rule.v1), hi = parseConst(rule.v2);
        const op = rule.op;
        if (typeof lo === "string" || typeof hi === "string" || typeof v === "string") {
            const x = toStr(v).toLowerCase(), y = toStr(lo).toLowerCase();
            if (op === "eq") return x === y; if (op === "neq") return x !== y;
            return false;
        }
        if (typeof v !== "number") return false;
        const x = norm15(v);
        switch (op) {
            case "gt": return x > lo; case "gte": return x >= lo; case "lt": return x < lo; case "lte": return x <= lo;
            case "eq": return x === lo; case "neq": return x !== lo;
            case "between": return x >= Math.min(lo, hi) && x <= Math.max(lo, hi);
            case "notbetween": return x < Math.min(lo, hi) || x > Math.max(lo, hi);
        }
        return false;
    }
    function cfStyleFor(wb, a) {
        for (const rule of wb.state.cf || []) {
            if (inRange(rule.range, a) && cfRuleMatches(wb, rule, a)) return Object.assign({ id: rule.fill }, CF_STYLES[rule.fill] || CF_STYLES.lightred);
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Pivot cədvəl
    // ------------------------------------------------------------------
    const AGG_NAMES = { sum: "Sum", count: "Count", average: "Average", min: "Min", max: "Max" };
    function pivotFields(wb, srcRange) {
        const g = parseRange(srcRange);
        if (!g) return [];
        const out = [];
        for (let c = g.c1; c <= g.c2; c++) { const v = wb.value(addr(c, g.r1)); if (v !== null && v !== "") out.push(toStr(v)); }
        return out;
    }
    function aggregate(vals, agg) {
        const nums = vals.filter((x) => typeof x === "number");
        switch (agg) {
            case "count": return vals.filter((x) => x !== null && x !== "").length;
            case "average": return nums.length ? sum(nums) / nums.length : null;
            case "min": return nums.length ? Math.min(...nums) : null;
            case "max": return nums.length ? Math.max(...nums) : null;
            default: return nums.length ? sum(nums) : (vals.length ? 0 : null);
        }
    }
    const keyCmp = (x, y) => (typeof x === "number" && typeof y === "number" ? x - y : compare(x, y));
    function pivotValueNames(def) {
        const seen = {};
        return (def.values || []).map((v) => {
            seen[v.field] = (seen[v.field] || 0) + 1;
            return (AGG_NAMES[v.agg] || "Sum") + " of " + v.field + (seen[v.field] > 1 ? seen[v.field] : "");
        });
    }
    function computePivot(wb, def) {
        const g = parseRange(def.src);
        const res = { ok: false, rows: [], rowKeys: [], colKeys: [], names: [], error: null, get: () => null, grand: () => null };
        if (!g) { res.error = "Invalid source range"; return res; }
        const fields = pivotFields(wb, def.src);
        const idx = {};
        fields.forEach((f, i) => { idx[f] = g.c1 + i; });
        if (!def.rows || !(def.rows in idx)) { res.error = "No Rows field selected"; return res; }
        if (!def.values || !def.values.length || def.values.some((v) => !(v.field in idx))) { res.error = "No Values field selected"; return res; }
        const recs = [];
        for (let r = g.r1 + 1; r <= g.r2; r++) {
            const rec = {};
            let empty = true;
            for (const f of fields) { const v = wb.value(addr(idx[f], r)); rec[f] = v; if (v !== null && v !== "") empty = false; }
            if (!empty) recs.push(rec);
        }
        const filt = def.filters || {};
        const rows = recs.filter((rec) => Object.keys(filt).every((f) => !filt[f] || !filt[f].length || filt[f].includes(toStr(rec[f]))));
        const uniq = (f) => { const m = new Map(); for (const rec of rows) { const v = rec[f]; if (v !== null && v !== "") m.set(toStr(v), v); } return [...m.values()].sort(keyCmp); };
        const rowKeys = uniq(def.rows);
        const hasCols = !!(def.cols && def.cols in idx && def.cols !== def.rows);
        const colKeys = hasCols ? uniq(def.cols) : [];
        const names = pivotValueNames(def);
        const multi = def.values.length;
        const sel = (rk, ck) => rows.filter((rec) => (rk === undefined || toStr(rec[def.rows]) === toStr(rk)) && (ck === undefined || toStr(rec[def.cols]) === toStr(ck)));
        const raw = (vi, rk, ck) => { const s = sel(rk, ck); return s.length ? aggregate(s.map((x) => x[def.values[vi].field]), def.values[vi].agg) : null; };
        const gt = def.values.map((v, vi) => raw(vi));
        const shown = (vi, rk, ck) => {
            const v = raw(vi, rk, ck);
            if (v === null) return null;
            if (def.values[vi].show === "pct_grand") return gt[vi] ? v / gt[vi] : null;
            return v;
        };
        const cell = (v, vi) => ({ v, pct: def.values[vi].show === "pct_grand" });
        const out = [];
        if (!hasCols) {
            out.push([{ v: "Row Labels", h: 1 }, ...names.map((n) => ({ v: n, h: 1 }))]);
            for (const rk of rowKeys) out.push([{ v: rk, l: 1 }, ...def.values.map((_, vi) => cell(shown(vi, rk), vi))]);
            out.push([{ v: "Grand Total", h: 1 }, ...def.values.map((_, vi) => cell(shown(vi), vi))]);
        } else {
            out.push([{ v: names[0], h: 1 }, { v: "Column Labels", h: 1 }, ...colKeys.slice(1).map(() => ({ v: null })), { v: null }]);
            out.push([{ v: "Row Labels", h: 1 }, ...colKeys.map((ck) => ({ v: ck, h: 1 })), { v: "Grand Total", h: 1 }]);
            for (const rk of rowKeys) out.push([{ v: rk, l: 1 }, ...colKeys.map((ck) => cell(shown(0, rk, ck), 0)), cell(shown(0, rk), 0)]);
            out.push([{ v: "Grand Total", h: 1 }, ...colKeys.map((ck) => cell(shown(0, undefined, ck), 0)), cell(shown(0), 0)]);
        }
        res.ok = true; res.rows = out; res.rowKeys = rowKeys; res.colKeys = colKeys; res.names = names; res.multi = multi;
        res.get = (rk, ck, vi = 0) => shown(vi, rk, ck);
        res.grand = (vi = 0) => shown(vi);
        res.rawGrand = (vi = 0) => gt[vi];
        return res;
    }

    // ------------------------------------------------------------------
    // Qrafik üçün verilənlər
    // ------------------------------------------------------------------
    function chartData(wb, def) {
        const g = parseRange(def.range);
        if (!g) return null;
        const cats = [], series = [];
        for (let r = g.r1 + 1; r <= g.r2; r++) cats.push(wb.value(addr(g.c1, r)));
        for (let c = g.c1 + 1; c <= g.c2; c++) {
            const name = toStr(wb.value(addr(c, g.r1)));
            const vals = [];
            for (let r = g.r1 + 1; r <= g.r2; r++) { const v = wb.value(addr(c, r)); vals.push(typeof v === "number" ? v : null); }
            const cfg = (def.series || []).find((s) => s.col === numToCol(c)) || {};
            series.push({ col: numToCol(c), name, vals, kind: cfg.kind || (def.type === "line" ? "line" : "column"), axis: cfg.axis || "primary" });
        }
        const xs = cats.map((v) => (typeof v === "number" ? v : null));
        let trend = null;
        if (def.type === "scatter" && series.length) {
            const pts = xs.map((x, i) => [x, series[0].vals[i]]).filter((p) => p[0] !== null && p[1] !== null);
            if (pts.length >= 2) {
                const n = pts.length, mx = sum(pts.map((p) => p[0])) / n, my = sum(pts.map((p) => p[1])) / n;
                const sxx = sum(pts.map((p) => (p[0] - mx) ** 2)), sxy = sum(pts.map((p) => (p[0] - mx) * (p[1] - my))), syy = sum(pts.map((p) => (p[1] - my) ** 2));
                if (sxx) trend = { slope: sxy / sxx, intercept: my - (sxy / sxx) * mx, r2: syy ? (sxy * sxy) / (sxx * syy) : 1 };
            }
        }
        return { cats, xs, series, trend };
    }

    // ------------------------------------------------------------------
    return {
        XlError, E, isErr, Rng, Workbook, parseFormula, tokenize, shiftFormula,
        colToNum, numToCol, addr, parseAddr, parseRange, cellsIn, rangeCovers, inRange,
        parseConst, toStr, toNum, norm15,
        formatNumber, formatValue,
        dvRulesFor, dvListItems, dvViolation,
        CF_STYLES, cfStyleFor, cfRuleMatches,
        pivotFields, computePivot, pivotValueNames, AGG_NAMES,
        chartData, FUNCS, MAX_ROWS,
    };
});
