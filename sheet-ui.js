/* ====================================================================
 * SDF 1 — Excel tipli cədvəl redaktoru (tələbə interfeysi)
 *   SheetUI.mount(host, { task, state, onChange(state) }) → { getState, reset }
 * Xanalar, düstur sətri, doldurma (fill handle), daxili kopyala/yapışdır (nisbi istinadlarla),
 * ədəd formatı, Data Validation, Conditional Formatting, PivotTable (slicer ilə), Chart pəncərələri.
 * ==================================================================== */
(function (root) {
    "use strict";
    const SE = root.SheetEngine;
    const { colToNum, numToCol, addr, parseAddr, parseRange, cellsIn } = SE;

    const CSS = `
:root{--xb:#c8cdd4;--xh:#eef0f3;--xs:#14305a}
.xl,.xl-ov{font-size:13px;color:#1b1f24}
.xl *{user-select:none;-webkit-user-select:none}.xl input,.xl select{user-select:text;-webkit-user-select:text}
.xl-tb{display:flex;flex-wrap:wrap;gap:4px;align-items:center;padding:6px;background:#f6f7f9;border:1px solid var(--xb);border-bottom:0;border-radius:6px 6px 0 0}
.xl-tb .sep{width:1px;height:20px;background:var(--xb);margin:0 3px}
.xl-b{font:inherit;font-size:12px;padding:4px 8px;border:1px solid var(--xb);background:#fff;border-radius:4px;cursor:pointer;color:#1b1f24}
.xl-b:hover{background:#e9eef7}.xl-b.pri{background:#14305a;color:#fff;border-color:#14305a}.xl-b.pri:hover{background:#0e2343}
.xl-b.dng{color:#9b1c1c}.xl-b:disabled{opacity:.45;cursor:default}
.xl-fb{display:flex;border:1px solid var(--xb);border-bottom:0;background:#fff}
.xl-nb{width:70px;padding:4px 6px;border-right:1px solid var(--xb);font-weight:600;text-align:center;background:#f6f7f9}
.xl-fx{padding:4px 8px;color:#5a6370;font-style:italic;border-right:1px solid var(--xb)}
.xl-fi{flex:1;border:0;padding:4px 8px;font:inherit;outline:0;min-width:0;font-family:Consolas,monospace}
.xl-wrap{position:relative;overflow:auto;border:1px solid var(--xb);background:#fff;max-height:470px;outline:0}
.xl table{border-collapse:collapse;table-layout:fixed}
.xl th{background:var(--xh);font-weight:600;color:#5a6370;border:1px solid var(--xb);padding:2px 4px;text-align:center;position:sticky;top:0;z-index:2;font-size:12px}
.xl th.rh{position:sticky;left:0;top:auto;z-index:1;width:34px;min-width:34px}.xl th.corner{left:0;z-index:3}
.xl th.hot{background:#d3dff2;color:#14305a}
.xl td{border:1px solid #dde0e5;height:22px;padding:0 4px;overflow:hidden;white-space:nowrap;text-overflow:clip;position:relative;cursor:cell;font-size:13px}
.xl td.num{text-align:right}.xl td.err{text-align:center;color:#9b1c1c}.xl td.bl{text-align:center}
.xl td.sel{background:#e3ecfa}.xl td.act{outline:2px solid var(--xs);outline-offset:-2px;background:#fff}
.xl td.cf{}.xl td input{position:absolute;inset:0;width:100%;height:100%;border:2px solid var(--xs);padding:0 3px;font:inherit;outline:0;background:#fff;z-index:5}
.xl-fh{position:absolute;width:8px;height:8px;background:var(--xs);border:1px solid #fff;cursor:crosshair;z-index:4;display:none}
.xl-dd{position:absolute;width:18px;height:20px;border:1px solid #888;background:#fff;z-index:4;display:none;cursor:pointer;text-align:center;font-size:10px;line-height:18px}
.xl-menu{position:absolute;background:#fff;border:1px solid #888;box-shadow:0 2px 8px rgba(0,0,0,.25);z-index:20;min-width:110px}
.xl-menu div{padding:4px 10px;cursor:pointer}.xl-menu div:hover{background:#d3dff2}
.xl-st{min-height:20px;padding:3px 8px;font-size:12px;color:#5a6370;border:1px solid var(--xb);border-top:0;border-radius:0 0 6px 6px;background:#fafbfc}
.xl-st.bad{color:#9b1c1c;font-weight:600}
.xl-objs{margin-top:10px;display:flex;flex-direction:column;gap:10px}
.xl-obj{border:1px solid var(--xb);border-radius:6px;background:#fff;overflow:hidden}
.xl-oh{display:flex;align-items:center;gap:6px;padding:5px 8px;background:#f6f7f9;border-bottom:1px solid var(--xb);font-weight:600;font-size:12px}
.xl-oh .sp{flex:1}.xl-ob{padding:8px;overflow:auto}
.xl-pt{border-collapse:collapse;font-size:13px}.xl-pt td{border:1px solid #dde0e5;padding:2px 8px;text-align:right;cursor:default;height:auto}
.xl-pt td.h{background:#dce6f5;font-weight:600;text-align:left}.xl-pt td.l{text-align:left}.xl-pt tr:last-child td{font-weight:700;background:#eef2f9}
.xl-sl{display:flex;flex-wrap:wrap;gap:4px;align-items:center;margin-bottom:8px;padding:6px;border:1px solid #cfd6e2;border-radius:4px;background:#f7f9fc}
.xl-sl b{font-size:12px;margin-right:4px}.xl-sl button{font:inherit;font-size:12px;padding:3px 12px;border:1px solid #9aa7bd;background:#fff;cursor:pointer;border-radius:3px}
.xl-sl button.on{background:#14305a;color:#fff;border-color:#14305a}
.xl-ov{position:fixed;inset:0;background:rgba(14,35,67,.45);z-index:9000;display:flex;align-items:center;justify-content:center;padding:16px}
.xl-dl{background:#fff;border-radius:8px;width:min(560px,100%);max-height:92vh;overflow:auto;box-shadow:0 10px 40px rgba(0,0,0,.35)}
.xl-dl h3{font-size:15px;padding:12px 16px;border-bottom:1px solid #e1e4e8;color:#14305a}
.xl-db{padding:12px 16px;display:flex;flex-direction:column;gap:9px}.xl-df{display:flex;gap:8px;justify-content:flex-end;padding:10px 16px;border-top:1px solid #e1e4e8}
.xl-r{display:flex;align-items:center;gap:8px}.xl-r>label{width:140px;font-weight:600;font-size:12px;flex:none}
.xl-r input[type=text],.xl-r select{flex:1;min-width:0;padding:5px 7px;border:1px solid var(--xb);border-radius:4px;font:inherit;background:#fff}
.xl-h{font-size:12px;color:#5a6370}.xl-li{display:flex;align-items:center;gap:8px;padding:4px 8px;border:1px solid #e1e4e8;border-radius:4px;font-size:12px}
.xl-li span{flex:1}.xl-sw{display:inline-block;width:22px;height:14px;border-radius:2px;border:1px solid #999;vertical-align:middle}
.xl-vr{display:flex;gap:6px;align-items:center;flex-wrap:wrap}.xl-vr select,.xl-vr input{padding:4px;border:1px solid var(--xb);border-radius:4px;font:inherit;background:#fff}
.xl svg text{font-family:inherit;user-select:none}
`;
    function injectCss() {
        if (document.getElementById("xl-css")) return;
        const s = document.createElement("style"); s.id = "xl-css"; s.textContent = CSS; document.head.appendChild(s);
    }
    function el(tag, attrs, ...kids) {
        const e = document.createElement(tag);
        for (const [k, v] of Object.entries(attrs || {})) {
            if (k === "class") e.className = v; else if (k === "text") e.textContent = v;
            else if (k.startsWith("on")) e.addEventListener(k.slice(2), v); else if (v !== false && v !== null && v !== undefined) e.setAttribute(k, v === true ? "" : v);
        }
        for (const kid of kids.flat()) if (kid !== null && kid !== undefined && kid !== false) e.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
        return e;
    }
    const NS = "http://www.w3.org/2000/svg";
    function sv(tag, attrs, text) {
        const e = document.createElementNS(NS, tag);
        for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v);
        if (text !== undefined) e.textContent = text;
        return e;
    }
    const clone = (o) => JSON.parse(JSON.stringify(o));
    const OPS = { between: "between", notbetween: "not between", eq: "equal to", neq: "not equal to", gt: "greater than", gte: "greater than or equal to", lt: "less than", lte: "less than or equal to" };
    const NUMFMTS = { general: "General", number: "Number", currency: "Currency ($)", percent: "Percentage" };

    function normRange(s) {
        const t = String(s || "").replace(/\$/g, "").trim().toUpperCase();
        if (/^[A-Z]{1,3}\d+$/.test(t)) return t + ":" + t;
        return /^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/.test(t) ? t : null;
    }
    const shortRange = (g) => (g.c1 === g.c2 && g.r1 === g.r2 ? addr(g.c1, g.r1) : addr(g.c1, g.r1) + ":" + addr(g.c2, g.r2));

    function mount(host, opts) {
        injectCss();
        const task = opts.task;
        const COLS = task.grid.cols, ROWS = task.grid.rows;
        let st = SE_normalize(opts.state || task.setup);
        let wb = new SE.Workbook(st);
        const undo = [];
        let sel = { a: { c: 1, r: 1 }, b: { c: 1, r: 1 } };
        let editing = null;       // {a, input, typed}
        let clip = null;
        let fillDrag = null;

        function SE_normalize(s) {
            const o = clone(s || {});
            return { cells: o.cells || {}, fmt: o.fmt || {}, dv: o.dv || [], cf: o.cf || [], pivots: o.pivots || [], charts: o.charts || [] };
        }
        function fireChange() { try { opts.onChange && opts.onChange(clone(st)); } catch (e) {} }
        function snap() { undo.push(JSON.stringify(st)); if (undo.length > 60) undo.shift(); }
        function setState(ns) { st = ns; wb = new SE.Workbook(st); }

        // ---------- DOM ----------
        host.innerHTML = "";
        const root_ = el("div", { class: "xl" });
        const tb = el("div", { class: "xl-tb" });
        const nameBox = el("div", { class: "xl-nb", text: "A1" });
        const fInput = el("input", { class: "xl-fi", type: "text", spellcheck: "false", autocomplete: "off" });
        const fb = el("div", { class: "xl-fb" }, nameBox, el("div", { class: "xl-fx", text: "fx" }), fInput);
        const wrap = el("div", { class: "xl-wrap", tabindex: "0" });
        const table = el("table");
        const handle = el("div", { class: "xl-fh" });
        const ddBtn = el("div", { class: "xl-dd", text: "▼" });
        const status = el("div", { class: "xl-st", text: "Select a cell and type. Formulas start with “=”. Drag the small square at the bottom-right of the selection to fill." });
        const objs = el("div", { class: "xl-objs" });
        wrap.append(table, handle, ddBtn);
        root_.append(tb, fb, wrap, status, objs);
        host.append(root_);

        function say(msg, bad) { status.textContent = msg; status.className = "xl-st" + (bad ? " bad" : ""); }

        // sütun enləri
        const widths = [];
        for (let c = 1; c <= COLS; c++) {
            let m = 0;
            for (const [a, raw] of Object.entries(task.setup.cells || {})) { const p = parseAddr(a); if (p.c === c) m = Math.max(m, String(raw).replace(/^'/, "").length); }
            widths[c] = Math.max(78, Math.min(230, Math.round(m * 7.2 + 18)));
        }
        const tds = {};
        (function build() {
            const thead = el("tr", {}, el("th", { class: "rh corner" }));
            for (let c = 1; c <= COLS; c++) {
                const th = el("th", { text: numToCol(c), "data-c": c, style: `width:${widths[c]}px;min-width:${widths[c]}px` });
                th.addEventListener("mousedown", (e) => { e.preventDefault(); sel = { a: { c, r: 1 }, b: { c, r: ROWS } }; endEdit(true); refresh(); wrap.focus(); });
                thead.append(th);
            }
            table.append(thead);
            for (let r = 1; r <= ROWS; r++) {
                const tr = el("tr");
                const rh = el("th", { class: "rh", text: r, "data-r": r });
                rh.addEventListener("mousedown", (e) => { e.preventDefault(); sel = { a: { c: 1, r }, b: { c: COLS, r } }; endEdit(true); refresh(); wrap.focus(); });
                tr.append(rh);
                for (let c = 1; c <= COLS; c++) {
                    const a = addr(c, r), td = el("td", { "data-a": a });
                    tds[a] = td; tr.append(td);
                }
                table.append(tr);
            }
        })();

        // ---------- seçim və göstərmə ----------
        const selRect = () => ({ c1: Math.min(sel.a.c, sel.b.c), c2: Math.max(sel.a.c, sel.b.c), r1: Math.min(sel.a.r, sel.b.r), r2: Math.max(sel.a.r, sel.b.r) });
        const selText = () => shortRange(selRect());
        const actAddr = () => addr(sel.a.c, sel.a.r);

        function cellText(a) {
            const v = wb.value(a);
            return SE.formatValue(v, st.fmt[a]);
        }
        function refresh() {
            wb.invalidate();
            const g = selRect();
            for (const a in tds) {
                if (editing && editing.a === a) continue;
                const td = tds[a], p = parseAddr(a);
                const v = wb.value(a);
                td.textContent = cellText(a);
                let cls = "";
                if (SE.isErr(v)) cls = "err"; else if (typeof v === "number") cls = "num"; else if (typeof v === "boolean") cls = "bl";
                if (p.c >= g.c1 && p.c <= g.c2 && p.r >= g.r1 && p.r <= g.r2) cls += (p.c === sel.a.c && p.r === sel.a.r) ? " act" : " sel";
                td.className = cls;
                const cf = (st.cf || []).length ? SE.cfStyleFor(wb, a) : null;
                td.style.background = cf ? cf.fill : ""; td.style.color = cf ? cf.font : "";
                td.title = typeof v === "string" && v.length > 14 ? v : "";
            }
            table.querySelectorAll("th[data-c]").forEach((th) => th.classList.toggle("hot", +th.dataset.c >= g.c1 && +th.dataset.c <= g.c2));
            table.querySelectorAll("th[data-r]").forEach((th) => th.classList.toggle("hot", +th.dataset.r >= g.r1 && +th.dataset.r <= g.r2));
            nameBox.textContent = (g.c1 === g.c2 && g.r1 === g.r2) ? actAddr() : selText();
            if (!editing) fInput.value = wb.raw(actAddr()) === "" ? "" : String(wb.raw(actAddr()));
            positionAux();
            renderObjects();
        }
        function positionAux() {
            const g = selRect(), br = tds[addr(g.c2, g.r2)];
            if (br && !editing) { handle.style.display = "block"; handle.style.left = br.offsetLeft + br.offsetWidth - 5 + "px"; handle.style.top = br.offsetTop + br.offsetHeight - 5 + "px"; }
            else handle.style.display = "none";
            const a = actAddr(), td = tds[a];
            const rule = (st.dv || []).find((r) => r.type === "list" && SE.inRange(r.range, a));
            if (rule && td && !editing) { ddBtn.style.display = "block"; ddBtn.style.left = td.offsetLeft + td.offsetWidth - 19 + "px"; ddBtn.style.top = td.offsetTop + 1 + "px"; ddBtn._rule = rule; }
            else ddBtn.style.display = "none";
        }
        function scrollTo(a) {
            const td = tds[a]; if (!td) return;
            const l = td.offsetLeft, t = td.offsetTop;
            if (l < wrap.scrollLeft + 36) wrap.scrollLeft = Math.max(0, l - 36);
            else if (l + td.offsetWidth > wrap.scrollLeft + wrap.clientWidth) wrap.scrollLeft = l + td.offsetWidth - wrap.clientWidth + 4;
            if (t < wrap.scrollTop + 24) wrap.scrollTop = Math.max(0, t - 24);
            else if (t + td.offsetHeight > wrap.scrollTop + wrap.clientHeight) wrap.scrollTop = t + td.offsetHeight - wrap.clientHeight + 4;
        }
        function moveTo(c, r, extend) {
            c = Math.max(1, Math.min(COLS, c)); r = Math.max(1, Math.min(ROWS, r));
            if (extend) sel.b = { c, r }; else sel = { a: { c, r }, b: { c, r } };
            refresh(); scrollTo(addr(c, r));
        }

        // ---------- redaktə ----------
        function startEdit(a, initial, typed) {
            if (editing) endEdit(true);
            const td = tds[a]; if (!td) return;
            const raw = initial !== undefined ? initial : String(wb.raw(a));
            td.textContent = "";
            const inp = el("input", { type: "text", spellcheck: "false", autocomplete: "off" });
            inp.value = raw;
            td.append(inp);
            editing = { a, input: inp, typed: !!typed };
            inp.focus();
            if (initial === undefined) inp.setSelectionRange(raw.length, raw.length);
            inp.addEventListener("input", () => { fInput.value = inp.value; });
            inp.addEventListener("keydown", (e) => {
                e.stopPropagation();
                if (e.key === "Enter") { e.preventDefault(); if (endEdit(false)) moveTo(sel.a.c, sel.a.r + (e.shiftKey ? -1 : 1)); }
                else if (e.key === "Tab") { e.preventDefault(); if (endEdit(false)) moveTo(sel.a.c + (e.shiftKey ? -1 : 1), sel.a.r); }
                else if (e.key === "Escape") { e.preventDefault(); endEdit(true); wrap.focus(); refresh(); }
                else if (editing && editing.typed && ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key)) {
                    e.preventDefault();
                    const d = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[e.key];
                    if (endEdit(false)) moveTo(sel.a.c + d[0], sel.a.r + d[1]);
                }
            });
            inp.addEventListener("blur", () => { if (editing && editing.input === inp) endEdit(false); });
            fInput.value = raw;
        }
        // cancel=true → ləğv; false → yadda saxla. Qaytarır: uğurlu olub-olmadığı
        function endEdit(cancel) {
            if (!editing) return true;
            const { a, input } = editing;
            const text = input.value;
            if (!cancel) {
                const ok = commitCell(a, text);
                if (!ok) { input.focus(); return false; }
            }
            editing = null;
            if (input.parentNode) input.remove();
            refresh();
            return true;
        }
        function showAlert(title, msg) {
            const ov = el("div", { class: "xl-ov" });
            const close = () => { ov.remove(); wrap.focus(); };
            ov.append(el("div", { class: "xl-dl", style: "width:min(420px,100%)" }, el("h3", { text: title }), el("div", { class: "xl-db" }, msg), el("div", { class: "xl-df" }, el("button", { class: "xl-b pri", text: "OK", onclick: close }))));
            document.body.append(ov);
            ov.querySelector("button").focus();
        }
        function autoFormat(a, text) {
            if (st.fmt[a]) return;
            let m;
            if ((m = /^-?\d+(\.(\d+))?%$/.exec(text))) st.fmt[a] = { t: "pct", d: m[2] ? m[2].length : 0 };
            else if ((m = /^-?\$\d[\d,]*(\.(\d+))?$/.exec(text))) st.fmt[a] = { t: "cur", d: m[2] ? m[2].length : 0 };
        }
        function commitCell(a, text) {
            text = String(text);
            if (text.trim() !== text && text[0] !== "'") text = text.trim();
            if (text[0] === "=") {
                const err = wb.checkFormula(text);
                if (err) { say("Invalid formula: " + text, true); return false; }
            }
            const old = String(wb.raw(a));
            if (old === text) return true;
            const rule = text !== "" ? SE.dvViolation(wb, a, text) : null;
            if (rule) {
                const msg = rule.msg || "This value doesn't match the data validation restrictions defined for this cell.";
                editing = null; const inp = tds[a].querySelector("input"); if (inp) inp.remove(); refresh();
                say("Data Validation: the entered value was rejected.", true);
                showAlert(rule.title || "Microsoft Excel", msg);
                return true;
            }
            snap();
            wb.set(a, text);
            autoFormat(a, text);
            say("");
            fireChange();
            return true;
        }

        fInput.addEventListener("focus", () => { if (!editing) { /* düstur sətrində redaktə */ } });
        fInput.addEventListener("keydown", (e) => {
            e.stopPropagation();
            if (e.key === "Enter") { e.preventDefault(); const a = actAddr(); if (commitCell(a, fInput.value)) { refresh(); moveTo(sel.a.c, sel.a.r + 1); wrap.focus(); } }
            else if (e.key === "Escape") { e.preventDefault(); refresh(); wrap.focus(); }
        });

        // ---------- siçan ----------
        let dragging = false;
        table.addEventListener("mousedown", (e) => {
            const td = e.target.closest("td"); if (!td || td.querySelector("input")) return;
            e.preventDefault();
            if (editing) { if (!endEdit(false)) return; }
            const p = parseAddr(td.dataset.a);
            if (e.shiftKey) sel.b = { c: p.c, r: p.r }; else sel = { a: { c: p.c, r: p.r }, b: { c: p.c, r: p.r } };
            dragging = true; refresh(); wrap.focus();
        });
        table.addEventListener("mouseover", (e) => {
            if (!dragging) return;
            const td = e.target.closest("td"); if (!td) return;
            const p = parseAddr(td.dataset.a); sel.b = { c: p.c, r: p.r }; refresh();
        });
        table.addEventListener("dblclick", (e) => { const td = e.target.closest("td"); if (td && !td.querySelector("input")) startEdit(td.dataset.a); });
        const onUp = () => {
            dragging = false;
            if (fillDrag) { const fd = fillDrag; fillDrag = null; doFill(fd.src, fd.target); }
        };
        document.addEventListener("mouseup", onUp);
        handle.addEventListener("mousedown", (e) => { e.preventDefault(); e.stopPropagation(); fillDrag = { src: selRect(), target: null }; });
        wrap.addEventListener("mousemove", (e) => {
            if (!fillDrag) return;
            const td = document.elementFromPoint(e.clientX, e.clientY);
            const cell = td && td.closest && td.closest("td[data-a]");
            if (cell) { fillDrag.target = parseAddr(cell.dataset.a); const s = fillDrag.src; sel = { a: { c: s.c1, r: s.r1 }, b: { c: Math.max(s.c2, fillDrag.target.c), r: Math.max(s.r2, fillDrag.target.r) } }; refresh(); handle.style.display = "block"; }
        });
        ddBtn.addEventListener("mousedown", (e) => {
            e.preventDefault(); e.stopPropagation();
            const rule = ddBtn._rule, a = actAddr(); if (!rule) return;
            const items = SE.dvListItems(wb, rule);
            const menu = el("div", { class: "xl-menu", style: `left:${ddBtn.offsetLeft - 90}px;top:${ddBtn.offsetTop + 22}px` },
                items.map((it) => el("div", { text: it, onmousedown: (ev) => { ev.preventDefault(); menu.remove(); if (commitCell(a, it)) { refresh(); } wrap.focus(); } })));
            wrap.append(menu);
            const close = (ev) => { if (!menu.contains(ev.target)) { menu.remove(); document.removeEventListener("mousedown", close, true); } };
            setTimeout(() => document.addEventListener("mousedown", close, true), 0);
        });

        // ---------- doldurma / kopyalama ----------
        function copyRaw(a, dr, dc) {
            const raw = String(wb.raw(a));
            return raw[0] === "=" ? SE.shiftFormula(raw, dr, dc) : raw;
        }
        function doFill(src, target) {
            if (!src || !target) { refresh(); return; }
            const downD = target.r - src.r2, rightD = target.c - src.c2;
            if (downD <= 0 && rightD <= 0) { sel = { a: { c: src.c1, r: src.r1 }, b: { c: src.c2, r: src.r2 } }; refresh(); return; }
            snap();
            const h = src.r2 - src.r1 + 1, w = src.c2 - src.c1 + 1;
            const cp = (a, sa, dr, dc) => { setRaw(a, copyRaw(sa, dr, dc)); if (st.fmt[sa]) st.fmt[a] = { ...st.fmt[sa] }; else delete st.fmt[a]; };
            if (downD >= rightD) {
                for (let c = src.c1; c <= src.c2; c++) for (let r = src.r2 + 1; r <= Math.min(ROWS, target.r); r++) {
                    const sr = src.r1 + ((r - src.r1) % h); cp(addr(c, r), addr(c, sr), r - sr, 0);
                }
                sel = { a: { c: src.c1, r: src.r1 }, b: { c: src.c2, r: Math.min(ROWS, target.r) } };
            } else {
                for (let r = src.r1; r <= src.r2; r++) for (let c = src.c2 + 1; c <= Math.min(COLS, target.c); c++) {
                    const sc = src.c1 + ((c - src.c1) % w); cp(addr(c, r), addr(sc, r), 0, c - sc);
                }
                sel = { a: { c: src.c1, r: src.r1 }, b: { c: Math.min(COLS, target.c), r: src.r2 } };
            }
            fireChange(); refresh();
        }
        function setRaw(a, raw) { if (raw === "") delete st.cells[a]; else st.cells[a] = raw; wb.invalidate(); }

        function fillDir(dir) {
            const g = selRect(); snap();
            if (dir === "down") {
                if (g.r1 === g.r2) { say("Select at least 2 rows to fill down (the top cell is copied)."); undo.pop(); return; }
                for (let c = g.c1; c <= g.c2; c++) for (let r = g.r1 + 1; r <= g.r2; r++) { const a = addr(c, r), sa = addr(c, g.r1); setRaw(a, copyRaw(sa, r - g.r1, 0)); st.fmt[sa] ? st.fmt[a] = { ...st.fmt[sa] } : delete st.fmt[a]; }
            } else {
                if (g.c1 === g.c2) { say("Select at least 2 columns to fill right (the left cell is copied)."); undo.pop(); return; }
                for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) { const a = addr(c, r), sa = addr(g.c1, r); setRaw(a, copyRaw(sa, 0, c - g.c1)); st.fmt[sa] ? st.fmt[a] = { ...st.fmt[sa] } : delete st.fmt[a]; }
            }
            fireChange(); refresh();
        }
        function doCopy() {
            const g = selRect(), cells = [], fmt = [];
            for (let r = g.r1; r <= g.r2; r++) { const rr = [], ff = []; for (let c = g.c1; c <= g.c2; c++) { const a = addr(c, r); rr.push(String(wb.raw(a))); ff.push(st.fmt[a] ? { ...st.fmt[a] } : null); } cells.push(rr); fmt.push(ff); }
            clip = { r0: g.r1, c0: g.c1, cells, fmt };
            say(`Copied: ${selText()} (paste with Ctrl+V; relative references in formulas shift).`);
        }
        function doPaste() {
            if (!clip) { say("Copy a cell first (Ctrl+C)."); return; }
            snap();
            const base = selRect(), ch = clip.cells.length, cw = clip.cells[0].length;
            const selH = base.r2 - base.r1 + 1, selW = base.c2 - base.c1 + 1;
            // seçim kopyalanan blokun tam misli qədərdirsə, blok seçimə təkrarlanır (Excel kimi)
            const H = selH % ch === 0 ? selH : ch, W = selW % cw === 0 ? selW : cw;
            for (let i = 0; i < H; i++) for (let j = 0; j < W; j++) {
                const r = base.r1 + i, c = base.c1 + j; if (r > ROWS || c > COLS) continue;
                const si = i % ch, sj = j % cw, raw = clip.cells[si][sj], a = addr(c, r);
                setRaw(a, raw[0] === "=" ? SE.shiftFormula(raw, r - (clip.r0 + si), c - (clip.c0 + sj)) : raw);
                if (clip.fmt[si][sj]) st.fmt[a] = { ...clip.fmt[si][sj] }; else delete st.fmt[a];
            }
            sel = { a: { c: base.c1, r: base.r1 }, b: { c: Math.min(COLS, base.c1 + W - 1), r: Math.min(ROWS, base.r1 + H - 1) } };
            fireChange(); refresh();
        }
        function doClear() {
            snap(); const g = selRect();
            for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) setRaw(addr(c, r), "");
            fireChange(); refresh();
        }
        function doUndo() {
            if (!undo.length) { say("Nothing to undo."); return; }
            setState(SE_normalize(JSON.parse(undo.pop())));
            fireChange(); refresh(); say("Undone.");
        }

        // ---------- format ----------
        function applyFmt(fn) {
            snap(); const g = selRect();
            for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const a = addr(c, r); const nf = fn(st.fmt[a]); nf ? st.fmt[a] = nf : delete st.fmt[a]; }
            fireChange(); refresh();
        }
        const decChange = (delta) => applyFmt((f) => {
            if (!f || f.t === "gen") { const v = wb.value(actAddr()); const cur = typeof v === "number" && !Number.isInteger(v) ? String(v).split(".")[1].length : 0; return { t: "num", d: Math.max(0, Math.min(10, cur + delta)) }; }
            return { ...f, d: Math.max(0, Math.min(10, (f.d ?? 2) + delta)) };
        });

        // ---------- klaviatura ----------
        wrap.addEventListener("keydown", (e) => {
            if (editing) return;
            const ctrl = e.ctrlKey || e.metaKey;
            const k = e.key;
            if (ctrl) {
                const lk = k.toLowerCase();
                if (lk === "c") { e.preventDefault(); doCopy(); return; }
                if (lk === "v") { e.preventDefault(); doPaste(); return; }
                if (lk === "x") { e.preventDefault(); doCopy(); doClear(); return; }
                if (lk === "z") { e.preventDefault(); doUndo(); return; }
                if (lk === "d") { e.preventDefault(); fillDir("down"); return; }
                if (lk === "r") { e.preventDefault(); fillDir("right"); return; }
                if (lk === "a") { e.preventDefault(); sel = { a: { c: 1, r: 1 }, b: { c: COLS, r: ROWS } }; refresh(); return; }
                return;
            }
            const sh = e.shiftKey;
            if (k === "ArrowDown") { e.preventDefault(); moveTo((sh ? sel.b.c : sel.a.c), (sh ? sel.b.r : sel.a.r) + 1, sh); }
            else if (k === "ArrowUp") { e.preventDefault(); moveTo((sh ? sel.b.c : sel.a.c), (sh ? sel.b.r : sel.a.r) - 1, sh); }
            else if (k === "ArrowLeft") { e.preventDefault(); moveTo((sh ? sel.b.c : sel.a.c) - 1, (sh ? sel.b.r : sel.a.r), sh); }
            else if (k === "ArrowRight") { e.preventDefault(); moveTo((sh ? sel.b.c : sel.a.c) + 1, (sh ? sel.b.r : sel.a.r), sh); }
            else if (k === "Tab") { e.preventDefault(); moveTo(sel.a.c + (sh ? -1 : 1), sel.a.r); }
            else if (k === "Enter") { e.preventDefault(); moveTo(sel.a.c, sel.a.r + (sh ? -1 : 1)); }
            else if (k === "F2") { e.preventDefault(); startEdit(actAddr()); }
            else if (k === "Delete" || k === "Backspace") { e.preventDefault(); doClear(); }
            else if (k.length === 1 && !e.altKey) { e.preventDefault(); startEdit(actAddr(), k, true); }
        });

        // ---------- alətlər paneli ----------
        const fmtSel = el("select", { class: "xl-b", title: "Number format", onchange: () => {
            const v = fmtSel.value; fmtSel.value = "";
            if (v === "gen") applyFmt(() => null);
            else if (v === "num") applyFmt(() => ({ t: "num", d: 2, sep: 1 }));
            else if (v === "cur") applyFmt(() => ({ t: "cur", d: 2 }));
            else if (v === "pct") applyFmt(() => ({ t: "pct", d: 0 }));
        } }, el("option", { value: "", text: "Format ▾" }), el("option", { value: "gen", text: "General" }), el("option", { value: "num", text: "Number" }), el("option", { value: "cur", text: "Currency ($)" }), el("option", { value: "pct", text: "Percentage" }));
        const B = (t, f, title, cls) => el("button", { class: "xl-b " + (cls || ""), text: t, title: title || "", type: "button", onclick: (e) => { e.preventDefault(); f(); wrap.focus(); } });
        const sep = () => el("span", { class: "sep" });
        tb.append(
            B("↶ Undo", doUndo, "Ctrl+Z"), B("Copy", doCopy, "Ctrl+C"), B("Paste", doPaste, "Ctrl+V"), sep(),
            fmtSel, B("$", () => applyFmt(() => ({ t: "cur", d: 2 })), "Currency"), B("%", () => applyFmt(() => ({ t: "pct", d: 0 })), "Percent Style"),
            B(".0←", () => decChange(-1), "Decrease Decimal"), B(".00→", () => decChange(1), "Increase Decimal"), sep(),
            B("Fill ↓", () => fillDir("down"), "Ctrl+D"), B("Fill →", () => fillDir("right"), "Ctrl+R"), B("Clear", doClear, "Delete"), sep(),
            B("Data Validation", () => dvDialog(), "Data → Data Validation"),
            B("Conditional Formatting", () => cfDialog(), "Home → Conditional Formatting"),
            B("PivotTable", () => pivotDialog(), "Insert → PivotTable"),
            B("Chart", () => chartDialog(), "Insert → Chart"), sep(),
            B("Reset", () => { if (confirm("Reset this task's spreadsheet to its initial state? Everything you typed will be erased.")) reset(); }, "", "dng"));

        // ---------- dialoq köməkçiləri ----------
        function modal(title, body, buttons) {
            const ov = el("div", { class: "xl-ov" });
            const close = () => { ov.remove(); wrap.focus(); };
            const foot = el("div", { class: "xl-df" }, buttons.map((b) => el("button", { class: "xl-b " + (b.cls || ""), type: "button", text: b.t, onclick: () => { const r = b.f ? b.f() : true; if (r !== false) close(); } })));
            ov.append(el("div", { class: "xl-dl" }, el("h3", { text: title }), el("div", { class: "xl-db" }, body), foot));
            document.body.append(ov);
            return close;
        }
        const row = (label, ctl) => el("div", { class: "xl-r" }, el("label", { text: label }), ctl);
        const txt = (v, ph) => el("input", { type: "text", value: v || "", placeholder: ph || "", spellcheck: "false" });
        const slc = (items, v) => { const s = el("select"); for (const [k, t] of items) s.append(el("option", { value: k, text: t })); s.value = v; return s; };
        const err = el("div", { class: "xl-h" });
        const bad = (box, m) => { box.textContent = m; box.style.color = "#9b1c1c"; return false; };

        // ---------- Data Validation ----------
        function dvDialog() {
            const defRange = selText();
            const rng = txt(defRange);
            const type = slc([["whole", "Whole number"], ["decimal", "Decimal"], ["list", "List"]], "whole");
            const op = slc(Object.entries(OPS), "between");
            const v1 = txt("", "Minimum"), v2 = txt("", "Maximum");
            const src = txt("", "e.g. $G$1:$G$4  or  A,B,C");
            const title = txt(""), msg = txt("");
            const lab1 = el("label", { text: "Minimum" }), lab2 = el("label", { text: "Maximum" });
            const rOp = row("Data", op), rV1 = el("div", { class: "xl-r" }, lab1, v1), rV2 = el("div", { class: "xl-r" }, lab2, v2), rSrc = row("Source", src);
            const msgBox = el("div", { class: "xl-h" });
            const existing = el("div", { style: "display:flex;flex-direction:column;gap:4px" });
            function sync() {
                const isList = type.value === "list";
                rOp.style.display = rV1.style.display = rV2.style.display = isList ? "none" : "flex";
                rSrc.style.display = isList ? "flex" : "none";
                rV2.style.display = !isList && ["between", "notbetween"].includes(op.value) ? "flex" : "none";
                lab1.textContent = ["between", "notbetween"].includes(op.value) ? "Minimum" : "Value";
            }
            type.onchange = op.onchange = sync; sync();
            function list() {
                existing.innerHTML = "";
                (st.dv || []).forEach((r, i) => existing.append(el("div", { class: "xl-li" },
                    el("span", { text: `${r.range} — ${r.type === "list" ? "List: " + r.source : (r.type === "whole" ? "Whole number" : "Decimal") + " " + OPS[r.op || "between"] + " " + r.min + (["between", "notbetween"].includes(r.op || "between") ? " and " + r.max : "")}` }),
                    el("button", { class: "xl-b dng", type: "button", text: "Delete", onclick: () => { snap(); st.dv.splice(i, 1); fireChange(); refresh(); list(); } }))));
                if (!(st.dv || []).length) existing.append(el("div", { class: "xl-h", text: "No rules yet." }));
            }
            list();
            modal("Data Validation", [
                row("Apply to range", rng), el("div", { class: "xl-h", text: "Settings" }), row("Allow", type), rOp, rV1, rV2, rSrc,
                el("div", { class: "xl-h", text: "Error Alert" }), row("Title", title), row("Error message", msg), msgBox,
                el("div", { class: "xl-h", text: "Existing rules:" }), existing,
            ], [{ t: "Close" }, {
                t: "OK — apply rule", cls: "pri", f: () => {
                    const rg = normRange(rng.value); if (!rg) return bad(msgBox, "Invalid range (e.g. B2:B10).");
                    const rule = { range: rg, type: type.value, msg: msg.value.trim(), title: title.value.trim(), style: "stop" };
                    if (type.value === "list") {
                        if (!src.value.trim()) return bad(msgBox, "Enter the list source.");
                        rule.source = src.value.trim();
                    } else {
                        const a = Number(v1.value), b = Number(v2.value);
                        if (v1.value.trim() === "" || isNaN(a)) return bad(msgBox, "Enter a numeric value.");
                        rule.op = op.value; rule.min = a;
                        if (["between", "notbetween"].includes(op.value)) { if (v2.value.trim() === "" || isNaN(b)) return bad(msgBox, "Enter the maximum value."); rule.max = b; } else rule.max = a;
                    }
                    snap(); st.dv = (st.dv || []).filter((r) => r.range !== rg); st.dv.push(rule);
                    fireChange(); refresh(); say(`Data Validation applied: ${rg}`);
                },
            }]);
        }

        // ---------- Conditional Formatting ----------
        function cfDialog() {
            const rng = txt(selText());
            const kind = slc([["cell", "Format only cells with: Cell Value"], ["formula", "Use a formula to determine which cells to format"]], "cell");
            const op = slc(Object.entries(OPS), "gt");
            const v1 = txt(""), v2 = txt("");
            const fm = txt("", "=COUNTIF($A$2:$A$5,A2)>1");
            const fill = slc(Object.entries(SE.CF_STYLES).map(([k, s]) => [k, s.name]), "lightred");
            const sw = el("span", { class: "xl-sw" });
            const msgBox = el("div", { class: "xl-h" });
            const rK = row("Rule type", kind), rOp = row("Condition", op), rV1 = row("Value", v1), rV2 = row("and", v2), rF = row("Formula", fm);
            const existing = el("div", { style: "display:flex;flex-direction:column;gap:4px" });
            function paint() { const s = SE.CF_STYLES[fill.value]; sw.style.background = s.fill; sw.style.borderColor = s.font; }
            function sync() {
                const f = kind.value === "formula";
                rOp.style.display = rV1.style.display = f ? "none" : "flex"; rF.style.display = f ? "flex" : "none";
                rV2.style.display = !f && ["between", "notbetween"].includes(op.value) ? "flex" : "none";
            }
            kind.onchange = op.onchange = sync; fill.onchange = paint; sync(); paint();
            function list() {
                existing.innerHTML = "";
                (st.cf || []).forEach((r, i) => existing.append(el("div", { class: "xl-li" },
                    el("span", { class: "xl-sw", style: `background:${(SE.CF_STYLES[r.fill] || {}).fill}` }),
                    el("span", { text: `${r.range} — ${r.kind === "formula" ? r.formula : "Cell value " + OPS[r.op] + " " + r.v1 + (r.op === "between" || r.op === "notbetween" ? " and " + r.v2 : "")}` }),
                    el("button", { class: "xl-b dng", type: "button", text: "Delete", onclick: () => { snap(); st.cf.splice(i, 1); fireChange(); refresh(); list(); } }))));
                if (!(st.cf || []).length) existing.append(el("div", { class: "xl-h", text: "No rules yet." }));
            }
            list();
            modal("Conditional Formatting — New Rule", [
                row("Apply to range", rng), rK, rOp, rV1, rV2, rF, row("Format with", el("div", { style: "flex:1;display:flex;gap:8px;align-items:center" }, fill, sw)), msgBox,
                el("div", { class: "xl-h", text: "Existing rules:" }), existing,
            ], [{ t: "Close" }, {
                t: "OK — add rule", cls: "pri", f: () => {
                    const rg = normRange(rng.value); if (!rg) return bad(msgBox, "Invalid range (e.g. D2:D4).");
                    const rule = { range: rg, kind: kind.value, fill: fill.value };
                    if (kind.value === "formula") {
                        let f = fm.value.trim(); if (!f) return bad(msgBox, "Enter a formula.");
                        if (f[0] !== "=") f = "=" + f;
                        const e1 = wb.checkFormula(f); if (e1) return bad(msgBox, "Invalid formula.");
                        rule.formula = f;
                    } else {
                        if (v1.value.trim() === "") return bad(msgBox, "Enter a value.");
                        rule.op = op.value; rule.v1 = v1.value.trim();
                        if (["between", "notbetween"].includes(op.value)) { if (v2.value.trim() === "") return bad(msgBox, "Enter the second value."); rule.v2 = v2.value.trim(); }
                    }
                    snap(); st.cf = st.cf || []; st.cf.push(rule); fireChange(); refresh(); say(`Conditional Formatting rule added: ${rg}`);
                },
            }]);
        }

        // ---------- PivotTable ----------
        function region(a) {   // aktiv xananı əhatə edən ardıcıl məlumat sahəsi (Ctrl+A kimi)
            const p = parseAddr(a); let c1 = p.c, c2 = p.c, r1 = p.r, r2 = p.r;
            const has = (c, r) => wb.raw(addr(c, r)) !== "";
            let grow = true;
            while (grow) {
                grow = false;
                const colHas = (c) => { for (let r = r1 - 1; r <= r2 + 1; r++) if (r >= 1 && r <= ROWS && has(c, r)) return true; return false; };
                const rowHas = (r) => { for (let c = c1 - 1; c <= c2 + 1; c++) if (c >= 1 && c <= COLS && has(c, r)) return true; return false; };
                if (c1 > 1 && colHas(c1 - 1)) { c1--; grow = true; }
                if (c2 < COLS && colHas(c2 + 1)) { c2++; grow = true; }
                if (r1 > 1 && rowHas(r1 - 1)) { r1--; grow = true; }
                if (r2 < ROWS && rowHas(r2 + 1)) { r2++; grow = true; }
            }
            return shortRange({ c1, c2, r1, r2 });
        }
        function pivotDialog(edit) {
            const g0 = selectedOnce();
            const def = edit ? clone(edit) : { id: "P" + (Date.now() % 100000), src: g0, rows: "", cols: "", values: [], filters: {}, slicers: [], numFmt: "general" };
            const src = txt(def.src);
            const fieldsBox = el("div", { class: "xl-h" });
            const rowsS = el("select"), colsS = el("select");
            const valBox = el("div", { style: "display:flex;flex-direction:column;gap:6px" });
            const slBox = el("div", { class: "xl-vr" });
            const nf = slc(Object.entries(NUMFMTS), def.numFmt || "general");
            const msgBox = el("div", { class: "xl-h" });
            let fields = [];
            function loadFields() {
                const rg = normRange(src.value);
                fields = rg ? SE.pivotFields(wb, rg) : [];
                fieldsBox.textContent = fields.length ? "Fields: " + fields.join(" · ") : "Enter a valid source range (including the header row).";
                const opt = (s, v, blank) => { const cur = v; s.innerHTML = ""; s.append(el("option", { value: "", text: blank })); fields.forEach((f) => s.append(el("option", { value: f, text: f }))); s.value = fields.includes(cur) ? cur : ""; };
                opt(rowsS, rowsS.value || def.rows, "— select —"); opt(colsS, colsS.value || def.cols, "— none —");
                drawVals(); drawSlicers();
            }
            function drawVals() {
                const keep = def.values;
                valBox.innerHTML = "";
                keep.forEach((v, i) => {
                    const f = slc([["", "— field —"], ...fields.map((x) => [x, x])], fields.includes(v.field) ? v.field : "");
                    const ag = slc(Object.entries(SE.AGG_NAMES), v.agg || "sum");
                    const sh = slc([["value", "Normal"], ["pct_grand", "% of Grand Total"]], v.show || "value");
                    f.onchange = () => (v.field = f.value); ag.onchange = () => (v.agg = ag.value); sh.onchange = () => (v.show = sh.value);
                    valBox.append(el("div", { class: "xl-vr" }, f, ag, sh, el("button", { class: "xl-b dng", type: "button", text: "✕", onclick: () => { def.values.splice(i, 1); drawVals(); } })));
                });
                valBox.append(el("button", { class: "xl-b", type: "button", text: "+ Add value field", onclick: () => { def.values.push({ field: "", agg: "sum", show: "value" }); drawVals(); } }));
            }
            function drawSlicers() {
                slBox.innerHTML = "";
                if (!fields.length) return;
                fields.forEach((f) => {
                    const cb = el("input", { type: "checkbox" }); cb.checked = (def.slicers || []).includes(f);
                    cb.onchange = () => { def.slicers = (def.slicers || []).filter((x) => x !== f); if (cb.checked) def.slicers.push(f); };
                    slBox.append(el("label", { style: "display:flex;gap:4px;align-items:center;font-size:12px" }, cb, f));
                });
            }
            src.addEventListener("change", loadFields);
            const sd = el("div", { class: "xl-h", text: "Fields to add a Slicer for (Insert Slicer):" });
            modal(edit ? "PivotTable — edit" : "PivotTable — create", [
                row("Table/Range", src), fieldsBox, row("Rows", rowsS), row("Columns", colsS),
                el("div", { class: "xl-h", text: "Values (several values are allowed only when Columns is empty):" }), valBox,
                row("Number format", nf), sd, slBox, msgBox,
            ], [{ t: "Cancel" }, {
                t: edit ? "OK" : "OK — create PivotTable", cls: "pri", f: () => {
                    const rg = normRange(src.value); if (!rg) return bad(msgBox, "Invalid source range.");
                    if (!rowsS.value) return bad(msgBox, "Select the Rows field.");
                    const vals = def.values.filter((v) => v.field); if (!vals.length) return bad(msgBox, "Select at least one Values field.");
                    if (colsS.value && vals.length > 1) return bad(msgBox, "With Columns selected only one Values field is allowed.");
                    def.src = rg; def.rows = rowsS.value; def.cols = colsS.value || ""; def.values = vals.map((v) => ({ field: v.field, agg: v.agg || "sum", show: v.show || "value" })); def.numFmt = nf.value;
                    def.filters = def.filters || {};
                    Object.keys(def.filters).forEach((k) => { if (!(def.slicers || []).includes(k)) delete def.filters[k]; });
                    snap();
                    const i = st.pivots.findIndex((p) => p.id === def.id);
                    if (i >= 0) st.pivots[i] = def; else st.pivots.push(def);
                    fireChange(); refresh(); say("PivotTable created — shown below the spreadsheet.");
                },
            }]);
            rowsS.value = def.rows; loadFields();
            if (def.rows) rowsS.value = def.rows; if (def.cols) colsS.value = def.cols;
            if (!def.values.length) { def.values.push({ field: "", agg: "sum", show: "value" }); drawVals(); }
        }
        function selectedOnce() {
            const g = selRect();
            if (g.c1 === g.c2 && g.r1 === g.r2) return region(actAddr());
            return selText();
        }

        // ---------- Chart ----------
        function chartDialog(edit) {
            const def = edit ? clone(edit) : { id: "C" + (Date.now() % 100000), type: "column", range: selectedOnce(), title: "", series: [], trendline: null };
            const type = slc([["column", "Column (clustered)"], ["line", "Line"], ["combo", "Combo (column + line)"], ["scatter", "Scatter (X, Y)"]], def.type);
            const rng = txt(def.range), title = txt(def.title);
            const serBox = el("div", { style: "display:flex;flex-direction:column;gap:6px" });
            const tl = el("input", { type: "checkbox" }), eq = el("input", { type: "checkbox" }), r2 = el("input", { type: "checkbox" });
            tl.checked = !!def.trendline; eq.checked = !!(def.trendline && def.trendline.eq); r2.checked = !!(def.trendline && def.trendline.r2);
            const trBox = el("div", { class: "xl-vr" }, el("label", {}, tl, " Linear Trendline"), el("label", {}, eq, " Display Equation on chart"), el("label", {}, r2, " Display R-squared value"));
            const msgBox = el("div", { class: "xl-h" });
            let cfg = {};
            (def.series || []).forEach((s) => (cfg[s.col] = { kind: s.kind, axis: s.axis }));
            function drawSeries() {
                serBox.innerHTML = "";
                const g = parseRange((normRange(rng.value) || "A1:A1"));
                const sc = type.value === "scatter";
                trBox.style.display = sc ? "flex" : "none";
                if (!g || g.c2 <= g.c1) { serBox.append(el("div", { class: "xl-h", text: "The range must have at least two columns (first column = category / X)." })); return; }
                serBox.append(el("div", { class: "xl-h", text: sc ? "Y series: " + SE.toStr(wb.value(addr(g.c1 + 1, g.r1))) : "Series:" }));
                if (sc) return;
                for (let c = g.c1 + 1; c <= g.c2; c++) {
                    const col = numToCol(c), cur = cfg[col] || { kind: type.value === "line" ? "line" : "column", axis: "primary" };
                    cfg[col] = cur;
                    const k = slc([["column", "Clustered Column"], ["line", "Line"]], cur.kind);
                    const ax = slc([["primary", "Primary axis"], ["secondary", "Secondary axis"]], cur.axis || "primary");
                    k.onchange = () => (cur.kind = k.value); ax.onchange = () => (cur.axis = ax.value);
                    const only = type.value === "combo";
                    if (!only) { k.disabled = true; ax.disabled = type.value !== "combo" && type.value !== "line" && type.value !== "column"; }
                    serBox.append(el("div", { class: "xl-vr" }, el("b", { text: SE.toStr(wb.value(addr(c, g.r1))) || col }), k, ax));
                }
            }
            type.onchange = () => { if (type.value !== "combo") for (const k in cfg) cfg[k].kind = type.value === "line" ? "line" : "column"; drawSeries(); };
            rng.addEventListener("change", drawSeries);
            modal(edit ? "Chart — edit" : "Insert Chart", [row("Chart type", type), row("Data range", rng), row("Chart title", title), serBox, trBox, msgBox], [{ t: "Cancel" }, {
                t: edit ? "OK" : "OK — add chart", cls: "pri", f: () => {
                    const rg = normRange(rng.value); if (!rg) return bad(msgBox, "Invalid range.");
                    const g = parseRange(rg); if (g.c2 <= g.c1 || g.r2 <= g.r1) return bad(msgBox, "The range must have at least 2 columns and 2 rows (including the header).");
                    const series = []; for (let c = g.c1 + 1; c <= g.c2; c++) { const col = numToCol(c), s = cfg[col] || { kind: "column", axis: "primary" }; series.push({ col, kind: type.value === "line" ? "line" : (type.value === "combo" ? s.kind : "column"), axis: type.value === "scatter" ? "primary" : (s.axis || "primary") }); }
                    def.type = type.value; def.range = rg; def.title = title.value.trim(); def.series = series;
                    def.trendline = type.value === "scatter" && tl.checked ? { type: "linear", eq: eq.checked, r2: r2.checked } : null;
                    snap(); const i = st.charts.findIndex((x) => x.id === def.id); if (i >= 0) st.charts[i] = def; else st.charts.push(def);
                    fireChange(); refresh(); say("Chart added — shown below the spreadsheet.");
                },
            }]);
            drawSeries();
        }

        // ---------- obyektlərin göstərilməsi ----------
        function pfmt(p, v, pct) {
            if (v === null || v === undefined) return "";
            if (pct) return SE.formatNumber(v, { t: "pct", d: 2 });
            if (typeof v !== "number") return String(v);
            const m = p.numFmt;
            if (m === "currency") return SE.formatNumber(v, { t: "cur", d: 2 });
            if (m === "number") return SE.formatNumber(v, { t: "num", d: 2, sep: 1 });
            if (m === "percent") return SE.formatNumber(v, { t: "pct", d: 2 });
            return SE.formatNumber(v, null);
        }
        function renderObjects() {
            objs.innerHTML = "";
            st.pivots.forEach((p, idx) => {
                const body = el("div", { class: "xl-ob" });
                const res = SE.computePivot(wb, p);
                (p.slicers || []).forEach((f) => {
                    const g = parseRange(p.src); const fields = SE.pivotFields(wb, p.src); const fi = fields.indexOf(f); if (fi < 0 || !g) return;
                    const vals = new Map();
                    for (let r = g.r1 + 1; r <= g.r2; r++) { const v = wb.value(addr(g.c1 + fi, r)); if (v !== null && v !== "") vals.set(SE.toStr(v), v); }
                    const cur = (p.filters || {})[f] || [];
                    const box = el("div", { class: "xl-sl" }, el("b", { text: f }));
                    [...vals.keys()].sort().forEach((k) => {
                        const on = cur.length === 0 || cur.includes(k);
                        box.append(el("button", { type: "button", class: on && cur.length ? "on" : (cur.length === 0 ? "" : ""), text: k, style: cur.length === 0 ? "" : "", onclick: (e) => {
                            snap(); p.filters = p.filters || {};
                            let sel_ = (p.filters[f] || []).slice();
                            if (e.ctrlKey || e.shiftKey) { sel_ = sel_.includes(k) ? sel_.filter((x) => x !== k) : sel_.concat(k); }
                            else sel_ = sel_.length === 1 && sel_[0] === k ? [] : [k];
                            if (sel_.length === 0 || sel_.length === vals.size) delete p.filters[f]; else p.filters[f] = sel_;
                            fireChange(); refresh();
                        } }));
                    });
                    box.append(el("button", { type: "button", text: "✕ Clear", onclick: () => { snap(); if (p.filters) delete p.filters[f]; fireChange(); refresh(); } }));
                    body.append(box);
                });
                if (!res.ok) body.append(el("div", { class: "xl-h", style: "color:#9b1c1c", text: res.error || "The PivotTable could not be built" }));
                else {
                    const t = el("table", { class: "xl-pt" });
                    res.rows.forEach((row_) => t.append(el("tr", {}, row_.map((cell, ci) => {
                        const isNum = !cell.h && !cell.l && cell.v !== undefined;
                        return el("td", { class: cell.h ? "h" : (cell.l ? "l" : ""), text: isNum ? pfmt(p, cell.v, cell.pct) : (cell.v === null || cell.v === undefined ? "" : (typeof cell.v === "number" && cell.l ? SE.formatNumber(cell.v, null) : cell.v)) });
                    }))));
                    body.append(t);
                }
                objs.append(el("div", { class: "xl-obj" },
                    el("div", { class: "xl-oh" }, el("span", { text: `PivotTable ${idx + 1}` }), el("span", { class: "sp" }), el("span", { class: "xl-h", text: `${p.src} · Rows: ${p.rows}${p.cols ? " · Columns: " + p.cols : ""}` }),
                        el("button", { class: "xl-b", type: "button", text: "Edit", onclick: () => pivotDialog(p) }),
                        el("button", { class: "xl-b dng", type: "button", text: "Delete", onclick: () => { snap(); st.pivots.splice(idx, 1); fireChange(); refresh(); } })), body));
            });
            st.charts.forEach((ch, idx) => {
                objs.append(el("div", { class: "xl-obj" },
                    el("div", { class: "xl-oh" }, el("span", { text: `Chart ${idx + 1}` }), el("span", { class: "sp" }), el("span", { class: "xl-h", text: `${ch.type} · ${ch.range}` }),
                        el("button", { class: "xl-b", type: "button", text: "Edit", onclick: () => chartDialog(ch) }),
                        el("button", { class: "xl-b dng", type: "button", text: "Delete", onclick: () => { snap(); st.charts.splice(idx, 1); fireChange(); refresh(); } })),
                    el("div", { class: "xl-ob" }, drawChart(ch))));
            });
        }

        function nice(max, min) {
            const span = max - min || 1, step0 = span / 5, mag = Math.pow(10, Math.floor(Math.log10(step0)));
            const f = step0 / mag, step = (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * mag;
            const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step, ticks = [];
            for (let v = lo; v <= hi + step / 2; v += step) ticks.push(+v.toPrecision(12));
            return { lo, hi: ticks[ticks.length - 1], ticks };
        }
        function drawChart(ch) {
            const d = SE.chartData(wb, ch);
            const W = 560, H = 300, L = 62, T = 36, Bm = 56;
            const svg = sv("svg", { viewBox: `0 0 ${W} ${H}`, width: "100%", style: "max-width:600px;background:#fff" });
            if (!d || !d.series.length) { svg.append(sv("text", { x: 10, y: 20, fill: "#9b1c1c" }, "No data found for the chart")); return svg; }
            const colors = ["#2563eb", "#ea580c", "#16a34a", "#7c3aed", "#0891b2"];
            const g0 = parseRange(ch.range);
            const colFmt = (col) => st.fmt[col + (g0.r1 + 1)] || null;
            const hasSec = ch.type !== "scatter" && d.series.some((s) => s.axis === "secondary");
            const R = hasSec ? 62 : 22;
            const pw = W - L - R, ph = H - T - Bm;
            svg.append(sv("text", { x: W / 2, y: 20, "text-anchor": "middle", "font-size": 14, "font-weight": 700, fill: "#14305a" }, ch.title || ""));
            const axisScale = (series) => {
                const all = series.flatMap((s) => s.vals).filter((v) => v !== null);
                if (!all.length) return nice(1, 0);
                const mn = Math.min(0, ...all), mx = Math.max(...all);
                return nice(mx === mn ? mn + 1 : mx, mn);
            };
            const drawAxis = (sc, series, x, right) => {
                const fm = colFmt(series[0] ? series[0].col : "B");
                sc.ticks.forEach((tv) => {
                    const y = T + ph - ((tv - sc.lo) / (sc.hi - sc.lo)) * ph;
                    if (!right) svg.append(sv("line", { x1: L, x2: L + pw, y1: y, y2: y, stroke: "#e1e4e8" }));
                    svg.append(sv("text", { x: right ? x + 6 : x - 6, y: y + 4, "text-anchor": right ? "start" : "end", "font-size": 11, fill: "#5a6370" }, SE.formatNumber(tv, fm && fm.t === "pct" ? { t: "pct", d: 0 } : (Math.abs(tv) >= 1000 ? { t: "num", d: 0, sep: 1 } : null))));
                });
            };
            if (ch.type === "scatter") {
                const s = d.series[0], pts = d.xs.map((x, i) => [x, s.vals[i]]).filter((p) => p[0] !== null && p[1] !== null);
                if (!pts.length) { svg.append(sv("text", { x: 10, y: 50, fill: "#9b1c1c" }, "Scatter needs numeric X and Y values")); return svg; }
                const xs = pts.map((p) => p[0]), xmin = Math.min(...xs), xmax = Math.max(...xs);
                const sx = nice(xmax, Math.min(xmin, xmin - (xmax - xmin) * 0.05)), sy = axisScale([s]);
                drawAxis(sy, [s], L, false);
                const X = (v) => L + ((v - sx.lo) / (sx.hi - sx.lo)) * pw, Y = (v) => T + ph - ((v - sy.lo) / (sy.hi - sy.lo)) * ph;
                sx.ticks.forEach((tv) => svg.append(sv("text", { x: X(tv), y: T + ph + 16, "text-anchor": "middle", "font-size": 11, fill: "#5a6370" }, SE.formatNumber(tv, Math.abs(tv) >= 1000 ? { t: "num", d: 0, sep: 1 } : null))));
                svg.append(sv("line", { x1: L, x2: L + pw, y1: T + ph, y2: T + ph, stroke: "#888" }), sv("line", { x1: L, x2: L, y1: T, y2: T + ph, stroke: "#888" }));
                pts.forEach((p) => svg.append(sv("circle", { cx: X(p[0]), cy: Y(p[1]), r: 4, fill: colors[0] })));
                if (ch.trendline && d.trend) {
                    const t = d.trend;
                    svg.append(sv("line", { x1: X(xmin), y1: Y(t.slope * xmin + t.intercept), x2: X(xmax), y2: Y(t.slope * xmax + t.intercept), stroke: "#555", "stroke-dasharray": "5 4", "stroke-width": 1.5 }));
                    const parts = [];
                    if (ch.trendline.eq) parts.push(`y = ${+t.slope.toPrecision(4)}x ${t.intercept < 0 ? "−" : "+"} ${+Math.abs(t.intercept).toPrecision(4)}`);
                    if (ch.trendline.r2) parts.push(`R² = ${t.r2.toFixed(4)}`);
                    parts.forEach((txtv, i) => svg.append(sv("text", { x: L + 12, y: T + 16 + i * 15, "font-size": 12, fill: "#333", "font-weight": 600 }, txtv)));
                }
                svg.append(sv("text", { x: L + pw / 2, y: H - 22, "text-anchor": "middle", "font-size": 11, fill: "#5a6370" }, SE.toStr(wb.value(addr(g0.c1, g0.r1)))));
                svg.append(sv("text", { x: 14, y: T + ph / 2, "font-size": 11, fill: "#5a6370", transform: `rotate(-90 14 ${T + ph / 2})`, "text-anchor": "middle" }, s.name));
                return svg;
            }
            const prim = d.series.filter((s) => s.axis !== "secondary"), secs = d.series.filter((s) => s.axis === "secondary");
            const sp = axisScale(prim.length ? prim : secs), ss = secs.length ? axisScale(secs) : null;
            drawAxis(sp, prim.length ? prim : secs, L, false);
            if (ss) drawAxis(ss, secs, L + pw, true);
            const n = d.cats.length, band = pw / Math.max(n, 1);
            const Yp = (v) => T + ph - ((v - sp.lo) / (sp.hi - sp.lo)) * ph;
            const Ys = (v) => T + ph - ((v - ss.lo) / (ss.hi - ss.lo)) * ph;
            svg.append(sv("line", { x1: L, x2: L + pw, y1: T + ph, y2: T + ph, stroke: "#888" }));
            d.cats.forEach((cv, i) => svg.append(sv("text", { x: L + band * (i + 0.5), y: T + ph + 16, "text-anchor": "middle", "font-size": 11, fill: "#5a6370" }, SE.toStr(cv))));
            const colsS = d.series.filter((s) => s.kind === "column");
            d.series.forEach((s, si) => {
                const color = colors[si % colors.length], Y = s.axis === "secondary" ? Ys : Yp;
                if (s.kind === "column") {
                    const ci = colsS.indexOf(s), gw = band * 0.7, bw = gw / colsS.length;
                    s.vals.forEach((v, i) => { if (v === null) return; const y0 = Y(0), y1 = Y(v); svg.append(sv("rect", { x: L + band * i + band * 0.15 + ci * bw, y: Math.min(y0, y1), width: bw - 1, height: Math.abs(y0 - y1), fill: color })); });
                } else {
                    const pts = s.vals.map((v, i) => v === null ? null : [L + band * (i + 0.5), Y(v)]).filter(Boolean);
                    svg.append(sv("polyline", { points: pts.map((p) => p.join(",")).join(" "), fill: "none", stroke: color, "stroke-width": 2.5 }));
                    pts.forEach((p) => svg.append(sv("circle", { cx: p[0], cy: p[1], r: 3.5, fill: color })));
                }
            });
            d.series.forEach((s, si) => {
                const x = L + si * 150;
                svg.append(sv("rect", { x, y: H - 18, width: 10, height: 10, fill: colors[si % colors.length] }), sv("text", { x: x + 14, y: H - 9, "font-size": 11, fill: "#333" }, s.name + (s.axis === "secondary" ? " (sec.)" : "")));
            });
            return svg;
        }

        function reset() {
            snap();
            setState(SE_normalize(task.setup));
            sel = { a: { c: 1, r: 1 }, b: { c: 1, r: 1 } };
            fireChange(); refresh(); say("Spreadsheet reset to its initial state.");
        }

        refresh();
        return { getState: () => clone(st), reset, destroy() { document.removeEventListener("mouseup", onUp); host.innerHTML = ""; } };
    }

    // ana səhifədə tapşırığın cavab vəziyyətini yoxlamaq üçün kiçik köməkçi
    function isChanged(task, stateStr) {
        if (!stateStr) return false;
        try { const s = JSON.parse(stateStr); return JSON.stringify(s) !== JSON.stringify({ cells: task.setup.cells || {}, fmt: task.setup.fmt || {}, dv: [], cf: [], pivots: [], charts: [] }); } catch (e) { return false; }
    }
    root.SheetUI = { mount, isChanged };
})(typeof self !== "undefined" ? self : this);
