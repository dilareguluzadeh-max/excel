#!/usr/bin/env node
/* SDF 1 — independent re-grading of the results on the instructor's computer.
 *
 * In the portal the score is computed in the student's browser. This script takes the ANSWERS stored in the database
 * and grades them again with the same grader (sheetengine.js + grader.js + exam_data.js), then compares with the portal's score.
 *
 * Usage (Node.js 18+):
 *   node regrade.js --password INSTRUCTOR_PASSWORD          # read from Supabase (URL/KEY from .env)
 *   node regrade.js --json neticeler.json              # read from a JSON file
 *   node regrade.js --password ... --csv yeni_ballar.csv
 */
"use strict";
const fs = require("fs"), path = require("path");
const Grader = require("./grader.js");

const HERE = __dirname;
function loadEnv() {
    const p = path.join(HERE, ".env");
    if (!fs.existsSync(p)) return;
    for (const line of fs.readFileSync(p, "utf8").split(/\r?\n/)) {
        const t = line.trim(); if (!t || t[0] === "#" || !t.includes("=")) continue;
        const i = t.indexOf("="), k = t.slice(0, i).trim(), v = t.slice(i + 1).trim().replace(/^['"]|['"]$/g, "");
        if (k && !(k in process.env)) process.env[k] = v;
    }
}
loadEnv();
const arg = (n, d) => { const i = process.argv.indexOf("--" + n); return i > 0 ? process.argv[i + 1] : d; };
const PASSWORD = arg("password", process.env.ADMIN_PASSWORD || "");
const URL_ = arg("url", process.env.SUPABASE_URL || ""), KEY = arg("key", process.env.SUPABASE_ANON_KEY || "");
const JSON_FILE = arg("json", ""), CSV = arg("csv", "");

function loadExamData() {
    const text = fs.readFileSync(path.join(HERE, "exam_data.js"), "utf8");
    return JSON.parse(text.slice(text.indexOf("=") + 1).trim().replace(/;$/, ""));
}
async function fetchSubmissions() {
    if (!PASSWORD) throw new Error("Instructor password required: --password or ADMIN_PASSWORD in .env");
    if (!URL_ || !KEY) throw new Error("Set SUPABASE_URL and SUPABASE_ANON_KEY in the .env file");
    const r = await fetch(`${URL_}/rest/v1/rpc/admin_list_submissions`, {
        method: "POST", headers: { apikey: KEY, Authorization: `Bearer ${KEY}`, "Content-Type": "application/json" }, body: JSON.stringify({ p_password: PASSWORD }),
    });
    if (!r.ok) throw new Error(`Supabase error ${r.status}: ${await r.text()}`);
    return r.json();
}

(async () => {
    const data = loadExamData();
    const subs = JSON_FILE ? JSON.parse(fs.readFileSync(JSON_FILE, "utf8")) : await fetchSubmissions();
    const rows = [];
    for (const s of subs) {
        const letter = String(s.variant || "").replace("Variant", "").trim();
        const variant = data.variants[letter];
        if (!variant || !s.answers || typeof s.answers !== "object") continue;
        const res = Grader.gradeVariant(variant, s.answers).results;
        const scores = {}; let total = 0;
        for (const t of variant.tasks) { scores[t.id] = res[t.id].pts; total += res[t.id].pts; }
        if (String(s.status || "").startsWith("KƏNARLAŞDIRILDI")) total = 0;
        rows.push({ seat: s.student_id ?? "", name: s.student_name ?? "", group: s.student_group ?? "", variant: letter, status: s.status ?? "", portal_score: s.score ?? 0, new_score: total, diff: total - Number(s.score || 0), ...scores });
    }
    console.log("Seat".padEnd(6) + "Name".padEnd(26) + "Var.".padEnd(6) + "Portal".padStart(8) + "New".padStart(7) + "Diff".padStart(7) + "  Status");
    console.log("-".repeat(80));
    for (const r of rows) console.log(String(r.seat).padEnd(6) + String(r.name).slice(0, 25).padEnd(26) + r.variant.padEnd(6) + String(r.portal_score).padStart(8) + String(r.new_score).padStart(7) + (r.diff > 0 ? "+" : "") + String(r.diff).padStart(r.diff > 0 ? 6 : 7) + "  " + r.status + (r.diff ? "   ⚠️" : ""));
    console.log("-".repeat(80));
    console.log(`Total ${rows.length} results; ${rows.filter((r) => r.diff).length} differ from the portal score.`);
    if (CSV && rows.length) {
        const keys = Object.keys(rows[0]);
        fs.writeFileSync(CSV, "﻿" + [keys.join(";"), ...rows.map((r) => keys.map((k) => r[k]).join(";"))].join("\n"));
        console.log("📄 CSV saved: " + CSV);
    }
})().catch((e) => { console.error("Error:", e.message); process.exit(1); });
