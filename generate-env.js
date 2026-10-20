// Vercel build addımı: env.js faylını Vercel mühit dəyişənlərindən yaradır.
// Lokal istifadə:  SUPABASE_URL=... SUPABASE_ANON_KEY=... ADMIN_PASSWORD_HASH=... node generate-env.js
// Şifrənin heşini almaq:  node generate-env.js --hash "MƏNİM_ŞİFRƏM"
const fs = require("fs");
const crypto = require("crypto");

if (process.argv[2] === "--hash") {
    console.log(crypto.createHash("sha256").update(String(process.argv[3] || ""), "utf8").digest("hex"));
    process.exit(0);
}

// Vercel-də yalnız ADMIN_PASSWORD yazmaq kifayətdir — heş burada (Vercel serverində) hesablanır.
if (!process.env.ADMIN_PASSWORD_HASH && process.env.ADMIN_PASSWORD) {
    process.env.ADMIN_PASSWORD_HASH = crypto.createHash("sha256").update(process.env.ADMIN_PASSWORD, "utf8").digest("hex");
}
const need = ["SUPABASE_URL", "SUPABASE_ANON_KEY", "ADMIN_PASSWORD_HASH"];
const missing = need.filter((k) => !process.env[k]);
if (missing.length) {
    console.error("XƏTA: Vercel → Settings → Environment Variables bölməsində bunlar yoxdur: " + missing.map((k) => (k === "ADMIN_PASSWORD_HASH" ? "ADMIN_PASSWORD" : k)).join(", "));
    process.exit(1);
}
const env = Object.fromEntries(need.map((k) => [k, process.env[k].trim()]));
fs.writeFileSync("env.js", "window.__ENV = " + JSON.stringify(env, null, 4) + ";\n", "utf8");
console.log("env.js yaradıldı (" + need.join(", ") + ")");
