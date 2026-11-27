const http = require("http");
const fs = require("fs");
const path = require("path");

const HOST = "127.0.0.1";
const PORT = 3000;

const DATA_DIR = "F:\\CITIZEN_DEV\\HTML\\DEVIS\\DATA";
const DB_FILE = path.join(DATA_DIR, "excelDB.json");
const CATALOG_DIR = path.join(DATA_DIR, "CATALOG");

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(CATALOG_DIR)) fs.mkdirSync(CATALOG_DIR, { recursive: true });

function sendJson(res, statusCode, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  });
  res.end(body);
}

function safeName(s) {
  return String(s || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim();
}

function collectBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", chunk => { body += chunk; });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

const server = http.createServer(async (req, res) => {
  // CORS preflight
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
    return res.end();
  }

  // GET /api/load-db
  if (req.method === "GET" && req.url === "/api/load-db") {
    try {
      if (!fs.existsSync(DB_FILE)) {
        return sendJson(res, 200, { ok: true, db: null });
      }
      const raw = fs.readFileSync(DB_FILE, "utf-8");
      const db = JSON.parse(raw);
      return sendJson(res, 200, { ok: true, db });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: e.message });
    }
  }

  // POST /api/save-db
  if (req.method === "POST" && req.url === "/api/save-db") {
    try {
      const body = await collectBody(req);
      const obj = body ? JSON.parse(body) : {};
      fs.writeFileSync(DB_FILE, JSON.stringify(obj, null, 2), "utf-8");
      return sendJson(res, 200, { ok: true, file: DB_FILE });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: e.message });
    }
  }

  // POST /api/save-supplier-file
  // body: { supplier, fileName, content }
  if (req.method === "POST" && req.url === "/api/save-supplier-file") {
    try {
      const body = await collectBody(req);
      const obj = body ? JSON.parse(body) : {};

      const supplier = safeName((obj.supplier || "INCONNU")).toUpperCase() || "INCONNU";
      const fileName = safeName(obj.fileName || `import_${Date.now()}.json`) || `import_${Date.now()}.json`;
      const content = obj.content || {};

      const supplierDir = path.join(CATALOG_DIR, supplier);
      if (!fs.existsSync(supplierDir)) fs.mkdirSync(supplierDir, { recursive: true });

      const baseNoExt = fileName.replace(/\.[^.]+$/, "");
      const targetFile = path.join(supplierDir, `${baseNoExt}.json`);

      fs.writeFileSync(targetFile, JSON.stringify(content, null, 2), "utf-8");
      return sendJson(res, 200, { ok: true, file: targetFile });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: e.message });
    }
  }

  // GET /api/list-supplier-files?supplier=AXIANS
  if (req.method === "GET" && req.url.startsWith("/api/list-supplier-files")) {
    try {
      const u = new URL(req.url, "http://127.0.0.1");
      const supplier = safeName((u.searchParams.get("supplier") || "").trim()).toUpperCase();
      if (!supplier) return sendJson(res, 400, { ok: false, error: "supplier requis" });

      const supplierDir = path.join(CATALOG_DIR, supplier);
      if (!fs.existsSync(supplierDir)) return sendJson(res, 200, { ok: true, files: [] });

      const files = fs.readdirSync(supplierDir)
        .filter(f => f.toLowerCase().endsWith(".json"))
        .sort((a, b) => a.localeCompare(b));

      return sendJson(res, 200, { ok: true, files });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: e.message });
    }
  }

  // GET /api/load-supplier-catalog?supplier=AXIANS
  if (req.method === "GET" && req.url.startsWith("/api/load-supplier-catalog")) {
    try {
      const u = new URL(req.url, "http://127.0.0.1");
      const supplierRaw = (u.searchParams.get("supplier") || "").trim();
      const supplier = safeName(supplierRaw).toUpperCase();

      if (!supplier) {
        return sendJson(res, 400, { ok: false, error: "supplier requis" });
      }

      const supplierDir = path.join(CATALOG_DIR, supplier);

      if (!fs.existsSync(supplierDir)) {
        return sendJson(res, 200, { ok: true, db: { [supplier]: {} }, files: [] });
      }

      const files = fs.readdirSync(supplierDir)
        .filter(f => f.toLowerCase().endsWith(".json"))
        .sort((a, b) => a.localeCompare(b));

      const merged = { [supplier]: {} };

      for (const f of files) {
        const p = path.join(supplierDir, f);
        let obj = {};
        try {
          obj = JSON.parse(fs.readFileSync(p, "utf-8"));
        } catch {
          continue; // ignore invalid JSON file
        }

        // Cas 1: { AXIANS: { CODE:[...] } }
        if (obj && obj[supplier] && typeof obj[supplier] === "object") {
          for (const code of Object.keys(obj[supplier])) {
            if (!merged[supplier][code]) merged[supplier][code] = [];
            const arr = Array.isArray(obj[supplier][code]) ? obj[supplier][code] : [];
            merged[supplier][code].push(...arr);
          }
          continue;
        }

        // Cas 2: { CODE:[...] }
        if (obj && typeof obj === "object") {
          for (const code of Object.keys(obj)) {
            if (!merged[supplier][code]) merged[supplier][code] = [];
            const arr = Array.isArray(obj[code]) ? obj[code] : [];
            merged[supplier][code].push(...arr);
          }
        }
      }

      return sendJson(res, 200, { ok: true, db: merged, files });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: e.message });
    }
  }

  // 404
  sendJson(res, 404, { ok: false, error: "Route not found" });
});

server.listen(PORT, HOST, () => {
  console.log(`API JSON démarrée sur http://${HOST}:${PORT}`);
  console.log(`DB globale : ${DB_FILE}`);
  console.log(`Catalog fournisseurs : ${CATALOG_DIR}`);
});
