"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const HOST = "0.0.0.0";
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, "DATA");
const DB_FILE = path.join(DATA_DIR, "excelDB.json");
const CATALOG_DIR = path.join(DATA_DIR, "CATALOG");
const BACKUP_DIR = path.join(DATA_DIR, "BACKUPS");
const MAX_BODY_BYTES = 50 * 1024 * 1024;

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(CATALOG_DIR, { recursive: true });
fs.mkdirSync(BACKUP_DIR, { recursive: true });

function sendJson(res, status, data) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store"
  });
  res.end(JSON.stringify(data));
}

function fail(message, status = 400) {
  const error = new Error(message);
  error.status = status;
  throw error;
}

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function validateName(value, label) {
  if (typeof value !== "string") {
    fail(`${label} requis`);
  }

  const name = value.trim();

  if (
    !name ||
    name === "." ||
    name === ".." ||
    name.length > 180 ||
    /[\\/:*?"<>|\x00-\x1F]/.test(name) ||
    /[. ]$/.test(name) ||
    /^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(name) ||
    ["__proto__", "prototype", "constructor"].includes(name.toLowerCase())
  ) {
    fail(`${label} invalide`);
  }

  return name;
}

function supplierName(value) {
  return validateName(value, "Fournisseur").toUpperCase();
}

function supplierDirectory(supplier) {
  const directory = path.resolve(CATALOG_DIR, supplier);

  if (path.dirname(directory) !== path.resolve(CATALOG_DIR)) {
    fail("Chemin fournisseur interdit");
  }

  if (fs.existsSync(directory)) {
    const stat = fs.lstatSync(directory);
    if (!stat.isDirectory() || stat.isSymbolicLink()) {
      fail("Dossier fournisseur invalide");
    }
  }

  return directory;
}

function catalogPath(supplier, value) {
  const fileName = validateName(value, "Nom du catalogue");

  if (!fileName.toLowerCase().endsWith(".json")) {
    fail("Le catalogue doit être un fichier .json");
  }

  const directory = supplierDirectory(supplier);
  const target = path.resolve(directory, fileName);

  if (path.dirname(target) !== directory) {
    fail("Chemin catalogue interdit");
  }

  return target;
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, "utf8").replace(/^\uFEFF/, ""));
}

function readGlobalDB() {
  if (!fs.existsSync(DB_FILE)) return {};

  const db = readJson(DB_FILE);
  if (!isObject(db)) {
    fail("excelDB.json est invalide", 500);
  }

  return db;
}

function atomicWriteJson(file, data) {
  const temporary = `${file}.${crypto.randomUUID()}.tmp`;

  try {
    fs.writeFileSync(
      temporary,
      JSON.stringify(data, null, 2),
      { encoding: "utf8", flag: "wx" }
    );
    fs.renameSync(temporary, file);
  } finally {
    if (fs.existsSync(temporary)) {
      fs.unlinkSync(temporary);
    }
  }
}

function collectBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    let exceeded = false;

    req.on("data", chunk => {
      size += chunk.length;

      if (size > MAX_BODY_BYTES) {
        if (!exceeded) {
          exceeded = true;
          const error = new Error("Requête trop volumineuse.");
          error.status = 413;
          reject(error);
        }
        return;
      }

      chunks.push(chunk);
    });

    req.on("end", () => {
      if (!exceeded) {
        resolve(Buffer.concat(chunks).toString("utf8"));
      }
    });

    req.on("error", reject);
  });
}

async function requestJson(req) {
  const body = await collectBody(req);

  try {
    const data = JSON.parse(body || "{}");
    if (!isObject(data)) fail("Objet JSON attendu");
    return data;
  } catch (error) {
    if (error.status) throw error;
    fail("Corps JSON invalide");
  }
}

function listSupplierFolders() {
  return fs.readdirSync(CATALOG_DIR, { withFileTypes: true })
    .filter(entry => entry.isDirectory() && !entry.isSymbolicLink())
    .map(entry => entry.name)
    .sort((a, b) => a.localeCompare(b));
}

function listSupplierFiles(supplier) {
  const directory = supplierDirectory(supplier);
  if (!fs.existsSync(directory)) return [];

  return fs.readdirSync(directory, { withFileTypes: true })
    .filter(entry =>
      entry.isFile() &&
      !entry.isSymbolicLink() &&
      entry.name.toLowerCase().endsWith(".json")
    )
    .map(entry => entry.name)
    .sort((a, b) => a.localeCompare(b));
}

function catalogEntries(document, supplier) {
  if (!isObject(document)) {
    fail("Structure du catalogue JSON invalide", 422);
  }

  const entries = Object.prototype.hasOwnProperty.call(document, supplier)
    ? document[supplier]
    : document;

  if (!isObject(entries)) {
    fail("Les références du catalogue doivent former un objet", 422);
  }

  for (const [code, values] of Object.entries(entries)) {
    if (!Array.isArray(values) || values.some(value => !isObject(value))) {
      fail(`Catalogue invalide pour la référence ${code}`, 422);
    }
  }

  return entries;
}

function safeReadCatalog(filePath) {
  try {
    const document = readJson(filePath);
    return document;
  } catch (error) {
    return { __error: error.message };
  }
}

function rebuildSupplierFromFiles(supplier) {
  const files = listSupplierFiles(supplier);
  const merged = Object.create(null);
  const ignoredFiles = [];

  for (const fileName of files) {
    const filePath = catalogPath(supplier, fileName);
    const document = safeReadCatalog(filePath);

    if (document.__error) {
      ignoredFiles.push({ file: fileName, error: document.__error });
      continue;
    }

    let entries;

    try {
      entries = catalogEntries(document, supplier);
    } catch (error) {
      ignoredFiles.push({ file: fileName, error: error.message });
      continue;
    }

    for (const [code, values] of Object.entries(entries)) {
      if (!merged[code]) merged[code] = [];
      merged[code].push(...values);
    }
  }

  return { entries: merged, files, ignoredFiles };
}

function rebuildGlobalDB() {
  const db = {};
  const ignored = [];

  for (const supplier of listSupplierFolders()) {
    const rebuilt = rebuildSupplierFromFiles(supplier);
    db[supplier] = rebuilt.entries;
    if (rebuilt.ignoredFiles.length) {
      ignored.push({ supplier, files: rebuilt.ignoredFiles });
    }
  }

  return { db, ignored };
}

function writeGlobalDB(db) {
  atomicWriteJson(DB_FILE, db);
}

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
    return res.end();
  }

  try {
    const url = new URL(req.url, `http://${HOST}:${PORT}`);
    const route = url.pathname;

    if (req.method === "GET" && route === "/api/health") {
      return sendJson(res, 200, {
        ok: true,
        service: "comparateur-pdf-excel",
        timestamp: new Date().toISOString()
      });
    }

    if (req.method === "GET" && route === "/api/load-db") {
      return sendJson(res, 200, {
        ok: true,
        db: readGlobalDB()
      });
    }

    if (req.method === "POST" && route === "/api/save-db") {
      const data = await requestJson(req);
      writeGlobalDB(data);

      return sendJson(res, 200, {
        ok: true,
        file: DB_FILE
      });
    }

    if (req.method === "GET" && route === "/api/list-suppliers") {
      return sendJson(res, 200, {
        ok: true,
        suppliers: listSupplierFolders()
      });
    }

    if (req.method === "GET" && route === "/api/list-supplier-files") {
      const supplier = supplierName(url.searchParams.get("supplier"));
      return sendJson(res, 200, {
        ok: true,
        files: listSupplierFiles(supplier)
      });
    }

    if (req.method === "GET" && route === "/api/load-supplier-catalog") {
      const supplier = supplierName(url.searchParams.get("supplier"));
      const rebuilt = rebuildSupplierFromFiles(supplier);

      return sendJson(res, 200, {
        ok: true,
        db: { [supplier]: rebuilt.entries },
        files: rebuilt.files,
        ignoredFiles: rebuilt.ignoredFiles
      });
    }

    if (req.method === "POST" && route === "/api/save-supplier-file") {
      const data = await requestJson(req);
      const supplier = supplierName(data.supplier);
      const target = catalogPath(supplier, data.fileName);

      let content = data.content;

      if (typeof content === "string") {
        try {
          content = JSON.parse(content);
        } catch {
          fail("Le contenu du catalogue n'est pas un JSON valide");
        }
      }

      const entries = catalogEntries(content, supplier);
      const directory = supplierDirectory(supplier);
      fs.mkdirSync(directory, { recursive: true });

      const existed = fs.existsSync(target);
      const previous = existed ? fs.readFileSync(target) : null;

      try {
        atomicWriteJson(target, { [supplier]: entries });

        const rebuilt = rebuildGlobalDB();
        writeGlobalDB(rebuilt.db);

        return sendJson(res, 200, {
          ok: true,
          file: target,
          db: rebuilt.db,
          ignoredFiles: rebuilt.ignored
        });
      } catch (error) {
        try {
          if (existed) fs.writeFileSync(target, previous);
          else if (fs.existsSync(target)) fs.unlinkSync(target);
        } catch (rollbackError) {
          console.error("Échec restauration catalogue :", rollbackError);
        }
        throw error;
      }
    }

    if (req.method === "POST" && route === "/api/delete-supplier-file") {
      const data = await requestJson(req);
      const supplier = supplierName(data.supplier);
      const target = catalogPath(supplier, data.fileName);

      if (!fs.existsSync(target)) {
        fail("Fichier catalogue introuvable", 404);
      }

      const backupName = `${path.basename(target, ".json")}.${crypto.randomUUID()}.deleted.json`;
      const backupPath = path.join(BACKUP_DIR, backupName);

      fs.copyFileSync(target, backupPath);
      fs.unlinkSync(target);

      const rebuilt = rebuildGlobalDB();
      writeGlobalDB(rebuilt.db);

      return sendJson(res, 200, {
        ok: true,
        file: target,
        db: rebuilt.db,
        backup: backupPath,
        ignoredFiles: rebuilt.ignored
      });
    }

    return sendJson(res, 404, {
      ok: false,
      error: "Route introuvable"
    });
  } catch (error) {
    console.error(error);
    return sendJson(res, error.status || 500, {
      ok: false,
      error: error.message
    });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`API démarrée : http://${HOST}:${PORT}`);
  console.log(`Contrôle API : http://${HOST}:${PORT}/api/health`);
  console.log(`Base globale : ${DB_FILE}`);
  console.log(`Catalogues : ${CATALOG_DIR}`);
});
