const http = require("http");
const fs = require("fs");
const path = require("path");

const HOST = "127.0.0.1";
const PORT = 3000;

const DATA_DIR = "C:\\CITIZEN_DEV\\HTML\\DEVIS\\DATA";
const DB_FILE = path.join(DATA_DIR, "excelDB.json");
const CATALOG_DIR = path.join(DATA_DIR, "CATALOG");

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

if (!fs.existsSync(CATALOG_DIR)) {
  fs.mkdirSync(CATALOG_DIR, { recursive: true });
}

function sendJson(res, statusCode, obj) {
  const body = JSON.stringify(obj);

  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store"
  });

  res.end(body);
}

function safeName(value) {
  return String(value || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .trim();
}

function collectBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";

    req.on("data", chunk => {
      body += chunk;
    });

    req.on("end", () => {
      resolve(body);
    });

    req.on("error", reject);
  });
}

const server = http.createServer(async (req, res) => {
  /*
   * CORS preflight
   */
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });

    return res.end();
  }

  /*
   * Contrôle de disponibilité de l'API
   * GET /api/health
   */
  if (req.method === "GET" && req.url === "/api/health") {
    return sendJson(res, 200, {
      ok: true,
      service: "comparateur-pdf-excel",
      timestamp: new Date().toISOString()
    });
  }

  /*
   * GET /api/load-db
   */
  if (req.method === "GET" && req.url === "/api/load-db") {
    try {
      if (!fs.existsSync(DB_FILE)) {
        return sendJson(res, 200, {
          ok: true,
          db: null
        });
      }

      const raw = fs.readFileSync(DB_FILE, "utf-8");
      const db = JSON.parse(raw);

      return sendJson(res, 200, {
        ok: true,
        db
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  /*
   * POST /api/save-db
   */
  if (req.method === "POST" && req.url === "/api/save-db") {
    try {
      const body = await collectBody(req);
      const obj = body ? JSON.parse(body) : {};

      fs.writeFileSync(
        DB_FILE,
        JSON.stringify(obj, null, 2),
        "utf-8"
      );

      return sendJson(res, 200, {
        ok: true,
        file: DB_FILE
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  /*
   * POST /api/save-supplier-file
   *
   * Body :
   * {
   *   supplier,
   *   fileName,
   *   content
   * }
   */
  if (
    req.method === "POST" &&
    req.url === "/api/save-supplier-file"
  ) {
    try {
      const body = await collectBody(req);
      const obj = body ? JSON.parse(body) : {};

      const supplier =
        safeName(obj.supplier || "INCONNU")
          .toUpperCase() || "INCONNU";

      const fileName =
        safeName(
          obj.fileName || `import_${Date.now()}.json`
        ) || `import_${Date.now()}.json`;

      const content = obj.content || {};

      const supplierDir = path.join(
        CATALOG_DIR,
        supplier
      );

      if (!fs.existsSync(supplierDir)) {
        fs.mkdirSync(supplierDir, {
          recursive: true
        });
      }

      const baseNoExt = fileName.replace(
        /\.[^.]+$/,
        ""
      );

      const targetFile = path.join(
        supplierDir,
        `${baseNoExt}.json`
      );

      fs.writeFileSync(
        targetFile,
        JSON.stringify(content, null, 2),
        "utf-8"
      );

      return sendJson(res, 200, {
        ok: true,
        file: targetFile
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  /*
   * GET /api/list-supplier-files?supplier=AXIANS
   */
  if (
    req.method === "GET" &&
    req.url.startsWith("/api/list-supplier-files")
  ) {
    try {
      const url = new URL(
        req.url,
        `http://${HOST}`
      );

      const supplier = safeName(
        url.searchParams.get("supplier") || ""
      ).toUpperCase();

      if (!supplier) {
        return sendJson(res, 400, {
          ok: false,
          error: "supplier requis"
        });
      }

      const supplierDir = path.join(
        CATALOG_DIR,
        supplier
      );

      if (!fs.existsSync(supplierDir)) {
        return sendJson(res, 200, {
          ok: true,
          files: []
        });
      }

      const files = fs.readdirSync(supplierDir)
        .filter(file =>
          file.toLowerCase().endsWith(".json")
        )
        .sort((a, b) => a.localeCompare(b));

      return sendJson(res, 200, {
        ok: true,
        files
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  /*
   * GET /api/load-supplier-catalog?supplier=AXIANS
   */
  if (
    req.method === "GET" &&
    req.url.startsWith("/api/load-supplier-catalog")
  ) {
    try {
      const url = new URL(
        req.url,
        `http://${HOST}`
      );

      const supplierRaw =
        url.searchParams.get("supplier") || "";

      const supplier = safeName(
        supplierRaw.trim()
      ).toUpperCase();

      if (!supplier) {
        return sendJson(res, 400, {
          ok: false,
          error: "supplier requis"
        });
      }

      const supplierDir = path.join(
        CATALOG_DIR,
        supplier
      );

      if (!fs.existsSync(supplierDir)) {
        return sendJson(res, 200, {
          ok: true,
          db: {
            [supplier]: {}
          },
          files: []
        });
      }

      const files = fs.readdirSync(supplierDir)
        .filter(file =>
          file.toLowerCase().endsWith(".json")
        )
        .sort((a, b) => a.localeCompare(b));

      const merged = {
        [supplier]: {}
      };

      for (const file of files) {
        const filePath = path.join(
          supplierDir,
          file
        );

        let obj = {};

        try {
          obj = JSON.parse(
            fs.readFileSync(filePath, "utf-8")
          );
        } catch (error) {
          console.warn(
            `Fichier JSON ignoré : ${file}`
          );
          continue;
        }

        /*
         * Cas 1 :
         * {
         *   AXIANS: {
         *     CODE: [...]
         *   }
         * }
         */
        if (
          obj &&
          obj[supplier] &&
          typeof obj[supplier] === "object"
        ) {
          for (const code of Object.keys(
            obj[supplier]
          )) {
            if (!merged[supplier][code]) {
              merged[supplier][code] = [];
            }

            const values = Array.isArray(
              obj[supplier][code]
            )
              ? obj[supplier][code]
              : [];

            merged[supplier][code].push(...values);
          }

          continue;
        }

        /*
         * Cas 2 :
         * {
         *   CODE: [...]
         * }
         */
        if (obj && typeof obj === "object") {
          for (const code of Object.keys(obj)) {
            if (!merged[supplier][code]) {
              merged[supplier][code] = [];
            }

            const values = Array.isArray(obj[code])
              ? obj[code]
              : [];

            merged[supplier][code].push(...values);
          }
        }
      }

      return sendJson(res, 200, {
        ok: true,
        db: merged,
        files
      });
    } catch (error) {
      return sendJson(res, 500, {
        ok: false,
        error: error.message
      });
    }
  }

  /*
   * Route inexistante
   */
  return sendJson(res, 404, {
    ok: false,
    error: "Route not found"
  });
});

server.listen(PORT, HOST, () => {
  console.log(
    `API JSON démarrée sur http://${HOST}:${PORT}`
  );

  console.log(`Contrôle API : http://${HOST}:${PORT}/api/health`);
  console.log(`DB globale : ${DB_FILE}`);
  console.log(`Catalog fournisseurs : ${CATALOG_DIR}`);
});
