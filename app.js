"use strict";

/* =========================================================
   Configuration et état
   ========================================================= */

const CONFIG = {
  API_BASE: "http://127.0.0.1:3000",
  API_TIMEOUT: 30000,
  API_CHECK_INTERVAL: 10000,
  STORAGE_KEYS: {
    SUPPLIERS: "devis_suppliers_v1",
    CACHE: "devis_cache_v1"
  },
  DEFAULT_SUPPLIERS: ["CGTI", "NEWCOM", "AXIANS"],
  DEBUG_CGTI: true
};

const STATE = {
  jsonDB: null,
  excelWorkbooks: [],
  suppliersList: [],
  lastAnalysisRows: [],
  lastAnalysisFileName: "",
  apiOnline: false,
  cache: new Map(),
  initialized: false,
  healthTimer: null,
  healthChecking: false,
  catalogRequestId: 0,
  excelRequestId: 0
};

/* =========================================================
   Utilitaires
   ========================================================= */

const Utils = {
  escapeHtml: (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;"),

  stripAccents: (value) => String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""),

  normalizeSupplierName: (value) => String(value ?? "")
    .trim()
    .toUpperCase(),

  normalizeKey: (value) => Utils.stripAccents(value)
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim(),

  compactKey: (value) => Utils.normalizeKey(value)
    .replace(/[^A-Z0-9]/g, ""),

  normalizeCode: (value) => String(value ?? "")
    .trim()
    .toUpperCase()
    .replace(/[\s_-]/g, ""),

  canonicalUnit: (value) => {
    const key = Utils.normalizeKey(value)
      .replace(/²/g, "2")
      .replace(/³/g, "3")
      .replace(/\./g, "")
      .replace(/\s+/g, "");

    const aliases = {
      U: "U",
      UNIT: "U",
      UNITE: "U",
      UNITES: "U",
      F: "FORFAIT",
      FORFAIT: "FORFAIT",
      FORFAITS: "FORFAIT",
      ML: "ML",
      M: "M",
      M2: "M2",
      M3: "M3",
      KG: "KG",
      G: "G",
      L: "L",
      H: "H",
      J: "J",
      JH: "JH",
      JOUR: "J",
      JOURS: "J",
      MOIS: "MOIS",
      ENS: "ENS",
      ENSEMBLE: "ENS",
      "%": "%",
      SURDEVIS: "SUR DEVIS"
    };

    return Object.hasOwn(aliases, key) ? aliases[key] : "";
  },

  normalizeUnit: (value) => {
    const canonical = Utils.canonicalUnit(value);
    if (canonical) return canonical;

    return Utils.normalizeKey(value)
      .replace(/²/g, "2")
      .replace(/³/g, "3")
      .replace(/\s+/g, " ")
      .trim();
  },

  parsePrice: (value) => {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : NaN;
    }

    if (value === null || value === undefined) return NaN;

    let text = String(value)
      .trim()
      .replace(/\u2212/g, "-")
      .replace(/€/g, "")
      .replace(/\bEUR\b/gi, "")
      .replace(/\s+/g, "");

    if (!text) return NaN;

    // Deux séparateurs : le dernier est le séparateur décimal.
    if (text.includes(",") && text.includes(".")) {
      if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
        text = text.replace(/\./g, "").replace(",", ".");
      } else {
        text = text.replace(/,/g, "");
      }
    } else if (text.includes(",")) {
      text = text.replace(",", ".");
    }

    if (!/^[+-]?\d+(?:\.\d+)?$/.test(text)) return NaN;

    const number = Number(text);
    return Number.isFinite(number) ? number : NaN;
  },

  parsePdfNumber: (value) => {
    const text = String(value ?? "").trim();
    const isPercentage = text.endsWith("%");
    const number = Utils.parsePrice(
      isPercentage ? text.slice(0, -1) : text
    );

    return isPercentage ? number / 100 : number;
  },

  sanitizeFileName: (value) => String(value ?? "")
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/^_+|_+$/g, ""),

  createTimestamp: () => {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, "0");

    return (
      `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
      `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`
    );
  },

  csvEscape: (value) => {
    let text = String(value ?? "");

    // Évite l'interprétation de texte comme formule dans Excel.
    if (/^[\s]*[=+@-]/.test(text) && !Number.isFinite(Utils.parsePrice(text))) {
      text = `'${text}`;
    }

    return /[;"\r\n]/.test(text)
      ? `"${text.replace(/"/g, '""')}"`
      : text;
  },

  formatPrice: (value) => Number.isFinite(value)
    ? value.toFixed(2).replace(".", ",")
    : "",

  isRecord: (value) => value !== null &&
    typeof value === "object" &&
    !Array.isArray(value),

  errorMessage: (error) => error instanceof Error
    ? error.message
    : String(error),

  uniqueRows: (rows) => {
    const seen = new Set();

    return rows.filter(row => {
      const key = [
        row.page, row.y, row.code, row.unit, row.pu, row.raw
      ].join("|");

      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
};

/* =========================================================
   Notifications
   ========================================================= */

const Toast = {
  show: (message, type = "info", duration = 3500) => {
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);

    setTimeout(() => {
      toast.style.animation = "slideOut 0.3s ease";
      setTimeout(() => toast.remove(), 300);
    }, duration);
  },

  success: (message) => Toast.show(message, "success"),
  error: (message) => Toast.show(message, "error", 6000),
  info: (message) => Toast.show(message, "info")
};

/* =========================================================
   API
   ========================================================= */

const API = {
  setStatus: (isOnline) => {
    STATE.apiOnline = isOnline;

    const status = document.getElementById("apiStatus");
    const text = document.getElementById("apiStatusText");

    if (status) {
      status.classList.toggle("online", isOnline);
      status.classList.toggle("offline", !isOnline);
    }

    if (text) {
      text.textContent = isOnline ? "API en ligne" : "API hors ligne";
    }
  },

  checkHealth: async () => {
    if (STATE.healthChecking) return;
    STATE.healthChecking = true;

    const controller = new AbortController();
    const timer = setTimeout(
      () => controller.abort(),
      CONFIG.API_TIMEOUT
    );

    try {
      const response = await window.fetch(
        `${CONFIG.API_BASE}/api/health`,
        { cache: "no-store", signal: controller.signal }
      );

      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const data = await response.json();
      API.setStatus(data?.ok === true);
    } catch {
      API.setStatus(false);
    } finally {
      clearTimeout(timer);
      STATE.healthChecking = false;
    }
  },

  startMonitoring: () => {
    if (STATE.healthTimer !== null) return;

    void API.checkHealth();
    STATE.healthTimer = setInterval(
      () => void API.checkHealth(),
      CONFIG.API_CHECK_INTERVAL
    );
  },

  fetch: async (endpoint, options = {}, timeout = CONFIG.API_TIMEOUT) => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeout);

    try {
      const response = await window.fetch(
        `${CONFIG.API_BASE}${endpoint}`,
        {
          ...options,
          cache: "no-store",
          signal: controller.signal
        }
      );

      const text = await response.text();
      let data;

      try {
        data = text ? JSON.parse(text) : {};
      } catch {
        throw new Error(
          `HTTP ${response.status} : réponse du serveur non JSON.`
        );
      }

      if (!response.ok || !Utils.isRecord(data) || data.ok !== true) {
        throw new Error(
          data?.error || `Réponse API invalide (HTTP ${response.status}).`
        );
      }

      API.setStatus(true);
      return data;
    } catch (error) {
      console.error(`Erreur API ${endpoint} :`, error);

      if (error?.name === "AbortError") {
        throw new Error(
          "Délai API dépassé. Actualisez les catalogues pour vérifier " +
          "si l'opération a été effectuée avant de réessayer."
        );
      }

      throw error;
    } finally {
      clearTimeout(timer);
    }
  },

  post: (endpoint, body) => API.fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  }),

  saveDB: (db) => API.post("/api/save-db", db),

  loadDB: () => API.fetch("/api/load-db"),

  saveSupplierFile: (supplier, fileName, content) => API.post(
    "/api/save-supplier-file",
    { supplier, fileName, content }
  ),

  deleteSupplierFile: (supplier, fileName) => API.post(
    "/api/delete-supplier-file",
    { supplier, fileName }
  ),

  loadSupplierCatalog: async (supplier) => {
    const name = encodeURIComponent(Utils.normalizeSupplierName(supplier));
    const data = await API.fetch(
      `/api/load-supplier-catalog?supplier=${name}`
    );

    return Utils.isRecord(data.db) ? data.db : {};
  },

  listSupplierFiles: async (supplier) => {
    const name = encodeURIComponent(Utils.normalizeSupplierName(supplier));
    const data = await API.fetch(
      `/api/list-supplier-files?supplier=${name}`
    );

    return Array.isArray(data.files) ? data.files : [];
  }
};

/* =========================================================
   Stockage
   ========================================================= */

const Storage = {
  loadSuppliers: () => {
    try {
      const raw = localStorage.getItem(CONFIG.STORAGE_KEYS.SUPPLIERS);

      if (raw !== null) {
        const values = JSON.parse(raw);

        if (Array.isArray(values)) {
          STATE.suppliersList = [...new Set(
            values
              .filter(value => typeof value === "string")
              .map(Utils.normalizeSupplierName)
              .filter(Boolean)
          )];
          return;
        }
      }
    } catch (error) {
      console.warn("Lecture du stockage local impossible :", error);
    }

    STATE.suppliersList = [...CONFIG.DEFAULT_SUPPLIERS];
    Storage.saveSuppliers();
  },

  saveSuppliers: () => {
    try {
      localStorage.setItem(
        CONFIG.STORAGE_KEYS.SUPPLIERS,
        JSON.stringify(STATE.suppliersList)
      );
    } catch (error) {
      console.warn("Sauvegarde du stockage local impossible :", error);
    }
  }
};

/* =========================================================
   Interface
   ========================================================= */

const UI = {
  html: (id, content) => {
    const element = document.getElementById(id);
    if (element) element.innerHTML = content;
  },

  text: (id, content) => {
    const element = document.getElementById(id);
    if (element) element.textContent = content;
  },

  value: (id, fallback = "") => (
    document.getElementById(id)?.value ?? fallback
  ),

  setLoading: (id, loading) => {
    const element = document.getElementById(id);
    if (element) element.disabled = loading;
  },

  showProgress: (message) => {
    const element = document.getElementById("buildProgress");
    if (!element) return;
    element.textContent = message;
    element.style.display = "block";
  },

  hideProgress: () => {
    const element = document.getElementById("buildProgress");
    if (element) element.style.display = "none";
  },

  refreshSupplierSelects: () => {
    for (const id of ["catalogSupplierSelect", "supplierSelect"]) {
      const select = document.getElementById(id);
      if (!select) continue;

      const previous = select.value;
      select.innerHTML = '<option value="">-- choisir --</option>';

      for (const supplier of STATE.suppliersList) {
        const option = document.createElement("option");
        const count = Object.keys(STATE.jsonDB?.[supplier] || {}).length;
        option.value = supplier;
        option.textContent = count ? supplier : `${supplier} (non chargé)`;
        select.appendChild(option);
      }

      if ([...select.options].some(option => option.value === previous)) {
        select.value = previous;
      }
    }
  },

  renderSupplierList: () => {
    const box = document.getElementById("supplierList");
    if (!box) return;

    if (!STATE.suppliersList.length) {
      box.innerHTML = '<span class="warn">Aucun fournisseur.</span>';
      return;
    }

    box.innerHTML = STATE.suppliersList.map((supplier, index) => `
      <div class="sheet-item">
        <div class="top">
          <div class="file-name">${Utils.escapeHtml(supplier)}</div>
          <div class="supplier">
            Codes chargés :
            <b>${Object.keys(STATE.jsonDB?.[supplier] || {}).length}</b>
          </div>
        </div>
        <div class="actions">
          <button type="button" class="secondary del-supplier-btn"
                  data-index="${index}">Supprimer</button>
        </div>
      </div>
    `).join("");

    box.querySelectorAll(".del-supplier-btn").forEach(button => {
      button.addEventListener("click", () => {
        const index = Number(button.dataset.index);
        const supplier = STATE.suppliersList[index];
        if (!supplier) return;

        STATE.suppliersList.splice(index, 1);
        if (STATE.jsonDB) delete STATE.jsonDB[supplier];

        Storage.saveSuppliers();
        UI.refreshAll();
        void UI.renderCatalogFiles();
      });
    });
  },

  renderDbInfo: () => {
    if (!STATE.jsonDB) {
      UI.text("dbInfo", "Base JSON non chargée.");
      return;
    }

    const suppliers = Object.keys(STATE.jsonDB);

    if (!suppliers.length) {
      UI.text("dbInfo", "Base JSON vide.");
      return;
    }

    const badges = suppliers.map(supplier => `
      <span class="pill">
        ${Utils.escapeHtml(supplier)} :
        ${Object.keys(STATE.jsonDB[supplier] || {}).length} codes
      </span>
    `).join("");

    UI.html("dbInfo", `
      <b>Base JSON chargée :</b>
      <div class="badges">${badges}</div>
      <details>
        <summary>Aperçu JSON</summary>
        <pre>${Utils.escapeHtml(
          JSON.stringify(STATE.jsonDB, null, 2).slice(0, 6000)
        )}</pre>
      </details>
    `);
  },

  refreshAll: () => {
    UI.refreshSupplierSelects();
    UI.renderSupplierList();
    UI.renderDbInfo();
  },

  renderCatalogFiles: async () => {
    const box = document.getElementById("catalogFilesInfo");
    if (!box) return;

    const requestId = ++STATE.catalogRequestId;
    const supplier = UI.value("catalogSupplierSelect").trim();

    if (!supplier) {
      box.textContent = "";
      return;
    }

    box.textContent = "Chargement des fichiers catalogue...";

    try {
      const files = await API.listSupplierFiles(supplier);

      if (requestId !== STATE.catalogRequestId) return;

      if (!files.length) {
        box.innerHTML = `
          <span class="warn">
            Aucun fichier catalogue pour ${Utils.escapeHtml(supplier)}.
          </span>
        `;
        return;
      }

      box.innerHTML = `
        <div class="sheet-item">
          <div class="top">
            <div class="file-name">
              Catalogues ${Utils.escapeHtml(supplier)}
            </div>
          </div>
          <div class="actions" style="display:block;">
            ${files.map(file => `
              <div style="display:flex;align-items:center;gap:8px;margin:6px 0;">
                <span>${Utils.escapeHtml(file)}</span>
                <button type="button"
                        class="secondary delete-catalog-btn"
                        data-file="${Utils.escapeHtml(file)}">
                  Supprimer
                </button>
              </div>
            `).join("")}
          </div>
        </div>
      `;

      box.querySelectorAll(".delete-catalog-btn").forEach(button => {
        button.addEventListener("click", async () => {
          const fileName = button.dataset.file;
          if (!window.confirm(`Supprimer le catalogue ${fileName} ?`)) return;

          button.disabled = true;

          try {
            const response = await API.deleteSupplierFile(supplier, fileName);

            if (Utils.isRecord(response.db)) {
              STATE.jsonDB = response.db;
            } else {
              const db = await API.loadSupplierCatalog(supplier);
              STATE.jsonDB ??= {};
              STATE.jsonDB[supplier] = db[supplier] || {};
            }

            UI.refreshAll();
            Toast.success("Catalogue supprimé.");
            await UI.renderCatalogFiles();
          } catch (error) {
            button.disabled = false;
            Toast.error(`Erreur suppression : ${Utils.errorMessage(error)}`);
          }
        });
      });
    } catch (error) {
      if (requestId !== STATE.catalogRequestId) return;

      box.innerHTML = `
        <span class="ko">
          Erreur chargement catalogues :
          ${Utils.escapeHtml(Utils.errorMessage(error))}
        </span>
      `;
    }
  }
};

/* =========================================================
   Import Excel
   ========================================================= */

const Excel = {
  prepare: async (files) => {
    if (!window.XLSX) {
      throw new Error("La bibliothèque XLSX n'est pas chargée.");
    }

    const result = [];

    for (const file of files) {
      const workbook = window.XLSX.read(
        await file.arrayBuffer(),
        { type: "array" }
      );

      result.push({
        file,
        supplier: Excel.getSupplierFromFileName(file.name),
        workbook,
        sheetNames: workbook.SheetNames || [],
        selectedSheet: workbook.SheetNames?.[0] || null
      });
    }

    return result;
  },

  getSupplierFromFileName: (fileName) => {
    const name = Utils.normalizeSupplierName(
      String(fileName || "").replace(/\.[^.]+$/, "")
    );

    const suppliers = [...STATE.suppliersList]
      .sort((a, b) => b.length - a.length);

    return suppliers.find(supplier => name.includes(supplier)) || "INCONNU";
  },

  renderSheetPicker: (workbooks) => {
    const card = document.getElementById("sheetPickerCard");
    const box = document.getElementById("sheetPicker");
    if (!card || !box) return;

    card.style.display = workbooks.length ? "block" : "none";

    box.innerHTML = workbooks.map((info, index) => `
      <div class="sheet-item">
        <div class="top">
          <div class="file-name">${Utils.escapeHtml(info.file.name)}</div>
          <div class="supplier">
            Fournisseur : <b>${Utils.escapeHtml(info.supplier)}</b>
          </div>
        </div>
        <div class="row">
          <div class="field" style="min-width:260px;">
            <label>Feuille à importer</label>
            <select class="sheet-select" data-index="${index}">
              ${info.sheetNames.map(sheet => `
                <option value="${Utils.escapeHtml(sheet)}"
                        ${sheet === info.selectedSheet ? "selected" : ""}>
                  ${Utils.escapeHtml(sheet)}
                </option>
              `).join("")}
            </select>
          </div>
        </div>
      </div>
    `).join("");

    box.querySelectorAll(".sheet-select").forEach(select => {
      select.addEventListener("change", () => {
        const info = STATE.excelWorkbooks[Number(select.dataset.index)];
        if (info) info.selectedSheet = select.value;
      });
    });
  },

  getValByFlexibleKey: (row, wantedKey) => {
    const target = Utils.compactKey(wantedKey);
    const key = Object.keys(row || {})
      .find(candidate => Utils.compactKey(candidate) === target);

    return key === undefined ? "" : row[key];
  },

  findPriceColumn: (row, supplier) => {
    const keys = Object.keys(row || {});
    const normalized = keys.map(Utils.compactKey);
    const supplierKey = Utils.compactKey(supplier);

    if (supplierKey === "CGTI") {
      const index = normalized.indexOf("CGTITARIF2024REUNION");

      if (index < 0) {
        throw new Error(
          'Catalogue CGTI : colonne "CGTI TARIF 2024 REUNION" introuvable.'
        );
      }

      return keys[index];
    }

    let index = normalized.findIndex(key =>
      key.includes(supplierKey) && key.includes("TARIF")
    );

    if (index >= 0) return keys[index];

    const wanted = Utils.compactKey(UI.value("colPU").trim());

    if (wanted) {
      index = normalized.findIndex(key => key === wanted);
      if (index < 0) index = normalized.findIndex(key => key.includes(wanted));
      if (index >= 0) return keys[index];
    }

    index = normalized.findIndex(key =>
      key.includes("TARIF") || key.includes("PRIX") || key.includes("PU")
    );

    return index >= 0 ? keys[index] : null;
  },

  toJsonDB: (workbooks) => {
    const db = Object.create(null);

    for (const info of workbooks) {
      const supplier = Utils.normalizeSupplierName(info.supplier);
      const sheetName = info.selectedSheet;

      if (supplier === "INCONNU") {
        throw new Error(
          `Fournisseur non reconnu pour "${info.file.name}". ` +
          "Le nom du fichier doit contenir le nom du fournisseur."
        );
      }

      const sheet = info.workbook.Sheets[sheetName];
      if (!sheet) throw new Error(`Feuille introuvable : ${sheetName}`);

      const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!rows.length) continue;

      const codeColumn = UI.value("colCode").trim() || "Items";
      const unitColumn = UI.value("colUnit").trim() || "Unité";
      const priceColumn = Excel.findPriceColumn(rows[0], supplier);

      if (!priceColumn) {
        throw new Error(`Colonne prix introuvable pour ${supplier}.`);
      }

      db[supplier] ??= Object.create(null);

      for (const row of rows) {
        const code = Utils.normalizeCode(
          Excel.getValByFlexibleKey(row, codeColumn)
        );

        if (!code) continue;

        const pu = Utils.parsePrice(row[priceColumn]);

        db[supplier][code] ??= [];
        db[supplier][code].push({
          unit: Utils.normalizeUnit(
            Excel.getValByFlexibleKey(row, unitColumn)
          ),
          // JSON ne peut pas stocker NaN.
          pu: Number.isFinite(pu) ? pu : null,
          raw: row,
          sheet: sheetName,
          file: info.file.name,
          priceCol: priceColumn
        });
      }
    }

    if (!Object.values(db).some(catalog => Object.keys(catalog).length)) {
      throw new Error(
        "Aucun code importé. Vérifiez les feuilles et la colonne des codes."
      );
    }

    return db;
  }
};

/* =========================================================
   Outils d'extraction PDF
   ========================================================= */

const PDFHelpers = {
  centerX: (item) => Number(item.x) + (Number(item.width) || 0) / 2,

  // Découpe un fragment en mots, avec une position estimée pour chacun.
  // Utile lorsqu'un même fragment PDF.js couvre plusieurs colonnes.
  splitItem: (item) => {
    const str = String(item.str ?? "");
    const words = [...str.matchAll(/\S+/g)];

    if (words.length <= 1) {
      return [{ ...item, str: str.trim(), source: item }];
    }

    const width = Number(item.width) || 0;

    return words.map(word => ({
      x: item.x + width * word.index / str.length,
      y: item.y,
      width: width * word[0].length / str.length,
      str: word[0],
      source: item
    }));
  },

  cell: (items, bounds) => {
    if (!bounds) return "";

    return items
      .filter(item => {
        const x = PDFHelpers.centerX(item);
        return x >= bounds.left && x < bounds.right;
      })
      .sort((a, b) => a.x - b.x)
      .map(item => item.str)
      .join(" ")
      .trim();
  },

  groupItems: (items, tolerance = 2.5) => {
    const groups = [];

    for (const item of [...items].sort(
      (a, b) => b.y - a.y || a.x - b.x
    )) {
      const last = groups[groups.length - 1];

      if (!last || Math.abs(last.y - item.y) > tolerance) {
        groups.push({ y: item.y, items: [item] });
      } else {
        last.items.push(item);
      }
    }

    return groups;
  },

  findHeaderPart: (items, aliases, maxFragments = 4) => {
    const sorted = [...items].sort((a, b) => a.x - b.x);
    const wanted = new Set(aliases.map(Utils.compactKey));

    for (let start = 0; start < sorted.length; start++) {
      for (let count = 1; count <= maxFragments; count++) {
        const fragments = sorted.slice(start, start + count);
        if (fragments.length !== count) continue;

        const key = Utils.compactKey(
          fragments.map(item => item.str).join(" ")
        );

        if (!wanted.has(key)) continue;

        const left = Math.min(...fragments.map(item => item.x));
        const right = Math.max(
          ...fragments.map(item => item.x + item.width)
        );

        return { left, right, center: (left + right) / 2 };
      }
    }

    return null;
  },

  findHeader: (items, schema, order) => {
    const found = {};

    for (const field of order) {
      const part = PDFHelpers.findHeaderPart(items, schema[field]);
      if (!part) return null;
      found[field] = part;
    }

    if (order.some((field, index) =>
      index > 0 && found[field].center <= found[order[index - 1]].center
    )) {
      return null;
    }

    return found;
  },

  layoutFromCenters: (found, order) => {
    const layout = {};

    order.forEach((field, index) => {
      const previous = order[index - 1];
      const next = order[index + 1];

      layout[field] = {
        left: previous
          ? (found[previous].center + found[field].center) / 2
          : -Infinity,
        right: next
          ? (found[field].center + found[next].center) / 2
          : Infinity
      };
    });

    return layout;
  },

  readUnit: (items, bounds, allowLocalFallback = false) => {
    const original = PDFHelpers.cell(items, bounds);
    const direct = Utils.canonicalUnit(original);
    if (direct) return direct;

    // Un nombre seul ne peut jamais devenir une unité.
    const candidates = new Set();
    const width = bounds.right - bounds.left;

    const nearby = items
      .filter(item => {
        const x = PDFHelpers.centerX(item);
        const left = allowLocalFallback && Number.isFinite(width)
          ? bounds.left - width * 0.5
          : bounds.left;

        return x >= left && x < bounds.right;
      })
      .sort((a, b) => a.x - b.x);

    for (let start = 0; start < nearby.length; start++) {
      for (let count = 1; count <= 3; count++) {
        const fragments = nearby.slice(start, start + count);
        if (fragments.length !== count) continue;

        const unit = Utils.canonicalUnit(
          fragments.map(item => item.str).join(" ")
        );

        if (unit) candidates.add(unit);
      }
    }

    return candidates.size === 1 ? [...candidates][0] : "";
  }
};

/* =========================================================
   Lecture et analyse des PDF
   ========================================================= */

const PDF = {
  toLines: async (file) => {
    const library = window.pdfjsLib;

    if (!library) {
      throw new Error("La bibliothèque PDF.js n'est pas chargée.");
    }

    library.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

    const pdf = await library.getDocument({
      data: await file.arrayBuffer()
    }).promise;

    const lines = [];

    try {
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
        const page = await pdf.getPage(pageNumber);
        const content = await page.getTextContent();
        const grouped = new Map();

        for (const item of content.items) {
          const str = String(item.str || "").trim();
          if (!str || !item.transform) continue;

          const x = Number(item.transform[4]) - page.view[0];
          const y = Number(item.transform[5]);
          const width = Number(item.width) || 0;
          const yKey = Math.round(y * 2) / 2;

          if (!grouped.has(yKey)) grouped.set(yKey, []);
          grouped.get(yKey).push({ x, y, width, str });
        }

        for (const y of [...grouped.keys()].sort((a, b) => b - a)) {
          const items = grouped.get(y).sort((a, b) => a.x - b.x);
          const text = items.map(item => item.str)
            .join(" ")
            .replace(/\s+/g, " ")
            .trim();

          if (text) {
            lines.push({
              page: pageNumber,
              pageWidth: page.view[2] - page.view[0],
              y,
              items,
              text
            });
          }
        }

        page.cleanup();
      }
    } finally {
      await pdf.destroy();
    }

    if (!lines.length) {
      throw new Error(
        "Aucun texte extrait du PDF. Une reconnaissance OCR est peut-être nécessaire."
      );
    }

    return lines;
  },

  extractQuoteNumber: (lines, fileName = "") => {
    const candidates = [];

    const clean = (value) => String(value ?? "")
      .replace(/\([^)]*\)/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/[,:;]+$/g, "")
      .trim();

    const label = /^(?:devis|quote|offre)\b\s*(?:(?:n\s*[°ºo]?|no)\s*)?[:#-]?\s*(.*)$/i;

    for (let index = 0; index < lines.length; index++) {
      const text = String(lines[index].text || "").trim();
      const match = text.match(label);
      if (!match) continue;

      if (match[1]) {
        candidates.push(clean(match[1]));
      } else {
        const next = lines[index + 1];

        if (next && next.page === lines[index].page) {
          const nextMatch = String(next.text || "")
            .match(/^([A-Z0-9][A-Z0-9._/-]*)/i);

          if (nextMatch) candidates.push(clean(nextMatch[1]));
        }
      }
    }

    const unique = [...new Set(
      candidates.map(Utils.sanitizeFileName).filter(Boolean)
    )];

    if (unique.length) {
      return unique.sort((a, b) => b.length - a.length)[0];
    }

    const baseName = String(fileName).replace(/\.[^.]+$/, "");
    const match = baseName.match(
      /(?:devis|quote|offre)[_\s-]*(?:n[°ºo]?[_\s-]*)?(.+)/i
    );

    return match ? Utils.sanitizeFileName(match[1]) : "";
  },

  detectSupplier: (lines, fileName = "") => {
    const text = Utils.normalizeKey(
      `${fileName}\n${lines.map(line => line.text || "").join("\n")}`
    ).replace(/_/g, " ");

    if (/\bAXIANS\b/.test(text)) return "AXIANS";
    if (/\bNEW\s*COM\b/.test(text)) return "NEWCOM";
    if (/\bCGTI\b/.test(text)) return "CGTI";

    return "UNKNOWN";
  },

  extractRows: (lines, fileName = "", selectedSupplier = "") => {
    const detected = PDF.detectSupplier(lines, fileName);
    const selected = Utils.normalizeSupplierName(selectedSupplier);

    if (detected !== "UNKNOWN" && selected && detected !== selected) {
      throw new Error(
        `Le PDF "${fileName}" semble provenir de ${detected}, ` +
        `mais le catalogue sélectionné est ${selected}.`
      );
    }

    const supplier = detected !== "UNKNOWN" ? detected : selected;

    if (supplier === "AXIANS") return PDF.parseAxiansRows(lines);
    if (supplier === "NEWCOM") return PDF.parseNewcomRows(lines);
    if (supplier === "CGTI") return PDF.parseCgtiRows(lines);

    return PDF.parseGenericRows();
  },

  parseAxiansRows: (lines) => PDF.parseStandardRows(lines, "AXIANS"),

  parseNewcomRows: (lines) => PDF.parseStandardRows(lines, "NEWCOM"),

  parseStandardRows: (lines, supplier) => {
    const ordered = [...lines].sort(
      (a, b) => a.page - b.page || b.y - a.y
    );

    const order = ["code", "designation", "unit", "quantity", "pu", "total"];
    const schema = supplier === "AXIANS"
      ? {
          code: ["REF", "REFERENCE"],
          designation: ["DESIGNATION"],
          unit: ["UNITE", "UNIT"],
          quantity: ["QTE", "QUANTITE"],
          pu: ["PUHT", "PRIXUNITAIREHT"],
          total: ["PTHT", "PRIXTOTALHT"]
        }
      : {
          code: ["CODE", "REF", "REFERENCE"],
          designation: ["DESIGNATION", "LIBELLE"],
          unit: ["UNIT", "UNITE"],
          quantity: ["QTE", "QUANTITE"],
          pu: ["PRIXUNITAIRE", "PU", "PUHT"],
          total: ["TOTAL", "PRIXHT", "PTHT", "PRIXTOTAL"]
        };

    const layouts = new Map();

    for (const line of ordered) {
      const items = ordered
        .filter(candidate =>
          candidate.page === line.page &&
          Math.abs(candidate.y - line.y) <= 3
        )
        .flatMap(candidate => candidate.items || []);

      const found = PDFHelpers.findHeader(items, schema, order);

      if (found) {
        layouts.set(
          line.page,
          PDFHelpers.layoutFromCenters(found, order)
        );
      }
    }

    if (!layouts.size) {
      console.warn(`${supplier} — en-tête non reconnu`, ordered);
      throw new Error(
        `En-tête ${supplier} non détecté. Consultez la console du navigateur.`
      );
    }

    const headerPages = [...layouts.keys()].sort((a, b) => a - b);
    const rows = [];

    for (const line of ordered) {
      const previousPage = headerPages.filter(page => page <= line.page).pop();
      const layout = layouts.get(line.page) ||
        layouts.get(previousPage ?? headerPages[0]);

      const text = String(line.text || "").trim();
      const items = line.items || [];

      const pattern = supplier === "AXIANS"
        ? /^\s*(HB|MOE01|MOE02|[A-Z]{2,6}[\s_-]*\d{1,4}(?:BIS|TER|[A-Z])?)\b/i
        : /^\s*(MOE01|MOE02|HB|[A-Z0-9]{2,12})\b/i;

      const codeMatch = text.match(pattern);
      if (!codeMatch || !items.length) continue;

      const code = Utils.normalizeCode(codeMatch[1]);
      const quantityText = PDFHelpers.cell(items, layout.quantity);
      const puText = PDFHelpers.cell(items, layout.pu);
      const totalText = PDFHelpers.cell(items, layout.total);
      const unit = PDFHelpers.readUnit(items, layout.unit, true);

      const designation = items
        .filter(item => PDFHelpers.centerX(item) < layout.unit.left)
        .sort((a, b) => a.x - b.x)
        .map(item => item.str)
        .join(" ")
        .replace(codeMatch[0], "")
        .trim();

      const base = {
        code,
        unit,
        designation,
        page: line.page,
        y: line.y,
        rowKey: `${line.page}_${Math.round(line.y * 10)}`,
        raw: text
      };

      if (code === "MOE01" || code === "MOE02") {
        const pctMatch = text.match(/(\d+(?:[.,]\d+)?)\s*%/);

        if (supplier === "AXIANS") {
          const percentage = pctMatch
            ? Utils.parsePrice(pctMatch[1]) / 100
            : NaN;

          if (!Number.isFinite(percentage)) continue;

          rows.push({
            ...base,
            pu: percentage,
            quantity: null,
            total: null,
            priceKind: "percentage",
            puText: pctMatch[0]
          });
        } else {
          const rate = unit === "%"
            ? Utils.parsePdfNumber(quantityText)
            : pctMatch
              ? Utils.parsePrice(pctMatch[1])
              : NaN;

          const pu = Utils.parsePdfNumber(puText);
          if (!Number.isFinite(rate) || !Number.isFinite(pu)) continue;

          rows.push({
            ...base,
            unit: unit || "%",
            pu,
            quantity: rate,
            total: Utils.parsePdfNumber(totalText),
            priceKind: "percentage",
            puText
          });
        }

        continue;
      }

      const pu = Utils.parsePdfNumber(puText);
      if (!Number.isFinite(pu)) continue;

      rows.push({
        ...base,
        pu,
        quantity: Utils.parsePdfNumber(quantityText),
        total: Utils.parsePdfNumber(totalText),
        priceKind: "amount",
        puText
      });
    }

    const result = Utils.uniqueRows(rows);

    if (!result.length) {
      throw new Error(`${supplier} : aucune ligne exploitable dans le PDF.`);
    }

    return result;
  },

    parseCgtiRows: (lines) => {
    const DEBUG = CONFIG.DEBUG_CGTI === true;
    const rows = [];
    const diagnostics = [];
    const pages = new Map();
    const { centerX, cell, groupItems } = PDFHelpers;

    const textByX = (items) => [...items]
      .sort((a, b) => a.x - b.x)
      .map(item => item.str)
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();

    const inColumn = (item, column) => {
      const x = centerX(item);
      return x >= column.left && x < column.right;
    };

    /*
     * Chaque valeur est lue dans SA colonne (unité, PU, quantité,
     * total), délimitée à partir de l'en-tête du tableau.
     * Le texte complet de la ligne n'est jamais analysé : le "3"
     * d'un "M3" (souvent en exposant) ne peut plus être pris pour
     * le début du prix unitaire.
     */
    const headerOrder = ["unit", "pu", "quantity", "total"];
    const headerSchema = {
      unit: ["UNITE", "UNIT", "U"],
      pu: [
        "PRIX UNITAIRE", "PRIX UNITAIRE HT", "PRIX UNIT", "PRIX U",
        "PU", "PU HT"
      ],
      quantity: ["QUANTITE", "QTE", "QTES", "QUANT"],
      total: [
        "TOTAL", "TOTAL HT", "MONTANT", "MONTANT HT",
        "PRIX TOTAL", "PRIX TOTAL HT", "PT HT"
      ]
    };

    const findLayout = (items) => {
      for (const tolerance of [2.5, 8]) {
        for (const group of groupItems(items, tolerance)) {
          const found = PDFHelpers.findHeader(
            group.items,
            headerSchema,
            headerOrder
          );
          if (!found) continue;

          const layout = PDFHelpers.layoutFromCenters(found, headerOrder);
          layout.unit.left = found.unit.center -
            (found.pu.center - found.unit.center) / 2;

          return {
            ...layout,
            headerY: Math.min(...group.items.map(item => item.y))
          };
        }
      }

      return null;
    };

    // Colonnes par défaut si aucun en-tête n'est reconnu.
    const defaultLayout = (width) => ({
      unit: { left: width * 0.685, right: width * 0.765 },
      pu: { left: width * 0.765, right: width * 0.845 },
      quantity: { left: width * 0.845, right: width * 0.905 },
      total: { left: width * 0.905, right: Infinity },
      headerY: Infinity
    });

    const codePattern =
      /^((?:2\s*ND|[A-Z]{2,6})[\s_-]*\d{1,4}(?:BIS|TER)?|H\.?\s*B\.?)$/i;
    const offeredPattern = /\b(?:OFFERTE?S?|GRATUITE?S?)\b/;
    const outsidePattern = /\bHORS[\s-]*BORDEREAU\b|^H\.?\s?B\b/;
    const subtotalPattern =
      /\b(?:SOUS[\s-]*TOTAL|TOTAL|TVA|TTC|NET A PAYER|ACOMPTE)\b/;

    // Texte original de la cellule "OFFERT" / "GRATUIT", sinon "".
    const offeredLabel = (...texts) => texts.find(text =>
      offeredPattern.test(Utils.normalizeKey(text))
    ) || "";

    // Préparer les pages à partir des fragments PDF.js.
    for (const line of lines) {
      if (!pages.has(line.page)) {
        pages.set(line.page, {
          width: null,
          items: [],
          seen: new Set()
        });
      }

      const page = pages.get(line.page);

      if (Number.isFinite(line.pageWidth) && line.pageWidth > 0) {
        page.width = line.pageWidth;
      }

      for (const source of line.items || []) {
        const item = {
          x: Number(source.x),
          y: Number(source.y ?? line.y),
          width: Number(source.width) || 0,
          str: String(source.str ?? "").trim()
        };

        if (
          !item.str ||
          !Number.isFinite(item.x) ||
          !Number.isFinite(item.y)
        ) {
          continue;
        }

        const key = [
          item.x, item.y, item.width, item.str
        ].join("|");

        if (page.seen.has(key)) continue;
        page.seen.add(key);
        page.items.push(item);
      }
    }

    let lastLayout = null;

    for (const [pageNumber, page] of [...pages].sort(
      (a, b) => a[0] - b[0]
    )) {
      if (!page.width) {
        throw new Error(
          `CGTI : largeur PDF absente, page ${pageNumber}.`
        );
      }

      const detected = findLayout(page.items);

      if (!detected && !lastLayout) {
        diagnostics.push({
          page: pageNumber,
          code: "",
          erreur: "En-tête du tableau non reconnu : colonnes par défaut",
          texte: ""
        });
      }

      const layout = detected ||
        (lastLayout ? { ...lastLayout, headerY: Infinity } : null) ||
        defaultLayout(page.width);

      if (detected) lastLayout = detected;

      // Ignorer tout ce qui précède l'en-tête (n° de devis, adresses...).
      const body = page.items.filter(item => item.y < layout.headerY - 0.5);
      const tokens = body.flatMap(PDFHelpers.splitItem);

      /*
       * 1. Codes : fragment autonome dans la colonne de gauche.
       * Accepte notamment RN25, ELE151, IC18, FO12, PYL100, 2ND16 et HB.
       */
      const codeAnchors = [];

      for (const group of groupItems(
        body.filter(item => item.x < page.width * 0.22),
        0.8
      )) {
        for (const item of group.items) {
          const match = item.str.match(codePattern);
          if (!match) continue;

          let code = Utils.normalizeCode(match[1]).replace(/\./g, "");
          if (/^ENTRE\d/.test(code)) continue;
          if (/^HB\d*$/.test(code)) code = "HB";

          codeAnchors.push({
            code,
            codeLabel: match[1],
            y: item.y,
            item,
            line: null,
            extraLines: []
          });

          break;
        }
      }

      codeAnchors.sort((a, b) => b.y - a.y);

      /*
       * 2. Lignes chiffrées : lues uniquement dans les colonnes
       * PU / quantité / total.
       */
      const numericColumns = [layout.pu, layout.quantity, layout.total];

      const amountLines = groupItems(
        tokens.filter(token =>
          numericColumns.some(column => inColumn(token, column))
        ),
        1.5
      )
        .map(group => {
          const puText = cell(group.items, layout.pu);
          const quantityText = cell(group.items, layout.quantity);
          const totalText = cell(group.items, layout.total);
          const leftText = textByX(tokens.filter(token =>
            Math.abs(token.y - group.y) <= 1.5 &&
            centerX(token) < layout.unit.left
          ));

          return {
            y: group.y,
            puText,
            quantityText,
            totalText,
            pu: Utils.parsePdfNumber(puText),
            quantity: Utils.parsePdfNumber(quantityText),
            total: Utils.parsePdfNumber(totalText),
            offered: offeredLabel(puText, totalText),
            leftText,
            outside: outsidePattern.test(Utils.normalizeKey(leftText))
          };
        })
        .filter(line =>
          Number.isFinite(line.pu) ||
          line.offered ||
          (Number.isFinite(line.quantity) && Number.isFinite(line.total))
        );

      // 3. Rattacher chaque ligne chiffrée au code le plus proche.
      const pairs = [];

      codeAnchors.forEach((anchor, anchorIndex) => {
        amountLines.forEach((line, lineIndex) => {
          const distance = Math.abs(anchor.y - line.y);
          if (distance > 12) return;

          // Une ligne "hors bordereau" ne peut pas être volée par un voisin.
          if (line.outside && anchor.code !== "HB" && distance > 1.5) return;

          pairs.push({ anchorIndex, lineIndex, distance });
        });
      });

      pairs.sort((a, b) => a.distance - b.distance);

      const usedLines = new Set();

      for (const pair of pairs) {
        const anchor = codeAnchors[pair.anchorIndex];
        if (anchor.line || usedLines.has(pair.lineIndex)) continue;

        anchor.line = amountLines[pair.lineIndex];
        usedLines.add(pair.lineIndex);
      }

      /*
       * 4. Lignes chiffrées sans code : tranches MOE (rattachées
       * à MOE01/MOE02) ou lignes hors bordereau (ancres HB).
       * Elles ne perturbent ainsi plus la lecture des voisines.
       */
      const outsideAnchors = [];

      amountLines.forEach((line, lineIndex) => {
        if (usedLines.has(lineIndex)) return;

        const owner = [...codeAnchors, ...outsideAnchors]
          .filter(anchor => anchor.y > line.y)
          .sort((a, b) => a.y - b.y)[0];

        if (!line.outside && owner && /^MOE0[12]$/.test(owner.code)) {
          owner.extraLines.push(line);
          return;
        }

        if (!line.outside) {
          const lineText = textByX(tokens.filter(token =>
            Math.abs(token.y - line.y) <= 1.5
          ));

          if (subtotalPattern.test(Utils.normalizeKey(lineText))) return;
          if (!Number.isFinite(line.pu) && !line.offered) return;
        }

        outsideAnchors.push({
          code: "HB",
          codeLabel: "",
          y: line.y,
          item: null,
          line,
          extraLines: []
        });
      });

      const anchors = [...codeAnchors, ...outsideAnchors]
        .map(anchor => {
          const ys = [
            anchor.y,
            anchor.line?.y ?? anchor.y,
            ...anchor.extraLines.map(line => line.y)
          ];

          return {
            ...anchor,
            refY: anchor.line?.y ?? anchor.y,
            top: Math.max(...ys),
            bottom: Math.min(...ys)
          };
        })
        .sort((a, b) => b.top - a.top);

      for (let index = 0; index < anchors.length; index++) {
        const anchor = anchors[index];
        const previous = anchors[index - 1];
        const next = anchors[index + 1];
        const line = anchor.line;

        // Limites verticales : aucune lecture sur la prestation voisine.
        const upper = previous
          ? Math.min(anchor.top + 12, (previous.bottom + anchor.top) / 2)
          : anchor.top + 12;

        const lower = next
          ? Math.max(anchor.bottom - 12, (anchor.bottom + next.top) / 2)
          : anchor.bottom - 12;

        const inBand = (item) => item.y > lower && item.y <= upper;
        const bandTokens = tokens.filter(inBand);

        // Unité : fragments de la colonne, exposants compris ("M" + "3").
        const unitTokens = bandTokens.filter(token =>
          inColumn(token, layout.unit) &&
          (
            Math.abs(token.y - anchor.refY) <= 4 ||
            Math.abs(token.y - anchor.y) <= 4
          )
        );

        const unitText = textByX(unitTokens);
        const unit = Utils.canonicalUnit(unitText);

        const designation = groupItems(
          bandTokens.filter(token =>
            centerX(token) < layout.unit.left &&
            (!anchor.item || token.source !== anchor.item)
          ),
          0.8
        )
          .map(group => textByX(group.items))
          .join(" ")
          .trim();

        const raw = groupItems(body.filter(inBand), 0.8)
          .map(group => textByX(group.items))
          .join(" ");

        const pu = line ? line.pu : NaN;
        const quantity = line ? line.quantity : NaN;
        const total = line ? line.total : NaN;

        const issues = [];

        if (!line) {
          issues.push("montants PU / quantité / total introuvables");
        } else {
          if (!Number.isFinite(pu)) {
            issues.push(`PU illisible («${line.puText || "vide"}»)`);
          }
          if (!Number.isFinite(quantity)) {
            issues.push(
              `quantité illisible («${line.quantityText || "vide"}»)`
            );
          }
          if (!Number.isFinite(total)) {
            issues.push(`total illisible («${line.totalText || "vide"}»)`);
          }

          if ([pu, quantity, total].every(Number.isFinite)) {
            const expected = pu * quantity;
            const limit = Math.max(0.05, Math.abs(total) * 0.005);

            if (Math.abs(expected - total) > limit) {
              issues.push(
                `PU × quantité = ${expected.toFixed(2)} ` +
                `≠ total ${total.toFixed(2)}`
              );
            }
          }
        }

        if (!unit) {
          issues.push(unitText
            ? `unité illisible ou ambiguë («${unitText}»)`
            : "unité absente");
        }

        const extractionIssue = issues.join(" ; ");
        const isTiered = /^MOE0[12]$/.test(anchor.code);

        if (extractionIssue && !isTiered) {
          diagnostics.push({
            page: pageNumber,
            code: anchor.code,
            erreur: extractionIssue,
            texte: raw
          });
        }

        /*
         * Conserver le code même si une cellule est illisible.
         * La comparaison signalera A_VERIFIER plutôt que de faire
         * disparaître la prestation du CSV.
         */
        rows.push({
          code: anchor.code,
          codeLabel: anchor.codeLabel,
          unit,
          unitText,
          pu,
          quantity,
          total,
          designation,
          priceKind: isTiered ? "tiered_amount" : "amount",
          puText: line ? line.puText : "",
          quantityText: line ? line.quantityText : "",
          totalText: line ? line.totalText : "",
          pdfValue: line ? line.offered : "",
          offered: Boolean(line?.offered),
          outsideCatalog: anchor.code === "HB",
          extractionIssue,
          page: pageNumber,
          y: anchor.y,
          rowKey: `${pageNumber}_${Math.round(anchor.y * 10)}`,
          raw
        });
      }

      if (DEBUG) {
        console.info(`CGTI — page ${pageNumber}`, {
          codesDetectes: anchors.map(anchor => anchor.code)
        });
      }
    }

    if (diagnostics.length) {
      console.warn("CGTI — extractions à vérifier :");
      console.table(diagnostics);
    }

    const result = Utils.uniqueRows(rows);

    if (!result.length) {
      throw new Error(
        "CGTI : aucun code de prestation reconnu. " +
        "Consultez la console du navigateur."
      );
    }

    if (DEBUG) {
      console.table(result.map(row => ({
        code: row.code,
        unite: row.unit,
        pu: row.pu,
        quantite: row.quantity,
        total: row.total,
        aVerifier: row.extractionIssue
      })));
    }

    return result;
  },

  parseGenericRows: () => {
    throw new Error(
      "Ce fournisseur ne dispose pas encore d'un analyseur PDF."
    );
  }
};

/* =========================================================
   Comparaison et export
   ========================================================= */

const Compare = {
  supplierToMap: (supplier) => {
    const source = STATE.jsonDB?.[supplier] || {};
    const map = new Map();

    for (const [code, candidates] of Object.entries(source)) {
      if (!Array.isArray(candidates)) continue;

      const key = Utils.normalizeCode(code);
      map.set(key, [...(map.get(key) || []), ...candidates]);
    }

    return map;
  },

    execute: (pdfRows, excelByCode, tolerance) => {
    if (!Number.isFinite(tolerance) || tolerance < 0) {
      throw new Error("La tolérance doit être un nombre positif ou nul.");
    }

    const output = [];
    const stats = {
      ok: 0,
      koAbsent: 0,
      koUnit: 0,
      koPrice: 0,
      aVerifier: 0,
      hb: 0,
      total: pdfRows.length
    };

    for (const row of pdfRows) {
      if (row.priceKind === "tiered_amount") {
        output.push({
          ...row,
          status: "A_VERIFIER",
          detail:
            "Montant de maîtrise d'œuvre calculé par tranches : " +
            "ne pas comparer directement ce montant au tarif catalogue. " +
            "Assiette, tranches et taux à contrôler." +
            (row.pdfValue ? ` ; Valeur PDF : ${row.pdfValue}` : "")
        });
        stats.aVerifier++;
        continue;
      }

      if (row.code === "HB") {
        const hasNumbers = [row.quantity, row.pu, row.total]
          .every(Number.isFinite);

        let check = "Contrôle QTE×PUHT impossible (valeurs manquantes)";

        if (hasNumbers) {
          const expected = row.quantity * row.pu;
          const delta = Math.abs(expected - row.total);
          const limit = Math.max(0.05, Math.abs(row.total) * 0.005);

          check = delta <= limit
            ? "QTE×PUHT cohérent avec PTHT"
            : `INCOHÉRENT : ${row.quantity} × ${row.pu.toFixed(2)} ` +
              `= ${expected.toFixed(2)} ≠ PTHT ${row.total.toFixed(2)}`;
        }

        output.push({
          ...row,
          status: "HB_A_VERIFIER",
          detail: "Hors bordereau à valider" +
            (row.designation ? ` ; ${row.designation}` : "") +
            ` ; ${check}`
        });

        stats.hb++;
        continue;
      }

      const candidates = excelByCode.get(Utils.normalizeCode(row.code));

      if (!candidates?.length) {
        output.push({
          ...row,
          status: "KO_ABSENT",
          detail: "Code absent dans Excel"
        });
        stats.koAbsent++;
        continue;
      }

      // Lecture PDF incertaine : ne pas conclure à un écart de catalogue.
      if (row.extractionIssue) {
        output.push({
          ...row,
          status: "A_VERIFIER",
          detail: `Extraction PDF incertaine : ${row.extractionIssue} ; ` +
            `Excel=${candidates.map(candidate =>
              `${candidate?.unit || ""} ${Number.isFinite(candidate?.pu)
                ? candidate.pu.toFixed(2)
                : "prix absent"}`.trim()
            ).join(", ")}`
        });
        stats.aVerifier++;
        continue;
      }

      const rowUnit = Utils.normalizeUnit(row.unit);

      const sameUnit = rowUnit
        ? candidates.filter(candidate =>
            Utils.normalizeUnit(candidate?.unit) === rowUnit
          )
        : [];

      if (!sameUnit.length) {
        output.push({
          ...row,
          status: "KO_UNITE",
          detail: `Unité PDF=${row.unit || "(non reconnue)"}, ` +
            `Excel=${candidates.map(candidate => candidate?.unit || "").join(", ")}`
        });
        stats.koUnit++;
        continue;
      }

      const match = sameUnit.find(candidate =>
        Number.isFinite(candidate.pu) &&
        Number.isFinite(row.pu) &&
        Math.abs(candidate.pu - row.pu) <= tolerance + 1e-9
      );

      if (!match) {
        output.push({
          ...row,
          status: "KO_PRIX",
          detail: `PU PDF=${Number.isFinite(row.pu) ? row.pu.toFixed(2) : "illisible"} ; ` +
            `Excel=${sameUnit.map(candidate =>
              Number.isFinite(candidate.pu)
                ? candidate.pu.toFixed(2)
                : "prix absent"
            ).join(", ")}`
        });
        stats.koPrice++;
        continue;
      }

      output.push({ ...row, status: "OK", detail: "Conforme" });
      stats.ok++;
    }

    return { output, stats };
  },

  renderResults: (rows) => {
    if (!rows.length) {
      UI.text("results", "Aucun résultat.");
      return;
    }

    const body = rows.map(row => {
      const cssClass = row.status === "OK"
        ? "ok"
        : ["KO_PRIX", "HB_A_VERIFIER", "A_VERIFIER"].includes(row.status)
          ? "warn"
          : "ko";

      return `
        <tr>
          <td>${Utils.escapeHtml(row.pdf)}</td>
          <td>${Utils.escapeHtml(row.quoteNumber)}</td>
          <td>${Utils.escapeHtml(row.code)}</td>
          <td>${Utils.escapeHtml(row.unit)}</td>
          <td>${Number.isFinite(row.pu) ? row.pu.toFixed(2) : ""}</td>
          <td class="${cssClass}">${Utils.escapeHtml(row.status)}</td>
          <td>${Utils.escapeHtml(row.detail)}</td>
          <td>${Utils.escapeHtml(row.page ?? "")}</td>
          <td>${Utils.escapeHtml(row.rowKey)}</td>
        </tr>
      `;
    }).join("");

    UI.html("results", `
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>PDF</th><th>Numéro devis</th><th>Code</th>
              <th>Unité</th><th>PU</th><th>Statut</th>
              <th>Détail</th><th>Page</th><th>RowKey</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    `);
  },

  exportCSV: (rows, fileName = "resultats-analyse") => {
    if (!rows?.length) {
      Toast.error("Aucun résultat à exporter.");
      return;
    }

    const headers = [
      "PDF", "Numéro devis", "Code", "Unité", "PU",
      "Statut", "Détail", "Page", "RowKey", "Texte original"
    ];

    const data = [
      headers,
      ...rows.map(row => [
        row.pdf || "",
        row.quoteNumber || "",
        row.code || "",
        row.unit || "",
        Utils.formatPrice(row.pu),
        row.status || "",
        row.detail || "",
        row.page ?? "",
        row.rowKey || "",
        row.raw || ""
      ])
    ];

    const csv = data
      .map(row => row.map(Utils.csvEscape).join(";"))
      .join("\r\n");

    const blob = new Blob(
      ["\uFEFF", csv],
      { type: "text/csv;charset=utf-8;" }
    );

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = Utils.sanitizeFileName(fileName) || "resultats-analyse";

    link.href = url;
    link.download = `${safeName}_${Utils.createTimestamp()}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => URL.revokeObjectURL(url), 1000);
    Toast.success("Fichier exporté.");
  }
};

/* =========================================================
   Glisser-déposer
   ========================================================= */

const DragDrop = {
  droppedFiles: new Map(),

  getFiles: (inputId) => {
    if (DragDrop.droppedFiles.has(inputId)) {
      return [...DragDrop.droppedFiles.get(inputId)];
    }

    return [...(document.getElementById(inputId)?.files || [])];
  },

  setup: (dropZoneId, inputId) => {
    const zone = document.getElementById(dropZoneId);
    const input = document.getElementById(inputId);

    if (!zone || !input || zone.dataset.dropReady === "true") return;
    zone.dataset.dropReady = "true";

    const accepted = (file) => inputId === "pdfFiles"
      ? /\.pdf$/i.test(file.name) || file.type === "application/pdf"
      : /\.(xlsx|xls|csv)$/i.test(file.name);

    const stop = (event) => {
      event.preventDefault();
      event.stopPropagation();
    };

    input.addEventListener("click", event => event.stopPropagation());

    input.addEventListener("change", event => {
      zone.classList.remove("dragover");

      if (!event.detail?.files) {
        DragDrop.droppedFiles.delete(inputId);
      }
    });

    zone.addEventListener("click", event => {
      if (input.disabled || event.target === input) return;

      if (
        event.target instanceof Element &&
        event.target.closest("button, a, select, textarea, label")
      ) {
        return;
      }

      input.click();
    });

    for (const eventName of ["dragenter", "dragover"]) {
      zone.addEventListener(eventName, event => {
        stop(event);
        if (input.disabled) return;

        zone.classList.add("dragover");
        if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
      });
    }

    zone.addEventListener("dragleave", event => {
      stop(event);

      if (
        !(event.relatedTarget instanceof Node) ||
        !zone.contains(event.relatedTarget)
      ) {
        zone.classList.remove("dragover");
      }
    });

    zone.addEventListener("drop", event => {
      stop(event);
      zone.classList.remove("dragover");
      if (input.disabled) return;

      const received = [...(event.dataTransfer?.files || [])];
      let files = received.filter(accepted);

      if (!input.multiple) files = files.slice(0, 1);

      if (!files.length) {
        Toast.error(
          inputId === "pdfFiles"
            ? "Déposez un fichier PDF valide."
            : "Déposez un fichier Excel ou CSV valide."
        );
        return;
      }

      input.value = "";
      DragDrop.droppedFiles.set(inputId, files);

      input.dispatchEvent(new CustomEvent("change", {
        bubbles: true,
        detail: { files }
      }));

      if (files.length < received.length) {
        Toast.info("Certains fichiers ont été ignorés.");
      }
    });
  }
};

/* =========================================================
   Actions de l'application
   ========================================================= */

const Actions = {
  excelChanged: async (event) => {
    const requestId = ++STATE.excelRequestId;
    const files = [...(event.detail?.files || event.target.files || [])];

    STATE.excelWorkbooks = [];
    Excel.renderSheetPicker([]);

    if (!files.length) {
      UI.renderDbInfo();
      return;
    }

    UI.setLoading("buildDbBtn", true);
    UI.text("dbInfo", "Lecture des classeurs en cours...");

    try {
      const workbooks = await Excel.prepare(files);
      if (requestId !== STATE.excelRequestId) return;

      STATE.excelWorkbooks = workbooks;
      Excel.renderSheetPicker(workbooks);

      UI.text(
        "dbInfo",
        "Fichiers Excel chargés. Sélectionnez les feuilles puis construisez la base JSON."
      );
    } catch (error) {
      if (requestId !== STATE.excelRequestId) return;

      UI.text("dbInfo", `Erreur lecture Excel : ${Utils.errorMessage(error)}`);
      Toast.error(Utils.errorMessage(error));
    } finally {
      if (requestId === STATE.excelRequestId) {
        UI.setLoading("buildDbBtn", false);
      }
    }
  },

  pdfChanged: (event) => {
    const files = [...(event.detail?.files || event.target.files || [])];

    STATE.lastAnalysisRows = [];
    STATE.lastAnalysisFileName = "";
    UI.setLoading("exportCsvBtn", true);

    if (!files.length) {
      UI.text("pdfFileList", "Aucun PDF sélectionné.");
      return;
    }

    UI.html("pdfFileList", `
      <b>${files.length} PDF sélectionné(s) :</b><br>
      ${files.map(file => Utils.escapeHtml(file.name)).join("<br>")}
    `);

    Toast.info(`${files.length} PDF ajouté(s)`);
  },

  buildDB: async () => {
    if (!STATE.excelWorkbooks.length) {
      Toast.error("Merci de sélectionner les fichiers Excel.");
      return;
    }

    UI.setLoading("buildDbBtn", true);
    UI.showProgress("Construction de la base JSON...");

    try {
      const imported = Excel.toJsonDB(STATE.excelWorkbooks);
      const merged = { ...(STATE.jsonDB || {}), ...imported };

      UI.showProgress("Sauvegarde de la base JSON globale...");
      const response = await API.saveDB(merged);

      STATE.jsonDB = merged;
      UI.refreshAll();

      const suppliers = Object.keys(imported);
      const timestamp = Utils.createTimestamp();

      for (let index = 0; index < suppliers.length; index++) {
        const supplier = suppliers[index];

        UI.showProgress(
          `Sauvegarde ${supplier} (${index + 1}/${suppliers.length})...`
        );

        await API.saveSupplierFile(
          supplier,
          `${Utils.sanitizeFileName(supplier)}_${timestamp}.json`,
          JSON.stringify({ [supplier]: imported[supplier] }, null, 2)
        );
      }

      UI.html("summary", `
        Base JSON sauvegardée :
        <b>${Utils.escapeHtml(response.file || "")}</b><br>
        Fournisseurs traités :
        <b>${Utils.escapeHtml(suppliers.join(", "))}</b>
      `);

      await UI.renderCatalogFiles();
      Toast.success("Base JSON construite et sauvegardée.");
    } catch (error) {
      UI.text("dbInfo", `Erreur : ${Utils.errorMessage(error)}`);

      Toast.error(
        `${Utils.errorMessage(error)} ` +
        "Si une sauvegarde a déjà commencé, actualisez les catalogues avant de réessayer."
      );
    } finally {
      UI.setLoading("buildDbBtn", false);
      UI.hideProgress();
    }
  },

  listSuppliers: async () => {
    UI.setLoading("listSuppliersBtn", true);

    try {
      const data = await API.fetch("/api/list-suppliers");

      const found = Array.isArray(data.suppliers)
        ? data.suppliers
            .filter(value => typeof value === "string")
            .map(Utils.normalizeSupplierName)
            .filter(Boolean)
        : [];

      STATE.suppliersList = [...new Set([
        ...STATE.suppliersList,
        ...found
      ])].sort();

      Storage.saveSuppliers();
      UI.refreshAll();

      Toast.info(
        found.length
          ? `Fournisseurs sur disque : ${found.join(", ")}`
          : "Aucun dossier fournisseur dans CATALOG."
      );
    } catch (error) {
      Toast.error(Utils.errorMessage(error));
    } finally {
      UI.setLoading("listSuppliersBtn", false);
    }
  },

  loadDB: async () => {
    UI.setLoading("loadDbBtn", true);
    UI.text("dbInfo", "Chargement du catalogue...");

    try {
      const selected = UI.value("catalogSupplierSelect");
      const suppliers = selected ? [selected] : [...STATE.suppliersList];

      if (!suppliers.length) {
        throw new Error("Aucun fournisseur à charger.");
      }

      const nextDB = { ...(STATE.jsonDB || {}) };
      const summary = [];

      for (const supplier of suppliers) {
        const key = Utils.normalizeSupplierName(supplier);
        const db = await API.loadSupplierCatalog(key);
        const catalog = Utils.isRecord(db[key]) ? db[key] : {};

        nextDB[key] = catalog;
        summary.push(`${key} : ${Object.keys(catalog).length} codes`);
      }

      STATE.jsonDB = nextDB;
      UI.refreshAll();
      await UI.renderCatalogFiles();

      const empty = suppliers.every(supplier =>
        !Object.keys(nextDB[Utils.normalizeSupplierName(supplier)] || {}).length
      );

      if (empty) {
        Toast.error("Catalogue vide : aucun code trouvé.");
      } else {
        Toast.success(`Catalogue chargé (${summary.join(" | ")})`);
      }
    } catch (error) {
      UI.text("dbInfo", `Erreur chargement : ${Utils.errorMessage(error)}`);
      Toast.error(Utils.errorMessage(error));
    } finally {
      UI.setLoading("loadDbBtn", false);
    }
  },

  run: async () => {
    STATE.lastAnalysisRows = [];
    STATE.lastAnalysisFileName = "";
    UI.setLoading("exportCsvBtn", true);

    const files = DragDrop.getFiles("pdfFiles");
    const supplier = UI.value("supplierSelect");
    const toleranceText = UI.value("tol").trim() || "0.01";
    const tolerance = Utils.parsePrice(toleranceText);

    if (!files.length) {
      Toast.error("Merci de sélectionner au moins un PDF.");
      return;
    }

    if (!supplier) {
      Toast.error("Merci de sélectionner un fournisseur.");
      return;
    }

    if (!Object.keys(STATE.jsonDB?.[supplier] || {}).length) {
      Toast.error("Merci de charger un catalogue non vide pour ce fournisseur.");
      return;
    }

    if (!Number.isFinite(tolerance) || tolerance < 0) {
      Toast.error("La tolérance doit être un nombre positif ou nul.");
      return;
    }

    UI.setLoading("runBtn", true);
    UI.text("summary", "Analyse en cours...");
    UI.text("results", "Traitement en cours...");

    try {
      const excelByCode = Compare.supplierToMap(supplier);
      const allRows = [];
      const quoteNumbers = [];
      const aggregate = {
        ok: 0,
        koAbsent: 0,
        koUnit: 0,
        koPrice: 0,
        aVerifier: 0,
        hb: 0,
        total: 0
       };

      for (const file of files) {
        UI.text("summary", `Analyse de ${file.name}...`);

        const lines = await PDF.toLines(file);
        const quoteNumber = PDF.extractQuoteNumber(lines, file.name);
        quoteNumbers.push(quoteNumber);

        const extracted = PDF.extractRows(lines, file.name, supplier)
          .map(row => ({
            ...row,
            pdf: file.name,
            quoteNumber
          }));

        const result = Compare.execute(extracted, excelByCode, tolerance);
        allRows.push(...result.output);

        for (const key of Object.keys(aggregate)) {
          aggregate[key] += result.stats[key];
        }
      }

      STATE.lastAnalysisRows = allRows;

      const unique = [...new Set(quoteNumbers.filter(Boolean))];
      STATE.lastAnalysisFileName = unique.length
        ? `devis_${unique.join("_")}`
        : "resultats-analyse";

      Compare.renderResults(allRows);
      UI.setLoading("exportCsvBtn", allRows.length === 0);

      UI.html("summary", `
        Fournisseur : <b>${Utils.escapeHtml(supplier)}</b><br>
        Lignes analysées : <b>${aggregate.total}</b><br>
        <span class="ok">OK : ${aggregate.ok}</span> |
        <span class="ko">KO code absent : ${aggregate.koAbsent}</span> |
        <span class="ko">KO unité : ${aggregate.koUnit}</span> |
        <span class="warn">KO prix : ${aggregate.koPrice}</span> |
        <span class="warn">À vérifier : ${aggregate.aVerifier}</span> |
        <span class="warn">HB à vérifier : ${aggregate.hb}</span>
      `);

      Toast.success("Analyse terminée.");
    } catch (error) {
      UI.text("summary", `Erreur : ${Utils.errorMessage(error)}`);
      UI.text("results", "Aucun résultat.");

      STATE.lastAnalysisRows = [];
      STATE.lastAnalysisFileName = "";

      UI.setLoading("exportCsvBtn", true);
      Toast.error(Utils.errorMessage(error));
    } finally {
      UI.setLoading("runBtn", false);
    }
  },

  addSupplier: () => {
    const input = document.getElementById("newSupplierInput");
    if (!input) return;

    const name = Utils.normalizeSupplierName(input.value);

    if (!name) {
      Toast.error("Merci de saisir un nom de fournisseur.");
      return;
    }

    if (STATE.suppliersList.includes(name)) {
      Toast.error("Ce fournisseur existe déjà.");
      return;
    }

    if (/[<>:"/\\|?*\x00-\x1F]/.test(name)) {
      Toast.error("Le nom contient des caractères non autorisés.");
      return;
    }

    STATE.suppliersList.push(name);
    STATE.suppliersList.sort();
    Storage.saveSuppliers();

    input.value = "";
    UI.refreshAll();
    Toast.success("Fournisseur ajouté.");
  }
};

/* =========================================================
   Initialisation unique, après chargement du HTML
   ========================================================= */

async function initializeApp() {
  if (STATE.initialized) return;
  STATE.initialized = true;

  const on = (id, event, handler) => {
    const element = document.getElementById(id);
    if (element) element.addEventListener(event, handler);
  };

  document.querySelectorAll(".nav-btn").forEach(button => {
    button.addEventListener("click", () => {
      const panel = document.getElementById(button.dataset.tab);
      if (!panel) return;

      document.querySelectorAll(".nav-btn").forEach(item => {
        item.classList.remove("active");
      });

      document.querySelectorAll(".tab-panel").forEach(item => {
        item.classList.remove("active");
      });

      button.classList.add("active");
      panel.classList.add("active");
    });
  });

  // Installer le nettoyage du glisser-déposer avant les événements métier.
  DragDrop.setup("excelDropZone", "excelFiles");
  DragDrop.setup("pdfDropZone", "pdfFiles");

  on("excelFiles", "change", Actions.excelChanged);
  on("pdfFiles", "change", Actions.pdfChanged);
  on("buildDbBtn", "click", Actions.buildDB);
  on("listSuppliersBtn", "click", Actions.listSuppliers);
  on("loadDbBtn", "click", Actions.loadDB);
  on("runBtn", "click", Actions.run);
  on("addSupplierBtn", "click", Actions.addSupplier);

  on("exportCsvBtn", "click", () => {
    Compare.exportCSV(
      STATE.lastAnalysisRows,
      STATE.lastAnalysisFileName
    );
  });

  on("catalogSupplierSelect", "change", () => {
    void UI.renderCatalogFiles();
  });

  on("refreshCatalogFilesBtn", "click", () => {
    void UI.renderCatalogFiles();
  });

  Storage.loadSuppliers();
  UI.refreshAll();
  UI.setLoading("exportCsvBtn", true);
  API.startMonitoring();

  // Évite qu'un chargement automatique écrase une action utilisateur.
  UI.setLoading("buildDbBtn", true);
  UI.setLoading("loadDbBtn", true);
  UI.setLoading("runBtn", true);

  try {
    const data = await API.loadDB();
    STATE.jsonDB = Utils.isRecord(data.db) ? data.db : {};

    STATE.suppliersList = [...new Set([
      ...STATE.suppliersList,
      ...Object.keys(STATE.jsonDB)
    ])].sort();

    Storage.saveSuppliers();
    UI.refreshAll();
    await UI.renderCatalogFiles();

    UI.text("summary", "Base JSON globale chargée automatiquement.");
  } catch (error) {
    UI.text(
      "summary",
      `Chargement automatique impossible : ${Utils.errorMessage(error)}`
    );
  } finally {
    UI.setLoading("buildDbBtn", false);
    UI.setLoading("loadDbBtn", false);
    UI.setLoading("runBtn", false);
  }
}

if (typeof document !== "undefined") {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initializeApp, { once: true });
  } else {
    void initializeApp();
  }
}

// Export pour les tests Node (sans effet dans le navigateur).
if (typeof module !== "undefined" && module.exports) {
  module.exports = { CONFIG, Utils, PDFHelpers, PDF, Compare };
}
