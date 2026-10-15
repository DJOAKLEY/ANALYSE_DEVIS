"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { test } = require("node:test");

const source = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8");

function parser(debug = false) {
  const messages = { warn: [], table: [], info: [] };
  const context = vm.createContext({
    document: { readyState: "loading", addEventListener() {} },
    console: Object.fromEntries(Object.keys(messages).map(key =>
      [key, (...args) => messages[key].push(args)]
    ))
  });
  vm.runInContext(source + `
    CONFIG.DEBUG_CGTI = ${debug};
    globalThis.parse = PDF.parseCgtiRows;
  `, context);
  return {
    parse: lines => structuredClone(context.parse(lines)),
    messages
  };
}

// Fixtures synthétiques, non extraites de PDF réels, au format PDF.toLines.
function fragment(str, x, y = 400, width = str.length * 4) {
  return { str, x, y, width };
}

function prestation({
  code = "RN25", y = 400, unit = "U",
  pu = "357,69", quantity = "1,00", total = "357,69",
  currencies = true, description = "Fourniture et pose"
} = {}) {
  return [
    fragment(code, 30, y),
    fragment(description, 100, y),
    fragment(unit, 350, y),
    fragment(pu, 410, y),
    ...(currencies ? [fragment("€", 460, y)] : []),
    fragment(quantity, 485, y),
    fragment(total, 550, y),
    ...(currencies ? [fragment("EUR", 610, y)] : [])
  ];
}

function lines(items, page = 1) {
  const groups = new Map();
  for (const item of items) {
    const y = Math.round(item.y * 2) / 2;
    if (!groups.has(y)) groups.set(y, []);
    groups.get(y).push(item);
  }
  return [...groups].sort((a, b) => b[0] - a[0]).map(([y, fragments]) => ({
    page, pageWidth: 650, y, items: fragments,
    text: fragments.map(item => item.str).join(" ")
  }));
}

function header(y = 450) {
  return [
    fragment("Unité", 350, y),
    fragment("P.U.", 410, y),
    fragment("Quantité", 485, y),
    fragment("Total HT", 550, y)
  ];
}

function values(row) {
  return [row.unit, row.pu, row.quantity, row.total];
}

test("cas nominal : ordre, champs de sortie, désignation et absence de mutation", () => {
  const { parse, messages } = parser();
  const input = lines(prestation());
  const before = structuredClone(input);
  input.forEach(line => {
    line.items.forEach(Object.freeze);
    Object.freeze(line.items);
    Object.freeze(line);
  });
  Object.freeze(input);
  const [row] = parse(input);
  assert.deepEqual(values(row), ["U", 357.69, 1, 357.69]);
  assert.deepEqual(Object.keys(row), [
    "code", "codeLabel", "unit", "pu", "quantity", "total", "designation",
    "priceKind", "puText", "page", "y", "rowKey", "raw"
  ]);
  assert.equal(row.code, "RN25");
  assert.equal(row.codeLabel, "RN25");
  assert.equal(row.designation, "Fourniture et pose");
  assert.equal(row.priceKind, "amount");
  assert.equal(row.puText, "357,69");
  assert.equal(row.page, 1);
  assert.equal(row.y, 400);
  assert.equal(row.rowKey, "1_4000");
  assert.match(row.raw, /RN25.*357,69.*EUR/);
  assert.deepEqual(input, before);
  assert.equal(messages.warn.length, 0);
  assert.equal(messages.info.length, 0);
  assert.equal(messages.table.length, 0);
});

test("aucune permutation PU/quantité, même avec PU = 1, zéro ou produit incohérent", () => {
  for (const [pu, quantity, total, expected] of [
    ["357,69", "2,00", "715,38", [357.69, 2, 715.38]],
    ["1,00", "357,69", "357,69", [1, 357.69, 357.69]],
    ["0,00", "2,00", "0,00", [0, 2, 0]],
    ["357,69", "0,00", "0,00", [357.69, 0, 0]],
    ["−12,50", "-2,00", "0,00", [-12.5, -2, 0]],
    ["10,00", "2,00", "999,00", [10, 2, 999]]
  ]) {
    const [row] = parser().parse(lines(prestation({ pu, quantity, total })));
    assert.deepEqual([row.pu, row.quantity, row.total], expected);
  }
});

test("milliers dans les trois colonnes : espaces ordinaires, insécables et fines", () => {
  for (const space of [" ", "\u00a0", "\u202f", "\u2009"]) {
    const [row] = parser().parse(lines(prestation({
      pu: `1${space}234,56`, quantity: `2${space}345,00`, total: `9${space}876,54`
    })));
    assert.deepEqual(values(row), ["U", 1234.56, 2345, 9876.54]);
  }
});

test("codes, unité accentuée, nombres et devises fragmentés avec petits décalages", () => {
  const items = [
    fragment("2", 30), fragment("ND", 35), fragment("16", 44),
    fragment("BIS", 53), fragment("Pose", 100),
    fragment("Uni", 350), fragment("té", 363, 400.4),
    fragment("357", 410), fragment(",", 423, 400.4), fragment("69", 428),
    fragment("€", 460, 400.6),
    fragment("1", 485), fragment(",", 490), fragment("00", 495),
    fragment("357,69", 550),
    fragment("E", 610), fragment("U", 615, 400.5), fragment("R", 620)
  ];
  const [row] = parser().parse(lines(items));
  assert.equal(row.code, "2ND16BIS");
  assert.deepEqual(values(row), ["U", 357.69, 1, 357.69]);
});

test("m² fragmenté : ne pas confondre une lecture partielle m avec une autre unité", () => {
  const items = prestation({ unit: "m" });
  items.push(fragment("²", 355, 401.5));
  const [row] = parser().parse(lines(items));
  assert.equal(row.unit, "M2");
});

test("unité seule décalée verticalement : compléter les lectures numériques équivalentes", () => {
  const items = prestation({ unit: "Unité" });
  items.find(item => item.str === "Unité").y += 4;
  const { parse, messages } = parser();
  assert.deepEqual(values(parse(lines(items))[0]), ["U", 357.69, 1, 357.69]);
  assert.equal(messages.warn.length, 0);
});

test("code et suffixe décalés : reconnaître le code complet une seule fois", () => {
  const items = prestation().filter(item => item.x !== 30);
  items.push(fragment("RN", 30), fragment("25", 39, 401.5), fragment("TER", 48, 402));
  const rows = parser().parse(lines(items));
  assert.equal(rows.length, 1);
  assert.equal(rows[0].code, "RN25TER");
  assert.deepEqual(values(rows[0]), ["U", 357.69, 1, 357.69]);
});

test("unité et PU dans un même fragment : préfixe d'unité complet uniquement", () => {
  const items = prestation().filter(item => item.x !== 350 && item.x !== 410);
  items.push(fragment("Unité 357,69", 370));
  assert.deepEqual(values(parser().parse(lines(items))[0]), ["U", 357.69, 1, 357.69]);
});

test("signe moins isolé et milliers découpés dans une cellule", () => {
  const items = prestation({ pu: "", total: "-1234,56" }).filter(item => item.x !== 410);
  items.push(fragment("−", 410), fragment("1", 416), fragment("234", 422),
    fragment(",", 435), fragment("56", 440));
  assert.deepEqual(values(parser().parse(lines(items))[0]), ["U", -1234.56, 1, -1234.56]);
});

test("prestations très proches : bandes disjointes et aucune récupération voisine", () => {
  const items = [
    ...prestation({ code: "RN25", y: 400, pu: "illisible", quantity: "2", total: "20" }),
    ...prestation({ code: "ELE151", y: 396, pu: "40", quantity: "3", total: "120" }),
    ...prestation({ code: "IC18", y: 392, pu: "7", quantity: "4", total: "28" })
  ];
  const rows = parser().parse(lines(items));
  assert.deepEqual(rows.map(row => row.code), ["RN25", "ELE151", "IC18"]);
  assert.deepEqual(values(rows[0]), ["U", NaN, 2, 20]);
  assert.deepEqual(values(rows[1]), ["U", 40, 3, 120]);
  assert.deepEqual(values(rows[2]), ["U", 7, 4, 28]);
});

test("codes supportés et répétitions légitimes, malgré les doublons de fragments", () => {
  const items = ["RN25", "ELE151", "IC18", "FO12", "PYL100", "2ND16", "HB",
    "RN25BIS", "FO12TER", "RN25"].flatMap((code, i) =>
    prestation({ code, y: 700 - i * 20 })
  );
  const rows = parser().parse([
    ...lines([...items, ...items]),
    ...lines(prestation({ code: "RN25" }), 2)
  ]);
  assert.equal(rows.length, 11);
  assert.equal(rows.filter(row => row.code === "RN25").length, 3);
  assert.equal(new Set(rows.map(row => row.rowKey)).size, 11);
});

test("lectures chiffrées incompatibles : préserver seulement les cellules en accord", () => {
  const items = prestation({ pu: "10", quantity: "2", total: "20" });
  items.push(...prestation({ y: 405, pu: "30", quantity: "2", total: "60" })
    .filter(item => item.x >= 410));
  const { parse, messages } = parser();
  assert.deepEqual(values(parse(lines(items))[0]), ["U", NaN, 2, NaN]);
  assert.match(messages.table[0][0][0].erreur, /incompatibles.*PU.*total/);
});

test("deux unités reconnues différentes : diagnostic plutôt que choix arbitraire", () => {
  const items = prestation();
  items.push(fragment("ML", 350, 404));
  const { parse, messages } = parser();
  assert.deepEqual(values(parse(lines(items))[0]), ["", 357.69, 1, 357.69]);
  assert.match(messages.table[0][0][0].erreur, /ambiguës.*unité/);
});

test("cellules illisibles : conservation indépendante du PU, de la quantité et du total", () => {
  for (const [field, expected] of [
    ["pu", ["U", NaN, 1, 357.69]],
    ["quantity", ["U", 357.69, NaN, 357.69]],
    ["total", ["U", 357.69, 1, NaN]]
  ]) {
    const { parse, messages } = parser();
    const [row] = parse(lines(prestation({ [field]: "illisible" })));
    assert.deepEqual(values(row), expected);
    assert.equal(row.code, "RN25");
    assert.equal(messages.warn.length, 1);
    assert.match(messages.table[0][0][0].erreur, /illisibles/);
  }
});

test("nombres bornés : aucun suffixe numérique récupéré dans une cellule invalide", () => {
  for (const pu of ["abc357,69", "1 23,45", "12,34,56", "1.234,56", "10 € texte"]) {
    const [row] = parser().parse(lines(prestation({ pu })));
    assert.ok(Number.isNaN(row.pu), pu);
  }
});

test("quantité/total textuels ambigus avec milliers : ne pas choisir une partition", () => {
  const items = prestation().filter(item => item.x !== 485 && item.x !== 550);
  items.push(fragment("1 234 567", 485));
  const { parse, messages } = parser();
  const [row] = parse(lines(items));
  assert.deepEqual(values(row), ["U", 357.69, NaN, NaN]);
  assert.match(messages.table[0][0][0].erreur, /ambiguës/);
});

test("quantité/total dans un même fragment non ambigu : lecture complète", () => {
  const items = prestation().filter(item => item.x !== 485 && item.x !== 550);
  items.push(fragment("1,00 357,69", 485));
  assert.deepEqual(values(parser().parse(lines(items))[0]), ["U", 357.69, 1, 357.69]);
});

test("sans devises ni en-tête fiable : conserver le code, pas trois nombres arbitraires", () => {
  const { parse, messages } = parser();
  const [row] = parse(lines(prestation({
    currencies: false, description: "U câble 123 456 789 m"
  })));
  assert.deepEqual(values(row), ["", NaN, NaN, NaN]);
  assert.equal(row.code, "RN25");
  assert.equal(messages.warn.length, 1);
});

test("colonnes explicitement reconnues : lecture sans devises et cellules indépendantes", () => {
  const { parse, messages } = parser();
  const [row] = parse(lines([...header(), ...prestation({
    currencies: false, unit: "Unité", pu: "1", quantity: "illisible", total: "0"
  })]));
  assert.deepEqual(values(row), ["U", 1, NaN, 0]);
  assert.equal(row.designation, "Fourniture et pose");
  assert.equal(messages.warn.length, 1);
});

test("pas d'unité piochée dans la description", () => {
  const [row] = parser().parse(lines(prestation({
    unit: "?", description: "Pose U câble m"
  })));
  assert.deepEqual(values(row), ["", 357.69, 1, 357.69]);
});

test("faux codes et nombres de description, d'en-tête ou de pied de page", () => {
  const items = [
    ...prestation(),
    fragment("Référence RN99", 30, 460),
    fragment("RN99", 110, 420),
    fragment("DEVIS2024", 30, 470),
    fragment("HB", 30, 25),
    fragment("12345", 30, 380),
    fragment("Page 2ND16", 30, 30),
    fragment("Référence : ELE151", 100, 400)
  ];
  assert.deepEqual(parser().parse(lines(items)).map(row => row.code), ["RN25"]);
});

test("largeur manquante et aucun code reconnu : erreurs explicites en français", () => {
  const input = lines(prestation());
  input.forEach(line => { delete line.pageWidth; });
  assert.throws(() => parser().parse(input), /largeur PDF absente/);
  assert.throws(() => parser().parse(lines([fragment("Description RN25", 100)])),
    /aucun code de prestation reconnu/);
});

test("DEBUG_CGTI contrôle le debug, pas les avertissements utiles", () => {
  const { parse, messages } = parser(true);
  parse(lines(prestation()));
  assert.equal(messages.info.length, 1);
  assert.equal(messages.table.length, 1);
  assert.equal(messages.warn.length, 0);
});
