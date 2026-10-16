"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");

const { CONFIG, PDF, Compare } = require("../app.js");
const fixture = require("./fixtures/cgti_d26_re060.js");

CONFIG.DEBUG_CGTI = false;

const quiet = (fn) => {
  const { warn, table, info } = console;
  console.warn = console.table = console.info = () => {};
  try {
    return fn();
  } finally {
    Object.assign(console, { warn, table, info });
  }
};

const parse = (lines = fixture.lines()) =>
  quiet(() => PDF.extractRows(lines, "D26 RE060.pdf", "CGTI"));

const analyse = () => Compare.execute(parse(), fixture.catalog, 0.01);

const byCode = (rows, code) => rows.filter(row => row.code === code);

const one = (rows, code) => {
  const found = byCode(rows, code);
  assert.equal(found.length, 1, `une seule ligne ${code} attendue`);
  return found[0];
};

test("CGTI : VRD02 lu en M3 à 84,11", () => {
  const row = one(parse(), "VRD02");

  assert.equal(row.unit, "M3");
  assert.equal(row.pu, 84.11);
  assert.equal(row.quantity, 12.5);
  assert.equal(row.total, 1051.38);
  assert.equal(row.extractionIssue, "");
});

test("CGTI : le 3 de M3 n'est pas absorbé par le PU (GRO03, GRO04)", () => {
  const rows = parse();

  const gro03 = one(rows, "GRO03");
  assert.equal(gro03.unit, "M3");
  assert.equal(gro03.pu, 145.6);
  assert.equal(gro03.quantity, 2);
  assert.equal(gro03.total, 291.2);

  const gro04 = one(rows, "GRO04");
  assert.equal(gro04.unit, "M3");
  assert.equal(gro04.pu, 182.3);
  assert.equal(gro04.quantity, 3.5);
  assert.equal(gro04.total, 638.05);
});

test("CGTI : RN25 n'est plus perturbé par la ligne hors bordereau voisine", () => {
  const row = one(parse(), "RN25");

  assert.equal(row.unit, "U");
  assert.equal(row.pu, 357.49);
  assert.equal(row.total, 357.49);
  assert.equal(row.extractionIssue, "");
});

test("CGTI : les lignes hors bordereau sont capturées comme HB", () => {
  const rows = parse();
  const hb = byCode(rows, "HB");

  assert.deepEqual(
    hb.map(row => [row.unit, row.pu, row.quantity, row.total]),
    [
      ["U", 120, 1, 120],
      ["ML", 25, 4, 100],
      ["FORFAIT", 450, 1, 450]
    ]
  );

  assert.match(hb[0].designation, /HORS BORDEREAU/);
  assert.match(hb[2].designation, /dépose de clôture/);
});

test("CGTI : aucun en-tête, titre ou total n'est pris pour une ligne", () => {
  const codes = parse().map(row => row.code);

  assert.deepEqual(codes, [
    "VRD02", "GRO03", "GRO04", "RN25", "HB",
    "2ND17", "HB", "MOE02", "HB"
  ]);
});

test("CGTI : MOE02 reste tiered_amount et conserve OFFERT", () => {
  const row = one(parse(), "MOE02");

  assert.equal(row.priceKind, "tiered_amount");
  assert.equal(row.pdfValue, "OFFERT");
  assert.equal(row.offered, true);
  assert.match(row.raw, /Tranche 1/);
});

test("Comparaison du devis D26 RE060", () => {
  const { output, stats } = analyse();
  const status = (code) => byCode(output, code).map(row => row.status);

  assert.deepEqual(status("VRD02"), ["OK"]);
  assert.deepEqual(status("GRO03"), ["OK"]);
  assert.deepEqual(status("GRO04"), ["OK"]);
  assert.deepEqual(status("RN25"), ["OK"]);

  // Écart réel de catalogue : U dans le PDF, KG dans Excel.
  assert.deepEqual(status("2ND17"), ["KO_UNITE"]);
  assert.match(one(output, "2ND17").detail, /PDF=U.*Excel=KG/);

  const moe = one(output, "MOE02");
  assert.equal(moe.status, "A_VERIFIER");
  assert.match(moe.detail, /OFFERT/);

  assert.deepEqual(status("HB"), [
    "HB_A_VERIFIER", "HB_A_VERIFIER", "HB_A_VERIFIER"
  ]);

  assert.equal(stats.ok, 4);
  assert.equal(stats.koUnit, 1);
  assert.equal(stats.koPrice, 0);
  assert.equal(stats.aVerifier, 1);
  assert.equal(stats.hb, 3);
  assert.equal(stats.total, 9);
});

test("CGTI : une extraction ambiguë donne A_VERIFIER et non KO", () => {
  const lines = fixture.lines();

  // Deux unités concurrentes dans la cellule de GRO03.
  const gro03 = lines.find(line => line.items.some(i => i.str === "GRO03"));
  gro03.items.push(fixture.item(fixture.COLUMNS.unit + 12, gro03.y, "U", 6));

  // PU de RN25 illisible.
  const rn25 = lines.find(line => line.items.some(i => i.str === "RN25"));
  rn25.items = rn25.items.map(value => value.str === "357,49 €"
    ? { ...value, str: "357,4#" }
    : value);

  const output = quiet(() => Compare.execute(
    PDF.extractRows(lines, "D26 RE060.pdf", "CGTI"),
    fixture.catalog,
    0.01
  )).output;

  const gro = one(output, "GRO03");
  assert.equal(gro.status, "A_VERIFIER");
  assert.match(gro.detail, /Extraction PDF incertaine/);

  const rn = one(output, "RN25");
  assert.equal(rn.status, "A_VERIFIER");
  assert.match(rn.detail, /PU/);
});

test("CGTI : incohérence PU × quantité ≠ total signalée A_VERIFIER", () => {
  const lines = fixture.lines();
  const vrd = lines.find(line => line.items.some(i => i.str === "VRD02"));
  vrd.items = vrd.items.map(value => value.str === "1 051,38 €"
    ? { ...value, str: "9 999,00 €" }
    : value);

  const output = quiet(() => Compare.execute(
    PDF.extractRows(lines, "D26 RE060.pdf", "CGTI"),
    fixture.catalog,
    0.01
  )).output;

  assert.equal(one(output, "VRD02").status, "A_VERIFIER");
});

test("CGTI : fragment unique couvrant plusieurs colonnes", () => {
  const lines = fixture.lines();
  const rn25 = lines.find(line => line.items.some(i => i.str === "RN25"));

  rn25.items = [
    ...rn25.items.filter(value => value.x < fixture.COLUMNS.unit),
    fixture.item(
      fixture.COLUMNS.unit,
      rn25.y,
      "U   357,49 €   1,00   357,49 €",
      fixture.COLUMNS.total + 40 - fixture.COLUMNS.unit
    )
  ];

  const row = one(parse(lines), "RN25");
  assert.equal(row.unit, "U");
  assert.equal(row.pu, 357.49);
  assert.equal(row.quantity, 1);
  assert.equal(row.total, 357.49);
});
