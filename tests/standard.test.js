"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const { PDF, Compare } = require("../app.js");

const item = (x, y, str, width = 24) => ({ x, y, str, width });

const line = (y, items, page = 1) => ({
  page,
  y,
  items,
  text: items.map(value => value.str).join(" ")
});

const header = (supplier) => line(750, [
  item(30, 750, supplier === "AXIANS" ? "Référence" : "Code"),
  item(100, 750, "Désignation", 200),
  item(360, 750, "Unité", 16),
  item(400, 750, "Qté", 20),
  item(450, 750, "PU HT", 40),
  item(530, 750, supplier === "AXIANS" ? "PT HT" : "Total", 40)
]);

const priced = (y, code, designation, unit, quantity, pu, total) => line(y, [
  ...(code ? [item(30, y, code, 40)] : []),
  item(100, y, designation, 200),
  item(360, y, unit, 16),
  item(400, y, quantity, 20),
  item(450, y, pu, 40),
  ...(total ? [item(530, y, total, 40)] : [])
]);

const newcomRows = () => [
  priced(710, "", "Acheminement du camion de Rivière des pluies vers ST PIERRE",
    "F", "2,00", "404,80", "809,60"),
  priced(680, "", "Installation de chantier avec Barrière de type Heras",
    "F", "1,00", "264,00", "264,00"),
  ...["PCE19", "PCE19", "IC27", "RN46", "RN25", "ANN11BIS"].map((code, i) =>
    priced(650 - i * 25, code, `Prestation ${i}`, "U", "1,00", "10,00", "10,00")
  )
];

for (const supplier of ["NEWCOM", "AXIANS"]) {
  const parse = (rows) => PDF.extractRows(
    [header(supplier), ...rows], `${supplier}.pdf`, supplier
  );

  test(`${supplier} : deux prestations sans code et six lignes codées distinctes`, () => {
    const rows = parse(newcomRows());
    assert.deepEqual(rows.map(row => row.code),
      ["HB", "HB", "PCE19", "PCE19", "IC27", "RN46", "RN25", "ANN11BIS"]);
    assert.deepEqual(rows.slice(0, 2).map(row =>
      [row.unit, row.pu, row.quantity, row.total, row.sourceCode, row.outsideCatalog]
    ), [
      ["FORFAIT", 404.8, 2, 809.6, "", true],
      ["FORFAIT", 264, 1, 264, "", true]
    ]);
    assert.match(rows[0].outsideLabel, /Acheminement/);
    assert.match(rows[1].designation, /Installation/);
  });

  test(`${supplier} : une désignation débordant dans la colonne code n'invente pas de code`, () => {
    const rows = newcomRows();
    rows[0].items[0] = item(30, rows[0].y, rows[0].items[0].str, 160);
    rows[0].text = rows[0].items.map(value => value.str).join(" ");
    const result = parse(rows);
    assert.equal(result.length, 8);
    assert.equal(result[0].code, "HB");
    assert.match(result[0].designation, /^Acheminement/);
  });

  test(`${supplier} : totaux, pourcentages et cellules incomplètes ne deviennent pas HB`, () => {
    const result = parse([
      ...newcomRows(),
      priced(400, "", "Sous-total", "F", "1", "100", "100"),
      priced(380, "", "TVA", "%", "20", "100", "20"),
      priced(360, "", "Remise", "F", "1", "10", "10"),
      priced(340, "", "Prestation sans quantité", "F", "", "100", "100"),
      priced(320, "", "Prestation sans PU", "F", "1", "", "100"),
      priced(300, "", "Prestation sans unité", "", "1", "100", "100"),
      priced(280, "", "Prestation sans total", "F", "1", "100", "")
    ]);
    assert.equal(result.length, 9);
    assert.equal(result.at(-1).code, "HB");
    assert.ok(Number.isNaN(result.at(-1).total));
  });

  test(`${supplier} : comparaison des codes, unités, prix et HB préservée`, () => {
    const rows = parse(newcomRows());
    const catalog = new Map([
      ["PCE19", [{ unit: "U", pu: 10 }]],
      ["IC27", [{ unit: "KG", pu: 10 }]],
      ["RN46", [{ unit: "U", pu: 20 }]],
      ["RN25", [{ unit: "U", pu: 10 }]],
      ["ANN11BIS", [{ unit: "U", pu: 10 }]]
    ]);
    const { output, stats } = Compare.execute(rows, catalog, 0.01);
    assert.deepEqual(output.map(row => row.status), [
      "HB_A_VERIFIER", "HB_A_VERIFIER", "OK", "OK",
      "KO_UNITE", "KO_PRIX", "OK", "OK"
    ]);
    assert.match(output[0].detail, /Acheminement.*cohérent/);
    assert.equal(stats.total, 8);
    assert.equal(stats.hb, 2);
  });

  test(`${supplier} : un code mentionné dans la désignation ne remplace pas un code absent`, () => {
    const [row] = parse([
      priced(710, "", "RN25 remplacement hors catalogue", "F", "1", "100", "100")
    ]);
    assert.equal(row.code, "HB");
    assert.equal(row.designation, "RN25 remplacement hors catalogue");
  });

  test(`${supplier} : HB explicite et mise en page héritée sur la page suivante`, () => {
    const secondPage = priced(710, "HB", "Prestation hors bordereau", "F", "2", "20", "40");
    secondPage.page = 2;
    const [row] = parse([secondPage]);
    assert.equal(row.code, "HB");
    assert.equal(row.page, 2);
    assert.equal(row.unit, "FORFAIT");
    const { output } = Compare.execute([row], new Map(), 0.01);
    assert.equal(output[0].status, "HB_A_VERIFIER");
    assert.match(output[0].detail, /Prestation hors bordereau.*cohérent/);
  });

  test(`${supplier} : aucune unité déduite d'une lettre dans la désignation sans code`, () => {
    const invalid = priced(710, "", "F", "", "1", "100", "100");
    invalid.items[0] = item(305, 710, "F", 8);
    assert.equal(parse([...newcomRows(), invalid]).length, 8);
  });

  test(`${supplier} : maîtrise d'œuvre conserve la lecture des pourcentages`, () => {
    const rows = parse([
      priced(710, "MOE01", "Maîtrise d'œuvre 5 %", "%", "5,00", "100,00", "5,00")
    ]);
    assert.equal(rows[0].code, "MOE01");
    assert.equal(rows[0].priceKind, "percentage");
    assert.equal(rows[0].pu, supplier === "AXIANS" ? 0.05 : 100);
    assert.equal(rows[0].quantity, supplier === "AXIANS" ? null : 5);
  });
}

test("NEWCOM : libellés HB échappés dans le tableau et conservés dans le CSV", async () => {
  const results = { innerHTML: "" };
  const link = { click() {}, remove() {} };
  let blob;
  const context = {
    module: { exports: {} },
    Blob,
    URL: {
      createObjectURL(value) {
        blob = value;
        return "blob:export";
      },
      revokeObjectURL() {}
    },
    setTimeout() {},
    document: {
      readyState: "loading",
      addEventListener() {},
      getElementById: () => results,
      createElement: (tag) => tag === "a" ? link : {},
      body: { appendChild() {} }
    }
  };
  vm.runInNewContext(fs.readFileSync(require.resolve("../app.js"), "utf8"), context);

  const { PDF: browserPDF, Compare: browserCompare } = context.module.exports;
  const rows = browserPDF.parseNewcomRows([
    header("NEWCOM"),
    priced(710, "", 'Installation <chantier>; "Heras"', "F", "1", "264", "264")
  ]);
  const { output } = browserCompare.execute(rows, new Map(), 0.01);
  browserCompare.renderResults(output);
  assert.match(results.innerHTML, /Installation &lt;chantier&gt;; &quot;Heras&quot;/);
  assert.doesNotMatch(results.innerHTML, /<chantier>/);
  assert.match(results.innerHTML, /HB_A_VERIFIER/);

  browserCompare.exportCSV(output, "devis-test");
  const csv = await blob.text();
  assert.match(csv, /PDF;Numéro devis;Code;Unité;PU;Statut;Détail;Page;RowKey;Texte original/);
  assert.match(csv, /;HB;FORFAIT;264,00;HB_A_VERIFIER;/);
  assert.match(csv, /Installation <chantier>; ""Heras""/);
  assert.match(link.download, /^devis-test_.*\.csv$/);
});
