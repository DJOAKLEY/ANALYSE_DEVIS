"use strict";

/*
 * Reconstitution des fragments PDF.js du devis CGTI "D26 RE060".
 *
 * Particularités reproduites :
 * - "M3" est rendu en deux fragments : "M" puis un "3" en exposant
 *   (ligne de base légèrement plus haute) ;
 * - des lignes "HORS BORDEREAU" sans code sont intercalées entre
 *   deux prestations codifiées ;
 * - MOE02 est "OFFERT" et suivi d'une ligne de tranche sans code ;
 * - le numéro de devis "RE060" ressemble à un code de prestation.
 */

const PAGE_WIDTH = 595;

const COLUMNS = {
  code: 30,
  designation: 90,
  unit: 420,
  pu: 458,
  quantity: 515,
  total: 548
};

const item = (x, y, str, width = Math.max(4, str.length * 4.4)) => ({
  x, y, width, str
});

const priced = (y, { code, designation, unit, pu, quantity, total }) => {
  const items = [];

  if (code) items.push(item(COLUMNS.code, y, code, 28));
  if (designation) items.push(item(COLUMNS.designation, y, designation, 260));

  if (unit === "M3") {
    items.push(item(COLUMNS.unit, y, "M", 7));
    items.push(item(COLUMNS.unit + 7, y + 2, "3", 3.5));
  } else if (unit) {
    items.push(item(COLUMNS.unit, y, unit, 8));
  }

  if (pu) items.push(item(COLUMNS.pu, y, pu, 40));
  if (quantity) items.push(item(COLUMNS.quantity, y, quantity, 20));
  if (total) items.push(item(COLUMNS.total, y, total, 40));

  return items;
};

const items = [
  item(30, 790, "CGTI", 30),
  item(30, 775, "Devis", 25),
  item(58, 775, "N°", 10),
  item(72, 775, "D26", 18),
  item(94, 775, "RE060", 28),

  // En-tête du tableau.
  item(30, 700, "Code", 22),
  item(90, 700, "Désignation", 52),
  item(416, 700, "Unité", 22),
  item(455, 700, "Prix", 18),
  item(475, 700, "unitaire", 32),
  item(512, 700, "Quantité", 34),
  item(552, 700, "Total HT", 34),

  // Titre de section : aucune valeur.
  item(90, 682, "TERRASSEMENT - VRD", 90),

  ...priced(664, {
    code: "VRD02",
    designation: "Déblais en terrain de toute nature",
    unit: "M3",
    pu: "84,11 €",
    quantity: "12,50",
    total: "1 051,38 €"
  }),
  ...priced(648, {
    code: "GRO03",
    designation: "Béton de propreté dosé à 150 kg",
    unit: "M3",
    pu: "145,60 €",
    quantity: "2,00",
    total: "291,20 €"
  }),
  ...priced(632, {
    code: "GRO04",
    designation: "Béton armé pour massif",
    unit: "M3",
    pu: "182,30 €",
    quantity: "3,50",
    total: "638,05 €"
  }),
  ...priced(616, {
    code: "RN25",
    designation: "Regard de visite 80x80",
    unit: "U",
    pu: "357,49 €",
    quantity: "1,00",
    total: "357,49 €"
  }),
  ...priced(606, {
    designation: "HORS BORDEREAU - Fourniture tampon fonte",
    unit: "U",
    pu: "120,00 €",
    quantity: "1,00",
    total: "120,00 €"
  }),
  ...priced(590, {
    code: "2ND17",
    designation: "Grillage avertisseur",
    unit: "U",
    pu: "2,15 €",
    quantity: "40,00",
    total: "86,00 €"
  }),
  ...priced(574, {
    code: "HB",
    designation: "Plus-value pour reprise d'enrobé",
    unit: "ML",
    pu: "25,00 €",
    quantity: "4,00",
    total: "100,00 €"
  }),
  ...priced(558, {
    code: "MOE02",
    designation: "Maîtrise d'œuvre - études d'exécution",
    unit: "F",
    pu: "OFFERT",
    quantity: "1,00",
    total: "OFFERT"
  }),
  ...priced(548, {
    designation: "Tranche 1 : de 0 à 50 000 € HT",
    pu: "0,00 €",
    quantity: "5,00",
    total: "0,00 €"
  }),
  ...priced(530, {
    designation: "Hors bordereau : dépose de clôture existante",
    unit: "F",
    pu: "450,00 €",
    quantity: "1,00",
    total: "450,00 €"
  }),

  // Pied de tableau.
  item(458, 500, "TOTAL HT", 40),
  item(COLUMNS.total, 500, "3 093,32 €", 40),
  item(458, 488, "TVA 8,50 %", 40),
  item(COLUMNS.total, 488, "262,93 €", 40),
  item(458, 476, "TOTAL TTC", 40),
  item(COLUMNS.total, 476, "3 356,25 €", 40)
];

// Même regroupement que PDF.toLines.
const toLines = (pageItems, page = 1) => {
  const grouped = new Map();

  for (const value of pageItems) {
    const yKey = Math.round(value.y * 2) / 2;
    if (!grouped.has(yKey)) grouped.set(yKey, []);
    grouped.get(yKey).push(value);
  }

  return [...grouped.keys()].sort((a, b) => b - a).map(y => {
    const lineItems = grouped.get(y).sort((a, b) => a.x - b.x);

    return {
      page,
      pageWidth: PAGE_WIDTH,
      y,
      items: lineItems,
      text: lineItems.map(value => value.str).join(" ")
    };
  });
};

const catalog = new Map([
  ["VRD02", [{ unit: "M3", pu: 84.11 }]],
  ["GRO03", [{ unit: "M3", pu: 145.6 }]],
  ["GRO04", [{ unit: "M3", pu: 182.3 }]],
  ["RN25", [{ unit: "U", pu: 357.49 }]],
  ["2ND17", [{ unit: "KG", pu: 2.15 }]],
  ["MOE02", [{ unit: "F", pu: 0 }]]
]);

module.exports = {
  PAGE_WIDTH,
  COLUMNS,
  item,
  priced,
  toLines,
  lines: () => toLines(items.map(value => ({ ...value }))),
  catalog
};
