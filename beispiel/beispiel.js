// SPDX-License-Identifier: Apache-2.0
//
// Beispiel: je eine Folie pro Masterfolie.
//   pnpm run beispiel                       →  beispiel/beispiel.pptx (ohne Logo)
//   pnpm run beispiel -- diday-logo.svg     →  mit Logo

const path = require("path");
const pptxgen = require("pptxgenjs");
const did = require("..");

async function build(file, logo) {
  const pres = new pptxgen();
  await did.setup(pres, { date: "01.11.2026", footer: "Datenburg & Brotfabrik", logo });
  const C = pres.SchemeColor;

  let s = pres.addSlide({ masterName: did.MASTER.TITEL });
  s.addText("Titel des Vortrags", { placeholder: "title" });
  s.addText("Untertitel – worum es geht", { placeholder: "body" });

  s = pres.addSlide({ masterName: did.MASTER.AUSSAGE });
  s.addText("Eine Kernaussage in der Sprechblase", { placeholder: "title" });
  s.addText("Erläuterung rechts neben der Spitze.", { placeholder: "body" });

  s = pres.addSlide({ masterName: did.MASTER.INHALT });
  s.addText("Inhaltsfolie", { placeholder: "title" });
  s.addText("Einleitungssatz im gelben Kopfband.", { placeholder: "lead" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.0, w: 5.9, h: 3.0, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" } });
  s.addText("Karte auf heller Fläche (Schemafarbe bg2)", { x: 0.9, y: 2.3, w: 5.3, h: 0.6, fontSize: 20, bold: true, color: C.text2 });
  s.addText([{ text: "Quellen: " , options: { bold: true } }, { text: "[1] di.day", options: { hyperlink: { url: "https://di.day/" }, color: did.QUELLE_COLOR.INHALT } }], { placeholder: "quelle" });

  s = pres.addSlide({ masterName: did.MASTER.INHALT_DUNKEL });
  s.addText("Dunkle Inhaltsfolie", { placeholder: "title" });
  s.addText("Für Folien, die herausstechen sollen.", { placeholder: "lead" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.0, w: 5.9, h: 3.0, rectRadius: 0.12, fill: { color: C.accent2 }, line: { type: "none" } });
  s.addText("Karte in Blau (accent2)", { x: 0.9, y: 2.3, w: 5.3, h: 0.6, fontSize: 20, bold: true, color: C.background1 });

  await pres.writeFile({ fileName: file });
  await did.finish(file);
  return file;
}

if (require.main === module) {
  const logo = process.argv[2];
  build(path.join(__dirname, "beispiel.pptx"), logo).then((f) => console.log("geschrieben:", f));
}

module.exports = { build };
