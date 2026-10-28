// SPDX-License-Identifier: Apache-2.0
//
// DIDAY-Theme für pptxgenjs – nach der Vorlage „Diday-vorlage_v1.3“ des
// Digital Independence Day: Farben, Schrift (Noto Sans) und vier Masterfolien
// im Format 16:9 (13,333 × 7,5 Zoll). Das DIDAY-Logo ist nicht enthalten; wer
// es verwenden darf, übergibt es als Option (SVG oder PNG, Pfad oder Buffer).
//
//   const pptxgen = require("pptxgenjs");
//   const did = require("did-theme-template");
//   const pres = new pptxgen();
//   await did.setup(pres, { date: "01.11.2026", footer: "Datenburg & Brotfabrik", logo: "diday-logo.svg" });
//   const s = pres.addSlide({ masterName: did.MASTER.INHALT });
//   s.addText("Titel", { placeholder: "title" });
//   s.addText("Einleitung", { placeholder: "lead" });
//   await pres.writeFile({ fileName: "vortrag.pptx" });
//   await did.finish("vortrag.pptx");
//
// Masterfolien und ihre Platzhalter:
//   TITEL          gelb/lila, großes Logo          title, body, quelle
//   AUSSAGE        Sprechblase auf Gelb             title, body, quelle
//   INHALT         weiß, gelbes Kopfband            title, lead, quelle
//   INHALT_DUNKEL  lila, gelbes Kopfband            title, lead, quelle
// Alle haben die Fußzeile der Vorlage: Datum, Text, Seitenzahl.
// QUELLE_COLOR nennt je Master die Schriftfarbe der Quellenzeile, damit Links
// dort dieselbe Farbe bekommen können.
//
// finish() ist nötig, weil pptxgenjs aus pres.theme nur die Schriften übernimmt
// und Master-Formen nur als Rechteck kennt: Es schreibt die Farbpalette ins
// Theme der PPTX und macht aus dem Rechteck die Sprechblase der Vorlage.

const fs = require("fs");
const sharp = require("sharp");
// jszip kommt mit pptxgenjs und wird von dort aus aufgelöst (pnpm legt es nicht flach ab).
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

const THEME = {
  name: "DIDAY",
  headFontFace: "Noto Sans",
  bodyFontFace: "Noto Sans",
  colors: {
    dk1: "1D1D1B", // Fließtext (Schwarz aus dem Logo)
    lt1: "FFFFFF",
    dk2: "595082", // Lila der Vorlage: dunkle Flächen, Überschriften
    lt2: "EFEDF7", // helle Kartenfläche
    accent1: "F8C662", // Gelb der Vorlage
    accent2: "4E46DC", // Blau aus dem Logo
    accent3: "C8443F", // Warnung
    accent4: "2E9361", // gut (Grün der Vorlage)
    accent5: "D6D2EE", // helles Lila: Nebentexte auf dunklen Flächen
    accent6: "6B6787", // Grau-Lila für Nebentexte
    hlink: "4E46DC",
    folHlink: "6B6787",
  },
};

const MASTER = { TITEL: "TITEL", AUSSAGE: "AUSSAGE", INHALT: "INHALT", INHALT_DUNKEL: "INHALT_DUNKEL" };
// Schemafarben der Quellenzeile je Master (Werte von pptxgenjs SchemeColor).
const QUELLE_COLOR = { TITEL: "accent5", AUSSAGE: "tx2", INHALT: "accent6", INHALT_DUNKEL: "accent5" };
const CALLOUT = "Sprechblase";
// Spitze unten links, weit nach unten gezogen – wie die Modifier in der Vorlage.
const CALLOUT_GEOM = '<a:prstGeom prst="wedgeRectCallout"><a:avLst><a:gd name="adj1" fmla="val -34174"/><a:gd name="adj2" fmla="val 130852"/></a:avLst></a:prstGeom>';

// Logo (Pfad oder Buffer, SVG oder PNG) als PNG-Daten für pptxgenjs.
async function logoPng(logo) {
  const input = Buffer.isBuffer(logo) ? logo : fs.readFileSync(logo);
  const png = await sharp(input, { density: 300 }).resize(1600).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

// Setzt Format, Schriften und Masterfolien. date und footer stehen in der
// Fußzeile; logo (optional) erscheint groß auf TITEL und klein oben rechts.
async function setup(pres, { date = "", footer = "", logo } = {}) {
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const logoData = logo ? await logoPng(logo) : null;

  // Maße aus der Vorlage (28 × 15,75 cm), umgerechnet auf 13,333 × 7,5 Zoll.
  const quelle = (color) => ({ placeholder: { options: { name: "quelle", type: "body", x: 0.6, y: 6.78, w: 12.1, h: 0.34, fontSize: 9, color, valign: "top", margin: 0 }, text: "" } });
  const FOOT = { y: 7.142, h: 0.358, fontSize: 16, bold: true, valign: "middle", margin: 0 };
  const foot = (color) => [
    { text: { text: date, options: Object.assign({ x: 0.475, w: 3.8, color, align: "left" }, FOOT) } },
    { text: { text: footer, options: Object.assign({ x: 4.523, w: 4.286, color, align: "center" }, FOOT) } },
  ];
  const pageNumber = (color) => Object.assign({ x: 11.9, w: 0.95, color, align: "left" }, FOOT);
  const band = (y, h) => ({ rect: { x: 0, y, w: 13.333, h, fill: { color: C.accent1 }, line: { type: "none" } } });
  // Ohne Logo bleiben die Plätze leer (die Objektlisten filtern null heraus).
  const logoSmall = logoData && { image: { x: 11.138, y: 0.424, w: 1.719, h: 0.766, data: logoData, altText: "DIDAY" } };

  pres.defineSlideMaster({
    title: MASTER.TITEL,
    background: { color: C.text2 },
    objects: [
      band(0, 5.001),
      logoData && { image: { x: 4.523, y: 1.185, w: 4.286, h: 1.91, data: logoData, altText: "DIDAY" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 3.35, w: 12.13, h: 0.95, fontSize: 32, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 5.3, w: 12.13, h: 1.2, fontSize: 22, color: C.accent1, align: "center", valign: "top", margin: 0 }, text: "" } },
      quelle(QUELLE_COLOR.TITEL),
      ...foot(C.accent1),
    ].filter(Boolean),
    slideNumber: pageNumber(C.accent1),
  });

  pres.defineSlideMaster({
    title: MASTER.AUSSAGE,
    background: { color: C.accent1 },
    objects: [
      { rect: { x: 3.332, y: 1.785, w: 6.667, h: 2.499, fill: { color: C.background1 }, line: { color: C.accent4, width: 5.67 }, objectName: CALLOUT } },
      logoSmall,
      { placeholder: { options: { name: "title", type: "title", x: 3.571, y: 1.963, w: 6.19, h: 2.143, fontSize: 28, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0 }, text: "" } },
      // rechts neben der Spitze der Sprechblase
      { placeholder: { options: { name: "body", type: "body", x: 6.3, y: 4.5, w: 6.4, h: 1.0, fontSize: 20, color: C.text2, align: "left", valign: "top", margin: 0 }, text: "" } },
      quelle(QUELLE_COLOR.AUSSAGE),
      ...foot(C.text2),
    ].filter(Boolean),
    slideNumber: pageNumber(C.text2),
  });

  const content = (title, background) =>
    pres.defineSlideMaster({
      title,
      background: { color: background },
      objects: [
        band(0, 1.608),
        band(7.142, 0.358),
        logoSmall,
        { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.38, w: 10.3, h: 0.82, fontSize: 30, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
        { placeholder: { options: { name: "lead", type: "body", x: 0.6, y: 1.2, w: 12.1, h: 0.36, fontSize: 15, color: C.text2, valign: "top", margin: 0 }, text: "" } },
        quelle(QUELLE_COLOR[title]),
        ...foot(C.text2),
      ].filter(Boolean),
      slideNumber: pageNumber(C.text2),
    });
  content(MASTER.INHALT, C.background1);
  content(MASTER.INHALT_DUNKEL, C.text2);
}

function xmlAttr(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Nach pres.writeFile(): Farbpalette ins Theme, Sprechblase in den AUSSAGE-Master.
async function finish(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));

  const PART = "ppt/theme/theme1.xml";
  const entry = zip.file(PART);
  if (!entry) throw new Error(`${PART} fehlt in ${file}`);
  let xml = await entry.async("string");
  for (const [slot, hex] of Object.entries(THEME.colors)) {
    const re = new RegExp(`<a:${slot}>[\\s\\S]*?</a:${slot}>`);
    if (!re.test(xml)) throw new Error(`Farbslot ${slot} fehlt in ${PART}`);
    xml = xml.replace(re, `<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`);
  }
  const name = xmlAttr(THEME.name);
  xml = xml.replace(/(<a:theme\b[^>]*?\sname=")[^"]*(")/, `$1${name}$2`).replace(/(<a:clrScheme\b[^>]*?\sname=")[^"]*(")/, `$1${name}$2`);
  zip.file(PART, xml);

  let callouts = 0;
  for (const layout of zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/)) {
    const before = await layout.async("string");
    const after = before.replace(new RegExp(`(name="${CALLOUT}"[\\s\\S]*?)<a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>`), (_, head) => {
      callouts++;
      return head + CALLOUT_GEOM;
    });
    if (after !== before) zip.file(layout.name, after);
  }
  if (callouts !== 1) throw new Error(`${callouts} Sprechblasen gefunden, erwartet 1`);

  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { THEME, MASTER, QUELLE_COLOR, setup, finish };
