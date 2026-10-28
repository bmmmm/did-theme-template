// SPDX-License-Identifier: Apache-2.0
const test = require("node:test");
const assert = require("node:assert");
const fs = require("fs");
const os = require("os");
const path = require("path");
const sharp = require("sharp");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
const did = require("..");
const { build } = require("../beispiel/beispiel");

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "did-theme-"));
const read = async (file) => JSZip.loadAsync(fs.readFileSync(file));
const text = (zip, name) => zip.file(name).async("string");

async function layouts(zip) {
  const out = {};
  for (const f of zip.file(/^ppt\/slideLayouts\/slideLayout\d+\.xml$/)) {
    const xml = await f.async("string");
    const name = (xml.match(/<p:cSld name="([^"]+)"/) || [])[1];
    out[name] = { xml, rels: await text(zip, f.name.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels") };
  }
  return out;
}

test("ohne Logo: Palette, Schrift, Master, Sprechblase, Fußzeile", async () => {
  const zip = await read(await build(path.join(tmp, "ohne.pptx")));
  const theme = await text(zip, "ppt/theme/theme1.xml");
  for (const [slot, hex] of Object.entries(did.THEME.colors)) assert.match(theme, new RegExp(`<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`), slot);
  assert.match(theme, /<a:clrScheme name="DIDAY">/);
  assert.match(theme, /<a:latin typeface="Noto Sans"/);

  const L = await layouts(zip);
  for (const m of Object.values(did.MASTER)) assert.ok(L[m], `Master ${m} fehlt`);
  assert.match(L.AUSSAGE.xml, /prst="wedgeRectCallout"><a:avLst><a:gd name="adj1" fmla="val -34174"\/><a:gd name="adj2" fmla="val 130852"\/>/);
  for (const m of Object.values(did.MASTER)) {
    assert.match(L[m].xml, /<a:t>01\.11\.2026<\/a:t>/, `Datum in ${m}`);
    assert.match(L[m].xml, /<a:t>Datenburg &amp; Brotfabrik<\/a:t>/, `Fußzeile in ${m}`);
    assert.match(L[m].xml, /type="sldNum"/, `Seitenzahl in ${m}`);
    assert.doesNotMatch(L[m].rels, /relationships\/image/, `kein Bild in ${m}`);
  }
});

test("mit Logo: großes Logo auf TITEL, kleines auf den anderen Mastern", async () => {
  const logo = await sharp({ create: { width: 400, height: 178, channels: 4, background: "#4E46DC" } }).png().toBuffer();
  const zip = await read(await build(path.join(tmp, "mit.pptx"), logo));
  const L = await layouts(zip);
  for (const m of Object.values(did.MASTER)) assert.match(L[m].rels, /relationships\/image/, `Logo in ${m}`);
});

test("Quellenlink nimmt die Farbe der Quellenzeile", async () => {
  const zip = await read(path.join(tmp, "ohne.pptx"));
  const slide = await text(zip, "ppt/slides/slide3.xml");
  assert.match(slide, /<a:hlinkClick [^>]*>[\s\S]*?hlinkClr[^>]*val="tx"/);
  assert.match(slide, /<a:schemeClr val="accent6"\/>[\s\S]*?<a:t>\[1\] di\.day<\/a:t>/);
});
