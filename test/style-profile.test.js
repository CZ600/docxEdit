"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const fs = require("node:fs");
const { loadDocx } = require("../src");
const {
  STYLES_XML_PATH,
  parseStylesXml,
  resolveStyleChain,
  resolveEffectiveStyle,
  extractStyleProfile,
  applyStyleProfileToData,
  mergeStyleObjects,
} = require("../src/core/styles-xml-model");
const { parseXmlString } = require("../src/core/document-part");

const SAMPLE_DOCX = path.resolve(__dirname, "..", "测试文档.docx");

describe("styles-xml-model: parseStylesXml", () => {
  it("parses styles.xml from a real document", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    assert.ok(doc.stylesData);
    assert.ok(doc.stylesData.styles instanceof Map);
    assert.ok(doc.stylesData.docDefaults);
  });

  it("extracts docDefaults with paragraph and run styles", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const defaults = doc.stylesData.docDefaults;

    assert.ok(defaults.paragraphStyle);
    assert.ok(defaults.paragraphStyle.spacing);
    assert.equal(defaults.paragraphStyle.spacing.after, "160");

    assert.ok(defaults.runStyle);
    assert.equal(defaults.runStyle.fontSize, "22");
  });

  it("parses named paragraph styles", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const styles = doc.stylesData.styles;

    assert.ok(styles.has("a"), "Normal style exists");
    assert.equal(styles.get("a").name, "Normal");
    assert.equal(styles.get("a").type, "paragraph");

    assert.ok(styles.has("1"), "heading 1 exists");
    const h1 = styles.get("1");
    assert.equal(h1.name, "heading 1");
    assert.equal(h1.type, "paragraph");
    assert.equal(h1.basedOn, "a");
    assert.equal(h1.paragraphStyle.keepNext, true);
    assert.equal(h1.paragraphStyle.spacing.before, "480");
    assert.equal(h1.runStyle.fontSize, "48");
    assert.equal(h1.runStyle.color, "2F5496");
  });

  it("parses heading 2-5 with correct font sizes", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const styles = doc.stylesData.styles;

    assert.equal(styles.get("2").runStyle.fontSize, "40");
    assert.equal(styles.get("3").runStyle.fontSize, "32");
    assert.equal(styles.get("4").runStyle.fontSize, "28");
    assert.equal(styles.get("5").runStyle.fontSize, "24");
  });

  it("parses character styles", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const styles = doc.stylesData.styles;

    const intenseEmphasis = styles.get("aa");
    assert.ok(intenseEmphasis);
    assert.equal(intenseEmphasis.type, "character");
    assert.equal(intenseEmphasis.runStyle.italic, true);
  });

  it("captures theme font attributes in fontFamily", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const h1 = doc.stylesData.styles.get("1");

    assert.ok(h1.runStyle.fontFamily);
    assert.equal(h1.runStyle.fontFamily.asciiTheme, "majorHAnsi");
    assert.equal(h1.runStyle.fontFamily.eastAsiaTheme, "majorEastAsia");
  });

  it("handles documents without styles.xml gracefully", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    doc.stylesData = null;

    assert.deepEqual(doc.getStyleProfile(), {
      defaults: { paragraphStyle: {}, runStyle: {} },
      styles: {},
    });
    assert.deepEqual(doc.resolveEffectiveStyle("1"), {
      paragraphStyle: {},
      runStyle: {},
    });
    assert.deepEqual(doc.getNamedStyles(), []);
    doc.applyStyleProfile({ styles: {} });
  });
});

describe("styles-xml-model: resolveStyleChain", () => {
  it("resolves heading 1 chain (Normal → heading 1)", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const resolved = resolveStyleChain(doc.stylesData, "1");

    assert.ok(resolved.paragraphStyle);
    assert.ok(resolved.runStyle);
    assert.equal(resolved.runStyle.fontSize, "48");
    assert.equal(resolved.runStyle.color, "2F5496");
  });

  it("resolves effective style including docDefaults", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const effective = resolveEffectiveStyle(doc.stylesData, "1");

    assert.equal(effective.paragraphStyle.spacing.after, "80");
    assert.equal(effective.paragraphStyle.spacing.line, "278");
    assert.equal(effective.runStyle.fontSize, "48");
  });

  it("resolves Normal style as just docDefaults", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const effective = resolveEffectiveStyle(doc.stylesData, "a");

    assert.equal(effective.paragraphStyle.spacing.after, "160");
    assert.equal(effective.runStyle.fontSize, "22");
  });

  it("returns empty for unknown styleId", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const resolved = resolveStyleChain(doc.stylesData, "nonexistent");
    assert.deepEqual(resolved, { paragraphStyle: {}, runStyle: {} });
  });

  it("handles circular references without infinite loop", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const resolved = resolveStyleChain(doc.stylesData, "1", new Set(["1"]));
    assert.deepEqual(resolved, { paragraphStyle: {}, runStyle: {} });
  });
});

describe("styles-xml-model: extractStyleProfile", () => {
  it("produces a JSON-serializable object", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const profile = doc.getStyleProfile();

    const json = JSON.stringify(profile);
    const parsed = JSON.parse(json);

    assert.ok(parsed.defaults);
    assert.ok(parsed.styles);
    assert.equal(typeof parsed.styles["1"].name, "string");
    assert.equal(typeof parsed.styles["1"].runStyle.fontSize, "string");
  });

  it("contains all named styles", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const profile = doc.getStyleProfile();
    const styleIds = Object.keys(profile.styles);

    assert.ok(styleIds.includes("a"), "Normal");
    assert.ok(styleIds.includes("1"), "heading 1");
    assert.ok(styleIds.includes("a3"), "Title");
    assert.ok(styleIds.length >= 10);
  });

  it("each style has required metadata fields", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const profile = doc.getStyleProfile();

    for (const [styleId, style] of Object.entries(profile.styles)) {
      assert.ok(style.name, `style ${styleId} has name`);
      assert.ok(style.type, `style ${styleId} has type`);
      assert.ok(style.paragraphStyle !== undefined, `style ${styleId} has paragraphStyle`);
      assert.ok(style.runStyle !== undefined, `style ${styleId} has runStyle`);
    }
  });
});

describe("styles-xml-model: applyStyleProfile", () => {
  it("updates an existing style definition", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);

    doc.applyStyleProfile({
      styles: {
        "1": {
          name: "heading 1",
          type: "paragraph",
          basedOn: "a",
          paragraphStyle: { keepNext: true, spacing: { before: "600" } },
          runStyle: { fontSize: "56", color: "FF0000" },
        },
      },
    });

    const profile = doc.getStyleProfile();
    assert.equal(profile.styles["1"].runStyle.color, "FF0000");
    assert.equal(profile.styles["1"].runStyle.fontSize, "56");
    assert.equal(profile.styles["1"].paragraphStyle.spacing.before, "600");
  });

  it("creates a new style definition if styleId does not exist", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);

    doc.applyStyleProfile({
      styles: {
        customStyle1: {
          name: "My Custom Style",
          type: "paragraph",
          basedOn: "a",
          paragraphStyle: { alignment: "center" },
          runStyle: { bold: true, color: "00FF00" },
        },
      },
    });

    const profile = doc.getStyleProfile();
    assert.ok(profile.styles.customStyle1);
    assert.equal(profile.styles.customStyle1.name, "My Custom Style");
    assert.equal(profile.styles.customStyle1.runStyle.bold, true);
    assert.equal(profile.styles.customStyle1.runStyle.color, "00FF00");
  });

  it("updates docDefaults", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);

    doc.applyStyleProfile({
      defaults: {
        runStyle: { fontSize: "24" },
        paragraphStyle: { spacing: { after: "200", line: "300", lineRule: "auto" } },
      },
    });

    const profile = doc.getStyleProfile();
    assert.equal(profile.defaults.runStyle.fontSize, "24");
    assert.equal(profile.defaults.paragraphStyle.spacing.after, "200");
  });

  it("persists changes through save and reload", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);

    doc.applyStyleProfile({
      styles: {
        "2": {
          name: "heading 2",
          type: "paragraph",
          basedOn: "a",
          runStyle: { fontSize: "44", color: "0000FF" },
        },
      },
    });

    const outputPath = path.resolve(__dirname, "output", "apply-profile.docx");
    await fs.promises.mkdir(path.dirname(outputPath), { recursive: true });
    await doc.saveAs(outputPath);

    const doc2 = await loadDocx(outputPath);
    const profile2 = doc2.getStyleProfile();

    assert.equal(profile2.styles["2"].runStyle.color, "0000FF");
    assert.equal(profile2.styles["2"].runStyle.fontSize, "44");

    await fs.promises.unlink(outputPath);
  });

  it("full round-trip: extract from A, apply to B", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const originalProfile = doc.getStyleProfile();

    const modifiedProfile = JSON.parse(JSON.stringify(originalProfile));
    modifiedProfile.styles["3"].runStyle.fontSize = "36";
    modifiedProfile.styles["3"].runStyle.color = "CC0000";

    doc.applyStyleProfile(modifiedProfile);

    const resultProfile = doc.getStyleProfile();
    assert.equal(resultProfile.styles["3"].runStyle.fontSize, "36");
    assert.equal(resultProfile.styles["3"].runStyle.color, "CC0000");
  });
});

describe("styles-xml-model: mergeStyleObjects", () => {
  it("merges flat properties", () => {
    const result = mergeStyleObjects({ a: "1", b: "2" }, { b: "3", c: "4" });
    assert.deepEqual(result, { a: "1", b: "3", c: "4" });
  });

  it("deep-merges nested objects", () => {
    const result = mergeStyleObjects(
      { spacing: { before: "100", after: "200" } },
      { spacing: { before: "300" } },
    );
    assert.deepEqual(result, { spacing: { before: "300", after: "200" } });
  });

  it("removes keys with null values", () => {
    const result = mergeStyleObjects({ a: "1", b: "2" }, { b: null });
    assert.deepEqual(result, { a: "1" });
  });

  it("handles empty inputs", () => {
    assert.deepEqual(mergeStyleObjects({}, {}), {});
    assert.deepEqual(mergeStyleObjects(null, { a: "1" }), { a: "1" });
    assert.deepEqual(mergeStyleObjects({ a: "1" }, null), { a: "1" });
  });
});

describe("VirtualWordDocument style APIs", () => {
  it("getNamedStyles returns style metadata list", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const named = doc.getNamedStyles();

    assert.ok(Array.isArray(named));
    assert.ok(named.length > 0);

    const h1 = named.find((s) => s.styleId === "1");
    assert.ok(h1);
    assert.equal(h1.name, "heading 1");
    assert.equal(h1.type, "paragraph");
    assert.equal(h1.basedOn, "a");
  });

  it("resolveEffectiveStyle combines defaults + chain", async () => {
    const doc = await loadDocx(SAMPLE_DOCX);
    const effective = doc.resolveEffectiveStyle("1");

    assert.equal(effective.runStyle.fontSize, "48");
    assert.ok(effective.paragraphStyle.spacing.line, "docDefaults line spacing inherited");
  });
});
