"use strict";

const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const {
  parseParagraphStyle,
  parseRunStyle,
  applyParagraphStyle,
  applyRunStyle,
  cloneStyle,
} = require("../src/core/style-model");

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function wrapInRun(rPrXml) {
  return `<w:r xmlns:w="${WORD_NS}">${rPrXml}<w:t>text</w:t></w:r>`;
}

function wrapInParagraph(pPrXml) {
  return `<w:p xmlns:w="${WORD_NS}">${pPrXml}<w:r><w:t>text</w:t></w:r></w:p>`;
}

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, "application/xml").documentElement;
}

function serialize(element) {
  return new XMLSerializer().serializeToString(element);
}

describe("run style: vertAlign (superscript/subscript)", () => {
  it("parses superscript", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`));
    const style = parseRunStyle(el);
    assert.equal(style.vertAlign, "superscript");
  });

  it("parses subscript", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>`));
    const style = parseRunStyle(el);
    assert.equal(style.vertAlign, "subscript");
  });

  it("applies superscript and round-trips", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:b/></w:rPr>`));
    applyRunStyle(el, { vertAlign: "superscript" });
    const parsed = parseRunStyle(el);
    assert.equal(parsed.vertAlign, "superscript");
  });

  it("removes vertAlign when set to null", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`));
    applyRunStyle(el, { vertAlign: null });
    const parsed = parseRunStyle(el);
    assert.equal(parsed.vertAlign, undefined);
  });
});

describe("run style: strike / doubleStrike", () => {
  it("parses strikethrough", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:strike/></w:rPr>`));
    const style = parseRunStyle(el);
    assert.equal(style.strike, true);
  });

  it("parses double strikethrough", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:dstrike/></w:rPr>`));
    const style = parseRunStyle(el);
    assert.equal(style.doubleStrike, true);
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInRun(`<w:rPr></w:rPr>`));
    applyRunStyle(el, { strike: true, doubleStrike: true });
    const parsed = parseRunStyle(el);
    assert.equal(parsed.strike, true);
    assert.equal(parsed.doubleStrike, true);
  });

  it("removes when false", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:strike/><w:dstrike/></w:rPr>`));
    applyRunStyle(el, { strike: false, doubleStrike: false });
    const parsed = parseRunStyle(el);
    assert.equal(parsed.strike, false);
    assert.equal(parsed.doubleStrike, false);
  });
});

describe("run style: smallCaps / caps", () => {
  it("parses smallCaps", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:smallCaps/></w:rPr>`));
    assert.equal(parseRunStyle(el).smallCaps, true);
  });

  it("parses caps", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:caps/></w:rPr>`));
    assert.equal(parseRunStyle(el).caps, true);
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInRun(`<w:rPr></w:rPr>`));
    applyRunStyle(el, { smallCaps: true, caps: true });
    const parsed = parseRunStyle(el);
    assert.equal(parsed.smallCaps, true);
    assert.equal(parsed.caps, true);
  });
});

describe("run style: hidden (vanish)", () => {
  it("parses hidden text", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:vanish/></w:rPr>`));
    assert.equal(parseRunStyle(el).hidden, true);
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInRun(`<w:rPr></w:rPr>`));
    applyRunStyle(el, { hidden: true });
    assert.equal(parseRunStyle(el).hidden, true);
  });
});

describe("run style: position (character offset)", () => {
  it("parses position value", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:position w:val="2"/></w:rPr>`));
    assert.equal(parseRunStyle(el).position, "2");
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInRun(`<w:rPr></w:rPr>`));
    applyRunStyle(el, { position: "-4" });
    assert.equal(parseRunStyle(el).position, "-4");
  });
});

describe("run style: characterScale (character width)", () => {
  it("parses w:w value", () => {
    const el = parseXml(wrapInRun(`<w:rPr><w:w w:val="150"/></w:rPr>`));
    assert.equal(parseRunStyle(el).characterScale, "150");
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInRun(`<w:rPr></w:rPr>`));
    applyRunStyle(el, { characterScale: "200" });
    assert.equal(parseRunStyle(el).characterScale, "200");
  });
});

describe("run style: combined footnote reference style", () => {
  it("parses a typical footnote reference marker", () => {
    const el = parseXml(wrapInRun(
      `<w:rPr><w:vertAlign w:val="superscript"/><w:sz w:val="18"/></w:rPr>`,
    ));
    const style = parseRunStyle(el);
    assert.equal(style.vertAlign, "superscript");
    assert.equal(style.fontSize, "18");
  });
});

describe("paragraph style: widowControl", () => {
  it("parses widowControl enabled", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr><w:widowControl/></w:pPr>`));
    assert.equal(parseParagraphStyle(el).widowControl, true);
  });

  it("parses widowControl disabled", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr><w:widowControl w:val="0"/></w:pPr>`));
    assert.equal(parseParagraphStyle(el).widowControl, false);
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr></w:pPr>`));
    applyParagraphStyle(el, { widowControl: false });
    assert.equal(parseParagraphStyle(el).widowControl, false);
  });
});

describe("paragraph style: outlineLevel", () => {
  it("parses outline level 0", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`));
    assert.equal(parseParagraphStyle(el).outlineLevel, "0");
  });

  it("parses outline level 2", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr><w:outlineLvl w:val="2"/></w:pPr>`));
    assert.equal(parseParagraphStyle(el).outlineLevel, "2");
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr></w:pPr>`));
    applyParagraphStyle(el, { outlineLevel: "1" });
    assert.equal(parseParagraphStyle(el).outlineLevel, "1");
  });
});

describe("paragraph style: borders (pBdr)", () => {
  it("parses paragraph borders", () => {
    const el = parseXml(wrapInParagraph(
      `<w:pPr><w:pBdr>` +
      `<w:top w:val="single" w:sz="4" w:space="10" w:color="FF0000"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="10" w:color="0000FF"/>` +
      `</w:pBdr></w:pPr>`,
    ));
    const style = parseParagraphStyle(el);
    assert.ok(style.borders);
    assert.equal(style.borders.top.val, "single");
    assert.equal(style.borders.top.color, "FF0000");
    assert.equal(style.borders.bottom.color, "0000FF");
  });

  it("applies and round-trips", () => {
    const el = parseXml(wrapInParagraph(`<w:pPr></w:pPr>`));
    applyParagraphStyle(el, {
      borders: {
        top: { val: "double", sz: "6", space: "12", color: "00FF00" },
        bottom: { val: "single", sz: "2", space: "4", color: "CCCCCC" },
      },
    });
    const parsed = parseParagraphStyle(el);
    assert.equal(parsed.borders.top.val, "double");
    assert.equal(parsed.borders.top.color, "00FF00");
    assert.equal(parsed.borders.bottom.val, "single");
  });

  it("removes borders when cleared", () => {
    const el = parseXml(wrapInParagraph(
      `<w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="FF0000"/></w:pBdr></w:pPr>`,
    ));
    applyParagraphStyle(el, { borders: {} });
    const parsed = parseParagraphStyle(el);
    assert.equal(parsed.borders, undefined);
  });
});

describe("paragraph style: combined heading style", () => {
  it("parses a heading with outlineLevel + keepNext + spacing", () => {
    const el = parseXml(wrapInParagraph(
      `<w:pPr>` +
      `<w:keepNext/>` +
      `<w:keepLines/>` +
      `<w:spacing w:before="480" w:after="80"/>` +
      `<w:outlineLvl w:val="0"/>` +
      `</w:pPr>`,
    ));
    const style = parseParagraphStyle(el);
    assert.equal(style.keepNext, true);
    assert.equal(style.keepLines, true);
    assert.equal(style.spacing.before, "480");
    assert.equal(style.outlineLevel, "0");
  });
});

describe("style round-trip: full run style with all new properties", () => {
  it("parse → apply → parse produces equivalent style", () => {
    const originalStyle = {
      bold: true,
      italic: true,
      strike: true,
      doubleStrike: false,
      smallCaps: true,
      caps: false,
      hidden: false,
      vertAlign: "superscript",
      position: "2",
      characterScale: "120",
      fontSize: "24",
      color: "FF0000",
      underline: "single",
      fontFamily: { ascii: "Arial", eastAsia: "黑体" },
    };

    const el = parseXml(wrapInRun(`<w:rPr></w:rPr>`));
    applyRunStyle(el, originalStyle);
    const result = parseRunStyle(el);

    assert.equal(result.bold, true);
    assert.equal(result.italic, true);
    assert.equal(result.strike, true);
    assert.equal(result.smallCaps, true);
    assert.equal(result.vertAlign, "superscript");
    assert.equal(result.position, "2");
    assert.equal(result.characterScale, "120");
    assert.equal(result.fontSize, "24");
    assert.equal(result.color, "FF0000");
    assert.equal(result.fontFamily.ascii, "Arial");
    assert.equal(result.fontFamily.eastAsia, "黑体");
  });
});

describe("style round-trip: full paragraph style with all new properties", () => {
  it("parse → apply → parse produces equivalent style", () => {
    const originalStyle = {
      alignment: "center",
      keepNext: true,
      keepLines: true,
      widowControl: false,
      outlineLevel: "1",
      spacing: { before: "240", after: "120", line: "360", lineRule: "auto" },
      indent: { firstLine: "480" },
      borders: {
        bottom: { val: "single", sz: "4", space: "10", color: "000000" },
      },
    };

    const el = parseXml(wrapInParagraph(`<w:pPr></w:pPr>`));
    applyParagraphStyle(el, originalStyle);
    const result = parseParagraphStyle(el);

    assert.equal(result.alignment, "center");
    assert.equal(result.keepNext, true);
    assert.equal(result.widowControl, false);
    assert.equal(result.outlineLevel, "1");
    assert.equal(result.spacing.before, "240");
    assert.equal(result.indent.firstLine, "480");
    assert.equal(result.borders.bottom.val, "single");
    assert.equal(result.borders.bottom.color, "000000");
  });
});
