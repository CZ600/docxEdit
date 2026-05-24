"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const fs = require("node:fs/promises");
const JSZip = require("jszip");
const { createVNode, loadDocx } = require("../src");

const TEST_DOC = path.resolve(__dirname, "test-docx", "角标和公式测试文档.docx");

// ============================================================
// 读取测试 (Reading Tests)
// ============================================================

test("reading: paragraph text includes footnote reference placeholders", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p0 = doc.getBody().getParagraph(0);
  const text = p0.getText();

  // 段落文本应包含脚注引用占位符
  assert.match(text, /\[\[FOOTNOTE_REF:\d+\]\]/);
  // 原始文本内容应保留
  assert.ok(text.includes("这是用来测试角标的一段文字"));
  assert.ok(text.includes("，更多的是"));
});

test("reading: footnote reference placeholders appear at correct positions", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p0 = doc.getBody().getParagraph(0);
  const text = p0.getText();

  // 脚注引用应出现在正确的位置：文字之后
  assert.ok(
    text.indexOf("这是用来测试角标的一段文字") < text.indexOf("[[FOOTNOTE_REF:"),
    "footnote reference should appear after the preceding text",
  );
  assert.ok(
    text.indexOf("[[FOOTNOTE_REF:") < text.indexOf("，更多的是"),
    "footnote reference should appear before the following text",
  );
});

test("reading: paragraph with two footnote references includes both", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p0 = doc.getBody().getParagraph(0);
  const text = p0.getText();

  // 应包含两个脚注引用占位符
  const matches = text.match(/\[\[FOOTNOTE_REF:\d+\]\]/g);
  assert.equal(matches && matches.length, 2);
});

test("reading: second paragraph with footnote references also works", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p2 = doc.getBody().getParagraph(2);
  const text = p2.getText();

  assert.match(text, /\[\[FOOTNOTE_REF:\d+\]\]/);
  assert.ok(text.includes("这是用来测试角标的一段文字"));
});

test("reading: math formulas are correctly parsed with correct count", async () => {
  const doc = await loadDocx(TEST_DOC);
  const maths = doc.getMaths();

  assert.equal(maths.length, 2);
});

test("reading: math formula text content is extracted", async () => {
  const doc = await loadDocx(TEST_DOC);
  const maths = doc.getMaths();

  assert.equal(maths[0].getText(), "x+yx=z");
  assert.equal(maths[1].getText(), "x+yx=z");
});

test("reading: math formula display mode is block", async () => {
  const doc = await loadDocx(TEST_DOC);
  const maths = doc.getMaths();

  assert.equal(maths[0].getDisplay(), "block");
  assert.equal(maths[1].getDisplay(), "block");
});

test("reading: paragraph text shows math placeholder for formula paragraphs", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p1 = doc.getBody().getParagraph(1);
  assert.equal(p1.getText(), "[[MATH:x+yx=z]]");
});

test("reading: footnote reference runs have correct style", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p0 = doc.getBody().getParagraph(0);
  const runs = p0.getRuns();

  // Run1 和 Run3 是脚注引用 run，应有 styleId
  assert.equal(runs[1].getStyle().styleId, "7");
  assert.equal(runs[3].getStyle().styleId, "7");
});

// ============================================================
// 写入测试 - Round-trip (保存后重新加载验证)
// ============================================================

test("round-trip: footnote references preserved after save and reload", async () => {
  const doc = await loadDocx(TEST_DOC);
  const buffer = await doc.toBuffer();

  // 直接读取原始 XML 内容验证
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");

  // 验证脚注引用元素存在于 XML 中
  assert.match(xml, /w:footnoteReference/);
  assert.match(xml, /w:footnoteReference\s[^>]*w:id="0"/);
  assert.match(xml, /w:footnoteReference\s[^>]*w:id="1"/);

  // 重新加载验证读取结果一致
  const reloaded = await loadDocx(buffer);
  const p0 = reloaded.getBody().getParagraph(0);
  const text = p0.getText();
  assert.match(text, /\[\[FOOTNOTE_REF:\d+\]\]/);
  assert.ok(text.includes("这是用来测试角标的一段文字"));
  assert.ok(text.includes("，更多的是"));
});

test("round-trip: math formulas preserved after save and reload", async () => {
  const doc = await loadDocx(TEST_DOC);
  const buffer = await doc.toBuffer();

  // 直接读取原始 XML 内容验证
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");

  // 验证数学公式元素存在
  assert.match(xml, /m:oMathPara/);
  assert.match(xml, /m:oMath/);
  // 验证上标结构 (m:sSup) 保留
  assert.match(xml, /m:sSup/);
  assert.match(xml, /m:t[^>]*>x\+</);
  assert.match(xml, /m:t[^>]*>=z/);

  // 重新加载验证
  const reloaded = await loadDocx(buffer);
  const maths = reloaded.getMaths();
  assert.equal(maths.length, 2);
  assert.equal(maths[0].getText(), "x+yx=z");
  assert.equal(maths[0].getDisplay(), "block");
});

test("round-trip: footnote reference count matches original", async () => {
  const doc = await loadDocx(TEST_DOC);
  const originalText0 = doc.getBody().getParagraph(0).getText();
  const originalText2 = doc.getBody().getParagraph(2).getText();

  const buffer = await doc.toBuffer();
  const reloaded = await loadDocx(buffer);

  const reloadedText0 = reloaded.getBody().getParagraph(0).getText();
  const reloadedText2 = reloaded.getBody().getParagraph(2).getText();

  // 脚注引用占位符数量应一致
  const originalMatches0 = (originalText0.match(/\[\[FOOTNOTE_REF:\d+\]\]/g) || []).length;
  const reloadedMatches0 = (reloadedText0.match(/\[\[FOOTNOTE_REF:\d+\]\]/g) || []).length;
  assert.equal(reloadedMatches0, originalMatches0);

  const originalMatches2 = (originalText2.match(/\[\[FOOTNOTE_REF:\d+\]\]/g) || []).length;
  const reloadedMatches2 = (reloadedText2.match(/\[\[FOOTNOTE_REF:\d+\]\]/g) || []).length;
  assert.equal(reloadedMatches2, originalMatches2);
});

test("round-trip: footnotes.xml preserved", async () => {
  const doc = await loadDocx(TEST_DOC);
  const buffer = await doc.toBuffer();

  const zip = await JSZip.loadAsync(buffer);
  const footnotesXml = await zip.file("word/footnotes.xml").async("string");

  // 验证脚注内容保留
  assert.match(footnotesXml, /w:footnote/);
  assert.match(footnotesXml, /w:footnote[^>]*w:id="0"/);
  assert.match(footnotesXml, /11111/);
});

// ============================================================
// 写入测试 - 新建文档 (Create New Document)
// ============================================================

test("write: new paragraph with superscript run style produces correct XML", async () => {
  const doc = await loadDocx(TEST_DOC);
  const tree = doc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");

  // 添加一个新段落，包含上角标文本
  body.children.push(createVNode({
    type: "paragraph",
    props: { text: "" },
    children: [
      createVNode({
        type: "run",
        props: { style: { vertAlign: "superscript" } },
        children: [
          createVNode({ type: "text", props: { text: "superscript text" }, children: [] }),
        ],
      }),
      createVNode({
        type: "run",
        props: {},
        children: [
          createVNode({ type: "text", props: { text: " normal text" }, children: [] }),
        ],
      }),
    ],
  }));

  doc.patch(tree);
  const buffer = await doc.toBuffer();

  // 直接读取原始 XML 验证
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");

  // 验证 vertAlign 元素正确生成
  assert.match(xml, /w:vertAlign\s+w:val="superscript"/);
  // 验证文本内容存在
  assert.match(xml, />superscript text</);
  assert.match(xml, /> normal text</);

  // 重新加载验证读取
  const reloaded = await loadDocx(buffer);
  const paragraphs = reloaded.getBody().getParagraphs();
  const newParagraph = paragraphs[paragraphs.length - 1];
  const runs = newParagraph.getRuns();
  assert.equal(runs[0].getStyle().vertAlign, "superscript");
  assert.equal(runs[1].getStyle().vertAlign, undefined);
});

test("write: new paragraph with subscript run style produces correct XML", async () => {
  const doc = await loadDocx(TEST_DOC);
  const tree = doc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");

  body.children.push(createVNode({
    type: "paragraph",
    props: { text: "" },
    children: [
      createVNode({
        type: "run",
        props: { style: { vertAlign: "subscript" } },
        children: [
          createVNode({ type: "text", props: { text: "subscript text" }, children: [] }),
        ],
      }),
    ],
  }));

  doc.patch(tree);
  const buffer = await doc.toBuffer();

  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");

  assert.match(xml, /w:vertAlign\s+w:val="subscript"/);

  const reloaded = await loadDocx(buffer);
  const paragraphs = reloaded.getBody().getParagraphs();
  const newParagraph = paragraphs[paragraphs.length - 1];
  assert.equal(newParagraph.getRuns()[0].getStyle().vertAlign, "subscript");
});

test("write: new paragraph with inline math formula produces correct XML", async () => {
  const doc = await loadDocx(TEST_DOC);
  const initialMathCount = doc.getMaths().length;
  const tree = doc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");

  body.children.push(createVNode({
    type: "paragraph",
    props: { text: "" },
    children: [
      createVNode({
        type: "run",
        props: {},
        children: [
          createVNode({ type: "text", props: { text: "Formula: " }, children: [] }),
        ],
      }),
      createVNode({
        type: "math",
        props: {
          text: "E=mc^2",
          style: {
            fontSize: "28",
            fontFamily: { ascii: "Cambria Math", hAnsi: "Cambria Math" },
          },
        },
        children: [],
      }),
    ],
  }));

  doc.patch(tree);
  const buffer = await doc.toBuffer();

  // 直接读取原始 XML 验证
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");

  // 验证内联数学公式存在（不是 m:oMathPara，而是直接的 m:oMath）
  assert.match(xml, /m:oMath[\s>]/);
  assert.match(xml, /E=mc\^2/);
  assert.match(xml, /Cambria Math/);

  // 重新加载验证
  const reloaded = await loadDocx(buffer);
  assert.equal(reloaded.getMaths().length, initialMathCount + 1);

  const maths = reloaded.getMaths();
  const newMath = maths[maths.length - 1];
  assert.equal(newMath.getText(), "E=mc^2");
  assert.equal(newMath.getDisplay(), "inline");
});

test("write: new paragraph with block math formula produces correct XML", async () => {
  const doc = await loadDocx(TEST_DOC);
  const initialMathCount = doc.getMaths().length;
  const tree = doc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");

  body.children.push(createVNode({
    type: "paragraph",
    props: { style: { alignment: "center" } },
    children: [
      createVNode({
        type: "math",
        props: {
          text: "a^2+b^2=c^2",
          display: "block",
          style: {
            justification: "center",
            bold: true,
            fontSize: "32",
          },
        },
        children: [],
      }),
    ],
  }));

  doc.patch(tree);
  const buffer = await doc.toBuffer();

  // 直接读取原始 XML 验证
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");

  // 验证块级数学公式
  assert.match(xml, /m:oMathPara/);
  assert.match(xml, /m:jc\s+m:val="center"/);
  assert.match(xml, /a\^2\+b\^2=c\^2/);
  assert.match(xml, /w:b\s+w:val="1"/);
  assert.match(xml, /w:sz\s+w:val="32"/);

  // 重新加载验证
  const reloaded = await loadDocx(buffer);
  assert.equal(reloaded.getMaths().length, initialMathCount + 1);

  const maths = reloaded.getMaths();
  const newMath = maths[maths.length - 1];
  assert.equal(newMath.getText(), "a^2+b^2=c^2");
  assert.equal(newMath.getDisplay(), "block");
});

// ============================================================
// 文本替换后的角标和公式保持测试 (Preservation After Text Replacement)
// ============================================================

test("text replacement: footnote references preserved when modifying surrounding text", async () => {
  const doc = await loadDocx(TEST_DOC);
  const p0 = doc.getBody().getParagraph(0);
  const originalText = p0.getText();

  // 替换部分文本（不涉及脚注引用部分）
  const count = p0.replaceAll("测试", "检验");
  assert.ok(count > 0);

  const buffer = await doc.toBuffer();

  // 验证脚注引用仍在 XML 中
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /w:footnoteReference/);

  // 重新加载验证
  const reloaded = await loadDocx(buffer);
  const reloadedP0 = reloaded.getBody().getParagraph(0);
  const text = reloadedP0.getText();

  // 脚注引用占位符仍然存在
  assert.match(text, /\[\[FOOTNOTE_REF:\d+\]\]/);
  // 文本替换生效
  assert.ok(text.includes("检验"));
  assert.ok(!text.includes("测试"));
});

test("text replacement: math formulas preserved when modifying paragraph text", async () => {
  const doc = await loadDocx(TEST_DOC);

  // 替换第一段文本（不影响数学公式段落）
  doc.replaceAll("测试", "验证");

  const buffer = await doc.toBuffer();
  const reloaded = await loadDocx(buffer);

  // 数学公式不受影响
  const maths = reloaded.getMaths();
  assert.equal(maths.length, 2);
  assert.equal(maths[0].getText(), "x+yx=z");
});

// ============================================================
// 新建脚注测试 (Create New Footnotes)
// ============================================================

test("addFootnote: creates a new footnote entry in existing document", async () => {
  const doc = await loadDocx(TEST_DOC);
  const initialCount = doc.getFootnotes().length;

  const id = doc.addFootnote("This is a brand new footnote");

  assert.equal(typeof id, "number");
  assert.ok(id >= 0);
  assert.equal(doc.getFootnotes().length, initialCount + 1);

  const newFootnote = doc.getFootnotes()[doc.getFootnotes().length - 1];
  assert.ok(newFootnote.getText().includes("This is a brand new footnote"));
});

test("addFootnote: returns sequential IDs for multiple calls", async () => {
  const doc = await loadDocx(TEST_DOC);

  const id1 = doc.addFootnote("First new footnote");
  const id2 = doc.addFootnote("Second new footnote");

  assert.ok(id2 > id1, "second footnote ID should be greater than first");
});

test("addFootnote: footnote can be referenced in a paragraph and saved to XML", async () => {
  const doc = await loadDocx(TEST_DOC);
  const id = doc.addFootnote("Dynamic footnote content");

  const tree = doc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");

  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: {},
          children: [createVNode({ type: "text", props: { text: "Paragraph with new footnote" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { styleId: "7" } },
          children: [createVNode({ type: "footnoteReference", props: { id: String(id) }, children: [] })],
        }),
      ],
    }),
  );

  doc.patch(tree);
  const buffer = await doc.toBuffer();

  // 直接读取原始 XML 验证
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file("word/document.xml").async("string");
  const fnXml = await zip.file("word/footnotes.xml").async("string");

  // document.xml 中应包含 footnoteReference
  assert.match(docXml, new RegExp(`w:footnoteReference[^>]*w:id="${id}"`));
  // footnotes.xml 中应包含新脚注内容
  assert.match(fnXml, /Dynamic footnote content/);
  assert.match(fnXml, new RegExp(`w:footnote[^>]*w:id="${id}"`));

  // 重新加载验证
  const reloaded = await loadDocx(buffer);
  const footnotes = reloaded.getFootnotes();
  assert.ok(footnotes.some((f) => f.getText().includes("Dynamic footnote content")));

  const paragraphs = reloaded.getBody().getParagraphs();
  const lastP = paragraphs[paragraphs.length - 1];
  assert.ok(lastP.getText().includes("Paragraph with new footnote"));
  assert.match(lastP.getText(), /\[\[FOOTNOTE_REF:\d+\]\]/);
});

test("addFootnote: works on a document without existing footnotes", async () => {
  // 创建一个没有脚注的最小文档
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
  );
  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body></w:body></w:document>',
  );

  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  const doc = await loadDocx(buffer);

  // 文档初始没有脚注
  assert.equal(doc.getFootnotes().length, 0);

  const id = doc.addFootnote("First footnote in a brand new document");
  assert.equal(id, 0);
  assert.equal(doc.getFootnotes().length, 1);

  // 创建段落引用这个脚注
  const tree = doc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: {},
          children: [createVNode({ type: "text", props: { text: "Text with new footnote" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: {},
          children: [createVNode({ type: "footnoteReference", props: { id: "0" }, children: [] })],
        }),
      ],
    }),
  );

  doc.patch(tree);
  const output = await doc.toBuffer();

  // 验证 XML
  const outZip = await JSZip.loadAsync(output);
  assert.ok(outZip.file("word/footnotes.xml"), "footnotes.xml should be created");

  const outDocXml = await outZip.file("word/document.xml").async("string");
  assert.match(outDocXml, /w:footnoteReference/);

  const outFnXml = await outZip.file("word/footnotes.xml").async("string");
  assert.match(outFnXml, /First footnote in a brand new document/);

  // 验证关系文件
  const outRels = await outZip.file("word/_rels/document.xml.rels").async("string");
  assert.match(outRels, /relationships\/footnotes/);

  // 重新加载完整验证
  const reloaded = await loadDocx(output);
  assert.equal(reloaded.getFootnotes().length, 1);
  assert.ok(reloaded.getFootnotes()[0].getText().includes("First footnote in a brand new document"));
});
