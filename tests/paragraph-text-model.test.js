"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser } = require("@xmldom/xmldom");
const { ParagraphTextModel } = require("../src/core/paragraph-text-model");

function parseParagraph(xml) {
  const documentNode = new DOMParser().parseFromString(
    `<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</w:root>`,
    "application/xml",
  );

  return documentNode.getElementsByTagName("w:p")[0];
}

test("ParagraphTextModel reads a full paragraph across split runs", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>地理科学学院</w:t></w:r>
      <w:r><w:t>25</w:t></w:r>
      <w:r><w:t>级硕士地信团支</w:t></w:r>
      <w:r><w:t>部</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  assert.equal(model.getText(), "地理科学学院25级硕士地信团支部");
});

test("ParagraphTextModel writes replacements back to split text nodes", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>躬行实践</w:t></w:r>
      <w:r><w:t>察民生，青春</w:t></w:r>
      <w:r><w:t>共议家国事</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  const count = model.replace("实践察民生，青春共议", "调研民情，青年同议");

  assert.equal(count, 1);
  assert.equal(model.getText(), "躬行调研民情，青年同议家国事");

  const textNodes = Array.from(paragraph.getElementsByTagName("w:t")).map((node) => node.textContent);
  assert.deepEqual(textNodes, ["躬行调研", "民情，青年同", "议家国事"]);
});

test("ParagraphTextModel preserves control nodes such as tabs", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>第一项</w:t></w:r>
      <w:r><w:tab /></w:r>
      <w:r><w:t>说明</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  model.setText("更新项\t详细说明");

  assert.equal(model.getText(), "更新项\t详细说明");
  assert.equal(paragraph.getElementsByTagName("w:tab").length, 1);
});
