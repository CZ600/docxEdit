"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser } = require("@xmldom/xmldom");
const { ParagraphTextModel } = require("../src/core/paragraph-text-model");

function parseParagraph(xml) {
  const documentNode = new DOMParser().parseFromString(
    `<w:root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${xml}</w:root>`,
    "application/xml",
  );

  return documentNode.getElementsByTagName("w:p")[0];
}

test("ParagraphTextModel reads a full paragraph across split runs", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>Hello </w:t></w:r>
      <w:r><w:t>virtual </w:t></w:r>
      <w:r><w:t>docx</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  assert.equal(model.getText(), "Hello virtual docx");
});

test("ParagraphTextModel writes replacements back to split text nodes", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>alpha </w:t></w:r>
      <w:r><w:t>beta </w:t></w:r>
      <w:r><w:t>gamma</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  const count = model.replace("beta gamma", "delta theta");

  assert.equal(count, 1);
  assert.equal(model.getText(), "alpha delta theta");

  const textNodes = Array.from(paragraph.getElementsByTagName("w:t")).map((node) => node.textContent);
  assert.deepEqual(textNodes, ["alpha ", "delta", " theta"]);
});

test("ParagraphTextModel preserves control nodes such as tabs", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>Item</w:t></w:r>
      <w:r><w:tab /></w:r>
      <w:r><w:t>Value</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  model.setText("Field\tResult");

  assert.equal(model.getText(), "Field\tResult");
  assert.equal(paragraph.getElementsByTagName("w:tab").length, 1);
});

test("ParagraphTextModel preserves math nodes while updating surrounding text", () => {
  const paragraph = parseParagraph(`
    <w:p>
      <w:r><w:t>Before </w:t></w:r>
      <m:oMath>
        <m:r><m:t>x+1</m:t></m:r>
      </m:oMath>
      <w:r><w:t> after</w:t></w:r>
    </w:p>
  `);

  const model = new ParagraphTextModel(paragraph);
  assert.equal(model.getText(), "Before [[MATH:x+1]] after");

  model.setText("Updated [[MATH:x+1]] done");

  assert.equal(model.getText(), "Updated [[MATH:x+1]] done");
  assert.equal(paragraph.getElementsByTagName("m:oMath").length, 1);
  assert.equal(paragraph.getElementsByTagName("m:t")[0].textContent, "x+1");
});
