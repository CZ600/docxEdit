"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");
const { createVNode, loadDocx } = require("../src");
const { findSampleDocx } = require("./helpers/sample-file");

const SAMPLE_FILE = findSampleDocx();

function getPart(tree, type) {
  return tree.children.find((child) => child.type === type);
}

test("supports inserting updating and styling math formulas", async () => {
  const doc = await loadDocx(SAMPLE_FILE);
  const initialMathCount = doc.getMaths().length;
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");

  body.children.push(
    createVNode({
      type: "paragraph",
      props: { style: { alignment: "center" } },
      children: [
        createVNode({
          type: "math",
          props: {
            text: "x^2+y^2=z^2",
            display: "block",
            style: {
              justification: "center",
              bold: true,
              color: "C00000",
              fontSize: "32",
              fontFamily: {
                ascii: "Cambria Math",
                hAnsi: "Cambria Math",
              },
            },
          },
          children: [],
        }),
      ],
    }),
  );

  doc.patch(tree);

  assert.equal(doc.getMaths().length, initialMathCount + 1);
  assert.equal(doc.getBody().getParagraphs().at(-1).getText(), "[[MATH:x^2+y^2=z^2]]");

  const math = doc.getMaths().at(-1);
  math.setText("E=mc^2");
  math.patchStyle({ italic: true });

  const buffer = await doc.toBuffer();
  const reloaded = await loadDocx(buffer);
  const reloadedMath = reloaded.getMaths().at(-1);

  assert.equal(reloadedMath.getText(), "E=mc^2");
  assert.equal(reloadedMath.getDisplay(), "block");
  assert.deepEqual(reloadedMath.getStyle(), {
    justification: "center",
    bold: true,
    italic: true,
    color: "C00000",
    fontSize: "32",
    fontFamily: {
      ascii: "Cambria Math",
      hAnsi: "Cambria Math",
    },
  });

  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.match(documentXml, /<m:oMathPara[\s\S]*<m:jc m:val="center"/);
  assert.match(documentXml, /<w:color w:val="C00000"/);
  assert.match(documentXml, /<w:sz w:val="32"/);
  assert.match(documentXml, /E=mc\^2/);
});
