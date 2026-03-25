"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const { createVNode, loadDocx } = require("../src");
const { findSampleDocx } = require("./helpers/sample-file");

const SAMPLE_FILE = findSampleDocx();
const OUTPUT_DIR = path.resolve(__dirname, "..", "examples", "output");
const OUTPUT_FILE = path.join(OUTPUT_DIR, "sample-math-edited.docx");

function getPart(tree, type) {
  return tree.children.find((child) => child.type === type);
}

test("writes a docx output sample with math formulas for manual inspection", async () => {
  const doc = await loadDocx(SAMPLE_FILE);
  const initialMathCount = doc.getMaths().length;
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");

  body.children.push(
    createVNode({
      type: "paragraph",
      props: { style: { alignment: "center", spacing: { before: "240", after: "240" } } },
      children: [
        createVNode({
          type: "math",
          props: {
            text: "a^2+b^2=c^2",
            display: "block",
            style: {
              justification: "center",
              bold: true,
              color: "1F4E79",
              fontSize: "36",
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

  body.children.push(
    createVNode({
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
            text: "f(x)=x^3+2x+1",
            style: {
              color: "C00000",
              fontSize: "28",
            },
          },
          children: [],
        }),
      ],
    }),
  );

  doc.patch(tree);

  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  await doc.saveAs(OUTPUT_FILE);

  const stat = await fs.stat(OUTPUT_FILE);
  const reloaded = await loadDocx(OUTPUT_FILE);
  const mathTexts = reloaded.getMaths().map((math) => math.getText());

  assert.ok(stat.size > 0);
  assert.equal(reloaded.getMaths().length, initialMathCount + 2);
  assert.ok(mathTexts.includes("a^2+b^2=c^2"));
  assert.ok(mathTexts.includes("f(x)=x^3+2x+1"));
});
