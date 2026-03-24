"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");
const { loadDocx } = require("../src");

async function buildStyledDocx() {
  const zip = new JSZip();

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="BodyText" />
        <w:jc w:val="center" />
        <w:spacing w:before="120" w:after="240" />
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b />
          <w:color w:val="FF0000" />
          <w:sz w:val="28" />
        </w:rPr>
        <w:t>Styled Alpha</w:t>
      </w:r>
      <w:r><w:t> Tail</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Plain Beta</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: "nodebuffer" });
}

function getBody(tree) {
  return tree.children.find((node) => node.type === "body");
}

test("style model parses existing paragraph and run styles and persists tree updates", async () => {
  const doc = await loadDocx(await buildStyledDocx());
  const body = doc.getBody();
  const firstParagraph = body.getParagraph(0);
  const firstRun = firstParagraph.getRun(0);

  assert.deepEqual(firstParagraph.getStyle(), {
    styleId: "BodyText",
    alignment: "center",
    spacing: {
      before: "120",
      after: "240",
    },
  });
  assert.deepEqual(firstRun.getStyle(), {
    bold: true,
    color: "FF0000",
    fontSize: "28",
  });

  const tree = doc.toComponentTree();
  const treeBody = getBody(tree);
  treeBody.children[0].props.style = {
    styleId: "BodyText2",
    alignment: "right",
    spacing: { before: "200", after: "300", line: "360", lineRule: "auto" },
    indent: { left: "240", firstLine: "240" },
  };
  treeBody.children[0].children[0].props.style = {
    bold: false,
    italic: true,
    underline: "single",
    color: "00AA00",
    fontSize: "24",
  };

  const result = doc.patch(tree);
  assert.ok(result.operations.some((item) => item.nodeType === "paragraph"));
  assert.ok(result.operations.some((item) => item.nodeType === "run"));
  assert.deepEqual(doc.getBody().getParagraph(0).getStyle(), {
    styleId: "BodyText2",
    alignment: "right",
    spacing: { before: "200", after: "300", line: "360", lineRule: "auto" },
    indent: { left: "240", firstLine: "240" },
  });
  assert.deepEqual(doc.getBody().getParagraph(0).getRun(0).getStyle(), {
    bold: false,
    italic: true,
    underline: "single",
    color: "00AA00",
    fontSize: "24",
  });

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.deepEqual(reloaded.getBody().getParagraph(0).getStyle(), {
    styleId: "BodyText2",
    alignment: "right",
    spacing: { before: "200", after: "300", line: "360", lineRule: "auto" },
    indent: { left: "240", firstLine: "240" },
  });
  assert.deepEqual(reloaded.getBody().getParagraph(0).getRun(0).getStyle(), {
    bold: false,
    italic: true,
    underline: "single",
    color: "00AA00",
    fontSize: "24",
  });
});

test("controller style APIs support add update clear and migration across components", async () => {
  const doc = await loadDocx(await buildStyledDocx());
  const paragraphA = doc.getBody().getParagraph(0);
  const paragraphB = doc.getBody().getParagraph(1);
  const runA = paragraphA.getRun(0);
  const runB = paragraphB.getRun(0);

  paragraphB.patchStyle({
    alignment: "both",
    spacing: { after: "180" },
  });
  runB.patchStyle({
    italic: true,
    color: "3333FF",
  });

  assert.deepEqual(paragraphB.getStyle(), {
    alignment: "both",
    spacing: { after: "180" },
  });
  assert.deepEqual(runB.getStyle(), {
    italic: true,
    color: "3333FF",
  });

  paragraphB.copyStyleFrom(paragraphA);
  runB.copyStyleFrom(runA);

  assert.deepEqual(paragraphB.getStyle(), paragraphA.getStyle());
  assert.deepEqual(runB.getStyle(), runA.getStyle());

  runB.setStyle({});
  assert.deepEqual(runB.getStyle(), {});

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.deepEqual(reloaded.getBody().getParagraph(1).getStyle(), reloaded.getBody().getParagraph(0).getStyle());
  assert.deepEqual(reloaded.getBody().getParagraph(1).getRun(0).getStyle(), {});
});
