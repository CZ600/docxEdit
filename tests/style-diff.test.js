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

async function buildStyledTableDocx() {
  const zip = new JSZip();

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid" />
        <w:tblW w:w="5000" w:type="dxa" />
        <w:tblLayout w:type="fixed" />
        <w:jc w:val="center" />
        <w:tblBorders>
          <w:top w:val="single" w:sz="12" w:color="000000" />
          <w:left w:val="single" w:sz="12" w:color="000000" />
          <w:bottom w:val="single" w:sz="12" w:color="000000" />
          <w:right w:val="single" w:sz="12" w:color="000000" />
          <w:insideH w:val="single" w:sz="8" w:color="666666" />
          <w:insideV w:val="single" w:sz="8" w:color="666666" />
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:trPr>
          <w:trHeight w:val="480" w:hRule="atLeast" />
          <w:tblHeader />
        </w:trPr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="2500" w:type="dxa" />
            <w:shd w:val="clear" w:fill="DDDDDD" w:color="auto" />
            <w:vAlign w:val="center" />
            <w:tcBorders>
              <w:top w:val="single" w:sz="8" w:color="333333" />
              <w:left w:val="single" w:sz="8" w:color="333333" />
              <w:bottom w:val="single" w:sz="8" w:color="333333" />
              <w:right w:val="single" w:sz="8" w:color="333333" />
            </w:tcBorders>
            <w:tcMar>
              <w:top w:w="80" w:type="dxa" />
              <w:left w:w="100" w:type="dxa" />
            </w:tcMar>
          </w:tcPr>
          <w:p><w:r><w:t>A1</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B1</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A2</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
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

test("table style model parses existing table row and cell styles and persists tree updates", async () => {
  const doc = await loadDocx(await buildStyledTableDocx());
  const table = doc.getTables()[0];
  const row = table.getRow(0);
  const cell = table.getCell(0, 0);

  assert.deepEqual(table.getStyle(), {
    styleId: "TableGrid",
    width: { w: "5000", type: "dxa" },
    layout: "fixed",
    alignment: "center",
    borders: {
      top: { val: "single", sz: "12", color: "000000" },
      left: { val: "single", sz: "12", color: "000000" },
      bottom: { val: "single", sz: "12", color: "000000" },
      right: { val: "single", sz: "12", color: "000000" },
      insideH: { val: "single", sz: "8", color: "666666" },
      insideV: { val: "single", sz: "8", color: "666666" },
    },
  });
  assert.deepEqual(row.getStyle(), {
    height: { val: "480", rule: "atLeast" },
    header: true,
  });
  assert.deepEqual(cell.getStyle(), {
    width: { w: "2500", type: "dxa" },
    shading: { val: "clear", fill: "DDDDDD", color: "auto" },
    verticalAlign: "center",
    borders: {
      top: { val: "single", sz: "8", color: "333333" },
      left: { val: "single", sz: "8", color: "333333" },
      bottom: { val: "single", sz: "8", color: "333333" },
      right: { val: "single", sz: "8", color: "333333" },
    },
    margins: {
      top: { w: "80", type: "dxa" },
      left: { w: "100", type: "dxa" },
    },
  });

  const tree = doc.toComponentTree();
  const tableNode = getBody(tree).children[0];
  tableNode.props.style = {
    styleId: "CustomTable",
    width: { w: "7000", type: "dxa" },
    layout: "autofit",
    alignment: "right",
    borders: {
      top: { val: "double", sz: "16", color: "FF0000" },
      insideH: { val: "dashed", sz: "6", color: "00AA00" },
    },
  };
  tableNode.children[0].props.style = {
    height: { val: "640", rule: "exact" },
    header: false,
    cantSplit: true,
  };
  tableNode.children[0].children[0].props.style = {
    width: { w: "3200", type: "dxa" },
    shading: { val: "clear", fill: "FFF2CC", color: "auto" },
    verticalAlign: "bottom",
    gridSpan: "2",
    vMerge: "restart",
    borders: {
      top: { val: "double", sz: "12", color: "AA5500" },
      left: { val: "single", sz: "8", color: "AA5500" },
    },
    margins: {
      top: { w: "120", type: "dxa" },
      right: { w: "140", type: "dxa" },
    },
  };

  const result = doc.patch(tree);
  assert.ok(result.operations.some((item) => item.nodeType === "table"));
  assert.ok(result.operations.some((item) => item.nodeType === "table-row"));
  assert.ok(result.operations.some((item) => item.nodeType === "table-cell"));

  assert.deepEqual(doc.getTables()[0].getStyle(), {
    styleId: "CustomTable",
    width: { w: "7000", type: "dxa" },
    layout: "autofit",
    alignment: "right",
    borders: {
      top: { val: "double", sz: "16", color: "FF0000" },
      insideH: { val: "dashed", sz: "6", color: "00AA00" },
    },
  });
  assert.deepEqual(doc.getTables()[0].getRow(0).getStyle(), {
    height: { val: "640", rule: "exact" },
    header: false,
    cantSplit: true,
  });
  assert.deepEqual(doc.getTables()[0].getCell(0, 0).getStyle(), {
    width: { w: "3200", type: "dxa" },
    shading: { val: "clear", fill: "FFF2CC", color: "auto" },
    verticalAlign: "bottom",
    gridSpan: "2",
    vMerge: "restart",
    borders: {
      top: { val: "double", sz: "12", color: "AA5500" },
      left: { val: "single", sz: "8", color: "AA5500" },
    },
    margins: {
      top: { w: "120", type: "dxa" },
      right: { w: "140", type: "dxa" },
    },
  });

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.deepEqual(reloaded.getTables()[0].getStyle(), doc.getTables()[0].getStyle());
  assert.deepEqual(reloaded.getTables()[0].getRow(0).getStyle(), doc.getTables()[0].getRow(0).getStyle());
  assert.deepEqual(reloaded.getTables()[0].getCell(0, 0).getStyle(), doc.getTables()[0].getCell(0, 0).getStyle());
});

test("table controller style APIs support add update clear and migration", async () => {
  const doc = await loadDocx(await buildStyledTableDocx());
  const tableA = doc.getTables()[0];
  const rowA = tableA.getRow(0);
  const rowB = tableA.getRow(1);
  const cellA = tableA.getCell(0, 0);
  const cellB = tableA.getCell(1, 1);

  rowB.patchStyle({
    height: { val: "500", rule: "atLeast" },
    cantSplit: true,
  });
  cellB.patchStyle({
    shading: { fill: "DDEBF7" },
    verticalAlign: "center",
  });

  assert.deepEqual(rowB.getStyle(), {
    height: { val: "500", rule: "atLeast" },
    cantSplit: true,
  });
  assert.deepEqual(cellB.getStyle(), {
    shading: { fill: "DDEBF7" },
    verticalAlign: "center",
  });

  tableA.setStyle({
    width: { w: "7200", type: "dxa" },
    alignment: "center",
  });
  rowB.copyStyleFrom(rowA);
  cellB.copyStyleFrom(cellA);

  assert.deepEqual(rowB.getStyle(), rowA.getStyle());
  assert.deepEqual(cellB.getStyle(), cellA.getStyle());

  cellB.setStyle({});
  assert.deepEqual(cellB.getStyle(), {});

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.deepEqual(reloaded.getTables()[0].getStyle(), {
    width: { w: "7200", type: "dxa" },
    alignment: "center",
  });
  assert.deepEqual(reloaded.getTables()[0].getRow(1).getStyle(), reloaded.getTables()[0].getRow(0).getStyle());
  assert.deepEqual(reloaded.getTables()[0].getCell(1, 1).getStyle(), {});
});
