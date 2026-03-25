"use strict";

const fs = require("node:fs/promises");
const path = require("node:path");
const { createVNode, loadDocx } = require("../src");
const { findSampleDocx } = require("../tests/helpers/sample-file");

function getPart(tree, type) {
  return tree.children.find((child) => child.type === type);
}

function createTable(rows, columns) {
  return createVNode({
    type: "table",
    props: {},
    children: Array.from({ length: rows }, () =>
      createVNode({
        type: "table-row",
        props: {},
        children: Array.from({ length: columns }, () =>
          createVNode({
            type: "table-cell",
            props: {},
            children: [createVNode({ type: "paragraph", props: { text: "" }, children: [] })],
          }),
        ),
      }),
    ),
  });
}

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const outputDir = path.join(projectRoot, "examples", "output");
  const inputPath = findSampleDocx();
  const outputPath = path.join(outputDir, "sample-table-style-edited.docx");

  await fs.mkdir(outputDir, { recursive: true });

  const doc = await loadDocx(inputPath);
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");

  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "表格格式编辑检查：" },
      children: [],
    }),
  );
  body.children.push(createTable(4, 3));

  doc.patch(tree);

  const table = doc.getTables()[0];
  table.fill(
    [
      ["检查项", "设置值", "说明"],
      ["表格对齐", "居中", "table.setStyle"],
      ["行高", "640", "row.setStyle"],
      ["单元格底纹", "FFF2CC", "cell.setStyle"],
    ],
    { startRow: 0 },
  );

  table.setStyle({
    styleId: "TableGrid",
    width: { w: "8400", type: "dxa" },
    layout: "fixed",
    alignment: "center",
    borders: {
      top: { val: "double", sz: "16", color: "C00000" },
      left: { val: "single", sz: "8", color: "C00000" },
      bottom: { val: "double", sz: "16", color: "C00000" },
      right: { val: "single", sz: "8", color: "C00000" },
      insideH: { val: "single", sz: "6", color: "7F7F7F" },
      insideV: { val: "single", sz: "6", color: "7F7F7F" },
    },
  });

  const headerRow = table.getRow(0);
  headerRow.setStyle({
    height: { val: "640", rule: "exact" },
    header: true,
    cantSplit: true,
  });

  for (let column = 0; column < 3; column += 1) {
    const cell = table.getCell(0, column);
    cell.setStyle({
      width: { w: "2800", type: "dxa" },
      shading: { val: "clear", fill: "D9EAF7", color: "auto" },
      verticalAlign: "center",
      borders: {
        top: { val: "single", sz: "10", color: "4F81BD" },
        left: { val: "single", sz: "10", color: "4F81BD" },
        bottom: { val: "single", sz: "10", color: "4F81BD" },
        right: { val: "single", sz: "10", color: "4F81BD" },
      },
      margins: {
        top: { w: "100", type: "dxa" },
        left: { w: "120", type: "dxa" },
        right: { w: "120", type: "dxa" },
        bottom: { w: "100", type: "dxa" },
      },
    });
    cell.getParagraph(0).setStyle({ alignment: "center" });
    cell.getParagraph(0).getRun(0).setStyle({ bold: true, color: "1F1F1F" });
  }

  const bodyRow = table.getRow(2);
  bodyRow.patchStyle({
    height: { val: "560", rule: "atLeast" },
  });

  const highlightCell = table.getCell(2, 1);
  highlightCell.setStyle({
    width: { w: "2800", type: "dxa" },
    shading: { val: "clear", fill: "FFF2CC", color: "auto" },
    verticalAlign: "bottom",
    borders: {
      top: { val: "double", sz: "12", color: "BF9000" },
      left: { val: "single", sz: "8", color: "BF9000" },
      bottom: { val: "double", sz: "12", color: "BF9000" },
      right: { val: "single", sz: "8", color: "BF9000" },
    },
    margins: {
      top: { w: "80", type: "dxa" },
      left: { w: "140", type: "dxa" },
    },
  });

  await doc.saveAs(outputPath);

  console.log("Generated table style inspection document:");
  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
