"use strict";

const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const { createVNode, loadDocx } = require("../src");

function findSampleDocx(projectRoot) {
  const fileName = fs
    .readdirSync(projectRoot)
    .find((name) => name.endsWith(".docx") && !name.endsWith(".modified.docx"));

  if (!fileName) {
    throw new Error("Unable to locate the sample .docx file.");
  }

  return path.join(projectRoot, fileName);
}

async function createReplaceDocx(inputPath, outputDir) {
  const doc = await loadDocx(inputPath);

  doc.replaceAll(
    "\u5730\u7406\u79D1\u5B66\u5B66\u966225\u7EA7\u7855\u58EB\u5730\u4FE1\u56E2\u652F\u90E8",
    "\u5730\u7406\u79D1\u5B66\u5B66\u96622026\u7EA7\u7855\u58EB\u5730\u4FE1\u56E2\u652F\u90E8",
  );

  const outputPath = path.join(outputDir, "text-replaced.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createVirtualTreeDocx(inputPath, outputDir) {
  const doc = await loadDocx(inputPath);
  const tree = doc.toComponentTree();
  const body = tree.children.find((node) => node.type === "body");
  const firstParagraph = body.children.find((node) => node.type === "paragraph" && (node.props.text || "").trim());

  if (!firstParagraph) {
    throw new Error("Unable to locate a non-empty body paragraph.");
  }

  firstParagraph.props.text = "\u8FD9\u662F\u901A\u8FC7\u865A\u62DF\u6811 patch \u4FEE\u6539\u540E\u7684\u6BB5\u843D\u5185\u5BB9\u3002";
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "\u8FD9\u662F\u65B0\u63D2\u5165\u7684\u68C0\u67E5\u6BB5\u843D\uFF0C\u7528\u4E8E\u4EBA\u5DE5\u786E\u8BA4\u865A\u62DF\u6811\u63D2\u5165\u6548\u679C\u3002" },
      children: [],
    }),
  );

  await doc.patch(tree);

  const outputPath = path.join(outputDir, "text-virtual-tree.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createTableHeaderFooterDocx(inputPath, outputDir) {
  const doc = await loadDocx(inputPath);
  const table = doc.getTables()[0];
  const header = doc.getHeaders()[0];
  const footer = doc.getFooters()[0];

  if (!table) {
    throw new Error("Unable to locate a table in the sample document.");
  }

  table.fill(
    [
      ["\u5730\u4FE1\u9752\u5E74\u56E2\u65E5\u6D3B\u52A8", "2026-03-24", "\u5F20\u4E09", "\u793A\u4F8B\u8868\u683C\u586B\u5145"],
      ["\u4E13\u9898\u5B66\u4E60\u5206\u4EAB", "2026-03-25", "\u674E\u56DB", "\u7B2C\u4E8C\u884C\u4EBA\u5DE5\u68C0\u67E5"],
    ],
    { startRow: 1 },
  );

  if (header) {
    header.getParagraph(0).setText("\u6587\u5B57\u68C0\u67E5\u9875\u7709-\u5DF2\u66F4\u65B0");
  }

  if (footer) {
    footer.getParagraph(0).setText("\u6587\u5B57\u68C0\u67E5\u9875\u811A-\u5DF2\u66F4\u65B0");
  }

  const outputPath = path.join(outputDir, "text-table-header-footer.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const outputDir = path.join(projectRoot, "examples", "output");
  const inputPath = findSampleDocx(projectRoot);

  await fsp.mkdir(outputDir, { recursive: true });

  const [replacePath, virtualTreePath, tablePath] = await Promise.all([
    createReplaceDocx(inputPath, outputDir),
    createVirtualTreeDocx(inputPath, outputDir),
    createTableHeaderFooterDocx(inputPath, outputDir),
  ]);

  console.log("Generated text inspection documents:");
  console.log(replacePath);
  console.log(virtualTreePath);
  console.log(tablePath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
