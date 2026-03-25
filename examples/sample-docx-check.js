"use strict";

const fs = require("node:fs/promises");
const path = require("node:path");
const { createVNode, loadDocx } = require("../src");
const { buildSampleDocxWithHeaderFooterParts, findSampleDocx } = require("../tests/helpers/sample-file");

const TEST_IMAGE_1 = path.resolve(__dirname, "..", "testImage.jpg");
const TEST_IMAGE_2 = path.resolve(__dirname, "..", "testImage2.jpg");

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

async function createTextDocx(inputPath, outputDir) {
  const doc = await loadDocx(inputPath);
  const title = doc.getBody().getParagraph(0);
  const intro = doc.getBody().getParagraph(1);

  title.replace("和类", "");
  intro.replace("一种面向对象的编程语言", "一种常见的面向对象的编程语言");
  intro.replace("基本概念", "核心概念");

  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");
  body.children = body.children.filter((child) => child.type !== "sectPr");
  body.children.forEach((child, index) => {
    child.key = child.id || `body-${index}`;
  });
  body.children.splice(
    2,
    0,
    createVNode({
      key: "inserted-text-paragraph",
      type: "paragraph",
      props: { text: "这是插入的测试段落。" },
      children: [],
    }),
  );

  doc.patch(tree);

  const outputPath = path.join(outputDir, "sample-text-edited.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createTableDocx(inputPath, outputDir) {
  const doc = await loadDocx(inputPath);
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");

  body.children.push(createTable(3, 2));
  doc.patch(tree);

  doc.getTables()[0].fill(
    [
      ["测试项", "结果"],
      ["表格插入", "通过"],
      ["表格填充", "通过"],
    ],
    { startRow: 0 },
  );

  const outputPath = path.join(outputDir, "sample-table-edited.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createImageDocx(inputPath, outputDir) {
  const [replacementBuffer, insertedBuffer] = await Promise.all([
    fs.readFile(TEST_IMAGE_2),
    fs.readFile(TEST_IMAGE_1),
  ]);

  const doc = await loadDocx(inputPath);
  doc.getImages()[0].replace({
    data: replacementBuffer,
    filename: "testImage2.jpg",
    contentType: "image/jpeg",
    width: "1110000",
    height: "888000",
    alt: "替换后的测试图片",
  });

  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: {},
          children: [
            createVNode({
              type: "image",
              props: {
                filename: "testImage.jpg",
                contentType: "image/jpeg",
                data: insertedBuffer,
                width: "555000",
                height: "444000",
                alt: "新增测试图片",
              },
              children: [],
            }),
          ],
        }),
      ],
    }),
  );

  doc.patch(tree);

  const outputPath = path.join(outputDir, "sample-image-edited.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createHeaderFooterDocx(outputDir) {
  const doc = await loadDocx(await buildSampleDocxWithHeaderFooterParts());
  const tree = doc.toComponentTree();
  const header = getPart(tree, "header");
  const footer = getPart(tree, "footer");

  header.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "这是插入的测试页眉" },
      children: [],
    }),
  );
  footer.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "这是插入的测试页脚" },
      children: [],
    }),
  );

  doc.patch(tree);

  const outputPath = path.join(outputDir, "sample-header-footer-edited.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const outputDir = path.join(projectRoot, "examples", "output");
  const inputPath = findSampleDocx();

  await fs.mkdir(outputDir, { recursive: true });

  const [textPath, tablePath, imagePath, headerFooterPath] = await Promise.all([
    createTextDocx(inputPath, outputDir),
    createTableDocx(inputPath, outputDir),
    createImageDocx(inputPath, outputDir),
    createHeaderFooterDocx(outputDir),
  ]);

  console.log("Generated sample inspection documents:");
  console.log(textPath);
  console.log(tablePath);
  console.log(imagePath);
  console.log(headerFooterPath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
