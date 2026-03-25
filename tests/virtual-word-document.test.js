"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const JSZip = require("jszip");
const { createVNode, loadDocx } = require("../src");
const { buildSampleDocxWithHeaderFooterParts, findSampleDocx, SAMPLE_DOCX_NAME } = require("./helpers/sample-file");

const SAMPLE_FILE = findSampleDocx();
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

test(`uses ${SAMPLE_DOCX_NAME} for text deletion insertion and replacement`, async () => {
  const doc = await loadDocx(SAMPLE_FILE);
  const title = doc.getBody().getParagraph(0);
  const intro = doc.getBody().getParagraph(1);

  assert.equal(title.getText(), "Java 对象和类");

  const deleteCount = title.replace("和类", "");
  const insertCount = intro.replace("一种面向对象的编程语言", "一种常见的面向对象的编程语言");
  const replaceCount = intro.replace("基本概念", "核心概念");

  assert.equal(deleteCount, 1);
  assert.equal(insertCount, 1);
  assert.equal(replaceCount, 1);

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

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.equal(reloaded.getBody().getParagraph(0).getText(), "Java 对象");
  assert.equal(
    reloaded.getBody().getParagraph(1).getText(),
    "Java 作为一种常见的面向对象的编程语言，支持以下核心概念：",
  );
  assert.equal(reloaded.getBody().getParagraph(2).getText(), "这是插入的测试段落。");
});

test(`uses ${SAMPLE_DOCX_NAME} for table insertion and fill`, async () => {
  const doc = await loadDocx(SAMPLE_FILE);
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");

  body.children.push(createTable(3, 2));
  doc.patch(tree);

  assert.equal(doc.getTables().length, 1);

  doc.getTables()[0].fill(
    [
      ["测试项", "结果"],
      ["表格插入", "通过"],
      ["表格填充", "通过"],
    ],
    { startRow: 0 },
  );

  const reloaded = await loadDocx(await doc.toBuffer());
  const table = reloaded.getTables()[0];
  assert.ok(table, "expected the inserted table to persist");
  assert.equal(table.getCell(0, 0).getText(), "测试项");
  assert.equal(table.getCell(0, 1).getText(), "结果");
  assert.equal(table.getCell(1, 0).getText(), "表格插入");
  assert.equal(table.getCell(1, 1).getText(), "通过");
  assert.equal(table.getCell(2, 0).getText(), "表格填充");
  assert.equal(table.getCell(2, 1).getText(), "通过");
});

test(`uses ${SAMPLE_DOCX_NAME} for image insertion and replacement`, async () => {
  const replacementBuffer = await fs.readFile(TEST_IMAGE_2);
  const insertedBuffer = await fs.readFile(TEST_IMAGE_1);
  const doc = await loadDocx(SAMPLE_FILE);

  assert.equal(doc.getImages().length, 1);

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

  const outputBuffer = await doc.toBuffer();
  const reloaded = await loadDocx(outputBuffer);
  assert.equal(reloaded.getImages().length, 2);
  assert.equal(reloaded.getImages()[0].getFilename(), "testImage2.jpg");
  assert.deepEqual(reloaded.getImages()[0].getSize(), { width: "1110000", height: "888000" });
  assert.equal(reloaded.getImages()[1].getFilename(), "testImage.jpg");
  assert.deepEqual(reloaded.getImages()[1].getSize(), { width: "555000", height: "444000" });

  const zip = await JSZip.loadAsync(outputBuffer);
  const contentTypesXml = await zip.file("[Content_Types].xml").async("string");
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.match(contentTypesXml, /Extension="jpg"/);
  assert.match(documentXml, /<w:p><w:r><w:drawing>[\s\S]*<\/w:p><w:sectPr/);
});

test(`uses ${SAMPLE_DOCX_NAME} for header and footer insertion`, async () => {
  const doc = await loadDocx(await buildSampleDocxWithHeaderFooterParts());

  assert.equal(doc.getHeaders().length, 1);
  assert.equal(doc.getFooters().length, 1);

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

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.equal(reloaded.getHeaders()[0].getParagraph(0).getText(), "这是插入的测试页眉");
  assert.equal(reloaded.getFooters()[0].getParagraph(0).getText(), "这是插入的测试页脚");
});
