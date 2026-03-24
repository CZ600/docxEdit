"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const JSZip = require("jszip");
const { createVNode, loadDocx } = require("../src");

const TEST_IMAGE_1 = path.resolve(__dirname, "..", "testImage.jpg");
const TEST_IMAGE_2 = path.resolve(__dirname, "..", "testImage2.jpg");

async function buildImageDocx() {
  const zip = new JSZip();
  const imageBuffer = await fs.readFile(TEST_IMAGE_1);

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Before</w:t>
      </w:r>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="990000" cy="792000"/>
            <wp:docPr id="1" name="testImage.jpg" descr="seed"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="testImage.jpg" descr="seed"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId1"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="990000" cy="792000"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"/>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>`,
  );

  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpg"/>
</Relationships>`,
  );

  zip.file("word/media/image1.jpg", imageBuffer);
  return zip.generateAsync({ type: "nodebuffer" });
}

test("image nodes can be replaced with testImage2.jpg and removed", async () => {
  const replacementBuffer = await fs.readFile(TEST_IMAGE_2);
  const doc = await loadDocx(await buildImageDocx());
  assert.equal(doc.getImages().length, 1);

  const image = doc.getImages()[0];
  assert.equal(image.getFilename(), "testImage.jpg");
  assert.equal(image.getContentType(), "image/jpeg");
  assert.deepEqual(image.getSize(), { width: "990000", height: "792000" });

  image.replace({
    data: replacementBuffer,
    filename: "testImage2.jpg",
    contentType: "image/jpeg",
    width: "1110000",
    height: "888000",
    alt: "替换后的图片",
  });

  doc.getBody().getParagraph(0).setText("Before");
  const replaceTree = doc.toComponentTree();
  const body = replaceTree.children.find((node) => node.type === "body");
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "图1 替换后的题注" },
      children: [],
    }),
  );
  doc.patch(replaceTree);

  const replaced = await loadDocx(await doc.toBuffer());
  assert.equal(replaced.getImages().length, 1);
  assert.equal(replaced.getImages()[0].getFilename(), "testImage2.jpg");
  assert.equal(replaced.getImages()[0].getContentType(), "image/jpeg");
  assert.deepEqual(replaced.getImages()[0].getSize(), { width: "1110000", height: "888000" });
  assert.equal(replaced.getBody().getParagraphs().at(-1).getText(), "图1 替换后的题注");

  replaced.getImages()[0].remove();
  const removed = await loadDocx(await replaced.toBuffer());
  assert.equal(removed.getImages().length, 0);
});

test("doc.patch can insert testImage.jpg and delete an existing image in the same document", async () => {
  const insertedBuffer = await fs.readFile(TEST_IMAGE_1);
  const doc = await loadDocx(await buildImageDocx());
  const tree = doc.toComponentTree();
  const body = tree.children.find((node) => node.type === "body");
  const paragraph = body.children[0];

  paragraph.children.push(
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
            alt: "新增图片",
          },
          children: [],
        }),
      ],
    }),
  );
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "图2 新增图片题注" },
      children: [],
    }),
  );

  doc.patch(tree);
  assert.equal(doc.getImages().length, 2);
  doc.getImages()[0].remove();
  assert.equal(doc.getImages().length, 1);

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.equal(reloaded.getImages().length, 1);
  assert.equal(reloaded.getImages()[0].getFilename(), "testImage.jpg");
  assert.equal(reloaded.getImages()[0].getContentType(), "image/jpeg");
  assert.deepEqual(reloaded.getImages()[0].getSize(), { width: "555000", height: "444000" });
  assert.equal(reloaded.getBody().getParagraphs().at(-1).getText(), "图2 新增图片题注");
});
