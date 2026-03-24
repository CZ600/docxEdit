"use strict";

const fs = require("node:fs/promises");
const path = require("node:path");
const JSZip = require("jszip");
const { createVNode, loadDocx } = require("../src");

async function buildSeedDocx(seedImageBuffer) {
  const zip = new JSZip();

  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`,
  );

  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`,
  );

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
        <w:t>\u56FE\u7247\u68C0\u67E5\u793A\u4F8B</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="990000" cy="792000"/>
            <wp:docPr id="1" name="testImage.jpg" descr="seed-image"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="testImage.jpg" descr="seed-image"/>
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
    <w:p>
      <w:r>
        <w:t>\u56FE1 \u521D\u59CB\u56FE\u7247\u9898\u6CE8</w:t>
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

  zip.file("word/media/image1.jpg", seedImageBuffer);
  return zip.generateAsync({ type: "nodebuffer" });
}

async function createReplaceDocx(seedBuffer, replacementBuffer, outputDir) {
  const doc = await loadDocx(seedBuffer);
  const image = doc.getImages()[0];

  image.replace({
    data: replacementBuffer,
    filename: "testImage2.jpg",
    contentType: "image/jpeg",
    width: "1110000",
    height: "888000",
    alt: "\u66FF\u6362\u540E\u7684\u56FE\u7247",
  });

  const tree = doc.toComponentTree();
  const body = tree.children.find((node) => node.type === "body");
  body.children[2].props.text = "\u56FE1 \u66FF\u6362\u540E\u7684\u56FE\u7247\u9898\u6CE8";
  await doc.patch(tree);

  const outputPath = path.join(outputDir, "image-replaced.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createRemoveDocx(seedBuffer, outputDir) {
  const doc = await loadDocx(seedBuffer);
  doc.getImages()[0].remove();

  const tree = doc.toComponentTree();
  const body = tree.children.find((node) => node.type === "body");
  body.children[1].props.text = "\u539F\u56FE\u7247\u5DF2\u5220\u9664";
  body.children[2].props.text = "\u56FE1 \u56FE\u7247\u5DF2\u5220\u9664";
  await doc.patch(tree);

  const outputPath = path.join(outputDir, "image-removed.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function createInsertDocx(seedBuffer, insertedBuffer, outputDir) {
  const doc = await loadDocx(seedBuffer);
  doc.getImages()[0].remove();

  const tree = doc.toComponentTree();
  const body = tree.children.find((node) => node.type === "body");
  const imageParagraph = body.children[1];
  const captionParagraph = body.children[2];

  imageParagraph.children = [
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
            width: "777000",
            height: "620000",
            alt: "\u65B0\u589E\u56FE\u7247",
          },
          children: [],
        }),
      ],
    }),
  ];
  imageParagraph.props.text = "";
  captionParagraph.props.text = "\u56FE1 \u65B0\u589E\u56FE\u7247\u9898\u6CE8";

  await doc.patch(tree);

  const outputPath = path.join(outputDir, "image-inserted.docx");
  await doc.saveAs(outputPath);
  return outputPath;
}

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const outputDir = path.join(projectRoot, "examples", "output");
  const image1Path = path.join(projectRoot, "testImage.jpg");
  const image2Path = path.join(projectRoot, "testImage2.jpg");

  const [image1Buffer, image2Buffer] = await Promise.all([
    fs.readFile(image1Path),
    fs.readFile(image2Path),
    fs.mkdir(outputDir, { recursive: true }),
  ]);

  const seedBuffer = await buildSeedDocx(image1Buffer);
  const [replacedPath, removedPath, insertedPath] = await Promise.all([
    createReplaceDocx(seedBuffer, image2Buffer, outputDir),
    createRemoveDocx(seedBuffer, outputDir),
    createInsertDocx(seedBuffer, image1Buffer, outputDir),
  ]);

  console.log("Generated image inspection documents:");
  console.log(replacedPath);
  console.log(removedPath);
  console.log(insertedPath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
