"use strict";

const fs = require("node:fs/promises");
const path = require("node:path");
const JSZip = require("jszip");
const { loadDocx } = require("../src");

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
      <w:r><w:t>图片布局检查示例</w:t></w:r>
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
      <w:r><w:t>图1 初始图片题注</w:t></w:r>
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

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const outputDir = path.join(projectRoot, "examples", "output");
  const image1Path = path.join(projectRoot, "testImage.jpg");
  const image2Path = path.join(projectRoot, "testImage2.jpg");
  const outputPath = path.join(outputDir, "image-layout-edited.docx");

  await fs.mkdir(outputDir, { recursive: true });

  const [image1Buffer, image2Buffer] = await Promise.all([
    fs.readFile(image1Path),
    fs.readFile(image2Path),
  ]);

  const doc = await loadDocx(await buildSeedDocx(image1Buffer));
  const image = doc.getImages()[0];

  image.replace({
    data: image2Buffer,
    filename: "testImage2.jpg",
    contentType: "image/jpeg",
    width: "1200000",
    height: "960000",
    alt: "紧密环绕测试图片",
    layout: {
      mode: "anchor",
      wrap: "tight",
      distances: {
        top: "0",
        bottom: "0",
        left: "114300",
        right: "114300",
      },
      positionH: {
        relativeFrom: "margin",
        align: "center",
      },
      positionV: {
        relativeFrom: "paragraph",
        offset: "0",
      },
      allowOverlap: true,
      layoutInCell: true,
    },
    paragraphAlignment: "center",
  });

  doc.getBody().getParagraph(2).setText("图1 已调整大小并设置为紧密环绕布局");
  await doc.saveAs(outputPath);

  console.log("Generated image layout inspection document:");
  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
