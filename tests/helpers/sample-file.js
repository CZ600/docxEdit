"use strict";

const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const JSZip = require("jszip");

const SAMPLE_DOCX_NAME = "测试文档.docx";

function findSampleDocx() {
  const projectRoot = path.resolve(__dirname, "..", "..");
  const samplePath = path.join(projectRoot, SAMPLE_DOCX_NAME);
  if (!fs.existsSync(samplePath)) {
    throw new Error(`Unable to locate ${SAMPLE_DOCX_NAME} in the project root.`);
  }

  return samplePath;
}

async function buildSampleDocxWithHeaderFooterParts() {
  const zip = await JSZip.loadAsync(await fsp.readFile(findSampleDocx()));

  zip.file(
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:hdr>`,
  );
  zip.file(
    "word/footer1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:ftr>`,
  );

  return zip.generateAsync({ type: "nodebuffer" });
}

module.exports = {
  buildSampleDocxWithHeaderFooterParts,
  findSampleDocx,
  SAMPLE_DOCX_NAME,
};
