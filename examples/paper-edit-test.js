"use strict";

const fs = require("node:fs/promises");
const path = require("node:path");
const { loadDocx } = require("../src");
const { ParagraphTextModel } = require("../src/core/paragraph-text-model");

const INPUT_DOCX = "C:\\Users\\admin\\WPSDrive\\598227887\\WPS云盘\\硕士课程\\研一\\道路论文\\论文初稿中文版.docx";
const TARGET_TEXT = "DFCNet";
const REPLACEMENT_TEXT = "DFC-Net";
const TARGET_SENTENCE = "使用了DoRA和adapter的联合微调策略";
const REPLACEMENT_IMAGE = path.resolve(__dirname, "..", "testImage.jpg");

function countOccurrences(text, searchValue) {
  if (!text || !searchValue) return 0;
  return text.split(searchValue).length - 1;
}

async function main() {
  await fs.access(INPUT_DOCX);
  const imageBuffer = await fs.readFile(REPLACEMENT_IMAGE);
  const outputPath = path.resolve(process.cwd(), "论文初稿中文版.modified.docx");

  const doc = await loadDocx(INPUT_DOCX);
  const stats = {
    replacedText: 0,
    removedSentence: 0,
    replacedImages: 0,
  };

  for (const paragraph of doc.getParagraphs()) {
    const model = new ParagraphTextModel(paragraph.vnode.source);
    stats.replacedText += model.replaceAll(TARGET_TEXT, REPLACEMENT_TEXT);
    stats.removedSentence += model.replaceAll(TARGET_SENTENCE, "");
  }
  doc.rebuildFromXml();

  for (const image of doc.getImages()) {
    const size = image.getSize();
    image.replace({
      data: imageBuffer,
      filename: path.basename(REPLACEMENT_IMAGE),
      contentType: "image/jpeg",
      width: size.width || undefined,
      height: size.height || undefined,
    });
    stats.replacedImages += 1;
  }

  await doc.saveAs(outputPath);

  const reloaded = await loadDocx(outputPath);
  const allText = reloaded.getParagraphs().map((paragraph) => paragraph.getText()).join("\n");
  const remainingSourceText = countOccurrences(allText, TARGET_TEXT);
  const remainingSentence = countOccurrences(allText, TARGET_SENTENCE);

  console.log(JSON.stringify({
    input: INPUT_DOCX,
    output: outputPath,
    replacedText: stats.replacedText,
    removedSentence: stats.removedSentence,
    replacedImages: stats.replacedImages,
    remainingSourceText,
    remainingSentence,
    outputImageCount: reloaded.getImages().length,
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
