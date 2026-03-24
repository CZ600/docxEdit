"use strict";

const path = require("node:path");
const { loadDocx } = require("../src");

async function main() {
  const inputPath = path.resolve(__dirname, "..", "2活动新闻稿.docx");
  const outputPath = path.resolve(__dirname, "..", "2活动新闻稿.modified.docx");
  const doc = await loadDocx(inputPath);
  const firstContentParagraph = doc.getParagraphs().find((paragraph) => paragraph.getText().trim());

  console.log("First paragraph:");
  console.log(firstContentParagraph.getText());

  doc.replaceAll("地理科学学院25级硕士地信团支部", "地理科学学院2025级硕士地信团支部");
  await doc.saveAs(outputPath);

  console.log(`Saved modified file to ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
