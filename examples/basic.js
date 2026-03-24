"use strict";

const fs = require("node:fs");
const path = require("node:path");
const { loadDocx } = require("../src");

function findSampleDocx(projectRoot) {
  const fileName = fs
    .readdirSync(projectRoot)
    .find((name) => name.endsWith(".docx") && !name.endsWith(".modified.docx"));

  if (!fileName) {
    throw new Error("Unable to locate the sample .docx file.");
  }

  return path.join(projectRoot, fileName);
}

async function main() {
  const projectRoot = path.resolve(__dirname, "..");
  const inputPath = findSampleDocx(projectRoot);
  const outputPath = inputPath.replace(/\.docx$/i, ".modified.docx");
  const doc = await loadDocx(inputPath);
  const firstContentParagraph = doc.getBody().getParagraphs().find((paragraph) => paragraph.getText().trim());
  const table = doc.getTables()[0];
  const header = doc.getHeaders()[0];
  const footer = doc.getFooters()[0];

  console.log("First body paragraph:");
  console.log(firstContentParagraph.getText());
  console.log("Header before:", header ? header.getParagraph(0).getText() : "<none>");
  console.log("Footer before:", footer ? footer.getParagraph(0).getText() : "<none>");

  table.fill(
    [
      ["地信青年团日活动", "2026-03-24", "张三", "示例填充"],
      ["两会专题研讨", "2026-03-25", "李四", "第二行示例"],
    ],
    { startRow: 1 },
  );

  if (header) {
    header.getParagraph(0).setText("测试页眉-已更新");
  }

  if (footer) {
    footer.getParagraph(0).setText("测试页脚-已更新");
  }

  await doc.saveAs(outputPath);

  console.log(`Saved modified file to ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
