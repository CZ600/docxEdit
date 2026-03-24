"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { loadDocx } = require("../src");
const { findSampleDocx } = require("./helpers/sample-file");

const SAMPLE_FILE = findSampleDocx();

test("VirtualWordDocument parses sample docx and persists cross-run replacements", async () => {
  const doc = await loadDocx(SAMPLE_FILE);
  const paragraph = doc
    .getParagraphs()
    .find((item) => item.getText().includes("地理科学学院25级硕士地信团支部"));

  assert.ok(paragraph, "expected to find the sample paragraph");
  assert.match(paragraph.getText(), /25级硕士地信团支部/);

  const replacedCount = doc.replaceAll("地理科学学院25级硕士地信团支部", "地理科学学院2025级硕士地信团支部");
  assert.equal(replacedCount, 1);

  const outputPath = path.join(os.tmpdir(), `docx-vcomponent-${Date.now()}.docx`);
  await doc.saveAs(outputPath);

  const reloaded = await loadDocx(outputPath);
  const updatedParagraph = reloaded
    .getParagraphs()
    .find((item) => item.getText().includes("地理科学学院2025级硕士地信团支部"));

  assert.ok(updatedParagraph, "expected replacement to persist after reload");

  await fs.unlink(outputPath);
});

test("VirtualWordDocument fills the sample table and edits header/footer content", async () => {
  const doc = await loadDocx(SAMPLE_FILE);
  const table = doc.getTables()[0];
  const header = doc.getHeaders()[0];
  const footer = doc.getFooters()[0];

  assert.ok(table, "expected the sample document to contain one table");
  assert.ok(header, "expected the sample document to contain a header");
  assert.ok(footer, "expected the sample document to contain a footer");

  table.fill(
    [
      ["地信青年团日活动", "2026-03-24", "张三", "第一轮填充"],
      ["两会专题研讨", "2026-03-25", "李四", "第二轮填充"],
    ],
    { startRow: 1 },
  );
  header.getParagraph(0).setText("测试页眉-已更新");
  footer.getParagraph(0).setText("测试页脚-已更新");

  const outputPath = path.join(os.tmpdir(), `docx-vcomponent-special-${Date.now()}.docx`);
  await doc.saveAs(outputPath);

  const reloaded = await loadDocx(outputPath);
  const reloadedTable = reloaded.getTables()[0];

  assert.equal(reloadedTable.getCell(1, 0).getText(), "地信青年团日活动");
  assert.equal(reloadedTable.getCell(1, 1).getText(), "2026-03-24");
  assert.equal(reloadedTable.getCell(1, 2).getText(), "张三");
  assert.equal(reloadedTable.getCell(1, 3).getText(), "第一轮填充");
  assert.equal(reloadedTable.getCell(2, 0).getText(), "两会专题研讨");
  assert.equal(reloaded.getHeaders()[0].getParagraph(0).getText(), "测试页眉-已更新");
  assert.equal(reloaded.getFooters()[0].getParagraph(0).getText(), "测试页脚-已更新");
  assert.equal(reloaded.getFootnotes().length, 0);

  await fs.unlink(outputPath);
});
