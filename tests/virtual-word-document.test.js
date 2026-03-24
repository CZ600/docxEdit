"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const { loadDocx } = require("../src");

const SAMPLE_FILE = path.resolve(__dirname, "..", "2活动新闻稿.docx");

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
