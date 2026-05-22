"use strict";

const path = require("node:path");
const fs = require("node:fs/promises");
const { loadDocx } = require("../src");

const SOURCE_DOCX = path.resolve(__dirname, "..", "tests", "test-docx", "测试文档1.docx");
const TARGET_DOCX = path.resolve(__dirname, "..", "tests", "test-docx", "格式测试文档.docx");
const OUTPUT_DOCX = path.resolve(__dirname, "..", "tests", "test-docx", "格式测试文档.formatted.docx");

async function main() {
  await fs.access(SOURCE_DOCX);
  await fs.access(TARGET_DOCX);

  const sourceDoc = await loadDocx(SOURCE_DOCX);
  const targetDoc = await loadDocx(TARGET_DOCX);

  const sourceProfile = sourceDoc.getStyleProfile();
  const sourceNamedStyles = sourceDoc.getNamedStyles();
  const targetNamedStyles = targetDoc.getNamedStyles();

  // Build source styleId lookup by (name, type)
  const sourceStyleByName = new Map();
  for (const s of sourceNamedStyles) {
    sourceStyleByName.set(`${s.name}\0${s.type}`, s.styleId);
  }

  // Match target styles to source styles by name, apply source formatting under target's styleId
  const matchedStyles = {};
  const matchDetails = [];
  let matchCount = 0;

  for (const tgt of targetNamedStyles) {
    const key = `${tgt.name}\0${tgt.type}`;
    const srcStyleId = sourceStyleByName.get(key);
    if (!srcStyleId || !sourceProfile.styles[srcStyleId]) continue;

    const srcDef = sourceProfile.styles[srcStyleId];
    matchedStyles[tgt.styleId] = {
      name: tgt.name,
      type: tgt.type,
      basedOn: srcDef.basedOn || tgt.basedOn,
      paragraphStyle: srcDef.paragraphStyle,
      runStyle: srcDef.runStyle,
    };
    matchDetails.push({ name: tgt.name, targetId: tgt.styleId, sourceId: srcStyleId });
    matchCount += 1;
  }

  console.log("匹配到的样式:");
  matchDetails.forEach((d) => console.log(`  ${d.name} (source:${d.sourceId} → target:${d.targetId})`));

  // Apply doc defaults from source
  targetDoc.applyStyleProfile({ defaults: sourceProfile.defaults });
  // Apply matched styles
  targetDoc.applyStyleProfile({ styles: matchedStyles });

  await targetDoc.saveAs(OUTPUT_DOCX);

  // Verify: compare reloaded output style definitions with source
  const reloaded = await loadDocx(OUTPUT_DOCX);
  const reloadedProfile = reloaded.getStyleProfile();

  console.log("\n验证结果:");
  let allMatch = true;
  for (const detail of matchDetails) {
    const srcDef = sourceProfile.styles[detail.sourceId];
    const outDef = reloadedProfile.styles[detail.targetId];
    const paraMatch = JSON.stringify(srcDef.paragraphStyle) === JSON.stringify((outDef && outDef.paragraphStyle) || {});
    const runMatch = JSON.stringify(srcDef.runStyle) === JSON.stringify((outDef && outDef.runStyle) || {});
    const ok = paraMatch && runMatch;
    if (!ok) allMatch = false;
    console.log(`  ${detail.name}: ${ok ? "PASS" : "FAIL"} (para:${paraMatch} run:${runMatch})`);
    if (!ok) {
      console.log(`    source para: ${JSON.stringify(srcDef.paragraphStyle)}`);
      console.log(`    output para: ${JSON.stringify((outDef && outDef.paragraphStyle) || {})}`);
      console.log(`    source run:  ${JSON.stringify(srcDef.runStyle)}`);
      console.log(`    output run:  ${JSON.stringify((outDef && outDef.runStyle) || {})}`);
    }
  }

  console.log(`\n匹配: ${matchCount} | 源样式数: ${sourceNamedStyles.length} | 目标样式数: ${targetNamedStyles.length}`);
  console.log(`输出文件: ${OUTPUT_DOCX}`);
  if (allMatch) console.log("所有匹配样式验证通过。");
  else console.log("部分样式验证失败，请检查上方详情。");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
