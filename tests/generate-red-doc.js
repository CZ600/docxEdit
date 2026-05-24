"use strict";

const { loadDocx, createVNode } = require("../src");
const JSZip = require("jszip");
const path = require("path");
const fs = require("fs");

/**
 * 将含 [[FOOTNOTE_REF:id]] 和 [[ENDNOTE_REF:id]] 占位符的文本
 * 拆分为多个 run VNode，脚注引用部分用 vertAlign: "superscript" 渲染为上角标数字。
 */
function buildRunsWithSuperscriptRefs(text, baseStyle) {
  const parts = text.split(/(\[\[FOOTNOTE_REF:\d+\]\]|\[\[ENDNOTE_REF:\d+\]\])/);
  return parts
    .filter((p) => p.length > 0)
    .map((part) => {
      const fnMatch = part.match(/^\[\[FOOTNOTE_REF:(\d+)\]\]$/);
      const enMatch = part.match(/^\[\[ENDNOTE_REF:(\d+)\]\]$/);
      if (fnMatch || enMatch) {
        // 脚注/尾注引用 → 上角标数字
        const id = fnMatch ? fnMatch[1] : enMatch[1];
        return createVNode({
          type: "run",
          props: { style: { ...baseStyle, vertAlign: "superscript" } },
          children: [createVNode({ type: "text", props: { text: id }, children: [] })],
        });
      }
      // 普通文本
      return createVNode({
        type: "run",
        props: { style: baseStyle },
        children: [createVNode({ type: "text", props: { text: part }, children: [] })],
      });
    });
}

async function main() {
  const srcPath = path.resolve(__dirname, "test-docx", "角标和公式测试文档.docx");
  const doc = await loadDocx(srcPath);

  // 读取原始内容
  console.log("=== 原始文档内容 ===");
  const paragraphs = doc.getBody().getParagraphs();
  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i].getText();
    if (!text) continue;
    console.log("段落" + i + ":", text);
  }
  const maths = doc.getMaths();
  console.log("\n数学公式:");
  maths.forEach((m, i) => console.log("  公式" + (i + 1) + ":", m.getText(), "| display:", m.getDisplay()));

  // 新建空文档
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
  );
  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body></w:body></w:document>',
  );

  const newBuffer = await zip.generateAsync({ type: "nodebuffer" });
  const newDoc = await loadDocx(newBuffer);
  const tree = newDoc.toComponentTree();
  const body = tree.children.find((c) => c.type === "body");

  // 1. 标题
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: { style: { color: "FF0000", bold: true, fontSize: "36" } },
          children: [createVNode({ type: "text", props: { text: "角标和公式测试 - 红色版本" }, children: [] })],
        }),
      ],
    }),
  );

  // 2. 提取的原始段落文本（红色），将 [[FOOTNOTE_REF:id]] 转为真正的上角标
  const originalText = doc.getBody().getParagraph(0).getText();
  const extractedRuns = buildRunsWithSuperscriptRefs(originalText, { color: "FF0000", fontSize: "24" });
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: { style: { color: "FF0000", fontSize: "24" } },
          children: [createVNode({ type: "text", props: { text: "提取的段落: " }, children: [] })],
        }),
        ...extractedRuns,
      ],
    }),
  );

  // 3. 上角标演示: x² + y³ = z
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: { style: { color: "FF0000" } },
          children: [createVNode({ type: "text", props: { text: "上角标演示: x" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { color: "FF0000", vertAlign: "superscript" } },
          children: [createVNode({ type: "text", props: { text: "2" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { color: "FF0000" } },
          children: [createVNode({ type: "text", props: { text: " + y" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { color: "FF0000", vertAlign: "superscript" } },
          children: [createVNode({ type: "text", props: { text: "3" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { color: "FF0000" } },
          children: [createVNode({ type: "text", props: { text: " = z" }, children: [] })],
        }),
      ],
    }),
  );

  // 4. 下角标演示: H₂O
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: { style: { color: "FF0000" } },
          children: [createVNode({ type: "text", props: { text: "下角标演示: H" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { color: "FF0000", vertAlign: "subscript" } },
          children: [createVNode({ type: "text", props: { text: "2" }, children: [] })],
        }),
        createVNode({
          type: "run",
          props: { style: { color: "FF0000" } },
          children: [createVNode({ type: "text", props: { text: "O 是水" }, children: [] })],
        }),
      ],
    }),
  );

  // 5. 内联数学公式
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "run",
          props: { style: { color: "FF0000" } },
          children: [createVNode({ type: "text", props: { text: "内联公式: " }, children: [] })],
        }),
        createVNode({
          type: "math",
          props: {
            text: "E=mc^2",
            style: { fontFamily: { ascii: "Cambria Math", hAnsi: "Cambria Math" } },
          },
          children: [],
        }),
      ],
    }),
  );

  // 6. 块级数学公式
  body.children.push(
    createVNode({
      type: "paragraph",
      props: { text: "" },
      children: [
        createVNode({
          type: "math",
          props: {
            text: "a^2+b^2=c^2",
            display: "block",
            style: {
              justification: "center",
              bold: true,
              fontSize: "32",
              fontFamily: { ascii: "Cambria Math", hAnsi: "Cambria Math" },
            },
          },
          children: [],
        }),
      ],
    }),
  );

  newDoc.patch(tree);
  const outputPath = path.resolve(__dirname, "test-docx", "红色角标和公式输出v2.docx");
  await newDoc.saveAs(outputPath);

  // 验证输出文件原始 XML
  const verifyBuf = fs.readFileSync(outputPath);
  const verifyZip = await JSZip.loadAsync(verifyBuf);
  const verifyXml = await verifyZip.file("word/document.xml").async("string");

  console.log("\n=== 输出文件验证 ===");
  console.log("文件:", outputPath);
  console.log("文件大小:", verifyBuf.length, "bytes");
  console.log("包含 vertAlign superscript:", /w:vertAlign\s+w:val="superscript"/.test(verifyXml));
  console.log("包含 vertAlign subscript:", /w:vertAlign\s+w:val="subscript"/.test(verifyXml));
  console.log("包含 color FF0000:", /w:color\s+w:val="FF0000"/.test(verifyXml));
  console.log("包含 m:oMath:", /m:oMath/.test(verifyXml));
  console.log("包含 m:oMathPara:", /m:oMathPara/.test(verifyXml));
  console.log("包含 E=mc^2:", verifyXml.includes("E=mc^2"));
  console.log("包含 a^2+b^2=c^2:", verifyXml.includes("a^2+b^2=c^2"));

  // 重新加载验证
  const reloaded = await loadDocx(outputPath);
  const rp = reloaded.getBody().getParagraphs();
  console.log("\n=== 重新加载验证 ===");
  for (let i = 0; i < rp.length; i++) {
    const text = rp[i].getText();
    if (!text) continue;
    console.log("段落" + i + ":", text);
  }
  const rm = reloaded.getMaths();
  console.log("公式数量:", rm.length);
  rm.forEach((m, i) => console.log("  公式" + (i + 1) + ":", m.getText(), "| display:", m.getDisplay()));

  console.log("\n文件已保存，请用 Word 打开查看效果。");
}

main().catch(console.error);
