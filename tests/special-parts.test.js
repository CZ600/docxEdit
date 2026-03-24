"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");
const { loadDocx } = require("../src");

async function buildSyntheticDocx() {
  const zip = new JSZip();

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p>
      <w:r><w:t>正文</w:t></w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape>
            <v:textbox>
              <w:txbxContent>
                <w:p>
                  <w:r><w:t>文本框原文</w:t></w:r>
                </w:p>
              </w:txbxContent>
            </v:textbox>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
  </w:body>
</w:document>`,
  );

  zip.file(
    "word/comments.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="tester">
    <w:p>
      <w:r><w:t>批注原文</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>`,
  );

  return zip.generateAsync({ type: "nodebuffer" });
}

test("VirtualWordDocument parses comments and text boxes from supported parts", async () => {
  const buffer = await buildSyntheticDocx();
  const doc = await loadDocx(buffer);

  assert.equal(doc.getTextBoxes().length, 1);
  assert.equal(doc.getComments().length, 1);
  assert.equal(doc.getTextBoxes()[0].getText(), "文本框原文");
  assert.equal(doc.getComments()[0].getText(), "批注原文");

  doc.getTextBoxes()[0].getParagraphs()[0].setText("文本框已更新");
  doc.getComments()[0].replaceAll("批注原文", "批注已更新");

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.equal(reloaded.getTextBoxes()[0].getText(), "文本框已更新");
  assert.equal(reloaded.getComments()[0].getText(), "批注已更新");
});
