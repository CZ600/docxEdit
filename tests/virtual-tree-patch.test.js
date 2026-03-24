"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");
const { createVNode, loadDocx } = require("../src");

async function buildPatchableDocx() {
  const zip = new JSZip();

  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
    <w:p>
      <w:r><w:t>Left</w:t></w:r>
      <w:r><w:tab /></w:r>
      <w:r><w:t>Middle</w:t></w:r>
      <w:r><w:br /></w:r>
      <w:r><w:t>Right</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>Omega</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape>
            <v:textbox>
              <w:txbxContent>
                <w:p><w:r><w:t>Box old</w:t></w:r></w:p>
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
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Header old</w:t></w:r></w:p>
</w:hdr>`,
  );

  zip.file(
    "word/footer1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Footer old</w:t></w:r></w:p>
</w:ftr>`,
  );

  zip.file(
    "word/comments.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0">
    <w:p><w:r><w:t>Comment old</w:t></w:r></w:p>
  </w:comment>
</w:comments>`,
  );

  return zip.generateAsync({ type: "nodebuffer" });
}

function getPart(tree, type) {
  return tree.children.find((child) => child.type === type);
}

function removeIds(node) {
  delete node.id;
  for (const child of node.children || []) {
    removeIds(child);
  }
  return node;
}

function findFirstNodeByType(node, type) {
  if (node.type === type) {
    return node;
  }

  for (const child of node.children || []) {
    const found = findFirstNodeByType(child, type);
    if (found) {
      return found;
    }
  }

  return null;
}

test("doc.patch updates paragraph text while preserving tab and break nodes", async () => {
  const doc = await loadDocx(await buildPatchableDocx());
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");
  const paragraph = body.children[1];

  paragraph.props.text = "Start\tCenter\nFinish";

  const result = doc.patch(tree);
  assert.ok(result.operations.some((item) => item.type === "PROPS/TEXT_UPDATE"));
  assert.equal(doc.getBody().getParagraph(1).getText(), "Start\tCenter\nFinish");
  assert.equal(doc.partsData[0].xmlDocument.getElementsByTagName("w:tab").length, 1);
  assert.equal(doc.partsData[0].xmlDocument.getElementsByTagName("w:br").length, 1);

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.equal(reloaded.getBody().getParagraph(1).getText(), "Start\tCenter\nFinish");
});

test("doc.patch supports keyed structural insert remove and move with refreshed paragraph indexes", async () => {
  const doc = await loadDocx(await buildPatchableDocx());
  const tree = doc.toComponentTree();
  const body = getPart(tree, "body");

  body.children[0].key = "alpha";
  body.children[1].key = "control";
  body.children[2].key = "omega";
  body.children[3].key = "table";
  body.children[4].key = "textbox";
  doc.patch(tree);

  const keyedTree = doc.toComponentTree();
  const keyedBody = getPart(keyedTree, "body");

  const inserted = createVNode({
    key: "inserted",
    type: "paragraph",
    props: { text: "Inserted" },
    children: [],
  });

  keyedBody.children = [keyedBody.children[2], inserted, keyedBody.children[0], keyedBody.children[4]];
  const result = doc.patch(keyedTree);

  assert.ok(result.operations.some((item) => item.type === "MOVE"));
  assert.ok(result.operations.some((item) => item.type === "INSERT"));
  assert.ok(result.operations.some((item) => item.type === "REMOVE"));

  const bodyParagraphs = doc.getBody().getParagraphs().map((paragraph) => paragraph.getText());
  assert.deepEqual(bodyParagraphs, ["Omega", "Inserted", "Alpha", "", "Box old"]);
  assert.equal(doc.getParagraphs().length, 8);

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.deepEqual(
    reloaded.getBody().getParagraphs().map((paragraph) => paragraph.getText()),
    ["Omega", "Inserted", "Alpha", "", "Box old"],
  );
});

test("doc.patch supports unkeyed sequential updates and table structure changes that remain compatible with fill()", async () => {
  const doc = await loadDocx(await buildPatchableDocx());
  const plainTree = removeIds(doc.toComponentTree().toJSON());
  const body = getPart(plainTree, "body");
  const table = body.children.find((child) => child.type === "table");

  body.children = [body.children[2], body.children[1], body.children[0], table, body.children[4]];
  table.children.push({
    type: "table-row",
    props: {},
    children: [
      {
        type: "table-cell",
        props: {},
        children: [{ type: "paragraph", props: { text: "R2C1" }, children: [] }],
      },
      {
        type: "table-cell",
        props: {},
        children: [{ type: "paragraph", props: { text: "R2C2" }, children: [] }],
      },
    ],
  });

  doc.patch(plainTree);
  assert.deepEqual(
    doc.getBody().getParagraphs().slice(0, 3).map((paragraph) => paragraph.getText()),
    ["Omega", "Left\tMiddle\nRight", "Alpha"],
  );

  const tableController = doc.getTables()[0];
  tableController.fill([["N1", "N2"]], { startRow: 2 });

  const reloaded = await loadDocx(await doc.toBuffer());
  const reloadedTable = reloaded.getTables()[0];
  assert.equal(reloadedTable.getRow(2).getCell(0).getText(), "N1");
  assert.equal(reloadedTable.getRow(2).getCell(1).getText(), "N2");
});

test("legacy write APIs and doc.patch can be mixed across header footer comments and text boxes", async () => {
  const doc = await loadDocx(await buildPatchableDocx());

  doc.getHeaders()[0].getParagraph(0).setText("Header via legacy API");

  const tree = doc.toComponentTree();
  getPart(tree, "footer").children[0].props.text = "Footer via patch";
  getPart(tree, "comments").children.push(
    createVNode({
      type: "comment",
      props: { id: "1" },
      children: [
        createVNode({
          type: "paragraph",
          props: { text: "Comment added" },
          children: [],
        }),
      ],
    }),
  );

  const body = getPart(tree, "body");
  const textBox = findFirstNodeByType(body.children[4], "text-box");
  const textBoxParagraph = findFirstNodeByType(textBox, "paragraph");
  textBoxParagraph.props.text = "Box via patch";

  doc.patch(tree);

  assert.equal(doc.getHeaders()[0].getParagraph(0).getText(), "Header via legacy API");
  assert.equal(doc.getFooters()[0].getParagraph(0).getText(), "Footer via patch");
  assert.equal(doc.getTextBoxes()[0].getText(), "Box via patch");
  assert.equal(doc.getComments().length, 2);
  assert.equal(doc.getComments()[1].getText(), "Comment added");

  const reloaded = await loadDocx(await doc.toBuffer());
  assert.equal(reloaded.getHeaders()[0].getParagraph(0).getText(), "Header via legacy API");
  assert.equal(reloaded.getFooters()[0].getParagraph(0).getText(), "Footer via patch");
  assert.equal(reloaded.getTextBoxes()[0].getText(), "Box via patch");
  assert.equal(reloaded.getComments().length, 2);
  assert.equal(reloaded.getComments()[1].getText(), "Comment added");
});
