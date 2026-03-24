# docx-vcomponent

一个基于 JavaScript 的 `.docx` 解析与修改库。

当前版本已经实现：

- 文档虚拟树 `diff / patch`
- 段落与 run 的样式级建模
- 样式新增、修改、清空
- 组件之间的样式迁移
- 旧控制器 API 与新虚拟树 API 并存

所有写操作最终都会统一收敛到虚拟树 patch，再同步回底层 OOXML。

## 特性

- 解析正文、页眉、页脚、批注、脚注、尾注
- 识别 `paragraph`、`run`、`text`、`table`、`table-row`、`table-cell`、`hyperlink`、`text-box`
- 支持段落跨多个 `w:t` 的整段文本读取和回写
- 支持真正的虚拟树 `diff / patch`
- 支持段落样式和 run 样式的建模、修改和迁移
- 兼容旧控制器 API，旧写接口内部自动转为虚拟树 patch
- 保存修改后的 `.docx`

## 安装

```bash
npm install
```

## 快速开始

```js
const { loadDocx } = require("./src");

async function main() {
  const doc = await loadDocx("./sample.docx");

  doc.replaceAll("旧词", "新词");
  await doc.saveAs("./sample.modified.docx");
}

main();
```

## 虚拟树模型

文档会被解析成一棵虚拟树，典型结构如下：

```txt
document
  body
    paragraph
      run
        text
    table
      table-row
        table-cell
          paragraph
  header
    paragraph
  footer
    paragraph
  comments
    comment
      paragraph
```

目前支持的节点类型：

- `document`
- `body`
- `header`
- `footer`
- `footnotes`
- `endnotes`
- `comments`
- `paragraph`
- `run`
- `text`
- `table`
- `table-row`
- `table-cell`
- `hyperlink`
- `tab`
- `break`
- `text-box`
- `comment`
- `footnote`
- `endnote`

## 内部写入流程

无论你调用的是旧控制器接口，还是直接使用 `doc.patch(nextTree)`，内部流程都是一致的：

1. 从当前文档生成一棵新的虚拟树副本
2. 在副本上修改目标节点
3. 调用 `doc.patch(nextTree)`
4. patch 引擎执行 `INSERT / REMOVE / REPLACE / MOVE / PROPS/TEXT_UPDATE`
5. 将结果同步回底层 OOXML
6. 从 XML 重新建树并重建索引

对于段落文本修改，仍然保留当前 `ParagraphTextModel` 的策略：

- 尽量保留原有 `w:r / w:t`
- 尽量保留 `tab / break`
- 只将新的文本重新分配回原有文本节点

## 样式模型

当前已经支持两层样式建模：

- `paragraph.props.style` 对应 `w:pPr`
- `run.props.style` 对应 `w:rPr`

### 已支持的段落样式字段

```js
{
  styleId: "BodyText",
  alignment: "center",
  keepNext: true,
  keepLines: true,
  pageBreakBefore: false,
  spacing: {
    before: "120",
    after: "240",
    line: "360",
    lineRule: "auto",
  },
  indent: {
    left: "240",
    right: "120",
    firstLine: "240",
    hanging: "240",
  },
}
```

### 已支持的 run 样式字段

```js
{
  styleId: "Emphasis",
  bold: true,
  italic: true,
  underline: "single",
  color: "FF0000",
  highlight: "yellow",
  fontSize: "28",
  fontFamily: {
    ascii: "Calibri",
    hAnsi: "Calibri",
    eastAsia: "宋体",
    cs: "Arial",
  },
}
```

## 导出 API

入口定义在 [src/index.js](/d:/project/pythonProject/homework/docxEdit/src/index.js)。

```js
const {
  loadDocx,
  VirtualWordDocument,
  VNode,
  createVNode,
  cloneVNode,
  DocumentPartController,
  ParagraphController,
  RunController,
  TableController,
  TableRowController,
  TableCellController,
  TextBoxController,
  StructuredEntryController,
} = require("./src");
```

## 文档 API

### `loadDocx(input)`

加载 `.docx` 文件。

- `input: string | Buffer`
- 返回：`Promise<VirtualWordDocument>`

```js
const doc = await loadDocx("./sample.docx");
```

### `doc.toComponentTree()`

返回当前文档虚拟树的副本。你可以在这棵树上修改，再传给 `doc.patch()`。

```js
const tree = doc.toComponentTree();
console.log(tree.type); // document
```

### `doc.patch(nextTree)`

对完整虚拟树执行 patch，并把结果同步到底层 XML。

- 根节点类型必须为 `document`
- 支持文本更新、结构新增、删除、替换、重排
- 支持段落样式和 run 样式修改
- 返回 patch 结果，包含执行的操作列表

```js
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");
body.children[0].props.text = "新的第一段";

const result = doc.patch(tree);
console.log(result.operations);
```

### `doc.toBuffer()`

返回修改后的 `.docx` 二进制内容。

```js
const buffer = await doc.toBuffer();
```

### `doc.saveAs(outputPath)`

保存文档到指定路径。

```js
await doc.saveAs("./sample.modified.docx");
```

### 文档级查询接口

```js
doc.getParts();
doc.getBody();
doc.getHeaders();
doc.getFooters();
doc.getParagraphs();
doc.getParagraph(0);
doc.getTables();
doc.getTextBoxes();
doc.getFootnotes();
doc.getEndnotes();
doc.getComments();
```

### `doc.replaceAll(searchValue, replacement, options?)`

全文替换段落文本。

- `searchValue: string | RegExp`
- `replacement: string | Function`
- `options.partTypes?: string[]`

```js
doc.replaceAll("活动", "主题活动");
doc.replaceAll(/2025/g, "2026");
doc.replaceAll("页眉", "新页眉", { partTypes: ["header"] });
```

## 控制器 API

旧控制器 API 仍然保留，但内部已经迁移到虚拟树 patch。

### `DocumentPartController`

常见来源：

```js
const body = doc.getBody();
const header = doc.getHeaders()[0];
const footer = doc.getFooters()[0];
```

可用方法：

```js
body.toComponentTree();
body.getParagraphs();
body.getParagraph(0);
body.getTables();
body.getTable(0);
body.getTextBoxes();
body.replaceAll("旧词", "新词");
```

对于 `comments / footnotes / endnotes` part，还可以：

```js
const commentsPart = doc.getParts().find((part) => part.type === "comments");
commentsPart.getEntries();
commentsPart.getEntries({ includeSpecial: true });
```

### `ParagraphController`

```js
const paragraph = doc.getBody().getParagraph(0);

paragraph.getText();
paragraph.setText("新的段落内容");
paragraph.replace("旧词", "新词");
paragraph.replaceAll("青年", "青年学生");
paragraph.getStyle();
paragraph.setStyle({ alignment: "center" });
paragraph.patchStyle({ spacing: { after: "240" } });
paragraph.getRuns();
paragraph.getRun(0);
```

### `RunController`

```js
const run = doc.getBody().getParagraph(0).getRun(0);

run.getText();
run.getStyle();
run.setStyle({
  bold: true,
  color: "FF0000",
  fontSize: "28",
});
run.patchStyle({
  italic: true,
  underline: "single",
});
```

### 样式迁移

```js
const paragraphA = doc.getBody().getParagraph(0);
const paragraphB = doc.getBody().getParagraph(1);

paragraphB.copyStyleFrom(paragraphA);

const runA = paragraphA.getRun(0);
const runB = paragraphB.getRun(0);
runB.copyStyleFrom(runA);
```

### `TableController`

```js
const table = doc.getTables()[0];

table.getRows();
table.getRow(0);
table.getCell(1, 2);

table.fill(
  [
    ["活动名称", "日期", "负责人", "备注"],
    ["分享会", "2026-03-24", "张三", "已确认"],
  ],
  { startRow: 0 },
);
```

### `TableRowController`

```js
const row = doc.getTables()[0].getRow(0);

row.getCells();
row.getCell(0);
```

### `TableCellController`

```js
const cell = doc.getTables()[0].getCell(1, 0);

cell.getParagraphs();
cell.getParagraph(0);
cell.getText();
cell.setText("新的单元格内容");
```

### `TextBoxController`

```js
const textBox = doc.getTextBoxes()[0];

textBox.getParagraphs();
textBox.getText();
```

### `StructuredEntryController`

用于 `comment / footnote / endnote`。

```js
const comment = doc.getComments()[0];

comment.getParagraphs();
comment.getText();
comment.replaceAll("原文", "新文");
```

## 虚拟树 API 调用说明

### 1. 修改已有段落文本

```js
const { loadDocx } = require("./src");

const doc = await loadDocx("./sample.docx");
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");

body.children[0].props.text = "这是更新后的第一段";

await doc.patch(tree);
await doc.saveAs("./sample.modified.docx");
```

### 2. 插入一个新段落

```js
const { createVNode, loadDocx } = require("./src");

const doc = await loadDocx("./sample.docx");
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");

body.children.splice(
  1,
  0,
  createVNode({
    type: "paragraph",
    props: { text: "这是新插入的段落" },
    children: [],
  }),
);

await doc.patch(tree);
await doc.saveAs("./sample.modified.docx");
```

### 3. 删除一个段落

```js
const doc = await loadDocx("./sample.docx");
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");

body.children.splice(0, 1);

await doc.patch(tree);
```

### 4. 使用 `key` 做稳定重排

如果你要频繁重排同层节点，建议设置 `key`。

```js
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");

body.children[0].key = "first";
body.children[1].key = "second";
body.children[2].key = "third";

await doc.patch(tree);

const nextTree = doc.toComponentTree();
const nextBody = nextTree.children.find((node) => node.type === "body");
nextBody.children = [nextBody.children[2], nextBody.children[0], nextBody.children[1]];

await doc.patch(nextTree);
```

### 5. 修改段落样式

```js
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");

body.children[0].props.style = {
  styleId: "BodyText",
  alignment: "center",
  spacing: {
    before: "120",
    after: "240",
  },
};

await doc.patch(tree);
```

### 6. 修改 run 样式

```js
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");
const firstRun = body.children[0].children[0];

firstRun.props.style = {
  bold: true,
  italic: true,
  color: "FF0000",
  underline: "single",
};

await doc.patch(tree);
```

### 7. 在组件之间迁移样式

```js
const tree = doc.toComponentTree();
const body = tree.children.find((node) => node.type === "body");

const sourceParagraphStyle = body.children[0].props.style;
body.children[1].props.style = sourceParagraphStyle;

const sourceRunStyle = body.children[0].children[0].props.style;
body.children[1].children[0].props.style = sourceRunStyle;

await doc.patch(tree);
```

### 8. 修改页眉、页脚、批注、文本框

```js
const tree = doc.toComponentTree();

const header = tree.children.find((node) => node.type === "header");
const footer = tree.children.find((node) => node.type === "footer");
const comments = tree.children.find((node) => node.type === "comments");

header.children[0].props.text = "新的页眉";
footer.children[0].props.text = "新的页脚";
comments.children[0].children[0].props.text = "新的批注内容";

await doc.patch(tree);
```

## `createVNode()` 说明

`createVNode()` 用来手动创建新节点。

```js
const node = createVNode({
  type: "paragraph",
  key: "intro",
  props: { text: "介绍段落" },
  children: [],
});
```

参数说明：

- `type`: 节点类型
- `key`: 可选，同层稳定重排时推荐提供
- `props`: 节点属性
- `children`: 子节点数组

注意：

- 根节点必须是 `document`
- patch 时必须保持已有 part 不变，不能随意删除 `body/header/footer/comments` 这些 part 根
- 新增节点时，要符合当前支持的父子关系
- 样式修改建议直接写到 `paragraph.props.style` 或 `run.props.style`

## 测试

运行示例脚本：

```bash
npm run example
```

运行测试：

```bash
npm test
```

当前测试覆盖：

- 段落整段读取和回写
- `tab / break` 保留
- 虚拟树文本 patch
- 虚拟树结构插入、删除、重排
- 表格 patch 与 `fill()` 混用
- header / footer / comment / text-box 持久化
- 段落样式和 run 样式解析
- 样式新增、修改、清空
- 样式在组件之间迁移
- 真实样本文档回归

## 已知边界

- 当前只覆盖常见文本相关 OOXML 节点，不是完整的 Word OOXML 实现
- 当前样式建模主要覆盖段落和 run 的常用属性
- 对未知节点的策略是尽量保留，而不是细粒度理解和编辑
- `doc.patch(nextTree)` 期望目标树是由当前树演化而来，不保证支持任意非法结构
