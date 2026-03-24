# docx-vcomponent

一个基于 JavaScript 的 `.docx` 解析与修改原型库。它不直接让上层代码操作 OOXML，而是先把 Word 内容转换成类似前端框架虚拟 DOM 的“虚拟组件树”，再通过段落级文本模型把整段文本修改映射回底层被切碎的 `w:r / w:t` 节点。

## 设计目标

- 把 Word 文档抽象成组件树：`document -> paragraph -> run -> text`
- 上层逻辑优先操作“整段文本”，而不是直接操作零碎 run
- 当一句话被 Word 拆成多个组件时，仍然能按完整句子读取和替换
- 修改完成后，把文本重新分发回原始 OOXML 节点，尽量保留原有结构和样式边界

## 当前支持

- 读取 `word/document.xml`
- 解析段落、run、text、tab、break、table、row、cell、hyperlink
- 构建虚拟组件树
- 获取段落完整文本
- 跨多个 `w:t` 的整句替换
- 保存回新的 `.docx`

## 示例

```js
const { loadDocx } = require("./src");

async function run() {
  const doc = await loadDocx("./2活动新闻稿.docx");

  console.log(doc.getParagraph(1).getText());

  doc.replaceAll(
    "地理科学学院25级硕士地信团支部",
    "地理科学学院2025级硕士地信团支部",
  );

  await doc.saveAs("./2活动新闻稿.modified.docx");
}
```

## 核心思路

### 1. 虚拟组件树

文档加载后，会先解析为一棵组件树：

```txt
document
  paragraph
    run
      text
    run
      text
```

这样上层逻辑可以围绕“段落”“表格”“超链接”“文本片段”等语义节点工作，而不是直接写 XML 遍历代码。

### 2. 段落全文模型

`ParagraphTextModel` 会把一个段落里分散在多个 `w:t` 中的文本拼成完整字符串。例如：

```txt
["地理科学学院", "25", "级硕士地信团支", "部"]
=> "地理科学学院25级硕士地信团支部"
```

读取时你看到的是完整句子，替换时也对完整句子操作。

### 3. 回写策略

修改后的文本不会直接覆盖整个段落 XML，而是按原始文本节点顺序重新分发：

- 每个原有 `w:t` 保留
- 文本内容按原始长度切分回填
- 最后一个可写文本节点接收剩余文本
- `tab` / `break` 这类控制节点保留原位

这使得“跨 run 替换”成为可能，同时不会粗暴打平整个段落。

## 已知边界

- 当前只处理正文 `word/document.xml`
- 暂未覆盖页眉、页脚、批注、脚注、文本框、域代码等复杂区域
- 如果改写后的文本删除了段落中的控制节点顺序，比如删除了原有 `tab`，当前版本会报错而不是自动重排
- 样式继承策略目前是“尽量保留原节点并回填文本”，不是富文本级 diff

## 测试

```bash
npm test
```

## 下一步建议

- 把 `replaceAll` 扩展为基于选择器的组件操作 API，例如按段落、表格单元格、标题样式筛选
- 增加真正的 diff 层，支持像前端框架一样对虚拟树做 patch
- 扩展到 `header/footer/footnotes/comments`
- 为 run 样式建立显式模型，支持“替换文本但继承命中的样式片段”
