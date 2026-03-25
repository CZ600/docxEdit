"use strict";

const { DOMParser } = require("@xmldom/xmldom");
const { ParagraphTextModel } = require("./paragraph-text-model");
const {
  cloneStyle,
  parseParagraphStyle,
  parseRunStyle,
  parseTableCellStyle,
  parseTableRowStyle,
  parseTableStyle,
} = require("./style-model");
const { replaceAllInText, replaceFirstInText } = require("./text-utils");
const { createVNode } = require("./vnode");
const { formatMathPlaceholder, isMathElement, parseMathNode } = require("./math-model");
const { childElements, isElement } = require("../shared/xml");

function getWordAttribute(element, localName) {
  return element.getAttribute(`w:${localName}`) || element.getAttribute(localName) || null;
}

function buildLocation(partType, ancestors) {
  if (ancestors.includes("w:txbxContent")) {
    return "textbox";
  }
  if (ancestors.includes("w:comment")) {
    return "comment";
  }
  if (ancestors.includes("w:footnote")) {
    return "footnote";
  }
  if (ancestors.includes("w:endnote")) {
    return "endnote";
  }
  if (ancestors.includes("w:tc")) {
    return "table-cell";
  }
  return partType;
}

function getNodeType(element) {
  if (isElement(element, "w:p")) return "paragraph";
  if (isElement(element, "w:r")) return "run";
  if (isElement(element, "w:t")) return "text";
  if (isElement(element, "w:tbl")) return "table";
  if (isElement(element, "w:tr")) return "table-row";
  if (isElement(element, "w:tc")) return "table-cell";
  if (isElement(element, "w:hyperlink")) return "hyperlink";
  if (isElement(element, "w:tab")) return "tab";
  if (isElement(element, "w:br") || isElement(element, "w:cr")) return "break";
  if (isElement(element, "w:txbxContent")) return "text-box";
  if (isElement(element, "w:comment")) return "comment";
  if (isElement(element, "w:footnote")) return "footnote";
  if (isElement(element, "w:endnote")) return "endnote";
  if (isElement(element, "w:drawing")) return "image";
  if (isMathElement(element)) return "math";
  return element.nodeName.replace(/^w:/, "");
}

class BaseController {
  constructor(doc, nodeId) {
    this.doc = doc;
    this.nodeId = nodeId;
  }

  get vnode() {
    return this.doc.getNodeById(this.nodeId);
  }

  get metadata() {
    return this.doc.getNodeMetadata(this.nodeId);
  }
}

class ParagraphController extends BaseController {
  getText() {
    return this.vnode.props.text || "";
  }

  setText(nextText) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const paragraph = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(paragraph, this.nodeId, "paragraph");
      paragraph.props.text = nextText;
    });
    return this;
  }

  replace(searchValue, replacement) {
    let count = 0;
    this.doc.patchWithMutableTree((nextRoot) => {
      const paragraph = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(paragraph, this.nodeId, "paragraph");
      const result = replaceFirstInText(paragraph.props.text || "", searchValue, replacement);
      count = result.count;
      if (count > 0) paragraph.props.text = result.text;
    }, { skipIfUnchanged: () => count === 0 });
    return count;
  }

  replaceAll(searchValue, replacement) {
    let count = 0;
    this.doc.patchWithMutableTree((nextRoot) => {
      const paragraph = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(paragraph, this.nodeId, "paragraph");
      const result = replaceAllInText(paragraph.props.text || "", searchValue, replacement);
      count = result.count;
      if (count > 0) paragraph.props.text = result.text;
    }, { skipIfUnchanged: () => count === 0 });
    return count;
  }

  getStyle() {
    return cloneStyle(this.vnode.props.style || {});
  }

  setStyle(nextStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const paragraph = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(paragraph, this.nodeId, "paragraph");
      paragraph.props.style = cloneStyle(nextStyle || {});
    });
    return this;
  }

  patchStyle(partialStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const paragraph = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(paragraph, this.nodeId, "paragraph");
      paragraph.props.style = mergeStyles(paragraph.props.style || {}, partialStyle || {});
    });
    return this;
  }

  copyStyleFrom(otherParagraph) {
    return this.setStyle(otherParagraph.getStyle());
  }

  getRuns() {
    return this.doc.getDescendantControllers(this.nodeId, "run");
  }

  getRun(index) {
    return this.getRuns()[index];
  }

  getMaths() {
    return this.doc.getDescendantControllers(this.nodeId, "math");
  }

  getMath(index) {
    return this.getMaths()[index];
  }
}

class TableCellController extends BaseController {
  getParagraphs() {
    return this.doc.getDescendantControllers(this.nodeId, "paragraph");
  }

  getParagraph(index) {
    return this.getParagraphs()[index];
  }

  getText({ separator = "\n" } = {}) {
    return this.getParagraphs().map((paragraph) => paragraph.getText()).join(separator);
  }

  setText(nextText) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const cell = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(cell, this.nodeId, "table-cell");
      let paragraphs = findDescendantNodesByType(cell, "paragraph");
      if (paragraphs.length === 0) {
        const paragraph = createVNode({ type: "paragraph", props: { text: "" }, children: [] });
        cell.children.push(paragraph);
        paragraph.parent = cell;
        paragraphs = [paragraph];
      }
      paragraphs[0].props.text = nextText;
      for (let index = 1; index < paragraphs.length; index += 1) {
        paragraphs[index].props.text = "";
      }
    });
    return this;
  }

  getStyle() {
    return cloneStyle(this.vnode.props.style || {});
  }

  setStyle(nextStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const cell = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(cell, this.nodeId, "table-cell");
      cell.props.style = cloneStyle(nextStyle || {});
    });
    return this;
  }

  patchStyle(partialStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const cell = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(cell, this.nodeId, "table-cell");
      cell.props.style = mergeStyles(cell.props.style || {}, partialStyle || {});
    });
    return this;
  }

  copyStyleFrom(otherCell) {
    return this.setStyle(otherCell.getStyle());
  }
}

class TableRowController extends BaseController {
  getCells() {
    return this.doc.getDirectChildControllers(this.nodeId, "table-cell");
  }

  getCell(index) {
    return this.getCells()[index];
  }

  getStyle() {
    return cloneStyle(this.vnode.props.style || {});
  }

  setStyle(nextStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const row = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(row, this.nodeId, "table-row");
      row.props.style = cloneStyle(nextStyle || {});
    });
    return this;
  }

  patchStyle(partialStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const row = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(row, this.nodeId, "table-row");
      row.props.style = mergeStyles(row.props.style || {}, partialStyle || {});
    });
    return this;
  }

  copyStyleFrom(otherRow) {
    return this.setStyle(otherRow.getStyle());
  }
}

class RunController extends BaseController {
  getText() {
    return collectInlineText(this.vnode);
  }

  getStyle() {
    return cloneStyle(this.vnode.props.style || {});
  }

  setStyle(nextStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const run = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(run, this.nodeId, "run");
      run.props.style = cloneStyle(nextStyle || {});
    });
    return this;
  }

  patchStyle(partialStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const run = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(run, this.nodeId, "run");
      run.props.style = mergeStyles(run.props.style || {}, partialStyle || {});
    });
    return this;
  }

  copyStyleFrom(otherRun) {
    return this.setStyle(otherRun.getStyle());
  }

  getImages() {
    return this.doc.getDescendantControllers(this.nodeId, "image");
  }

  getImage(index) {
    return this.getImages()[index];
  }

  getMaths() {
    return this.doc.getDescendantControllers(this.nodeId, "math");
  }

  getMath(index) {
    return this.getMaths()[index];
  }
}

class ImageController extends BaseController {
  getRelId() {
    return this.vnode.props.relId || null;
  }

  getFilename() {
    return this.vnode.props.filename || null;
  }

  getContentType() {
    return this.vnode.props.contentType || null;
  }

  getSize() {
    return { width: this.vnode.props.width || null, height: this.vnode.props.height || null };
  }

  getLayout() {
    return cloneStyle(this.vnode.props.layout || {});
  }

  async getData() {
    return this.doc.getImageData(this.nodeId);
  }

  replace({ data, filename, contentType, width, height, alt, layout, paragraphAlignment }) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const image = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(image, this.nodeId, "image");
      if (data !== undefined) image.props.data = data;
      if (filename !== undefined) image.props.filename = filename;
      if (contentType !== undefined) image.props.contentType = contentType;
      if (width !== undefined) image.props.width = width;
      if (height !== undefined) image.props.height = height;
      if (alt !== undefined) image.props.alt = alt;
      if (layout !== undefined) image.props.layout = cloneStyle(layout || {});
      if (paragraphAlignment !== undefined) {
        const paragraph = findAncestorNodeByType(image, "paragraph");
        if (paragraph) {
          paragraph.props.style = mergeStyles(paragraph.props.style || {}, { alignment: paragraphAlignment });
        }
      }
    });
    return this;
  }

  setLayout(nextLayout) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const image = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(image, this.nodeId, "image");
      image.props.layout = cloneStyle(nextLayout || {});
    });
    return this;
  }

  patchLayout(partialLayout) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const image = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(image, this.nodeId, "image");
      image.props.layout = mergeStyles(image.props.layout || {}, partialLayout || {});
    });
    return this;
  }

  setParagraphAlignment(alignment) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const image = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(image, this.nodeId, "image");
      const paragraph = findAncestorNodeByType(image, "paragraph");
      if (!paragraph) return;
      paragraph.props.style = mergeStyles(paragraph.props.style || {}, { alignment });
    });
    return this;
  }

  remove() {
    this.doc.patchWithMutableTree((nextRoot) => {
      removeNodeById(nextRoot, this.nodeId);
    });
    return this;
  }
}

class MathController extends BaseController {
  getText() {
    return this.vnode.props.text || "";
  }

  setText(nextText) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const math = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(math, this.nodeId, "math");
      math.props.text = String(nextText || "");
      math.props.placeholder = formatMathPlaceholder(math.props.text);
    });
    return this;
  }

  getDisplay() {
    return this.vnode.props.display || "inline";
  }

  setDisplay(nextDisplay) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const math = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(math, this.nodeId, "math");
      math.props.display = nextDisplay === "block" ? "block" : "inline";
    });
    return this;
  }

  getStyle() {
    return cloneStyle(this.vnode.props.style || {});
  }

  setStyle(nextStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const math = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(math, this.nodeId, "math");
      math.props.style = cloneStyle(nextStyle || {});
    });
    return this;
  }

  patchStyle(partialStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const math = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(math, this.nodeId, "math");
      math.props.style = mergeStyles(math.props.style || {}, partialStyle || {});
    });
    return this;
  }

  replace({ text, display, style }) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const math = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(math, this.nodeId, "math");
      if (text !== undefined) {
        math.props.text = String(text || "");
        math.props.placeholder = formatMathPlaceholder(math.props.text);
      }
      if (display !== undefined) {
        math.props.display = display === "block" ? "block" : "inline";
      }
      if (style !== undefined) {
        math.props.style = cloneStyle(style || {});
      }
    });
    return this;
  }
}

class TableController extends BaseController {
  getStyle() {
    return cloneStyle(this.vnode.props.style || {});
  }

  setStyle(nextStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const table = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(table, this.nodeId, "table");
      table.props.style = cloneStyle(nextStyle || {});
    });
    return this;
  }

  patchStyle(partialStyle) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const table = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(table, this.nodeId, "table");
      table.props.style = mergeStyles(table.props.style || {}, partialStyle || {});
    });
    return this;
  }

  copyStyleFrom(otherTable) {
    return this.setStyle(otherTable.getStyle());
  }

  getRows() {
    return this.doc.getDirectChildControllers(this.nodeId, "table-row");
  }

  getRow(index) {
    return this.getRows()[index];
  }

  getCell(rowIndex, cellIndex) {
    const row = this.getRow(rowIndex);
    return row ? row.getCell(cellIndex) : undefined;
  }

  fill(rows, { startRow = 0 } = {}) {
    this.doc.patchWithMutableTree((nextRoot) => {
      const table = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(table, this.nodeId, "table");
      const tableRows = findDirectChildNodesByType(table, "table-row");
      for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
        const row = tableRows[startRow + rowOffset];
        if (!row) continue;
        const cells = findDirectChildNodesByType(row, "table-cell");
        const rowValues = rows[rowOffset];
        for (let cellIndex = 0; cellIndex < rowValues.length; cellIndex += 1) {
          const cell = cells[cellIndex];
          if (!cell) continue;
          let paragraphs = findDescendantNodesByType(cell, "paragraph");
          if (paragraphs.length === 0) {
            const paragraph = createVNode({ type: "paragraph", props: { text: "" }, children: [] });
            cell.children.push(paragraph);
            paragraph.parent = cell;
            paragraphs = [paragraph];
          }
          paragraphs[0].props.text = String(rowValues[cellIndex] ?? "");
          for (let index = 1; index < paragraphs.length; index += 1) {
            paragraphs[index].props.text = "";
          }
        }
      }
    });
    return this;
  }
}

class TextBoxController extends BaseController {
  getParagraphs() {
    return this.doc.getDescendantControllers(this.nodeId, "paragraph");
  }

  getText({ separator = "\n" } = {}) {
    return this.getParagraphs().map((paragraph) => paragraph.getText()).join(separator);
  }
}

class StructuredEntryController extends BaseController {
  getParagraphs() {
    return this.doc.getDescendantControllers(this.nodeId, "paragraph");
  }

  getText({ separator = "\n" } = {}) {
    return this.getParagraphs().map((paragraph) => paragraph.getText()).join(separator);
  }

  replaceAll(searchValue, replacement) {
    let count = 0;
    this.doc.patchWithMutableTree((nextRoot) => {
      const entry = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(entry, this.nodeId, this.vnode.type);
      const paragraphs = findDescendantNodesByType(entry, "paragraph");
      for (const paragraph of paragraphs) {
        const result = replaceAllInText(paragraph.props.text || "", searchValue, replacement);
        if (result.count > 0) {
          paragraph.props.text = result.text;
          count += result.count;
        }
      }
    }, { skipIfUnchanged: () => count === 0 });
    return count;
  }
}

class DocumentPartController extends BaseController {
  constructor(doc, nodeId, path, type) {
    super(doc, nodeId);
    this.path = path;
    this.type = type;
  }

  toComponentTree() {
    return this.doc.cloneNodeById(this.nodeId);
  }

  getParagraphs() {
    return this.doc.getDescendantControllers(this.nodeId, "paragraph");
  }

  getParagraph(index) {
    return this.getParagraphs()[index];
  }

  getTables() {
    return this.doc.getDescendantControllers(this.nodeId, "table");
  }

  getTable(index) {
    return this.getTables()[index];
  }

  getTextBoxes() {
    return this.doc.getDescendantControllers(this.nodeId, "text-box");
  }

  getImages() {
    return this.doc.getDescendantControllers(this.nodeId, "image");
  }

  getEntries({ includeSpecial = false } = {}) {
    const entryType = getEntryTypeForPart(this.type);
    if (!entryType) return [];
    const entries = this.doc.getDirectChildControllers(this.nodeId, entryType);
    return includeSpecial ? entries : entries.filter((entry) => !entry.metadata.specialType);
  }

  replaceAll(searchValue, replacement) {
    let count = 0;
    this.doc.patchWithMutableTree((nextRoot) => {
      const partNode = findNodeById(nextRoot, this.nodeId);
      ensureExistingNode(partNode, this.nodeId, this.type);
      const paragraphs = findDescendantNodesByType(partNode, "paragraph");
      for (const paragraph of paragraphs) {
        const result = replaceAllInText(paragraph.props.text || "", searchValue, replacement);
        if (result.count > 0) {
          paragraph.props.text = result.text;
          count += result.count;
        }
      }
    }, { skipIfUnchanged: () => count === 0 });
    return count;
  }
}

function getEntryTypeForPart(partType) {
  if (partType === "comments") return "comment";
  if (partType === "footnotes") return "footnote";
  if (partType === "endnotes") return "endnote";
  return null;
}

function getPartRootElement(xmlDocument, type) {
  const documentElement = xmlDocument.documentElement;
  if (type === "body") {
    return childElements(documentElement).find((node) => isElement(node, "w:body")) || documentElement;
  }
  return documentElement;
}

function parseNode(element, context, ancestors = []) {
  if (!element || element.nodeType !== 1) return null;

  const nextAncestors = ancestors.concat(element.nodeName);

  if (isElement(element, "w:drawing")) {
    const imageVNode = parseImageNode(element, context);
    if (imageVNode) {
      registerMetadata(context, imageVNode, ancestors);
      return imageVNode;
    }
  }

  if (isMathElement(element)) {
    const mathProps = parseMathNode(element);
    const mathVNode = createVNode({
      id: element.__vnodeId || null,
      type: "math",
      props: {
        text: mathProps.text,
        display: mathProps.display,
        style: mathProps.style,
        placeholder: formatMathPlaceholder(mathProps.text),
      },
      children: [],
      source: element,
    });
    registerMetadata(context, mathVNode, ancestors);
    return mathVNode;
  }

  if (isElement(element, "w:p")) {
    const textModel = new ParagraphTextModel(element);
    const paragraphVNode = createVNode({
      id: element.__vnodeId || null,
      type: "paragraph",
      props: { style: parseParagraphStyle(element), text: textModel.getText() },
      children: childElements(element)
        .filter((child) => !isElement(child, "w:pPr"))
        .map((child) => parseNode(child, context, nextAncestors))
        .filter(Boolean),
      source: element,
    });
    registerMetadata(context, paragraphVNode, ancestors);
    return paragraphVNode;
  }

  const vnode = createVNode({
    id: element.__vnodeId || null,
    type: getNodeType(element),
    props: {},
    children: childElements(element)
      .filter((child) => !(isElement(element, "w:r") && isElement(child, "w:rPr")))
      .filter((child) => !(isElement(element, "w:tbl") && isElement(child, "w:tblPr")))
      .filter((child) => !(isElement(element, "w:tr") && isElement(child, "w:trPr")))
      .filter((child) => !(isElement(element, "w:tc") && isElement(child, "w:tcPr")))
      .map((child) => parseNode(child, context, nextAncestors))
      .filter(Boolean),
    source: element,
  });

  if (isElement(element, "w:r")) {
    vnode.props.style = parseRunStyle(element);
  }
  if (isElement(element, "w:tbl")) {
    vnode.props.style = parseTableStyle(element);
  }
  if (isElement(element, "w:tr")) {
    vnode.props.style = parseTableRowStyle(element);
  }
  if (isElement(element, "w:tc")) {
    vnode.props.style = parseTableCellStyle(element);
  }
  if (isElement(element, "w:t")) {
    vnode.props.text = element.textContent || "";
  }
  if (isElement(element, "w:tab")) {
    vnode.props.text = "\t";
  }
  if (isElement(element, "w:br") || isElement(element, "w:cr")) {
    vnode.props.text = "\n";
  }
  if (isElement(element, "w:comment") || isElement(element, "w:footnote") || isElement(element, "w:endnote")) {
    vnode.props.id = getWordAttribute(element, "id");
    vnode.props.specialType = getWordAttribute(element, "type");
  }

  registerMetadata(context, vnode, ancestors);
  return vnode;
}

function parseImageNode(element, context) {
  const blip = element.getElementsByTagName("a:blip")[0];
  if (!blip) {
    return null;
  }

  const relId = blip.getAttribute("r:embed") || blip.getAttribute("embed") || blip.getAttribute("r:link") || null;
  if (!relId) {
    return null;
  }

  const descriptor = context.doc.describeImage(context.partPath, relId);
  const extent = element.getElementsByTagName("wp:extent")[0] || null;
  const docPr = element.getElementsByTagName("wp:docPr")[0] || null;
  const cNvPr = element.getElementsByTagName("pic:cNvPr")[0] || null;
  const inline = element.getElementsByTagName("wp:inline")[0] || null;
  const anchor = element.getElementsByTagName("wp:anchor")[0] || null;
  const container = anchor || inline || null;
  const name = (docPr && docPr.getAttribute("name")) || (cNvPr && cNvPr.getAttribute("name")) || null;
  const alt = (docPr && (docPr.getAttribute("descr") || docPr.getAttribute("title"))) ||
    (cNvPr && (cNvPr.getAttribute("descr") || cNvPr.getAttribute("title"))) ||
    "";

  return createVNode({
    id: element.__vnodeId || null,
    type: "image",
    props: {
      relId,
      filename: name || (descriptor ? descriptor.filename : null),
      contentType: descriptor ? descriptor.contentType : null,
      target: descriptor ? descriptor.target : null,
      mediaPath: descriptor ? descriptor.mediaPath : null,
      width: extent ? extent.getAttribute("cx") : null,
      height: extent ? extent.getAttribute("cy") : null,
      alt,
      layout: parseImageLayout(container),
    },
    children: [],
    source: element,
  });
}

function parseImageLayout(container) {
  if (!container) {
    return { mode: "inline" };
  }

  const layout = {
    mode: isElement(container, "wp:anchor") ? "anchor" : "inline",
  };

  if (layout.mode === "anchor") {
    layout.wrap = parseWrap(container);
    layout.distances = compactObject({
      top: container.getAttribute("distT") || null,
      bottom: container.getAttribute("distB") || null,
      left: container.getAttribute("distL") || null,
      right: container.getAttribute("distR") || null,
    });
    layout.behindDoc = parseBooleanAttribute(container.getAttribute("behindDoc"));
    layout.allowOverlap = parseBooleanAttribute(container.getAttribute("allowOverlap"));
    layout.layoutInCell = parseBooleanAttribute(container.getAttribute("layoutInCell"));

    const positionH = container.getElementsByTagName("wp:positionH")[0] || null;
    const positionV = container.getElementsByTagName("wp:positionV")[0] || null;
    if (positionH) {
      layout.positionH = parsePosition(positionH);
    }
    if (positionV) {
      layout.positionV = parsePosition(positionV);
    }
  }

  return compactObject(layout);
}

function parseWrap(container) {
  if (container.getElementsByTagName("wp:wrapNone")[0]) return "none";
  if (container.getElementsByTagName("wp:wrapSquare")[0]) return "square";
  if (container.getElementsByTagName("wp:wrapTight")[0]) return "tight";
  if (container.getElementsByTagName("wp:wrapTopAndBottom")[0]) return "topAndBottom";
  return "none";
}

function parsePosition(positionElement) {
  const align = positionElement.getElementsByTagName("wp:align")[0];
  const posOffset = positionElement.getElementsByTagName("wp:posOffset")[0];
  return compactObject({
    relativeFrom: positionElement.getAttribute("relativeFrom") || null,
    align: align ? align.textContent : null,
    offset: posOffset ? posOffset.textContent : null,
  });
}

function parseBooleanAttribute(value) {
  if (value == null || value === "") return null;
  return !["0", "false", "off"].includes(String(value).toLowerCase());
}

function compactObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return value;
  }

  const result = {};
  for (const [key, innerValue] of Object.entries(value)) {
    if (innerValue == null) continue;
    if (typeof innerValue === "object" && !Array.isArray(innerValue)) {
      const nested = compactObject(innerValue);
      if (nested && Object.keys(nested).length > 0) result[key] = nested;
      continue;
    }
    result[key] = innerValue;
  }
  return result;
}

function findAncestorNodeByType(node, type) {
  let current = node ? node.parent : null;
  while (current) {
    if (current.type === type) return current;
    current = current.parent;
  }
  return null;
}

function registerMetadata(context, vnode, ancestors) {
  context.metadataById.set(vnode.id, {
    partPath: context.partPath,
    partType: context.partType,
    location: buildLocation(context.partType, ancestors),
    specialType: vnode.props.specialType || null,
  });
}

function buildPartController({ doc, path, type, xmlDocument, metadataById }) {
  const rootElement = getPartRootElement(xmlDocument, type);
  const context = { doc, partPath: path, partType: type, metadataById };
  const rootVNode = createVNode({
    id: rootElement.__vnodeId || null,
    type,
    props: { path, type },
    children: childElements(rootElement).map((child) => parseNode(child, context)).filter(Boolean),
    source: rootElement,
  });

  metadataById.set(rootVNode.id, {
    partPath: path,
    partType: type,
    location: type,
    specialType: null,
  });

  return { path, type, rootVNode };
}

function parseXmlString(xml) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function findNodeById(root, nodeId) {
  if (!root) return null;
  if (root.id === nodeId) return root;
  for (const child of root.children) {
    const found = findNodeById(child, nodeId);
    if (found) return found;
  }
  return null;
}

function removeNodeById(root, nodeId) {
  for (let index = 0; index < root.children.length; index += 1) {
    const child = root.children[index];
    if (child.id === nodeId) {
      root.children.splice(index, 1);
      return true;
    }
    if (removeNodeById(child, nodeId)) return true;
  }
  return false;
}

function findDirectChildNodesByType(node, type) {
  return node.children.filter((child) => child.type === type);
}

function findDescendantNodesByType(node, type, result = []) {
  for (const child of node.children) {
    if (child.type === type) result.push(child);
    findDescendantNodesByType(child, type, result);
  }
  return result;
}

function ensureExistingNode(node, nodeId, expectedType) {
  if (!node) throw new Error(`Node ${nodeId} is no longer available.`);
  if (expectedType && node.type !== expectedType) {
    throw new Error(`Node ${nodeId} is not a ${expectedType}.`);
  }
}

function mergeStyles(baseStyle, partialStyle) {
  const merged = cloneStyle(baseStyle || {});
  for (const [key, value] of Object.entries(partialStyle || {})) {
    if (value == null) {
      delete merged[key];
      continue;
    }
    if (value && typeof value === "object" && !Array.isArray(value)) {
      merged[key] = mergeStyles(merged[key] || {}, value);
      if (Object.keys(merged[key]).length === 0) delete merged[key];
      continue;
    }
    merged[key] = value;
  }
  return merged;
}

function collectInlineText(node) {
  let text = "";
  for (const child of node.children) {
    if (child.type === "text" || child.type === "tab" || child.type === "break") {
      text += child.props.text || "";
      continue;
    }
    if (child.type === "math") {
      text += child.props.placeholder || formatMathPlaceholder(child.props.text || "");
      continue;
    }
    text += collectInlineText(child);
  }
  return text;
}

module.exports = {
  DocumentPartController,
  ImageController,
  MathController,
  ParagraphController,
  RunController,
  StructuredEntryController,
  TableCellController,
  TableController,
  TableRowController,
  TextBoxController,
  buildPartController,
  buildLocation,
  ensureExistingNode,
  findDescendantNodesByType,
  findDirectChildNodesByType,
  findNodeById,
  getEntryTypeForPart,
  getNodeType,
  getPartRootElement,
  mergeStyles,
  parseXmlString,
  removeNodeById,
};
