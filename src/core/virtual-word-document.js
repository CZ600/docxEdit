"use strict";

const fs = require("node:fs/promises");
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const { createVNode } = require("./vnode");
const { ParagraphTextModel } = require("./paragraph-text-model");
const { childElements, isElement } = require("../shared/xml");

const MAIN_DOCUMENT_PATH = "word/document.xml";

function parseNode(element, paragraphControllers) {
  if (isElement(element, "w:p")) {
    const vnode = createVNode({
      type: "paragraph",
      props: {},
      children: childElements(element)
        .filter((child) => !isElement(child, "w:pPr"))
        .map((child) => parseNode(child, paragraphControllers)),
      source: element,
    });

    const textModel = new ParagraphTextModel(element);
    const controller = new ParagraphController(vnode, textModel);
    vnode.props.text = textModel.getText();
    paragraphControllers.push(controller);
    return vnode;
  }

  if (isElement(element, "w:r")) {
    return createVNode({
      type: "run",
      props: {},
      children: childElements(element).map((child) => parseNode(child, paragraphControllers)),
      source: element,
    });
  }

  if (isElement(element, "w:t")) {
    return createVNode({
      type: "text",
      props: { text: element.textContent || "" },
      children: [],
      source: element,
    });
  }

  if (isElement(element, "w:tbl")) {
    return createVNode({
      type: "table",
      props: {},
      children: childElements(element).map((child) => parseNode(child, paragraphControllers)),
      source: element,
    });
  }

  if (isElement(element, "w:tr")) {
    return createVNode({
      type: "table-row",
      props: {},
      children: childElements(element).map((child) => parseNode(child, paragraphControllers)),
      source: element,
    });
  }

  if (isElement(element, "w:tc")) {
    return createVNode({
      type: "table-cell",
      props: {},
      children: childElements(element).map((child) => parseNode(child, paragraphControllers)),
      source: element,
    });
  }

  if (isElement(element, "w:hyperlink")) {
    return createVNode({
      type: "hyperlink",
      props: {},
      children: childElements(element).map((child) => parseNode(child, paragraphControllers)),
      source: element,
    });
  }

  if (isElement(element, "w:tab")) {
    return createVNode({
      type: "tab",
      props: { text: "\t" },
      children: [],
      source: element,
    });
  }

  if (isElement(element, "w:br") || isElement(element, "w:cr")) {
    return createVNode({
      type: "break",
      props: { text: "\n" },
      children: [],
      source: element,
    });
  }

  return createVNode({
    type: element.nodeName.replace(/^w:/, ""),
    props: {},
    children: childElements(element).map((child) => parseNode(child, paragraphControllers)),
    source: element,
  });
}

class ParagraphController {
  constructor(vnode, textModel) {
    this.vnode = vnode;
    this.textModel = textModel;
  }

  getText() {
    const text = this.textModel.getText();
    this.vnode.props.text = text;
    return text;
  }

  setText(nextText) {
    this.textModel.setText(nextText);
    this.vnode.props.text = this.textModel.getText();
    return this;
  }

  replace(searchValue, replacement) {
    const count = this.textModel.replace(searchValue, replacement);
    this.vnode.props.text = this.textModel.getText();
    return count;
  }

  replaceAll(searchValue, replacement) {
    const count = this.textModel.replaceAll(searchValue, replacement);
    this.vnode.props.text = this.textModel.getText();
    return count;
  }
}

class VirtualWordDocument {
  constructor({ zip, xmlDocument, rootVNode, paragraphControllers }) {
    this.zip = zip;
    this.xmlDocument = xmlDocument;
    this.rootVNode = rootVNode;
    this.paragraphControllers = paragraphControllers;
  }

  static async load(input) {
    const sourceBuffer = Buffer.isBuffer(input) ? input : await fs.readFile(input);
    const zip = await JSZip.loadAsync(sourceBuffer);
    const documentXml = await zip.file(MAIN_DOCUMENT_PATH).async("string");
    const xmlDocument = new DOMParser().parseFromString(documentXml, "application/xml");
    const paragraphControllers = [];
    const documentElement = xmlDocument.documentElement;
    const bodyElement = childElements(documentElement).find((node) => isElement(node, "w:body"));

    const rootVNode = createVNode({
      type: "document",
      props: {},
      children: bodyElement
        ? childElements(bodyElement).map((child) => parseNode(child, paragraphControllers))
        : [],
      source: documentElement,
    });

    return new VirtualWordDocument({
      zip,
      xmlDocument,
      rootVNode,
      paragraphControllers,
    });
  }

  toComponentTree() {
    return this.rootVNode;
  }

  getParagraphs() {
    return this.paragraphControllers;
  }

  getParagraph(index) {
    return this.paragraphControllers[index];
  }

  replaceAll(searchValue, replacement) {
    let count = 0;

    for (const paragraph of this.paragraphControllers) {
      count += paragraph.replaceAll(searchValue, replacement);
    }

    return count;
  }

  async toBuffer() {
    const xml = new XMLSerializer().serializeToString(this.xmlDocument);
    this.zip.file(MAIN_DOCUMENT_PATH, xml);
    return this.zip.generateAsync({ type: "nodebuffer" });
  }

  async saveAs(outputPath) {
    const buffer = await this.toBuffer();
    await fs.writeFile(outputPath, buffer);
    return outputPath;
  }
}

module.exports = {
  MAIN_DOCUMENT_PATH,
  ParagraphController,
  VirtualWordDocument,
};
