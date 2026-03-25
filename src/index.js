"use strict";

const { VirtualWordDocument } = require("./core/virtual-word-document");
const {
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
} = require("./core/document-part");
const { VNode, cloneVNode, createVNode } = require("./core/vnode");

async function loadDocx(input) {
  return VirtualWordDocument.load(input);
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
  VNode,
  VirtualWordDocument,
  cloneVNode,
  createVNode,
  loadDocx,
};
