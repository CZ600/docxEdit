"use strict";

const { VirtualWordDocument } = require("./core/virtual-word-document");

async function loadDocx(input) {
  return VirtualWordDocument.load(input);
}

module.exports = {
  VirtualWordDocument,
  loadDocx,
};
