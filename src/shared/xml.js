"use strict";

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function isElement(node, qualifiedName) {
  return Boolean(node && node.nodeType === 1 && node.nodeName === qualifiedName);
}

function childElements(node) {
  const result = [];

  if (!node || !node.childNodes) {
    return result;
  }

  for (let i = 0; i < node.childNodes.length; i += 1) {
    const child = node.childNodes[i];
    if (child.nodeType === 1) {
      result.push(child);
    }
  }

  return result;
}

function setElementText(element, text) {
  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }

  element.appendChild(element.ownerDocument.createTextNode(text));

  if (/^\s|\s$/.test(text)) {
    element.setAttribute("xml:space", "preserve");
  } else {
    element.removeAttribute("xml:space");
  }
}

function createWordElement(documentNode, name) {
  return documentNode.createElementNS(WORD_NS, name);
}

module.exports = {
  WORD_NS,
  childElements,
  createWordElement,
  isElement,
  setElementText,
};
