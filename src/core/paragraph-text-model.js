"use strict";

const {
  childElements,
  createWordElement,
  isElement,
  setElementText,
} = require("../shared/xml");
const { replaceAllInText, replaceFirstInText } = require("./text-utils");

function isWritableToken(token) {
  return token.kind === "text";
}

function collectRunTokens(runElement, tokens) {
  for (const child of childElements(runElement)) {
    if (isElement(child, "w:t")) {
      tokens.push({
        kind: "text",
        value: child.textContent || "",
        element: child,
      });
      continue;
    }

    if (isElement(child, "w:tab")) {
      tokens.push({
        kind: "tab",
        value: "\t",
        element: child,
      });
      continue;
    }

    if (isElement(child, "w:br") || isElement(child, "w:cr")) {
      tokens.push({
        kind: "break",
        value: "\n",
        element: child,
      });
    }
  }
}

function collectParagraphTokens(node, tokens) {
  if (!node) {
    return;
  }

  if (isElement(node, "w:pPr")) {
    return;
  }

  if (isElement(node, "w:r")) {
    collectRunTokens(node, tokens);
    return;
  }

  for (const child of childElements(node)) {
    collectParagraphTokens(child, tokens);
  }
}

function groupWritableTokens(tokens) {
  const groups = [];
  let current = [];

  for (const token of tokens) {
    if (isWritableToken(token)) {
      current.push(token);
      continue;
    }

    groups.push(current);
    current = [];
  }

  groups.push(current);
  return groups;
}

function getFixedTokens(tokens) {
  return tokens.filter((token) => !isWritableToken(token));
}

function splitTextByFixedTokens(text, fixedTokens) {
  if (fixedTokens.length === 0) {
    return [text];
  }

  const chunks = [];
  let cursor = 0;

  for (const token of fixedTokens) {
    const index = text.indexOf(token.value, cursor);

    if (index === -1) {
      throw new Error(
        `The rewritten paragraph text must preserve control token "${token.value}" in its original order.`,
      );
    }

    chunks.push(text.slice(cursor, index));
    cursor = index + token.value.length;
  }

  chunks.push(text.slice(cursor));
  return chunks;
}

function distributeTextAcrossTokens(text, tokens) {
  if (tokens.length === 0) {
    if (text.length > 0) {
      throw new Error(
        "The rewritten paragraph produced text for a control-only region. This version only supports preserving control nodes in place.",
      );
    }

    return;
  }

  const originalLengths = tokens.map((token) => token.value.length);
  let cursor = 0;

  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    const nextValue =
      index === tokens.length - 1
        ? text.slice(cursor)
        : text.slice(cursor, cursor + originalLengths[index]);

    cursor += nextValue.length;
    token.value = nextValue;
    setElementText(token.element, nextValue);

    if (token.element.__vnode) {
      token.element.__vnode.props.text = nextValue;
    }
  }
}

class ParagraphTextModel {
  constructor(paragraphElement) {
    this.paragraphElement = paragraphElement;
  }

  buildTokens() {
    const tokens = [];
    collectParagraphTokens(this.paragraphElement, tokens);

    if (tokens.length === 0) {
      const runElement = createWordElement(this.paragraphElement.ownerDocument, "w:r");
      const textElement = createWordElement(this.paragraphElement.ownerDocument, "w:t");
      runElement.appendChild(textElement);
      this.paragraphElement.appendChild(runElement);

      tokens.push({
        kind: "text",
        value: "",
        element: textElement,
      });
    }

    return tokens;
  }

  getText() {
    return this.buildTokens()
      .map((token) => token.value)
      .join("");
  }

  setText(nextText) {
    const tokens = this.buildTokens();
    const fixedTokens = getFixedTokens(tokens);
    const groups = groupWritableTokens(tokens);
    const chunks = splitTextByFixedTokens(nextText, fixedTokens);

    if (groups.length !== chunks.length) {
      throw new Error("Internal paragraph grouping mismatch.");
    }

    for (let index = 0; index < groups.length; index += 1) {
      distributeTextAcrossTokens(chunks[index], groups[index]);
    }

    return this;
  }

  replaceAll(searchValue, replacement) {
    const currentText = this.getText();
    const result = replaceAllInText(currentText, searchValue, replacement);

    if (result.count > 0) {
      this.setText(result.text);
    }

    return result.count;
  }

  replace(searchValue, replacement) {
    const currentText = this.getText();
    const result = replaceFirstInText(currentText, searchValue, replacement);

    if (result.count > 0) {
      this.setText(result.text);
    }

    return result.count;
  }
}

module.exports = {
  ParagraphTextModel,
};
