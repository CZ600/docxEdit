"use strict";

const { childElements, createWordElement, setElementText } = require("../shared/xml");
const { cloneStyle } = require("./style-model");

const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

function isMathElement(node) {
  return Boolean(node && node.nodeType === 1 && (node.nodeName === "m:oMath" || node.nodeName === "m:oMathPara"));
}

function createMathElement(documentNode, name) {
  return documentNode.createElementNS(MATH_NS, name);
}

function collectMathText(node) {
  if (!node) {
    return "";
  }

  let text = "";
  if (node.nodeType === 1 && node.nodeName === "m:t") {
    text += node.textContent || "";
  }

  if (node.childNodes) {
    for (let index = 0; index < node.childNodes.length; index += 1) {
      text += collectMathText(node.childNodes[index]);
    }
  }

  return text;
}

function formatMathPlaceholder(text) {
  return `[[MATH:${text || ""}]]`;
}

function parseMathNode(element) {
  return {
    text: collectMathText(element),
    display: element.nodeName === "m:oMathPara" ? "block" : "inline",
    style: parseMathStyle(element),
  };
}

function parseMathStyle(element) {
  const style = {};
  const runStyle = parseFirstMathRunStyle(element);
  const justification = parseMathJustification(element);

  if (justification) {
    style.justification = justification;
  }

  return Object.assign(style, runStyle);
}

function parseMathJustification(element) {
  if (!element || element.nodeName !== "m:oMathPara") {
    return null;
  }

  const paraProps = childElements(element).find((child) => child.nodeName === "m:oMathParaPr");
  if (!paraProps) {
    return null;
  }

  const jc = childElements(paraProps).find((child) => child.nodeName === "m:jc");
  return jc ? getAttribute(jc, "val") : null;
}

function parseFirstMathRunStyle(element) {
  const run = findFirstDescendant(element, "m:r");
  if (!run) {
    return {};
  }

  const mathRunProps = childElements(run).find((child) => child.nodeName === "m:rPr");
  if (!mathRunProps) {
    return {};
  }

  const wordRunProps = childElements(mathRunProps).find((child) => child.nodeName === "w:rPr");
  if (!wordRunProps) {
    return {};
  }

  return parseWordRunProperties(wordRunProps);
}

function parseWordRunProperties(rPr) {
  const style = {};
  const rStyle = childElements(rPr).find((child) => child.nodeName === "w:rStyle");
  const underline = childElements(rPr).find((child) => child.nodeName === "w:u");
  const color = childElements(rPr).find((child) => child.nodeName === "w:color");
  const highlight = childElements(rPr).find((child) => child.nodeName === "w:highlight");
  const size = childElements(rPr).find((child) => child.nodeName === "w:sz");
  const fonts = childElements(rPr).find((child) => child.nodeName === "w:rFonts");

  if (rStyle) {
    style.styleId = getAttribute(rStyle, "val");
  }

  assignOnOff(style, "bold", childElements(rPr).find((child) => child.nodeName === "w:b"));
  assignOnOff(style, "italic", childElements(rPr).find((child) => child.nodeName === "w:i"));

  if (underline) {
    style.underline = getAttribute(underline, "val") || "single";
  }

  if (color) {
    style.color = getAttribute(color, "val");
  }

  if (highlight) {
    style.highlight = getAttribute(highlight, "val");
  }

  if (size) {
    style.fontSize = getAttribute(size, "val");
  }

  if (fonts) {
    style.fontFamily = compactObject({
      ascii: getAttribute(fonts, "ascii"),
      hAnsi: getAttribute(fonts, "hAnsi"),
      eastAsia: getAttribute(fonts, "eastAsia"),
      cs: getAttribute(fonts, "cs"),
    });
  }

  return compactObject(style);
}

function updateMathElement(element, props) {
  while (element.firstChild) {
    element.removeChild(element.firstChild);
  }

  const normalizedProps = normalizeMathProps(props);
  const doc = element.ownerDocument;

  if (element.nodeName === "m:oMathPara") {
    if (normalizedProps.style.justification) {
      const paraProps = createMathElement(doc, "m:oMathParaPr");
      const jc = createMathElement(doc, "m:jc");
      jc.setAttribute("m:val", normalizedProps.style.justification);
      paraProps.appendChild(jc);
      element.appendChild(paraProps);
    }
    element.appendChild(createMathCore(doc, normalizedProps));
    return;
  }

  const math = createMathCore(doc, normalizedProps);
  while (math.firstChild) {
    element.appendChild(math.firstChild);
  }
}

function createMathCore(doc, props) {
  const math = createMathElement(doc, "m:oMath");
  const run = createMathElement(doc, "m:r");
  const runProps = createMathRunProperties(doc, props.style);

  if (runProps) {
    run.appendChild(runProps);
  }

  const text = createMathElement(doc, "m:t");
  setElementText(text, props.text);
  run.appendChild(text);
  math.appendChild(run);
  return math;
}

function createMathRunProperties(doc, style) {
  const normalizedStyle = normalizeMathStyle(style);
  const wordStyle = Object.assign({}, normalizedStyle);
  delete wordStyle.justification;

  if (Object.keys(wordStyle).length === 0) {
    return null;
  }

  const mathRunProps = createMathElement(doc, "m:rPr");
  const wordRunProps = buildWordRunProperties(doc, wordStyle);

  if (!wordRunProps) {
    return null;
  }

  mathRunProps.appendChild(wordRunProps);
  return mathRunProps;
}

function buildWordRunProperties(doc, style) {
  const normalized = compactObject(cloneStyle(style || {}));
  if (!normalized || Object.keys(normalized).length === 0) {
    return null;
  }

  const rPr = createWordElement(doc, "w:rPr");

  appendValueChild(rPr, "w:rStyle", normalized.styleId);
  appendOnOffChild(rPr, "w:b", normalized.bold);
  appendOnOffChild(rPr, "w:i", normalized.italic);
  appendValueChild(rPr, "w:u", normalized.underline);
  appendValueChild(rPr, "w:color", normalized.color);
  appendValueChild(rPr, "w:highlight", normalized.highlight);
  appendValueChild(rPr, "w:sz", normalized.fontSize);
  appendAttributeChild(rPr, "w:rFonts", normalized.fontFamily, ["ascii", "hAnsi", "eastAsia", "cs"]);

  return childElements(rPr).length > 0 ? rPr : null;
}

function normalizeMathProps(props) {
  return {
    text: String((props && props.text) || ""),
    display: props && props.display === "block" ? "block" : "inline",
    style: normalizeMathStyle(props && props.style),
  };
}

function normalizeMathStyle(style) {
  return compactObject(cloneStyle(style || {}));
}

function findFirstDescendant(node, nodeName) {
  if (!node || !node.childNodes) {
    return null;
  }

  for (let index = 0; index < node.childNodes.length; index += 1) {
    const child = node.childNodes[index];
    if (child.nodeType !== 1) {
      continue;
    }
    if (child.nodeName === nodeName) {
      return child;
    }
    const found = findFirstDescendant(child, nodeName);
    if (found) {
      return found;
    }
  }

  return null;
}

function appendValueChild(container, childName, value) {
  if (value == null) {
    return;
  }

  const child = createWordElement(container.ownerDocument, childName);
  child.setAttribute("w:val", String(value));
  container.appendChild(child);
}

function appendOnOffChild(container, childName, value) {
  if (value == null) {
    return;
  }

  const child = createWordElement(container.ownerDocument, childName);
  child.setAttribute("w:val", value ? "1" : "0");
  container.appendChild(child);
}

function appendAttributeChild(container, childName, value, attributeNames) {
  if (!value || Object.keys(value).length === 0) {
    return;
  }

  const child = createWordElement(container.ownerDocument, childName);
  for (const attributeName of attributeNames) {
    if (value[attributeName] == null) {
      continue;
    }
    child.setAttribute(`w:${attributeName}`, String(value[attributeName]));
  }

  if (child.attributes.length > 0) {
    container.appendChild(child);
  }
}

function getAttribute(element, localName) {
  return element.getAttribute(`m:${localName}`) ||
    element.getAttribute(`w:${localName}`) ||
    element.getAttribute(localName) ||
    null;
}

function assignOnOff(style, key, element) {
  if (!element) {
    return;
  }
  style[key] = parseOnOffValue(element);
}

function parseOnOffValue(element) {
  const raw = element.getAttribute("w:val") || element.getAttribute("val");
  if (raw == null) {
    return true;
  }
  return !["0", "false", "off"].includes(String(raw).toLowerCase());
}

function compactObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return value;
  }

  const result = {};
  for (const [key, innerValue] of Object.entries(value)) {
    if (innerValue == null) {
      continue;
    }
    if (typeof innerValue === "object" && !Array.isArray(innerValue)) {
      const nested = compactObject(innerValue);
      if (nested && Object.keys(nested).length > 0) {
        result[key] = nested;
      }
      continue;
    }
    result[key] = innerValue;
  }
  return result;
}

module.exports = {
  MATH_NS,
  collectMathText,
  createMathElement,
  formatMathPlaceholder,
  isMathElement,
  normalizeMathProps,
  parseMathNode,
  updateMathElement,
};
