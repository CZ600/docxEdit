"use strict";

const { childElements, createWordElement, isElement } = require("../shared/xml");

function parseParagraphStyle(paragraphElement) {
  const pPr = getDirectChild(paragraphElement, "w:pPr");
  if (!pPr) {
    return {};
  }

  const style = {};
  const pStyle = getDirectChild(pPr, "w:pStyle");
  const alignment = getDirectChild(pPr, "w:jc");
  const spacing = getDirectChild(pPr, "w:spacing");
  const indent = getDirectChild(pPr, "w:ind");

  if (pStyle) {
    style.styleId = getWordAttribute(pStyle, "val");
  }

  if (alignment) {
    style.alignment = getWordAttribute(alignment, "val");
  }

  assignOnOffStyle(style, "keepNext", getDirectChild(pPr, "w:keepNext"));
  assignOnOffStyle(style, "keepLines", getDirectChild(pPr, "w:keepLines"));
  assignOnOffStyle(style, "pageBreakBefore", getDirectChild(pPr, "w:pageBreakBefore"));

  if (spacing) {
    style.spacing = compactObject({
      before: getWordAttribute(spacing, "before"),
      after: getWordAttribute(spacing, "after"),
      line: getWordAttribute(spacing, "line"),
      lineRule: getWordAttribute(spacing, "lineRule"),
    });
  }

  if (indent) {
    style.indent = compactObject({
      left: getWordAttribute(indent, "left"),
      right: getWordAttribute(indent, "right"),
      firstLine: getWordAttribute(indent, "firstLine"),
      hanging: getWordAttribute(indent, "hanging"),
    });
  }

  return compactObject(style);
}

function parseRunStyle(runElement) {
  const rPr = getDirectChild(runElement, "w:rPr");
  if (!rPr) {
    return {};
  }

  const style = {};
  const rStyle = getDirectChild(rPr, "w:rStyle");
  const underline = getDirectChild(rPr, "w:u");
  const color = getDirectChild(rPr, "w:color");
  const highlight = getDirectChild(rPr, "w:highlight");
  const size = getDirectChild(rPr, "w:sz");
  const fonts = getDirectChild(rPr, "w:rFonts");

  if (rStyle) {
    style.styleId = getWordAttribute(rStyle, "val");
  }

  assignOnOffStyle(style, "bold", getDirectChild(rPr, "w:b"));
  assignOnOffStyle(style, "italic", getDirectChild(rPr, "w:i"));

  if (underline) {
    style.underline = getWordAttribute(underline, "val") || "single";
  }

  if (color) {
    style.color = getWordAttribute(color, "val");
  }

  if (highlight) {
    style.highlight = getWordAttribute(highlight, "val");
  }

  if (size) {
    style.fontSize = getWordAttribute(size, "val");
  }

  if (fonts) {
    style.fontFamily = compactObject({
      ascii: getWordAttribute(fonts, "ascii"),
      hAnsi: getWordAttribute(fonts, "hAnsi"),
      eastAsia: getWordAttribute(fonts, "eastAsia"),
      cs: getWordAttribute(fonts, "cs"),
    });
  }

  return compactObject(style);
}

function applyParagraphStyle(paragraphElement, style) {
  const normalized = compactObject(cloneStyle(style || {}));
  const pPr = ensurePropertyContainer(paragraphElement, "w:pPr", normalized);
  if (!pPr) {
    return false;
  }

  const changes = [
    syncValueChild(pPr, "w:pStyle", normalized.styleId),
    syncValueChild(pPr, "w:jc", normalized.alignment),
    syncOnOffChild(pPr, "w:keepNext", normalized.keepNext),
    syncOnOffChild(pPr, "w:keepLines", normalized.keepLines),
    syncOnOffChild(pPr, "w:pageBreakBefore", normalized.pageBreakBefore),
    syncAttributeChild(pPr, "w:spacing", normalized.spacing, ["before", "after", "line", "lineRule"]),
    syncAttributeChild(pPr, "w:ind", normalized.indent, ["left", "right", "firstLine", "hanging"]),
  ];

  cleanupEmptyPropertyContainer(paragraphElement, pPr);
  return changes.some(Boolean);
}

function applyRunStyle(runElement, style) {
  const normalized = compactObject(cloneStyle(style || {}));
  const rPr = ensurePropertyContainer(runElement, "w:rPr", normalized);
  if (!rPr) {
    return false;
  }

  const changes = [
    syncValueChild(rPr, "w:rStyle", normalized.styleId),
    syncOnOffChild(rPr, "w:b", normalized.bold),
    syncOnOffChild(rPr, "w:i", normalized.italic),
    syncValueChild(rPr, "w:u", normalized.underline),
    syncValueChild(rPr, "w:color", normalized.color),
    syncValueChild(rPr, "w:highlight", normalized.highlight),
    syncValueChild(rPr, "w:sz", normalized.fontSize),
    syncAttributeChild(rPr, "w:rFonts", normalized.fontFamily, ["ascii", "hAnsi", "eastAsia", "cs"]),
  ];

  cleanupEmptyPropertyContainer(runElement, rPr);
  return changes.some(Boolean);
}

function cloneStyle(style) {
  if (Array.isArray(style)) {
    return style.map((item) => cloneStyle(item));
  }

  if (style && typeof style === "object") {
    const cloned = {};
    for (const [key, value] of Object.entries(style)) {
      cloned[key] = cloneStyle(value);
    }
    return cloned;
  }

  return style;
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

function assignOnOffStyle(style, key, element) {
  if (!element) {
    return;
  }

  style[key] = parseOnOffValue(element);
}

function parseOnOffValue(element) {
  const raw = getWordAttribute(element, "val");
  if (raw == null) {
    return true;
  }

  return !["0", "false", "off"].includes(String(raw).toLowerCase());
}

function ensurePropertyContainer(hostElement, containerName, style) {
  const hasStyle = style && Object.keys(style).length > 0;
  const existing = getDirectChild(hostElement, containerName);

  if (existing) {
    if (!hasStyle && childElements(existing).length === 0) {
      hostElement.removeChild(existing);
      return null;
    }

    return existing;
  }

  if (!hasStyle) {
    return null;
  }

  const container = createWordElement(hostElement.ownerDocument, containerName);
  const insertBefore = childElements(hostElement)[0] || null;
  hostElement.insertBefore(container, insertBefore);
  return container;
}

function cleanupEmptyPropertyContainer(hostElement, container) {
  if (container && childElements(container).length === 0) {
    hostElement.removeChild(container);
  }
}

function syncValueChild(container, childName, value) {
  return syncAttributeChild(container, childName, value == null ? null : { val: String(value) }, ["val"]);
}

function syncOnOffChild(container, childName, value) {
  if (value == null) {
    return removeDirectChild(container, childName);
  }

  const child = getOrCreateDirectChild(container, childName);
  const nextValue = value ? "1" : "0";
  return setWordAttribute(child, "val", nextValue);
}

function syncAttributeChild(container, childName, value, attributeNames) {
  if (!value || Object.keys(value).length === 0) {
    return removeDirectChild(container, childName);
  }

  const child = getOrCreateDirectChild(container, childName);
  let changed = false;

  for (const attributeName of attributeNames) {
    const nextValue = value[attributeName] ?? null;
    if (nextValue == null) {
      if (child.hasAttribute(`w:${attributeName}`) || child.hasAttribute(attributeName)) {
        child.removeAttribute(`w:${attributeName}`);
        child.removeAttribute(attributeName);
        changed = true;
      }
      continue;
    }

    changed = setWordAttribute(child, attributeName, String(nextValue)) || changed;
  }

  if (Array.from(child.attributes).length === 0) {
    container.removeChild(child);
    return true;
  }

  return changed;
}

function getOrCreateDirectChild(container, childName) {
  return getDirectChild(container, childName) || appendDirectChild(container, childName);
}

function appendDirectChild(container, childName) {
  const child = createWordElement(container.ownerDocument, childName);
  container.appendChild(child);
  return child;
}

function removeDirectChild(container, childName) {
  const child = getDirectChild(container, childName);
  if (!child) {
    return false;
  }

  container.removeChild(child);
  return true;
}

function getDirectChild(node, qualifiedName) {
  return childElements(node).find((child) => isElement(child, qualifiedName)) || null;
}

function setWordAttribute(element, localName, value) {
  const current = getWordAttribute(element, localName);
  if (current === value) {
    return false;
  }

  element.setAttribute(`w:${localName}`, value);
  return true;
}

function getWordAttribute(element, localName) {
  return element.getAttribute(`w:${localName}`) || element.getAttribute(localName) || null;
}

function stylesEqual(left, right) {
  return JSON.stringify(compactObject(cloneStyle(left || {}))) === JSON.stringify(compactObject(cloneStyle(right || {})));
}

module.exports = {
  applyParagraphStyle,
  applyRunStyle,
  cloneStyle,
  parseParagraphStyle,
  parseRunStyle,
  stylesEqual,
};
