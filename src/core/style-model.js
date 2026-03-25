"use strict";

const { childElements, createWordElement, isElement } = require("../shared/xml");

const BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"];
const CELL_MARGIN_SIDES = ["top", "left", "bottom", "right"];

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

function parseTableStyle(tableElement) {
  const tblPr = getDirectChild(tableElement, "w:tblPr");
  if (!tblPr) {
    return {};
  }

  const style = {};
  const tblStyle = getDirectChild(tblPr, "w:tblStyle");
  const tblW = getDirectChild(tblPr, "w:tblW");
  const tblLayout = getDirectChild(tblPr, "w:tblLayout");
  const jc = getDirectChild(tblPr, "w:jc");
  const borders = getDirectChild(tblPr, "w:tblBorders");

  if (tblStyle) {
    style.styleId = getWordAttribute(tblStyle, "val");
  }

  if (tblW) {
    style.width = compactObject({
      w: getWordAttribute(tblW, "w"),
      type: getWordAttribute(tblW, "type"),
    });
  }

  if (tblLayout) {
    style.layout = getWordAttribute(tblLayout, "type");
  }

  if (jc) {
    style.alignment = getWordAttribute(jc, "val");
  }

  if (borders) {
    style.borders = parseBorders(borders, BORDER_SIDES);
  }

  return compactObject(style);
}

function parseTableRowStyle(rowElement) {
  const trPr = getDirectChild(rowElement, "w:trPr");
  if (!trPr) {
    return {};
  }

  const style = {};
  const trHeight = getDirectChild(trPr, "w:trHeight");

  if (trHeight) {
    style.height = compactObject({
      val: getWordAttribute(trHeight, "val"),
      rule: getWordAttribute(trHeight, "hRule"),
    });
  }

  assignOnOffStyle(style, "header", getDirectChild(trPr, "w:tblHeader"));
  assignOnOffStyle(style, "cantSplit", getDirectChild(trPr, "w:cantSplit"));

  return compactObject(style);
}

function parseTableCellStyle(cellElement) {
  const tcPr = getDirectChild(cellElement, "w:tcPr");
  if (!tcPr) {
    return {};
  }

  const style = {};
  const tcW = getDirectChild(tcPr, "w:tcW");
  const shading = getDirectChild(tcPr, "w:shd");
  const verticalAlign = getDirectChild(tcPr, "w:vAlign");
  const gridSpan = getDirectChild(tcPr, "w:gridSpan");
  const vMerge = getDirectChild(tcPr, "w:vMerge");
  const tcBorders = getDirectChild(tcPr, "w:tcBorders");
  const tcMar = getDirectChild(tcPr, "w:tcMar");

  if (tcW) {
    style.width = compactObject({
      w: getWordAttribute(tcW, "w"),
      type: getWordAttribute(tcW, "type"),
    });
  }

  if (shading) {
    style.shading = compactObject({
      val: getWordAttribute(shading, "val"),
      fill: getWordAttribute(shading, "fill"),
      color: getWordAttribute(shading, "color"),
    });
  }

  if (verticalAlign) {
    style.verticalAlign = getWordAttribute(verticalAlign, "val");
  }

  if (gridSpan) {
    style.gridSpan = getWordAttribute(gridSpan, "val");
  }

  if (vMerge) {
    style.vMerge = getWordAttribute(vMerge, "val") || "continue";
  }

  if (tcBorders) {
    style.borders = parseBorders(tcBorders, BORDER_SIDES.filter((side) => !side.startsWith("inside")));
  }

  if (tcMar) {
    style.margins = parseMargins(tcMar);
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

function applyTableStyle(tableElement, style) {
  const normalized = compactObject(cloneStyle(style || {}));
  const tblPr = ensurePropertyContainer(tableElement, "w:tblPr", normalized);
  if (!tblPr) {
    return false;
  }

  const changes = [
    syncValueChild(tblPr, "w:tblStyle", normalized.styleId),
    syncWidthChild(tblPr, "w:tblW", normalized.width),
    syncAttributeChild(tblPr, "w:tblLayout", normalized.layout == null ? null : { type: normalized.layout }, ["type"]),
    syncValueChild(tblPr, "w:jc", normalized.alignment),
    syncBordersChild(tblPr, "w:tblBorders", normalized.borders, BORDER_SIDES),
  ];

  cleanupEmptyPropertyContainer(tableElement, tblPr);
  return changes.some(Boolean);
}

function applyTableRowStyle(rowElement, style) {
  const normalized = compactObject(cloneStyle(style || {}));
  const trPr = ensurePropertyContainer(rowElement, "w:trPr", normalized);
  if (!trPr) {
    return false;
  }

  const changes = [
    syncAttributeChild(trPr, "w:trHeight", normalized.height, ["val", "rule"], { rule: "hRule" }),
    syncOnOffChild(trPr, "w:tblHeader", normalized.header),
    syncOnOffChild(trPr, "w:cantSplit", normalized.cantSplit),
  ];

  cleanupEmptyPropertyContainer(rowElement, trPr);
  return changes.some(Boolean);
}

function applyTableCellStyle(cellElement, style) {
  const normalized = compactObject(cloneStyle(style || {}));
  const tcPr = ensurePropertyContainer(cellElement, "w:tcPr", normalized);
  if (!tcPr) {
    return false;
  }

  const changes = [
    syncWidthChild(tcPr, "w:tcW", normalized.width),
    syncAttributeChild(tcPr, "w:shd", normalized.shading, ["val", "fill", "color"]),
    syncValueChild(tcPr, "w:vAlign", normalized.verticalAlign),
    syncValueChild(tcPr, "w:gridSpan", normalized.gridSpan),
    syncValueChild(tcPr, "w:vMerge", normalized.vMerge),
    syncBordersChild(tcPr, "w:tcBorders", normalized.borders, BORDER_SIDES.filter((side) => !side.startsWith("inside"))),
    syncMarginsChild(tcPr, "w:tcMar", normalized.margins),
  ];

  cleanupEmptyPropertyContainer(cellElement, tcPr);
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

function syncAttributeChild(container, childName, value, attributeNames, aliases = null) {
  if (!value || Object.keys(value).length === 0) {
    return removeDirectChild(container, childName);
  }

  const child = getOrCreateDirectChild(container, childName);
  let changed = false;

  for (const attributeName of attributeNames) {
    const nextValue = value[attributeName] ?? null;
    const xmlAttributeName = aliases && aliases[attributeName] ? aliases[attributeName] : attributeName;
    if (nextValue == null) {
      if (child.hasAttribute(`w:${xmlAttributeName}`) || child.hasAttribute(xmlAttributeName)) {
        child.removeAttribute(`w:${xmlAttributeName}`);
        child.removeAttribute(xmlAttributeName);
        changed = true;
      }
      continue;
    }

    changed = setWordAttribute(child, xmlAttributeName, String(nextValue)) || changed;
  }

  if (Array.from(child.attributes).length === 0) {
    container.removeChild(child);
    return true;
  }

  return changed;
}

function syncWidthChild(container, childName, value) {
  return syncAttributeChild(container, childName, value, ["w", "type"]);
}

function parseBorders(container, sides) {
  const result = {};
  for (const side of sides) {
    const border = getDirectChild(container, `w:${side}`);
    if (!border) continue;
    result[side] = compactObject({
      val: getWordAttribute(border, "val"),
      sz: getWordAttribute(border, "sz"),
      color: getWordAttribute(border, "color"),
      space: getWordAttribute(border, "space"),
    });
  }
  return compactObject(result);
}

function syncBordersChild(container, childName, borders, sides) {
  if (!borders || Object.keys(borders).length === 0) {
    return removeDirectChild(container, childName);
  }

  const borderContainer = getOrCreateDirectChild(container, childName);
  let changed = false;

  for (const side of sides) {
    changed = syncAttributeChild(borderContainer, `w:${side}`, borders[side] || null, ["val", "sz", "color", "space"]) || changed;
  }

  if (childElements(borderContainer).length === 0) {
    container.removeChild(borderContainer);
    return true;
  }

  return changed;
}

function parseMargins(container) {
  const result = {};
  for (const side of CELL_MARGIN_SIDES) {
    const margin = getDirectChild(container, `w:${side}`);
    if (!margin) continue;
    result[side] = compactObject({
      w: getWordAttribute(margin, "w"),
      type: getWordAttribute(margin, "type"),
    });
  }
  return compactObject(result);
}

function syncMarginsChild(container, childName, margins) {
  if (!margins || Object.keys(margins).length === 0) {
    return removeDirectChild(container, childName);
  }

  const marginContainer = getOrCreateDirectChild(container, childName);
  let changed = false;

  for (const side of CELL_MARGIN_SIDES) {
    changed = syncWidthChild(marginContainer, `w:${side}`, margins[side] || null) || changed;
  }

  if (childElements(marginContainer).length === 0) {
    container.removeChild(marginContainer);
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
  applyTableCellStyle,
  applyTableRowStyle,
  applyTableStyle,
  cloneStyle,
  parseParagraphStyle,
  parseRunStyle,
  parseTableCellStyle,
  parseTableRowStyle,
  parseTableStyle,
  stylesEqual,
};
