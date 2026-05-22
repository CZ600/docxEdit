"use strict";

const { XMLSerializer } = require("@xmldom/xmldom");
const { childElements, createWordElement, isElement } = require("../shared/xml");
const {
  applyParagraphStyle,
  applyRunStyle,
  cloneStyle,
  compactObject,
  parseParagraphStyle,
  parseRunStyle,
} = require("./style-model");

const STYLES_XML_PATH = "word/styles.xml";

function findDirectChild(parent, qualifiedName) {
  if (!parent) return null;
  return childElements(parent).find((child) => isElement(child, qualifiedName)) || null;
}

function parseStylesXml(xmlDocument) {
  const root = xmlDocument.documentElement;
  const result = {
    docDefaults: { paragraphStyle: {}, runStyle: {} },
    styles: new Map(),
    xmlDocument,
  };

  const docDefaults = findDirectChild(root, "w:docDefaults");
  if (docDefaults) {
    const rPrDefault = findDirectChild(docDefaults, "w:rPrDefault");
    if (rPrDefault) {
      result.docDefaults.runStyle = compactObject(parseRunStyle(rPrDefault));
    }
    const pPrDefault = findDirectChild(docDefaults, "w:pPrDefault");
    if (pPrDefault) {
      result.docDefaults.paragraphStyle = compactObject(parseParagraphStyle(pPrDefault));
    }
  }

  for (const child of childElements(root)) {
    if (!isElement(child, "w:style")) continue;
    const styleId = child.getAttribute("w:styleId") || child.getAttribute("styleId");
    if (!styleId) continue;

    const type = child.getAttribute("w:type") || child.getAttribute("type") || "paragraph";
    const nameEl = findDirectChild(child, "w:name");
    const name = nameEl
      ? nameEl.getAttribute("w:val") || nameEl.getAttribute("val")
      : styleId;
    const basedOnEl = findDirectChild(child, "w:basedOn");
    const basedOn = basedOnEl
      ? basedOnEl.getAttribute("w:val") || basedOnEl.getAttribute("val")
      : null;

    result.styles.set(styleId, {
      styleId,
      type,
      name,
      basedOn,
      paragraphStyle: type === "paragraph" || type === "table" ? compactObject(parseParagraphStyle(child)) : {},
      runStyle: type === "paragraph" || type === "character" ? compactObject(parseRunStyle(child)) : {},
      element: child,
    });
  }

  return result;
}

function resolveStyleChain(stylesData, styleId, visited = new Set()) {
  if (!styleId || visited.has(styleId)) {
    return { paragraphStyle: {}, runStyle: {} };
  }
  visited.add(styleId);

  const style = stylesData.styles.get(styleId);
  if (!style) {
    return { paragraphStyle: {}, runStyle: {} };
  }

  const parent = resolveStyleChain(stylesData, style.basedOn, visited);
  return {
    paragraphStyle: mergeStyleObjects(parent.paragraphStyle, style.paragraphStyle),
    runStyle: mergeStyleObjects(parent.runStyle, style.runStyle),
  };
}

function resolveEffectiveStyle(stylesData, styleId) {
  const chain = resolveStyleChain(stylesData, styleId);
  return {
    paragraphStyle: compactObject(mergeStyleObjects(stylesData.docDefaults.paragraphStyle, chain.paragraphStyle)),
    runStyle: compactObject(mergeStyleObjects(stylesData.docDefaults.runStyle, chain.runStyle)),
  };
}

function extractStyleProfile(stylesData) {
  const profile = {
    defaults: {
      paragraphStyle: cloneStyle(stylesData.docDefaults.paragraphStyle || {}),
      runStyle: cloneStyle(stylesData.docDefaults.runStyle || {}),
    },
    styles: {},
  };

  for (const [styleId, style] of stylesData.styles) {
    profile.styles[styleId] = {
      name: style.name,
      type: style.type,
      basedOn: style.basedOn,
      paragraphStyle: cloneStyle(style.paragraphStyle || {}),
      runStyle: cloneStyle(style.runStyle || {}),
    };
  }

  return profile;
}

function applyStyleProfileToData(stylesData, profile) {
  if (profile.defaults) {
    updateDocDefaults(stylesData, profile.defaults);
  }

  if (profile.styles) {
    for (const [styleId, styleDef] of Object.entries(profile.styles)) {
      updateStyleDefinition(stylesData, styleId, styleDef);
    }
  }
}

function updateDocDefaults(stylesData, defaults) {
  const root = stylesData.xmlDocument.documentElement;

  if (defaults.runStyle && Object.keys(defaults.runStyle).length > 0) {
    let docDefaults = findDirectChild(root, "w:docDefaults");
    if (!docDefaults) {
      docDefaults = createWordElement(root.ownerDocument, "w:docDefaults");
      root.insertBefore(docDefaults, root.firstChild);
    }
    let rPrDefault = findDirectChild(docDefaults, "w:rPrDefault");
    if (!rPrDefault) {
      rPrDefault = createWordElement(root.ownerDocument, "w:rPrDefault");
      docDefaults.appendChild(rPrDefault);
    }
    applyRunStyle(rPrDefault, defaults.runStyle);
    stylesData.docDefaults.runStyle = cloneStyle(defaults.runStyle);
  }

  if (defaults.paragraphStyle && Object.keys(defaults.paragraphStyle).length > 0) {
    let docDefaults = findDirectChild(root, "w:docDefaults");
    if (!docDefaults) {
      docDefaults = createWordElement(root.ownerDocument, "w:docDefaults");
      root.insertBefore(docDefaults, root.firstChild);
    }
    let pPrDefault = findDirectChild(docDefaults, "w:pPrDefault");
    if (!pPrDefault) {
      pPrDefault = createWordElement(root.ownerDocument, "w:pPrDefault");
      docDefaults.appendChild(pPrDefault);
    }
    applyParagraphStyle(pPrDefault, defaults.paragraphStyle);
    stylesData.docDefaults.paragraphStyle = cloneStyle(defaults.paragraphStyle);
  }
}

function updateStyleDefinition(stylesData, styleId, styleDef) {
  let style = stylesData.styles.get(styleId);

  if (!style) {
    const root = stylesData.xmlDocument.documentElement;
    const element = createWordElement(root.ownerDocument, "w:style");
    element.setAttribute("w:styleId", styleId);
    element.setAttribute("w:type", styleDef.type || "paragraph");

    const nameEl = createWordElement(root.ownerDocument, "w:name");
    nameEl.setAttribute("w:val", styleDef.name || styleId);
    element.appendChild(nameEl);

    if (styleDef.basedOn) {
      const basedOnEl = createWordElement(root.ownerDocument, "w:basedOn");
      basedOnEl.setAttribute("w:val", styleDef.basedOn);
      element.appendChild(basedOnEl);
    }

    root.appendChild(element);

    style = {
      styleId,
      type: styleDef.type || "paragraph",
      name: styleDef.name || styleId,
      basedOn: styleDef.basedOn || null,
      paragraphStyle: {},
      runStyle: {},
      element,
    };
    stylesData.styles.set(styleId, style);
  }

  if (styleDef.paragraphStyle && (style.type === "paragraph" || style.type === "table")) {
    applyParagraphStyle(style.element, styleDef.paragraphStyle);
    style.paragraphStyle = cloneStyle(styleDef.paragraphStyle);
  }

  if (styleDef.runStyle) {
    applyRunStyle(style.element, styleDef.runStyle);
    style.runStyle = cloneStyle(styleDef.runStyle);
  }
}

function serializeStylesXml(stylesData) {
  return new XMLSerializer().serializeToString(stylesData.xmlDocument);
}

function mergeStyleObjects(base, overlay) {
  const result = cloneStyle(base || {});
  for (const [key, value] of Object.entries(overlay || {})) {
    if (value == null) {
      delete result[key];
      continue;
    }
    if (value && typeof value === "object" && !Array.isArray(value)) {
      result[key] = mergeStyleObjects(result[key] || {}, value);
      if (Object.keys(result[key]).length === 0) delete result[key];
      continue;
    }
    result[key] = value;
  }
  return result;
}

module.exports = {
  STYLES_XML_PATH,
  applyStyleProfileToData,
  extractStyleProfile,
  mergeStyleObjects,
  parseStylesXml,
  resolveEffectiveStyle,
  resolveStyleChain,
  serializeStylesXml,
};
