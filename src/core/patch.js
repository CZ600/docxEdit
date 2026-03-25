"use strict";

const { ParagraphTextModel } = require("./paragraph-text-model");
const {
  applyParagraphStyle,
  applyRunStyle,
  applyTableCellStyle,
  applyTableRowStyle,
  applyTableStyle,
  stylesEqual,
} = require("./style-model");
const { assignVNodeSource, createVNode } = require("./vnode");
const { childElements, createWordElement, isElement, setElementText } = require("../shared/xml");

const PART_TYPES = new Set(["body", "header", "footer", "footnotes", "endnotes", "comments"]);
const SPECIAL_ENTRY_TYPES = new Set(["comment", "footnote", "endnote"]);
const ALLOWED_CHILDREN = new Map([
  ["document", new Set(["body", "header", "footer", "footnotes", "endnotes", "comments"])],
  ["body", new Set(["paragraph", "table"])],
  ["header", new Set(["paragraph", "table"])],
  ["footer", new Set(["paragraph", "table"])],
  ["comments", new Set(["comment"])],
  ["footnotes", new Set(["footnote"])],
  ["endnotes", new Set(["endnote"])],
  ["comment", new Set(["paragraph", "table"])],
  ["footnote", new Set(["paragraph", "table"])],
  ["endnote", new Set(["paragraph", "table"])],
  ["paragraph", new Set(["run", "hyperlink"])],
  ["run", new Set(["text", "tab", "break", "text-box", "image"])],
  ["hyperlink", new Set(["run"])],
  ["text-box", new Set(["paragraph", "table"])],
  ["table", new Set(["table-row"])],
  ["table-row", new Set(["table-cell"])],
  ["table-cell", new Set(["paragraph", "table"])],
]);

function validateChildType(parentType, childType) {
  const allowed = ALLOWED_CHILDREN.get(parentType);
  if (allowed && !allowed.has(childType)) {
    throw new Error(`Unsupported child type "${childType}" under "${parentType}".`);
  }
}

function normalizeNextTree(nextRoot) {
  return createVNode(nextRoot);
}

function applyDocumentPatch({ currentRoot, nextRoot, partRootByPath, doc }) {
  const normalizedRoot = normalizeNextTree(nextRoot);
  if (normalizedRoot.type !== "document") {
    throw new Error('doc.patch(nextTree) requires a root vnode with type "document".');
  }

  const currentParts = currentRoot.children;
  const nextParts = normalizedRoot.children;
  const currentByPath = new Map(currentParts.map((node) => [node.props.path, node]));
  if (currentParts.length !== nextParts.length) {
    throw new Error("doc.patch(nextTree) must preserve the existing document parts.");
  }

  const operations = [];
  for (const nextPart of nextParts) {
    const path = nextPart.props && nextPart.props.path;
    const currentPart = currentByPath.get(path);
    const hostRoot = partRootByPath.get(path);
    if (!currentPart || !hostRoot) throw new Error(`Unknown document part "${path}".`);
    if (currentPart.type !== nextPart.type) {
      throw new Error(`Document part "${path}" cannot change type from "${currentPart.type}" to "${nextPart.type}".`);
    }
    patchNode(currentPart, nextPart, hostRoot, operations, { doc, partPath: path });
  }

  return { operations, nextRoot: normalizedRoot };
}

function patchNode(currentNode, nextNode, hostElement, operations, context) {
  nextNode.id = nextNode.id || currentNode.id;
  assignVNodeSource(nextNode, hostElement);
  hostElement.__vnodeKey = nextNode.key;

  const skipChildren = syncNodeProps(currentNode, nextNode, hostElement, operations, context);
  if (!skipChildren) {
    patchChildren(currentNode, nextNode, hostElement, operations, context);
  }
}

function syncNodeProps(currentNode, nextNode, hostElement, operations, context) {
  let changed = false;

  if (nextNode.type === "paragraph") {
    if (!stylesEqual(currentNode.props.style, nextNode.props.style)) {
      changed = applyParagraphStyle(hostElement, nextNode.props.style) || changed;
    }
    const nextText = nextNode.props.text;
    const currentText = currentNode.props.text;
    if (typeof nextText === "string" && nextText !== currentText) {
      new ParagraphTextModel(hostElement).setText(nextText);
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
      return true;
    }
    if (changed) {
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return false;
  }

  if (nextNode.type === "run") {
    if (!stylesEqual(currentNode.props.style, nextNode.props.style)) {
      changed = applyRunStyle(hostElement, nextNode.props.style) || changed;
    }
    if (changed) {
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return false;
  }

  if (nextNode.type === "table") {
    if (!stylesEqual(currentNode.props.style, nextNode.props.style)) {
      changed = applyTableStyle(hostElement, nextNode.props.style) || changed;
    }
    if (changed) {
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return false;
  }

  if (nextNode.type === "table-row") {
    if (!stylesEqual(currentNode.props.style, nextNode.props.style)) {
      changed = applyTableRowStyle(hostElement, nextNode.props.style) || changed;
    }
    if (changed) {
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return false;
  }

  if (nextNode.type === "table-cell") {
    if (!stylesEqual(currentNode.props.style, nextNode.props.style)) {
      changed = applyTableCellStyle(hostElement, nextNode.props.style) || changed;
    }
    if (changed) {
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return false;
  }

  if (nextNode.type === "text") {
    const nextText = nextNode.props.text || "";
    if ((currentNode.props.text || "") !== nextText) {
      setElementText(hostElement, nextText);
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return false;
  }

  if (nextNode.type === "image") {
    changed = syncImageNode(currentNode, nextNode, hostElement, context) || changed;
    if (changed) {
      operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
    }
    return true;
  }

  if (SPECIAL_ENTRY_TYPES.has(nextNode.type)) {
    changed = syncSpecialAttributes(currentNode, nextNode, hostElement) || changed;
  }

  if (changed) {
    operations.push({ type: "PROPS/TEXT_UPDATE", nodeId: currentNode.id, nodeType: currentNode.type });
  }

  return false;
}

function syncSpecialAttributes(currentNode, nextNode, hostElement) {
  let changed = false;
  const currentId = currentNode.props.id ?? null;
  const nextId = nextNode.props.id ?? currentId;
  const currentSpecialType = currentNode.props.specialType ?? null;
  const nextSpecialType = nextNode.props.specialType ?? currentSpecialType;
  if (nextId !== currentId) {
    hostElement.setAttribute("w:id", String(nextId));
    changed = true;
  }
  if (nextSpecialType !== currentSpecialType) {
    if (nextSpecialType == null) hostElement.removeAttribute("w:type");
    else hostElement.setAttribute("w:type", String(nextSpecialType));
    changed = true;
  }
  return changed;
}

function syncImageNode(currentNode, nextNode, hostElement, context) {
  const before = JSON.stringify({
    filename: currentNode.props.filename,
    width: currentNode.props.width,
    height: currentNode.props.height,
    alt: currentNode.props.alt,
    contentType: currentNode.props.contentType,
    relId: currentNode.props.relId,
    mediaPath: currentNode.props.mediaPath,
    target: currentNode.props.target,
  });
  const after = JSON.stringify({
    filename: nextNode.props.filename,
    width: nextNode.props.width,
    height: nextNode.props.height,
    alt: nextNode.props.alt,
    contentType: nextNode.props.contentType,
    relId: nextNode.props.relId,
    mediaPath: nextNode.props.mediaPath,
    target: nextNode.props.target,
  });

  const hasBinaryUpdate = nextNode.props.data != null;
  const hasStructuralUpdate = before !== after;

  if (!hasBinaryUpdate && !hasStructuralUpdate) {
    return false;
  }

  context.doc.createOrUpdateImage(context.partPath, nextNode);
  updateDrawingElement(hostElement, nextNode.props);
  return true;
}

function patchChildren(currentParent, nextParent, hostElement, operations, context) {
  const currentChildren = currentParent.children.slice();
  const nextChildren = nextParent.children.slice();
  const matchedCurrentIds = new Set();
  const realizedChildren = [];
  let sequentialCursor = 0;
  let needsReorder = false;

  for (const nextChild of nextChildren) {
    const matchedCurrent = findMatchingChild(currentChildren, nextChild, matchedCurrentIds, sequentialCursor);
    if (matchedCurrent) {
      matchedCurrentIds.add(matchedCurrent.node.id);
      if (matchedCurrent.mode === "sequential") sequentialCursor = matchedCurrent.index + 1;

      if (matchedCurrent.node.type !== nextChild.type) {
        validateChildType(nextParent.type, nextChild.type);
        const replacement = createHostSubtree(nextChild, hostElement.ownerDocument, operations, context);
        cleanupRemovedNode(matchedCurrent.node, context);
        hostElement.replaceChild(replacement, matchedCurrent.node.source);
        operations.push({ type: "REPLACE", nodeId: matchedCurrent.node.id, nodeType: nextChild.type });
        realizedChildren.push(nextChild);
        continue;
      }

      if (matchedCurrent.index !== realizedChildren.length) {
        operations.push({ type: "MOVE", nodeId: matchedCurrent.node.id, nodeType: matchedCurrent.node.type });
        needsReorder = true;
      }

      patchNode(matchedCurrent.node, nextChild, matchedCurrent.node.source, operations, context);
      realizedChildren.push(nextChild);
      continue;
    }

    validateChildType(nextParent.type, nextChild.type);
    const inserted = createHostSubtree(nextChild, hostElement.ownerDocument, operations, context);
    hostElement.insertBefore(inserted, getTrailingAnchorNode(hostElement));
    operations.push({ type: "INSERT", nodeId: nextChild.id, nodeType: nextChild.type });
    realizedChildren.push(nextChild);
    needsReorder = true;
  }

  for (const currentChild of currentChildren) {
    if (!matchedCurrentIds.has(currentChild.id)) {
      cleanupRemovedNode(currentChild, context);
      if (currentChild.source && currentChild.source.parentNode === hostElement) {
        hostElement.removeChild(currentChild.source);
      }
      operations.push({ type: "REMOVE", nodeId: currentChild.id, nodeType: currentChild.type });
      needsReorder = true;
    }
  }

  if (needsReorder) {
    reorderChildren(hostElement, realizedChildren);
  }
}

function cleanupRemovedNode(node, context) {
  if (node.type === "image") {
    context.doc.removeImageReference(context.partPath, node.props.relId);
  }
  for (const child of node.children || []) {
    cleanupRemovedNode(child, context);
  }
}

function findMatchingChild(currentChildren, nextChild, matchedCurrentIds, sequentialCursor) {
  if (nextChild.key != null) {
    const keyedIndex = currentChildren.findIndex(
      (child) => !matchedCurrentIds.has(child.id) && child.type === nextChild.type && child.key === nextChild.key,
    );
    if (keyedIndex !== -1) return { node: currentChildren[keyedIndex], index: keyedIndex, mode: "keyed" };
  }

  if (nextChild.id != null) {
    const byIdIndex = currentChildren.findIndex(
      (child) => !matchedCurrentIds.has(child.id) && child.id === nextChild.id,
    );
    if (byIdIndex !== -1) return { node: currentChildren[byIdIndex], index: byIdIndex, mode: "id" };
  }

  for (let index = sequentialCursor; index < currentChildren.length; index += 1) {
    const child = currentChildren[index];
    if (!matchedCurrentIds.has(child.id) && child.key == null) {
      return { node: child, index, mode: "sequential" };
    }
  }

  return null;
}

function reorderChildren(hostElement, nextChildren) {
  const anchorNode = getTrailingAnchorNode(hostElement);
  const realizedChildren = nextChildren.filter((child) => child.source && child.source.parentNode === hostElement);

  for (let index = 0; index < realizedChildren.length; index += 1) {
    const child = realizedChildren[index];
    const currentElements = childElements(hostElement).filter((element) => element !== anchorNode);
    const currentAtIndex = currentElements[index] || null;
    if (currentAtIndex === child.source) continue;
    hostElement.insertBefore(child.source, currentAtIndex || anchorNode);
  }
}

function getTrailingAnchorNode(hostElement) {
  if (isElement(hostElement, "w:body")) {
    return childElements(hostElement).find((child) => isElement(child, "w:sectPr")) || null;
  }
  return null;
}

function createHostSubtree(vnode, ownerDocument, operations, context) {
  if (PART_TYPES.has(vnode.type) || vnode.type === "document") {
    throw new Error(`Cannot create host subtree for root node type "${vnode.type}".`);
  }

  const element = createHostElement(vnode, ownerDocument, context);
  assignVNodeSource(vnode, element);

  if (vnode.type === "paragraph") {
    applyParagraphStyle(element, vnode.props.style);
    if (vnode.children.length > 0) {
      for (const child of vnode.children) {
        validateChildType(vnode.type, child.type);
        element.appendChild(createHostSubtree(child, ownerDocument, operations, context));
      }
    } else {
      new ParagraphTextModel(element).setText(vnode.props.text || "");
    }
    return element;
  }

  if (vnode.type === "text") {
    setElementText(element, vnode.props.text || "");
    return element;
  }

  if (vnode.type === "run") {
    applyRunStyle(element, vnode.props.style);
  }

  if (vnode.type === "table") {
    applyTableStyle(element, vnode.props.style);
  }

  if (vnode.type === "table-row") {
    applyTableRowStyle(element, vnode.props.style);
  }

  if (vnode.type === "table-cell") {
    applyTableCellStyle(element, vnode.props.style);
  }

  if (vnode.type === "image") {
    context.doc.createOrUpdateImage(context.partPath, vnode);
    updateDrawingElement(element, vnode.props);
    return element;
  }

  for (const child of vnode.children) {
    validateChildType(vnode.type, child.type);
    element.appendChild(createHostSubtree(child, ownerDocument, operations, context));
  }

  if (vnode.type === "table-cell" && childElements(element).length === 0) {
    const paragraphVNode = createVNode({ type: "paragraph", props: { text: "" }, children: [] });
    element.appendChild(createHostSubtree(paragraphVNode, ownerDocument, operations, context));
    vnode.children.push(paragraphVNode);
  }

  if (SPECIAL_ENTRY_TYPES.has(vnode.type)) {
    syncSpecialAttributes({ props: {} }, vnode, element);
  }

  return element;
}

function createHostElement(vnode, ownerDocument) {
  switch (vnode.type) {
    case "paragraph": return createWordElement(ownerDocument, "w:p");
    case "run": return createWordElement(ownerDocument, "w:r");
    case "text": return createWordElement(ownerDocument, "w:t");
    case "table": return createWordElement(ownerDocument, "w:tbl");
    case "table-row": return createWordElement(ownerDocument, "w:tr");
    case "table-cell": return createWordElement(ownerDocument, "w:tc");
    case "hyperlink": return createWordElement(ownerDocument, "w:hyperlink");
    case "tab": return createWordElement(ownerDocument, "w:tab");
    case "break": return createWordElement(ownerDocument, "w:br");
    case "text-box": return createWordElement(ownerDocument, "w:txbxContent");
    case "comment": return createWordElement(ownerDocument, "w:comment");
    case "footnote": return createWordElement(ownerDocument, "w:footnote");
    case "endnote": return createWordElement(ownerDocument, "w:endnote");
    case "image": return ownerDocument.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:drawing");
    default: throw new Error(`Unsupported vnode type "${vnode.type}" for host creation.`);
  }
}

function updateDrawingElement(drawingElement, props) {
  while (drawingElement.firstChild) {
    drawingElement.removeChild(drawingElement.firstChild);
  }

  const doc = drawingElement.ownerDocument;
  const wpNs = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  const aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const picNs = "http://schemas.openxmlformats.org/drawingml/2006/picture";
  const rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const inline = doc.createElementNS(wpNs, "wp:inline");
  const extent = doc.createElementNS(wpNs, "wp:extent");
  extent.setAttribute("cx", String(props.width || "990000"));
  extent.setAttribute("cy", String(props.height || "792000"));
  inline.appendChild(extent);

  const docPr = doc.createElementNS(wpNs, "wp:docPr");
  docPr.setAttribute("id", "1");
  docPr.setAttribute("name", props.filename || "image");
  if (props.alt) {
    docPr.setAttribute("descr", props.alt);
  }
  inline.appendChild(docPr);

  const graphic = doc.createElementNS(aNs, "a:graphic");
  const graphicData = doc.createElementNS(aNs, "a:graphicData");
  graphicData.setAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture");
  const pic = doc.createElementNS(picNs, "pic:pic");
  const nvPicPr = doc.createElementNS(picNs, "pic:nvPicPr");
  const cNvPr = doc.createElementNS(picNs, "pic:cNvPr");
  cNvPr.setAttribute("id", "0");
  cNvPr.setAttribute("name", props.filename || "image");
  if (props.alt) cNvPr.setAttribute("descr", props.alt);
  nvPicPr.appendChild(cNvPr);
  nvPicPr.appendChild(doc.createElementNS(picNs, "pic:cNvPicPr"));
  pic.appendChild(nvPicPr);

  const blipFill = doc.createElementNS(picNs, "pic:blipFill");
  const blip = doc.createElementNS(aNs, "a:blip");
  blip.setAttributeNS(rNs, "r:embed", props.relId);
  blipFill.appendChild(blip);
  const stretch = doc.createElementNS(aNs, "a:stretch");
  stretch.appendChild(doc.createElementNS(aNs, "a:fillRect"));
  blipFill.appendChild(stretch);
  pic.appendChild(blipFill);

  const spPr = doc.createElementNS(picNs, "pic:spPr");
  const xfrm = doc.createElementNS(aNs, "a:xfrm");
  const off = doc.createElementNS(aNs, "a:off");
  off.setAttribute("x", "0");
  off.setAttribute("y", "0");
  const ext = doc.createElementNS(aNs, "a:ext");
  ext.setAttribute("cx", String(props.width || "990000"));
  ext.setAttribute("cy", String(props.height || "792000"));
  xfrm.appendChild(off);
  xfrm.appendChild(ext);
  spPr.appendChild(xfrm);
  spPr.appendChild(doc.createElementNS(aNs, "a:prstGeom")).setAttribute("prst", "rect");
  pic.appendChild(spPr);

  graphicData.appendChild(pic);
  graphic.appendChild(graphicData);
  inline.appendChild(graphic);
  drawingElement.appendChild(inline);
}

module.exports = {
  applyDocumentPatch,
};
