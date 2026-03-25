"use strict";

const fs = require("node:fs/promises");
const JSZip = require("jszip");
const { XMLSerializer } = require("@xmldom/xmldom");
const { applyDocumentPatch } = require("./patch");
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
  buildPartController,
  buildLocation,
  parseXmlString,
} = require("./document-part");
const {
  createImageDescriptor,
  ensureImageBuffer,
  getRelsPath,
  loadRelationships,
  readImageBuffer,
  resolveRelationshipTarget,
} = require("./image-model");
const { cloneVNode, createVNode, visitVNode } = require("./vnode");
const { replaceAllInText } = require("./text-utils");

const MAIN_DOCUMENT_PATH = "word/document.xml";
const CONTENT_TYPES_PATH = "[Content_Types].xml";

class VirtualWordDocument {
  constructor({ zip, partsData, relationshipsByPartPath }) {
    this.zip = zip;
    this.partsData = partsData;
    this.relationshipsByPartPath = relationshipsByPartPath;
    this.rootVNode = null;
    this.parts = [];
    this.nodeById = new Map();
    this.metadataById = new Map();
    this.partRootByPath = new Map();
    this.documentNodeId = null;
    this.rebuildFromXml();
  }

  static async load(input) {
    const sourceBuffer = Buffer.isBuffer(input) ? input : await fs.readFile(input);
    const zip = await JSZip.loadAsync(sourceBuffer);
    const partDescriptors = await loadSupportedPartDescriptors(zip);
    const partsData = [];
    const relationshipsByPartPath = new Map();

    for (const descriptor of partDescriptors) {
      const xml = await zip.file(descriptor.path).async("string");
      partsData.push({
        path: descriptor.path,
        type: descriptor.type,
        xmlDocument: parseXmlString(xml),
      });

      const relsPath = getRelsPath(descriptor.path);
      const relsXml = zip.file(relsPath) ? await zip.file(relsPath).async("string") : null;
      const relationships = loadRelationships(descriptor.path, relsXml);
      for (const relationship of relationships.relationships.values()) {
        relationship.partPath = descriptor.path;
      }
      relationshipsByPartPath.set(descriptor.path, relationships);
    }

    return new VirtualWordDocument({ zip, partsData, relationshipsByPartPath });
  }

  rebuildFromXml() {
    this.metadataById = new Map();
    this.partRootByPath = new Map();

    const builtParts = this.partsData.map((partData) =>
      buildPartController({
        doc: this,
        path: partData.path,
        type: partData.type,
        xmlDocument: partData.xmlDocument,
        metadataById: this.metadataById,
      }),
    );

    this.rootVNode = createVNode({
      id: this.documentNodeId,
      type: "document",
      props: {},
      children: builtParts.map((part) => part.rootVNode),
      source: null,
    });
    this.documentNodeId = this.rootVNode.id;

    this.nodeById = new Map();
    visitVNode(this.rootVNode, (node, parent) => {
      node.parent = parent;
      this.nodeById.set(node.id, node);
      if (!this.metadataById.has(node.id)) {
        if (node.type === "document") {
          this.metadataById.set(node.id, {
            partPath: null,
            partType: "document",
            location: "document",
            specialType: null,
          });
          return;
        }
        const parentMetadata = parent ? this.metadataById.get(parent.id) : null;
        if (parentMetadata) {
          this.metadataById.set(node.id, {
            partPath: parentMetadata.partPath,
            partType: parentMetadata.partType,
            location: buildLocation(parentMetadata.partType, []),
            specialType: node.props.specialType || null,
          });
        }
      }
    });

    this.parts = builtParts.map((part) => {
      const controller = new DocumentPartController(this, part.rootVNode.id, part.path, part.type);
      this.partRootByPath.set(part.path, part.rootVNode.source);
      return controller;
    });
  }

  describeImage(partPath, relId) {
    const relationships = this.relationshipsByPartPath.get(partPath);
    if (!relationships) return null;
    const relationship = relationships.relationships.get(relId);
    if (!relationship) return null;
    relationship.partPath = partPath;
    return createImageDescriptor({ relId, relationship, zip: this.zip });
  }

  async getImageData(nodeId) {
    const node = this.getNodeById(nodeId);
    const mediaPath = node.props.mediaPath || null;
    if (!mediaPath) return null;
    return readImageBuffer(this.zip, mediaPath);
  }

  toComponentTree() {
    return cloneVNode(this.rootVNode);
  }

  cloneNodeById(nodeId) {
    return cloneVNode(this.getNodeById(nodeId));
  }

  getNodeById(nodeId) {
    const node = this.nodeById.get(nodeId);
    if (!node) throw new Error(`Node ${nodeId} is no longer available.`);
    return node;
  }

  getNodeMetadata(nodeId) {
    return this.metadataById.get(nodeId) || null;
  }

  patchWithMutableTree(mutator, { skipIfUnchanged = null } = {}) {
    const nextRoot = this.toComponentTree();
    mutator(nextRoot);
    if (skipIfUnchanged && skipIfUnchanged()) return this;
    return this.patch(nextRoot);
  }

  patch(nextTree) {
    const result = applyDocumentPatch({
      currentRoot: this.rootVNode,
      nextRoot: nextTree,
      partRootByPath: this.partRootByPath,
      doc: this,
    });
    this.rebuildFromXml();
    return result;
  }

  getParts() {
    return this.parts.slice();
  }

  getBody() {
    return this.parts.find((part) => part.type === "body");
  }

  getHeaders() {
    return this.parts.filter((part) => part.type === "header");
  }

  getFooters() {
    return this.parts.filter((part) => part.type === "footer");
  }

  getParagraphs() {
    return this.getDescendantControllers(this.rootVNode.id, "paragraph");
  }

  getParagraph(index) {
    return this.getParagraphs()[index];
  }

  getTables() {
    return this.getDescendantControllers(this.rootVNode.id, "table");
  }

  getTextBoxes() {
    return this.getDescendantControllers(this.rootVNode.id, "text-box");
  }

  getImages() {
    return this.getDescendantControllers(this.rootVNode.id, "image");
  }

  getMaths() {
    return this.getDescendantControllers(this.rootVNode.id, "math");
  }

  getFootnotes() {
    return this.getDescendantControllers(this.rootVNode.id, "footnote").filter((entry) => !entry.metadata.specialType);
  }

  getEndnotes() {
    return this.getDescendantControllers(this.rootVNode.id, "endnote").filter((entry) => !entry.metadata.specialType);
  }

  getComments() {
    return this.getDescendantControllers(this.rootVNode.id, "comment").filter((entry) => !entry.metadata.specialType);
  }

  replaceAll(searchValue, replacement, { partTypes = null } = {}) {
    let count = 0;
    this.patchWithMutableTree((nextRoot) => {
      visitVNode(nextRoot, (node) => {
        if (node.type !== "paragraph") return;
        const metadata = this.metadataById.get(node.id);
        if (!metadata) return;
        if (partTypes && !partTypes.includes(metadata.partType)) return;
        const result = replaceAllInText(node.props.text || "", searchValue, replacement);
        if (result.count > 0) {
          node.props.text = result.text;
          count += result.count;
        }
      });
    }, { skipIfUnchanged: () => count === 0 });
    return count;
  }

  getDirectChildControllers(nodeId, type) {
    const node = this.getNodeById(nodeId);
    return node.children.filter((child) => child.type === type).map((child) => this.createControllerForNode(child.id));
  }

  getDescendantControllers(nodeId, type) {
    const root = this.getNodeById(nodeId);
    const results = [];
    for (const child of root.children) {
      collectControllersByType(this, child, type, results);
    }
    return results;
  }

  createControllerForNode(nodeId) {
    const node = this.getNodeById(nodeId);
    switch (node.type) {
      case "paragraph":
        return new ParagraphController(this, nodeId);
      case "run":
        return new RunController(this, nodeId);
      case "image":
        return new ImageController(this, nodeId);
      case "math":
        return new MathController(this, nodeId);
      case "table":
        return new TableController(this, nodeId);
      case "table-row":
        return new TableRowController(this, nodeId);
      case "table-cell":
        return new TableCellController(this, nodeId);
      case "text-box":
        return new TextBoxController(this, nodeId);
      case "comment":
      case "footnote":
      case "endnote":
        return new StructuredEntryController(this, nodeId);
      default:
        throw new Error(`Unsupported controller node type "${node.type}".`);
    }
  }

  getRelationships(partPath) {
    let relationships = this.relationshipsByPartPath.get(partPath);
    if (!relationships) {
      relationships = loadRelationships(partPath, null);
      this.relationshipsByPartPath.set(partPath, relationships);
    }
    return relationships;
  }

  createOrUpdateImage(partPath, imageNode) {
    const relationships = this.getRelationships(partPath);
    const buffer = ensureImageBuffer(imageNode.props.data || Buffer.alloc(0));
    let relId = imageNode.props.relId;
    let mediaPath = imageNode.props.mediaPath;

    if (!relId) {
      relId = nextRelationshipId(relationships);
    }

    if (!mediaPath) {
      mediaPath = this.allocateMediaPath(imageNode.props.filename, imageNode.props.contentType);
    }

    const target = makeRelativeTarget(partPath, mediaPath);
    relationships.relationships.set(relId, {
      id: relId,
      type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      target,
      targetMode: null,
      partPath,
    });

    this.zip.file(mediaPath, buffer);
    imageNode.props.relId = relId;
    imageNode.props.mediaPath = mediaPath;
    imageNode.props.target = target;
    imageNode.props.filename = imageNode.props.filename || mediaPath.split("/").pop();
    return imageNode.props;
  }

  removeImageReference(partPath, relId) {
    if (!relId) return;
    const relationships = this.getRelationships(partPath);
    const relationship = relationships.relationships.get(relId);
    if (!relationship) return;
    relationships.relationships.delete(relId);
    const mediaPath = resolveRelationshipTarget(partPath, relationship.target);
    if (!this.isMediaPathReferenced(mediaPath)) {
      this.zip.remove(mediaPath);
    }
  }

  isMediaPathReferenced(mediaPath) {
    for (const [partPath, relationships] of this.relationshipsByPartPath.entries()) {
      for (const relationship of relationships.relationships.values()) {
        if (resolveRelationshipTarget(partPath, relationship.target) === mediaPath) {
          return true;
        }
      }
    }
    return false;
  }

  allocateMediaPath(filename, contentType) {
    const ext = (filename && /\.[^.]+$/.exec(filename)) ? filename.split(".").pop() : extensionForContentType(contentType);
    let index = 1;
    while (this.zip.file(`word/media/image${index}.${ext}`)) {
      index += 1;
    }
    return `word/media/image${index}.${ext}`;
  }

  async toBuffer() {
    for (const part of this.partsData) {
      const xml = new XMLSerializer().serializeToString(part.xmlDocument);
      this.zip.file(part.path, xml);
    }

    for (const relationships of this.relationshipsByPartPath.values()) {
      syncRelationshipsXml(relationships);
      const xml = new XMLSerializer().serializeToString(relationships.xmlDocument);
      this.zip.file(relationships.relsPath, xml);
    }

    await syncContentTypesXml(this.zip);

    return this.zip.generateAsync({ type: "nodebuffer" });
  }

  async saveAs(outputPath) {
    const buffer = await this.toBuffer();
    await fs.writeFile(outputPath, buffer);
    return outputPath;
  }
}

function collectControllersByType(doc, node, type, results) {
  if (node.type === type) results.push(doc.createControllerForNode(node.id));
  for (const child of node.children) {
    collectControllersByType(doc, child, type, results);
  }
}

function nextRelationshipId(relationships) {
  let max = 0;
  for (const relId of relationships.relationships.keys()) {
    const match = /^rId(\d+)$/.exec(relId);
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `rId${max + 1}`;
}

function makeRelativeTarget(partPath, absolutePath) {
  const partDir = partPath.split("/").slice(0, -1);
  const targetParts = absolutePath.split("/");
  while (partDir.length > 0 && targetParts.length > 0 && partDir[0] === targetParts[0]) {
    partDir.shift();
    targetParts.shift();
  }
  return `${partDir.map(() => "..").join("/")}${partDir.length > 0 ? "/" : ""}${targetParts.join("/")}`;
}

function extensionForContentType(contentType) {
  if (contentType === "image/png") return "png";
  if (contentType === "image/jpeg") return "jpg";
  if (contentType === "image/gif") return "gif";
  return "bin";
}

function syncRelationshipsXml(relationships) {
  const documentElement = relationships.xmlDocument.documentElement;
  while (documentElement.firstChild) {
    documentElement.removeChild(documentElement.firstChild);
  }
  const sorted = Array.from(relationships.relationships.values()).sort((a, b) => a.id.localeCompare(b.id));
  for (const relationship of sorted) {
    const node = relationships.xmlDocument.createElementNS(documentElement.namespaceURI, "Relationship");
    node.setAttribute("Id", relationship.id);
    node.setAttribute("Type", relationship.type);
    node.setAttribute("Target", relationship.target);
    if (relationship.targetMode) {
      node.setAttribute("TargetMode", relationship.targetMode);
    }
    documentElement.appendChild(node);
  }
}

async function loadSupportedPartDescriptors(zip) {
  const fileNames = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const descriptors = [];
  if (zip.file(MAIN_DOCUMENT_PATH)) descriptors.push({ path: MAIN_DOCUMENT_PATH, type: "body" });
  for (const path of fileNames.filter((name) => /^word\/header\d+\.xml$/.test(name))) descriptors.push({ path, type: "header" });
  for (const path of fileNames.filter((name) => /^word\/footer\d+\.xml$/.test(name))) descriptors.push({ path, type: "footer" });
  for (const path of fileNames.filter((name) => /^word\/footnotes\d*\.xml$/.test(name))) descriptors.push({ path, type: "footnotes" });
  for (const path of fileNames.filter((name) => /^word\/endnotes\d*\.xml$/.test(name))) descriptors.push({ path, type: "endnotes" });
  for (const path of fileNames.filter((name) => /^word\/comments\d*\.xml$/.test(name))) descriptors.push({ path, type: "comments" });
  return descriptors;
}

async function syncContentTypesXml(zip) {
  const file = zip.file(CONTENT_TYPES_PATH);
  if (!file) return;

  const xmlDocument = parseXmlString(await file.async("string"));
  const documentElement = xmlDocument.documentElement;
  const files = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const defaults = new Map();

  for (const node of Array.from(documentElement.getElementsByTagName("Default"))) {
    defaults.set((node.getAttribute("Extension") || "").toLowerCase(), node);
  }

  for (const name of files) {
    const extensionMatch = /\.([^.]+)$/.exec(name);
    if (!extensionMatch) continue;
    const extension = extensionMatch[1].toLowerCase();
    const contentType = defaultContentTypeForExtension(extension);
    if (!contentType || defaults.has(extension)) continue;

    const node = xmlDocument.createElementNS(documentElement.namespaceURI, "Default");
    node.setAttribute("Extension", extension);
    node.setAttribute("ContentType", contentType);
    documentElement.appendChild(node);
    defaults.set(extension, node);
  }

  const xml = new XMLSerializer().serializeToString(xmlDocument);
  zip.file(CONTENT_TYPES_PATH, xml);
}

function defaultContentTypeForExtension(extension) {
  if (extension === "png") return "image/png";
  if (extension === "jpg" || extension === "jpeg") return "image/jpeg";
  if (extension === "gif") return "image/gif";
  return null;
}

module.exports = {
  CONTENT_TYPES_PATH,
  MAIN_DOCUMENT_PATH,
  VirtualWordDocument,
};
