"use strict";

const { DOMParser } = require("@xmldom/xmldom");

const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const PNG_SIGNATURE = "89504e47";
const JPG_SIGNATURE = "ffd8ff";
const GIF_SIGNATURE = "47494638";

function getRelsPath(partPath) {
  const segments = partPath.split("/");
  const fileName = segments.pop();
  return `${segments.join("/")}/_rels/${fileName}.rels`;
}

function loadRelationships(partPath, xml) {
  const relsPath = getRelsPath(partPath);
  const xmlDocument = new DOMParser().parseFromString(
    xml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"></Relationships>`,
    "application/xml",
  );

  const relationships = new Map();
  const nodes = Array.from(xmlDocument.getElementsByTagName("Relationship"));
  for (const node of nodes) {
    relationships.set(node.getAttribute("Id"), {
      id: node.getAttribute("Id"),
      type: node.getAttribute("Type"),
      target: node.getAttribute("Target"),
      targetMode: node.getAttribute("TargetMode") || null,
    });
  }

  return {
    partPath,
    relsPath,
    xmlDocument,
    relationships,
  };
}

function resolveRelationshipTarget(partPath, target) {
  if (!target) {
    return null;
  }

  if (target.startsWith("/")) {
    return target.replace(/^\//, "");
  }

  const segments = partPath.split("/");
  segments.pop();
  const baseDir = segments.join("/");
  const combined = `${baseDir}/${target}`;
  return normalizePath(combined);
}

function createImageDescriptor({ relId, relationship, zip }) {
  if (!relationship || relationship.type !== IMAGE_REL_TYPE || relationship.targetMode) {
    return null;
  }

  const mediaPath = resolveRelationshipTarget(relationship.partPath, relationship.target);
  const file = zip.file(mediaPath);
  if (!file) {
    return null;
  }

  const buffer = file._data && file._data.compressedContent
    ? Buffer.from(file._data.compressedContent)
    : null;

  const extension = getExtension(mediaPath);
  return {
    relId,
    target: relationship.target,
    mediaPath,
    filename: mediaPath.split("/").pop(),
    extension,
    contentType: detectContentType(file.name, buffer),
    data: null,
  };
}

async function readImageBuffer(zip, mediaPath) {
  const file = zip.file(mediaPath);
  if (!file) {
    return null;
  }

  return file.async("nodebuffer");
}

function ensureImageBuffer(value) {
  if (Buffer.isBuffer(value)) {
    return value;
  }

  if (typeof value === "string") {
    return Buffer.from(value, "base64");
  }

  throw new Error("Image data must be a Buffer or base64 string.");
}

function detectContentType(filename, buffer) {
  const lower = filename ? filename.toLowerCase() : "";
  if (lower.endsWith(".png")) {
    return "image/png";
  }
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (lower.endsWith(".gif")) {
    return "image/gif";
  }

  if (buffer && buffer.length >= 4) {
    const header = buffer.subarray(0, 4).toString("hex");
    if (header === PNG_SIGNATURE) {
      return "image/png";
    }
    if (header.startsWith(JPG_SIGNATURE)) {
      return "image/jpeg";
    }
    if (header === GIF_SIGNATURE) {
      return "image/gif";
    }
  }

  return "application/octet-stream";
}

function extensionFromContentType(contentType) {
  if (contentType === "image/png") {
    return "png";
  }
  if (contentType === "image/jpeg") {
    return "jpg";
  }
  if (contentType === "image/gif") {
    return "gif";
  }
  return "bin";
}

function getExtension(path) {
  const match = /\.([^.]+)$/.exec(path || "");
  return match ? match[1].toLowerCase() : "";
}

function normalizePath(path) {
  const output = [];
  for (const part of path.split("/")) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      output.pop();
      continue;
    }
    output.push(part);
  }
  return output.join("/");
}

module.exports = {
  IMAGE_REL_TYPE,
  createImageDescriptor,
  detectContentType,
  ensureImageBuffer,
  extensionFromContentType,
  getRelsPath,
  loadRelationships,
  readImageBuffer,
  resolveRelationshipTarget,
};
