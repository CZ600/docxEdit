"use strict";

const SKIP_TYPES = new Set(["bookmarkStart", "bookmarkEnd", "sectPr", "tblGrid", "tblPr", "trPr", "tcPr", "gridCol"]);

function escapeHtml(text) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function extractHtml(doc, { partTypes = ["body"] } = {}) {
  const parts = [];

  for (const partType of partTypes) {
    let partList;
    if (partType === "body") {
      const body = doc.getBody();
      partList = body ? [body] : [];
    } else if (partType === "headers") {
      partList = doc.getHeaders();
    } else if (partType === "footers") {
      partList = doc.getFooters();
    } else {
      partList = [];
    }

    for (const part of partList) {
      for (const child of part.vnode.children) {
        if (SKIP_TYPES.has(child.type)) continue;
        const html = renderNode(doc, child);
        if (html) parts.push(html);
      }
    }
  }

  return parts.join("\n");
}

function renderNode(doc, node) {
  switch (node.type) {
    case "paragraph":
      return renderParagraph(doc, node);
    case "table":
      return renderTable(doc, node);
    default:
      return "";
  }
}

function renderParagraph(doc, node) {
  const style = node.props.style || {};
  const headingLevel = doc.resolveHeadingLevel(style.styleId);
  const content = renderInlineChildren(doc, node.children);

  if (headingLevel && headingLevel <= 6) {
    return `<h${headingLevel}>${content}</h${headingLevel}>`;
  }
  return `<p>${content}</p>`;
}

function renderInlineChildren(doc, children) {
  let html = "";
  for (const child of children) {
    switch (child.type) {
      case "run":
        html += renderRun(doc, child);
        break;
      case "math":
        html += renderMath(child);
        break;
      case "image":
        html += renderImage(child);
        break;
      case "tab":
        html += "\t";
        break;
      case "break":
        html += "<br>";
        break;
      case "hyperlink":
        html += renderInlineChildren(doc, child.children);
        break;
      default:
        break;
    }
  }
  return html;
}

function renderRun(doc, node) {
  const style = node.props.style || {};
  let text = "";

  for (const child of node.children) {
    switch (child.type) {
      case "text":
        text += escapeHtml(child.props.text || "");
        break;
      case "tab":
        text += "\t";
        break;
      case "break":
        text += "<br>";
        break;
      case "math":
        text += renderMath(child);
        break;
      case "image":
        text += renderImage(child);
        break;
      default:
        break;
    }
  }

  if (style.strike) text = `<s>${text}</s>`;
  if (style.underline && style.underline !== "none") text = `<u>${text}</u>`;
  if (style.italic) text = `<em>${text}</em>`;
  if (style.bold) text = `<strong>${text}</strong>`;

  return text;
}

function renderMath(node) {
  const text = escapeHtml(node.props.text || "");
  const display = node.props.display || "inline";
  if (display === "block") {
    return `<div class="math-block">${text}</div>`;
  }
  return `<span class="math">${text}</span>`;
}

function renderImage(node) {
  const alt = escapeHtml(node.props.alt || "");
  return `<img alt="${alt}" />`;
}

function renderTable(doc, node) {
  const rows = node.children.filter((c) => c.type === "table-row");
  let html = "<table>";

  for (const row of rows) {
    html += "<tr>";
    const cells = row.children.filter((c) => c.type === "table-cell");
    for (const cell of cells) {
      const cellStyle = cell.props.style || {};
      const attrs = [];
      if (cellStyle.gridSpan) attrs.push(` colspan="${cellStyle.gridSpan}"`);

      const paragraphs = cell.children.filter((c) => c.type === "paragraph");
      const content = paragraphs.map((p) => renderInlineChildren(doc, p.children)).join("<br>");

      html += `<td${attrs.join("")}>${content}</td>`;
    }
    html += "</tr>";
  }

  html += "</table>";
  return html;
}

module.exports = { extractHtml };
