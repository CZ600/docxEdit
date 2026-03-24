"use strict";

const fs = require("node:fs");
const path = require("node:path");

function findSampleDocx() {
  const projectRoot = path.resolve(__dirname, "..", "..");
  const fileName = fs
    .readdirSync(projectRoot)
    .find((name) => name.endsWith(".docx") && !name.endsWith(".modified.docx"));

  if (!fileName) {
    throw new Error("Unable to locate the sample .docx file in the project root.");
  }

  return path.join(projectRoot, fileName);
}

module.exports = {
  findSampleDocx,
};
