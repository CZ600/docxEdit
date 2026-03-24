"use strict";

function countStringOccurrences(text, searchValue) {
  if (searchValue.length === 0) {
    throw new Error("searchValue must not be empty.");
  }

  let count = 0;
  let cursor = 0;

  while (true) {
    const index = text.indexOf(searchValue, cursor);
    if (index === -1) {
      return count;
    }

    count += 1;
    cursor = index + searchValue.length;
  }
}

function replaceAllInText(text, searchValue, replacement) {
  if (searchValue instanceof RegExp) {
    if (!searchValue.global) {
      throw new Error("RegExp searchValue must use the global flag.");
    }

    let count = 0;
    const nextText = text.replace(searchValue, (...args) => {
      count += 1;
      return typeof replacement === "function" ? replacement(...args) : replacement;
    });

    return { count, text: nextText };
  }

  const count = countStringOccurrences(text, searchValue);
  return {
    count,
    text: count === 0 ? text : text.split(searchValue).join(replacement),
  };
}

function replaceFirstInText(text, searchValue, replacement) {
  if (searchValue instanceof RegExp) {
    throw new Error("replace() currently supports string searchValue only. Use replaceAll() for RegExp.");
  }

  if (typeof searchValue !== "string" || searchValue.length === 0) {
    throw new Error("searchValue must be a non-empty string.");
  }

  const index = text.indexOf(searchValue);
  if (index === -1) {
    return { count: 0, text };
  }

  return {
    count: 1,
    text: text.slice(0, index) + replacement + text.slice(index + searchValue.length),
  };
}

module.exports = {
  replaceAllInText,
  replaceFirstInText,
};
