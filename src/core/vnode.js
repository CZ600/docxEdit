"use strict";

let nextId = 1;

class VNode {
  constructor({ id = null, key = null, type, props = {}, children = [], source = null, parent = null }) {
    this.id = id || createVNodeId();
    this.key = key;
    this.type = type;
    this.props = props;
    this.children = [];
    this.source = source;
    this.parent = parent;

    for (const child of children) {
      this.appendChild(child);
    }
  }

  appendChild(child) {
    if (!child) {
      return;
    }

    child.parent = this;
    this.children.push(child);
  }

  toJSON() {
    return {
      id: this.id,
      key: this.key,
      type: this.type,
      props: clonePlainValue(this.props),
      children: this.children.map((child) =>
        child && typeof child.toJSON === "function" ? child.toJSON() : child,
      ),
    };
  }
}

function createVNodeId() {
  const id = nextId;
  nextId += 1;
  return id;
}

function createVNode(definition, { preserveId = true } = {}) {
  const normalizedChildren = (definition.children || []).map((child) =>
    child instanceof VNode ? child : createVNode(child, { preserveId }),
  );
  const resolvedKey =
    definition.key !== undefined && definition.key !== null
      ? definition.key
      : definition.source && typeof definition.source === "object"
        ? definition.source.__vnodeKey ?? null
        : null;
  const vnode = new VNode({
    id: preserveId ? definition.id : null,
    key: resolvedKey,
    type: definition.type,
    props: clonePlainValue(definition.props || {}),
    children: normalizedChildren,
    source: definition.source || null,
    parent: definition.parent || null,
  });

  if (vnode.source && typeof vnode.source === "object") {
    vnode.source.__vnode = vnode;
    vnode.source.__vnodeId = vnode.id;
  }

  return vnode;
}

function cloneVNode(vnode) {
  return createVNode(vnode.toJSON());
}

function clonePlainValue(value) {
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(value)) {
    return Buffer.from(value);
  }

  if (Array.isArray(value)) {
    return value.map((item) => clonePlainValue(item));
  }

  if (value && typeof value === "object") {
    const cloned = {};
    for (const [key, innerValue] of Object.entries(value)) {
      cloned[key] = clonePlainValue(innerValue);
    }
    return cloned;
  }

  return value;
}

function assignVNodeSource(vnode, source) {
  vnode.source = source;

  if (source && typeof source === "object") {
    source.__vnode = vnode;
    source.__vnodeId = vnode.id;
    source.__vnodeKey = vnode.key;
  }
}

function visitVNode(vnode, visitor, parent = null) {
  if (!vnode) {
    return;
  }

  visitor(vnode, parent);

  for (const child of vnode.children) {
    visitVNode(child, visitor, vnode);
  }
}

module.exports = {
  VNode,
  assignVNodeSource,
  cloneVNode,
  clonePlainValue,
  createVNode,
  createVNodeId,
  visitVNode,
};
