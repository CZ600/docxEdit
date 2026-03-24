"use strict";

let nextId = 1;

class VNode {
  constructor({ type, props = {}, children = [], source = null }) {
    this.id = nextId;
    nextId += 1;
    this.type = type;
    this.props = props;
    this.children = children;
    this.source = source;
  }

  toJSON() {
    return {
      id: this.id,
      type: this.type,
      props: this.props,
      children: this.children.map((child) =>
        child && typeof child.toJSON === "function" ? child.toJSON() : child,
      ),
    };
  }
}

function createVNode(definition) {
  const vnode = new VNode(definition);

  if (vnode.source && typeof vnode.source === "object") {
    vnode.source.__vnode = vnode;
  }

  return vnode;
}

module.exports = {
  VNode,
  createVNode,
};
