import { SaxesParser, type SaxesTagNS } from "saxes";
import { LIMITS } from "./zip";
export const W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export type XmlNode = {
  name: string;
  local: string;
  uri: string;
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  text: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  parent?: XmlNode;
  special: boolean;
};
export function parseXml(xml: string): XmlNode {
  if (xml.length > LIMITS.xml)
    throw new Error("XML part exceeds the 8 MB limit");
  if (/<!\s*(DOCTYPE|ENTITY)/i.test(xml))
    throw new Error("XML DTDs and entities are not allowed");
  const parser = new SaxesParser({ xmlns: true, position: true });
  let root: XmlNode | undefined;
  const stack: XmlNode[] = [];
  let count = 0;
  parser.on("opentag", (tag: SaxesTagNS) => {
    if (++count > 150000 || stack.length > 128)
      throw new Error("XML complexity exceeds safety limits");
    const node: XmlNode = {
      name: tag.name,
      local: tag.local,
      uri: tag.uri,
      start: xml.lastIndexOf("<", parser.position - 1),
      openEnd: parser.position,
      closeStart: parser.position,
      end: parser.position,
      text: "",
      attrs: {},
      children: [],
      parent: stack.at(-1),
      special: false,
    };
    for (const a of Object.values(tag.attributes))
      node.attrs[`{${a.uri}}${a.local}`] = a.value;
    if (node.parent) node.parent.children.push(node);
    else root = node;
    stack.push(node);
  });
  parser.on("text", (s) => {
    if (stack.length) stack.at(-1)!.text += s;
  });
  parser.on("cdata", (s) => {
    if (stack.length) stack.at(-1)!.text += s;
  });
  parser.on("comment", () => {
    for (const n of stack) n.special = true;
  });
  parser.on("processinginstruction", () => {
    for (const n of stack) n.special = true;
  });
  parser.on("doctype", () => {
    throw new Error("XML DTDs are not allowed");
  });
  parser.on("closetag", (tag: SaxesTagNS) => {
    const n = stack.pop()!;
    n.end = parser.position;
    n.closeStart = tag.isSelfClosing
      ? n.openEnd
      : xml.lastIndexOf("</", parser.position - 1);
  });
  parser.write(xml).close();
  if (!root) throw new Error("Empty XML part");
  return root;
}
export function walk(root: XmlNode): XmlNode[] {
  const nodes: XmlNode[] = [];
  const todo = [root];
  while (todo.length) {
    const n = todo.pop()!;
    nodes.push(n);
    todo.push(...[...n.children].reverse());
  }
  return nodes;
}
export const isW = (n: XmlNode, local: string) =>
  n.uri === W_NS && n.local === local;
