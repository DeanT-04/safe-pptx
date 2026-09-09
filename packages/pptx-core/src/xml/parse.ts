import { DOMParser, XMLSerializer, MIME_TYPE } from '@xmldom/xmldom';
import type { Document, Element, Node } from '@xmldom/xmldom';
import { NS, type NsPrefix } from './namespaces.js';

export type { Document, Element, Node };

export class XmlParseError extends Error {
  constructor(message: string, readonly partPath: string) {
    super(`XML parse error in ${partPath}: ${message}`);
  }
}

export function parseXml(text: string, partPath = '<unknown>'): Document {
  try {
    const doc = new DOMParser({
      onError: (severity, message) => {
        if (severity === 'fatalError') {
          throw new XmlParseError(message, partPath);
        }
      },
    }).parseFromString(text, MIME_TYPE.XML_TEXT);
    if (!doc.documentElement) {
      throw new XmlParseError('no root element', partPath);
    }
    return doc;
  } catch (err) {
    if (err instanceof XmlParseError) throw err;
    throw new XmlParseError(err instanceof Error ? err.message : String(err), partPath);
  }
}

export function serializeXml(node: Node): string {
  return new XMLSerializer().serializeToString(node);
}

/**
 * Direct children of `node` matching the qualified name. Null-tolerant.
 * Comparison uses namespaceURI + localName — xmldom's `tagName` keeps the
 * source prefix (e.g. `p:sldIdLst`), which may differ across documents.
 */
export function childrenByName(node: Node | null, prefix: NsPrefix, local: string): Element[] {
  if (!node) return [];
  const out: Element[] = [];
  const uri = NS[prefix];
  let child = node.firstChild;
  while (child) {
    if (
      child.nodeType === 1 &&
      (child as Element).localName === local &&
      (child as Element).namespaceURI === uri
    ) {
      out.push(child as Element);
    }
    child = child.nextSibling;
  }
  return out;
}

export function firstChildByName(node: Node | null, prefix: NsPrefix, local: string): Element | null {
  return childrenByName(node, prefix, local)[0] ?? null;
}

/** Descendants matching the qualified name (document order). Null-tolerant. */
export function descendantsByName(node: Node | null, prefix: NsPrefix, local: string): Element[] {
  if (!node) return [];
  const out: Element[] = [];
  const uri = NS[prefix];
  const walk = (n: Node): void => {
    let child = n.firstChild;
    while (child) {
      if (child.nodeType === 1) {
        if ((child as Element).localName === local && (child as Element).namespaceURI === uri) {
          out.push(child as Element);
        }
        walk(child);
      }
      child = child.nextSibling;
    }
  };
  walk(node);
  return out;
}

export function firstDescendantByName(node: Node | null, prefix: NsPrefix, local: string): Element | null {
  return descendantsByName(node, prefix, local)[0] ?? null;
}

/** Concatenated text of `el`'s direct child text nodes (not descendants). */
export function directText(el: Node | null): string {
  if (!el) return '';
  let out = '';
  let child = el.firstChild;
  while (child) {
    if (child.nodeType === 3 || child.nodeType === 4) out += child.nodeValue ?? '';
    child = child.nextSibling;
  }
  return out;
}

/** Concatenated text of all descendant text nodes. */
export function deepText(el: Node | null): string {
  if (!el) return '';
  let out = '';
  const walk = (n: Node): void => {
    let child = n.firstChild;
    while (child) {
      if (child.nodeType === 3 || child.nodeType === 4) out += child.nodeValue ?? '';
      else walk(child);
      child = child.nextSibling;
    }
  };
  walk(el);
  return out;
}

export function getAttr(el: Element, prefix: NsPrefix, local: string): string | null {
  return el.getAttributeNS(NS[prefix], local);
}

/** Attribute with no namespace (most OOXML attrs are unnamespaced). */
export function getAttrNs(el: Element, local: string): string | null {
  return el.getAttribute(local);
}

export function setAttrNs(el: Element, local: string, value: string): void {
  el.setAttribute(local, value);
}

export function hasName(el: Element, prefix: NsPrefix, local: string): boolean {
  return el.localName === local && el.namespaceURI === NS[prefix];
}

/** Create `<prefix:local>text</prefix:local>` under `doc` (or detached if doc omitted). */
export function makeElement(doc: Document, prefix: NsPrefix, local: string, text?: string): Element {
  const el = doc.createElementNS(NS[prefix], `${prefix}:${local}`);
  if (text !== undefined) el.appendChild(doc.createTextNode(text));
  return el;
}
