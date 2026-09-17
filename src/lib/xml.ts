/**
 * XML layer — adapter cho 2 backend:
 *  - Trình duyệt: DOMParser/XMLSerializer native (chuẩn W3C, round-trip an toàn).
 *  - Node/CLI: @xmldom/xmldom (không có native DOM).
 *
 * Mọi module khác chỉ dùng kiểu XNode/XElement/XDocument (structural).
 * KHÔNG dựng lại document từ text — giữ nguyên OMML math, ảnh, định dạng run.
 */

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** Kiểu DOM chung (structural) — đủ cho các thao tác mình dùng */
export interface XNode {
  nodeType: number;
  nodeName: string;
  nodeValue: string | null;
  childNodes: XNode[];
  firstChild: XNode | null;
  nextSibling: XNode | null;
  parentNode: XNode | null;
  ownerDocument: XDocument;
  appendChild(n: XNode): XNode;
  removeChild(n: XNode): XNode;
  insertBefore(n: XNode, ref: XNode | null): XNode;
  cloneNode(deep?: boolean): XNode;
}

export interface XElement extends XNode {
  setAttributeNS(ns: string | null, name: string, value: string): void;
  getAttributeNS?(ns: string | null, name: string): string | null;
  /** setAttribute — dùng cho attribute có prefix OOXML (w:val, xml:space...) */
  setAttribute(name: string, value: string): void;
}

/** Kiểu lỏng cho errorHandler của xmldom */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type NeverErrorHandler = (level: unknown, msg: unknown) => unknown;

export interface XDocument extends XNode {
  documentElement: XElement | null;
  createElementNS(ns: string, name: string): XElement;
  createTextNode(text: string): XNode;
}

const IS_NATIVE = typeof (globalThis as { DOMParser?: unknown }).DOMParser !== "undefined";

export function isNativeDom(): boolean {
  return IS_NATIVE;
}

/** Cast node tự nhiên (native/xmldom) về XNode */
export function asXNode(n: unknown): XNode {
  return n as unknown as XNode;
}

export function asXElement(n: unknown): XElement {
  return n as unknown as XElement;
}

export function asXDocument(n: unknown): XDocument {
  return n as unknown as XDocument;
}

export function parseXml(xml: string): XDocument {
  if (IS_NATIVE) {
    const doc = new (globalThis as any).DOMParser().parseFromString(xml, "application/xml");
    const err = doc.querySelector("parsererror");
    if (err) throw new Error("Lỗi XML: " + (err.textContent || "").slice(0, 200));
    return asXDocument(doc);
  }
  const { DOMParser: XDOMParser } = require("@xmldom/xmldom") as typeof import("@xmldom/xmldom");
  const doc = new XDOMParser({
    // xmldom errorHandler nhận (level, msg); bỏ qua để không in lỗi vào console
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    errorHandler: ((level: unknown, msg: unknown) => {
      void level;
      void msg;
    }) as unknown as NonNullable<ConstructorParameters<typeof XDOMParser>[0]>["errorHandler"],
  }).parseFromString(xml, "application/xml");
  return asXDocument(doc);
}

export function serializeXml(doc: XDocument): string {
  // KHÔNG strip namespace: xmldom có thể chèn xmlns:prefix cục bộ trên từng element
  // (mỗi khối drawing một bản). Strip toàn file theo "bản đầu tiên" làm mất khai báo
  // của các khối sau → XML hỏng. Word/lxml chấp nhận khai báo lặp.
  return IS_NATIVE
    ? new (globalThis as any).XMLSerializer().serializeToString(asXNode(doc))
    : new (require("@xmldom/xmldom").XMLSerializer)().serializeToString(asXNode(doc));
}

/**
 * Dọn namespace trùng (giá trị giống) khai báo trên element con.
 * Chỉ cần cho backend xmldom; native serializer đã chuẩn.
 */
export function stripRedundantXmlns(xml: string): string {
  const seen = new Set<string>();
  return xml.replace(/ xmlns(:\w+)?="([^"]*)"/g, (m, prefix: string | undefined, value: string) => {
    const key = `${prefix || ""}|${value}`;
    if (seen.has(key)) return "";
    seen.add(key);
    return m;
  });
}

export function localName(node: XNode): string {
  return (node.nodeName || "").replace(/^.*:/, "");
}

export function isW(node: XNode | null | undefined, name: string): boolean {
  if (!node || !node.nodeName) return false;
  const ln = localName(node);
  // Với namespace đúng: node có thể là {w:nodeName} dạng "w:p" — kiểm tra localName
  return ln === name || node.nodeName === name;
}

export function childrenOf(node: XNode | null): XElement[] {
  const out: XElement[] = [];
  if (!node || !node.childNodes) return out;
  for (let i = 0; i < node.childNodes.length; i++) {
    const c = node.childNodes[i];
    if (c.nodeType === 1) out.push(c as XElement);
  }
  return out;
}

/** Nối text của paragraph: w:t giữ nguyên, w:tab → \t, w:br|w:cr → \n */
export function paragraphText(p: XElement): string {
  let out = "";
  const walk = (n: XNode) => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      const ln = localName(c);
      if (c.nodeType === 3) out += c.nodeValue ?? "";
      else if (ln === "tab") out += "\t";
      else if (ln === "br" || ln === "cr") out += "\n";
      else walk(c);
    }
  };
  walk(p);
  return out;
}

/** Paragraph có chứa công thức toán OMML? */
export function hasMath(p: XElement): boolean {
  const walk = (n: XNode): boolean => {
    const ln = localName(n);
    if (ln === "oMath" || ln === "oMathPara") return true;
    for (let i = 0; i < n.childNodes.length; i++) {
      if (walk(n.childNodes[i])) return true;
    }
    return false;
  };
  return walk(p);
}

export function setTextNodeText(el: XElement, text: string): void {
  const txt = el.ownerDocument.createTextNode(text);
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(txt);
}

/** Tạo element w:name với namespace chuẩn */
export function createWElement(doc: XDocument, name: string): XElement {
  return doc.createElementNS(W_NS, `w:${name}`);
}