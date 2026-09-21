/**
 * Tự động thêm tag "[HOCMAI] " sau nhãn "Câu N." cho câu hỏi CHƯA chuẩn hóa.
 *
 * Chuẩn (theo mẫu Bài 13 — bản PHT/Live đã chuẩn hóa):
 *   "Câu 1. [HOCMAI] Nhiệt dung riêng có đơn vị đo là ..."
 * Tag là run RIÊNG: bold + màu template (đỏ C00000 cho TSA), theo sau nhãn.
 *
 * Quy tắc:
 *  - Chỉ xử lý paragraph phân loại "question" (bắt đầu "Câu N." / "Câu N:").
 *  - Nếu sau nhãn (bỏ qua khoảng trắng) đã có "[HOCMAI]" → bỏ qua (đã chuẩn).
 *  - Chèn run mới NGAY SAU dấu chấm của nhãn; nếu nhãn nằm giữa một run
 *    (Word tách run tùy ý), split run đó làm 2, chèn tag giữa hai phần.
 *  - Không đụng OMML math, drawing, bảng — chỉ thao tác run chứa chữ nhãn.
 */

import { localName, childrenOf, isW, createWElement, paragraphText } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";

/** Nhãn "Câu N." / "Câu N:" — tìm KHÔNG anchor đầu (paragraph có tab đầu) */
const LABEL_RE = /Câu\s*(\d{1,3})\s*[:.]/;

/** Đã có tag ngay sau nhãn? (chấp nhận khoảng trắng giữa) */
const HAS_TAG_RE = /^[\s\u00A0]*\[HOCMAI\]/;

/** Text (logical) của một run w:r — GIỐNG paragraphText (tính cả tab/br/space) */
function runText(r: XElement): string {
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
  walk(r);
  return out;
}

/** Clone run nhưng chỉ giữ rPr + phần KHÔNG phải text (an toàn cho t/math khác) */
function cloneRunShell(src: XElement, doc: XDocument): XElement {
  const r = createWElement(doc, "r");
  for (const c of childrenOf(src)) {
    const ln = localName(c);
    if (ln === "t" || ln === "tab" || ln === "br" || ln === "cr") continue;
    r.appendChild(c.cloneNode(true));
  }
  return r;
}

/** Run mới với text chỉ định, giữ rPr của run nguồn */
function runWithText(src: XElement, doc: XDocument, text: string): XElement {
  const r = cloneRunShell(src, doc);
  const t = createWElement(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

/** Run tag "[HOCMAI] " — bold + màu template (giống mẫu chuẩn hóa) */
function makeTagRun(src: XElement, doc: XDocument, color: string): XElement {
  const r = cloneRunShell(src, doc);
  let rPr = childrenOf(r).find((c) => isW(c, "rPr"));
  if (!rPr) {
    rPr = createWElement(doc, "rPr");
    r.insertBefore(rPr, r.firstChild);
  }
  const b = createWElement(doc, "b");
  const bCs = createWElement(doc, "bCs");
  const col = createWElement(doc, "color");
  col.setAttribute("w:val", color);
  // b/bCs/color phải nằm TRƯỚC sz trong rPr theo schema Word (order)
  // rPr gốc thường có rFonts, b?, color?, sz... — chèn sau rFonts nếu có
  const fonts = childrenOf(rPr).find((c) => isW(c, "rFonts"));
  if (fonts && fonts.nextSibling) {
    rPr.insertBefore(b, fonts.nextSibling as XNode);
    rPr.insertBefore(bCs, b.nextSibling as XNode);
    rPr.insertBefore(col, bCs.nextSibling as XNode);
  } else {
    rPr.appendChild(b);
    rPr.appendChild(bCs);
    rPr.appendChild(col);
  }
  const t = createWElement(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode("[HOCMAI] "));
  r.appendChild(t);
  return r;
}

/**
 * Thêm tag "[HOCMAI] " cho mọi câu hỏi chưa có tag.
 * Thao tác trực tiếp trên body; trả về số tag đã thêm.
 */
export function addHocmaiTags(
  blocks: DocBlock[],
  body: XElement,
  doc: XDocument,
  color: string,
): number {
  // Duyệt song song body + blocks để có tham chiếu element chính xác
  let blkIdx = 0;
  let added = 0;
  const children = childrenOf(body);
  for (const el of children) {
    if (blkIdx >= blocks.length) break;
    const b = blocks[blkIdx];
    const elIsP = isW(el, "p");
    const elIsTbl = isW(el, "tbl");
    if (b.kind === "table") {
      blkIdx++;
      continue;
    }
    // b là paragraph — el tương ứng phải là paragraph
    if (elIsP) {
      if (b.kind === "question") {
        added += tagOneParagraph(el as XElement, doc, color);
      }
      blkIdx++;
    } else if (elIsTbl) {
      // không nên xảy ra (block table đã xử lý trên), tịnh tiến để giữ đồng bộ
      blkIdx++;
    }
  }
  return added;
}

/**
 * Thêm tag cho MỘT paragraph câu hỏi; trả 1 nếu đã thêm, 0 nếu đã có tag/skip.
 *
 * Span được tính trên MỌI phần tử trực tiếp của paragraph (dùng paragraphText
 * cho từng element con) để khớp CHÍNH XÁC chuỗi raw của paragraphText(p) —
 * nếu chỉ đếm run trực tiếp sẽ lệch khi có tab đầu / hyperlink / bookmark
 * (lỗi từng xảy ra: tag chèn giữa chữ "(Tr[HOCMAI]ích").
 */
function tagOneParagraph(p: XElement, doc: XDocument, color: string): number {
  const raw = paragraphText(p);
  const m = LABEL_RE.exec(raw);
  if (!m) return 0;
  let end = m.index + m[0].length; // vị trí logical sau dấu chấm nhãn

  // Đã có tag [HOCMAI] ngay sau nhãn (bỏ qua khoảng trắng)?
  if (HAS_TAG_RE.test(raw.slice(end))) return 0;

  // Mở rộng qua whitespace liền sau nhãn để tag đứng ngay sau khoảng trắng
  // (khớp mẫu "Câu 1. [HOCMAI] ..." — space nằm run riêng trong file GV).
  while (end < raw.length && (raw[end] === " " || raw[end] === "\t" || raw[end] === "\u00A0")) {
    end++;
  }

  // Spans của từng CON TRỰC TIẾP (element/text) — khớp paragraphText
  const children = childrenOf(p);
  type Span = { st: number; en: number; el: XElement; isRun: boolean };
  const spans: Span[] = [];
  let pos = 0;
  for (const c of children) {
    const isRun = isW(c, "r");
    if (isRun) {
      const t = runText(c);
      spans.push({ st: pos, en: pos + t.length, el: c, isRun });
      pos += t.length;
    } else {
      // element khác (bookmark, hyperlink, pPr...) — đóng góp 0 ký tự text
      // (pPr/bookmark không có w:t). w:hyperlink có w:t nhưng hiếm ở nhãn.
      const t = paragraphText(c);
      if (t.length > 0) {
        spans.push({ st: pos, en: pos + t.length, el: c, isRun });
        pos += t.length;
      }
    }
  }
  if (end > pos) return 0;

  // Tìm phần tử chứa vị trí end
  let si = -1;
  for (let i = 0; i < spans.length; i++) {
    if (end <= spans[i].en) { si = i; break; }
  }
  if (si < 0) return 0;
  const sp = spans[si];
  const el = sp.el;
  const tagRun = makeTagRun(el, doc, color);

  // Chèn chính xác theo ranh giới
  if (end === sp.st) {
    // ngay đầu phần tử → chèn trước
    p.insertBefore(tagRun, el);
    return 1;
  }
  if (!sp.isRun) {
    // end nằm GIỮA phần tử không phải run (hyperlink...) → bỏ qua an toàn
    return 0;
  }
  const r = el;
  if (end === sp.en) {
    // cuối run → chèn ngay sau
    if (r.nextSibling) p.insertBefore(tagRun, r.nextSibling as XNode);
    else p.appendChild(tagRun);
    return 1;
  }

  // end nằm giữa run → split: [st,end) giữ format, [end,en) run mới
  const txt = runText(r);
  const splitAt = end - sp.st;
  const headRun = runWithText(r, doc, txt.slice(0, splitAt));
  const tailRun = runWithText(r, doc, txt.slice(splitAt));
  p.replaceChild(headRun, r);
  p.insertBefore(tailRun, headRun.nextSibling as XNode);
  p.insertBefore(tagRun, tailRun);
  return 1;
}