/**
 * Tự động chuẩn hóa tag "[HOCMAI]" cho câu hỏi.
 *
 * Chuẩn (theo mẫu PHT/Live Bài 13):
 *   "Câu 1. [HOCMAI] Nội dung câu hỏi ..."
 *
 * QUY TẮC BẤT BIẾN (yêu cầu chuẩn hóa tài liệu):
 *  1. KHÔNG hard-code màu. Tìm run [HOCMAI] CHUẨN trong chính tài liệu
 *     (run đầu tiên ngay sau nhãn "Câu N.") và CLONE toàn bộ rPr của nó
 *     (font, size, bold/italic, color RGB/theme, baseline, lang...).
 *  2. Nếu câu đã có [HOCMAI] nhưng SAI định dạng (sai màu) → CHUẨN HÓA:
 *     thay rPr của run tag hiện tại bằng bản clone từ run chuẩn.
 *  3. KHÔNG đổi màu/format của "Câu N." hay nội dung câu hỏi.
 *  4. Chỉ thao tác run — không đụng OMML math, drawing, bảng.
 */

import { localName, childrenOf, isW, createWElement, paragraphText } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";

/** Nhãn "Câu N." / "Câu N:" / "Câu N [HOCMAI]" — tìm KHÔNG anchor đầu (paragraph có tab đầu) */
const LABEL_RE = /Câu\s*(\d{1,3})\s*(?:[:.]|\s*\[(?:HOCMAI|Trắc nghiệm)[^\]]*\]|\s*$)/;

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

/** Text của cả paragraph (dùng để dò run chứa nhãn) */
function paraText(p: XElement): string {
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

/**
 * Tìm RUN [HOCMAI] chuẩn trong tài liệu:
 *  - ưu tiên run ngay sau nhãn "Câu N." (đúng vị trí mẫu);
 *  - nếu không có, lấy run [HOCMAI] đầu tiên bất kỳ (có rPr).
 * Trả về element run mẫu (hoặc null).
 */
export function findCanonicalHocmaiRun(blocks: DocBlock[], body: XElement): XElement | null {
  // 1) Duyệt các paragraph câu hỏi, tìm run [HOCMAI] ngay sau nhãn
  const children = childrenOf(body);
  for (const el of children) {
    if (!isW(el, "p")) continue;
    const p = el as XElement;
    const raw = paragraphText(p);
    const m = LABEL_RE.exec(raw);
    if (!m) continue;
    const after = raw.slice(m.index + m[0].length, m.index + m[0].length + 40);
    if (!/^\s*\[HOCMAI\]/.test(after)) continue;
    // tìm run chứa [HOCMAI] trong paragraph này (run trực tiếp)
    for (const r of childrenOf(p)) {
      if (!isW(r, "r")) continue;
      if (runText(r).includes("[HOCMAI]")) {
        const rPr = childrenOf(r).find((c) => isW(c, "rPr"));
        if (rPr) return r;
      }
    }
  }
  // 2) Fallback: run [HOCMAI] đầu tiên có rPr trong toàn body
  const walk = (n: XNode): XElement | null => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      if (isW(c, "r")) {
        if (runText(c as XElement).includes("[HOCMAI]")) {
          const rPr = childrenOf(c as XElement).find((x) => isW(x, "rPr"));
          if (rPr) return c as XElement;
        }
      }
      const found = walk(c);
      if (found) return found;
    }
    return null;
  };
  return walk(body);
}

/** Clone rPr của run mẫu sang run đích (thay toàn bộ rPr cũ) */
function cloneRPr(src: XElement, dst: XElement, doc: XDocument): void {
  const srcRPr = childrenOf(src).find((c) => isW(c, "rPr"));
  const dstRPr = childrenOf(dst).find((c) => isW(c, "rPr"));
  if (!srcRPr) return; // không có rPr mẫu → giữ nguyên
  const clone = srcRPr.cloneNode(true) as XElement;
  if (dstRPr) {
    dst.replaceChild(clone as XNode, dstRPr as XNode);
  } else {
    dst.insertBefore(clone as XNode, dst.firstChild as XNode | null);
  }
}

/** Run mới với text chỉ định, clone rPr từ run mẫu */
function runWithTextClone(src: XElement, doc: XDocument, text: string): XElement {
  const r = createWElement(doc, "r");
  // copy mọi thứ trừ rPr (rPr được set riêng) và text/tab/br
  for (const c of childrenOf(src)) {
    const ln = localName(c);
    if (ln === "t" || ln === "tab" || ln === "br" || ln === "cr") continue;
    r.appendChild(c.cloneNode(true));
  }
  cloneRPr(src, r, doc);
  const t = createWElement(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

/** Run tag "[HOCMAI]" (text + clone rPr từ run mẫu) — giữ space cuối như mẫu */
function makeTagRunClone(src: XElement, doc: XDocument): XElement {
  return runWithTextClone(src, doc, "[HOCMAI] ");
}

/**
 * Chuẩn hóa toàn bộ câu hỏi:
 *  - chưa có [HOCMAI] → thêm (clone rPr từ run chuẩn);
 *  - có nhưng sai rPr (khác mẫu: màu/font/size/bold...) → thay rPr bằng clone.
 * Trả về { added, normalized }.
 */
export function addHocmaiTags(
  blocks: DocBlock[],
  body: XElement,
  doc: XDocument,
  _color: string, // giữ API cũ; KHÔNG dùng nữa — clone từ tài liệu
): number {
  const canonical = findCanonicalHocmaiRun(blocks, body);
  if (!canonical) return addHocmaiTagsFallback(blocks, body, doc, _color);

  let blkIdx = 0;
  let fixed = 0;
  const children = childrenOf(body);
  for (const el of children) {
    if (blkIdx >= blocks.length) break;
    const b = blocks[blkIdx];
    if (b.kind === "table") { blkIdx++; continue; }
    if (isW(el, "p")) {
      if (b.kind === "question") {
        fixed += tagOneParagraph(el as XElement, doc, canonical);
      }
      blkIdx++;
    } else if (isW(el, "tbl")) {
      blkIdx++;
    }
  }
  return fixed;
}

/** Fallback khi không tìm thấy run mẫu: dùng màu template như cũ (giữ API) */
function addHocmaiTagsFallback(
  blocks: DocBlock[],
  body: XElement,
  doc: XDocument,
  color: string,
): number {
  const fallbackSrc = makeFallbackRun(doc, color);
  let blkIdx = 0;
  let fixed = 0;
  const children = childrenOf(body);
  for (const el of children) {
    if (blkIdx >= blocks.length) break;
    const b = blocks[blkIdx];
    if (b.kind === "table") { blkIdx++; continue; }
    if (isW(el, "p")) {
      if (b.kind === "question") {
        fixed += tagOneParagraph(el as XElement, doc, fallbackSrc);
      }
      blkIdx++;
    } else if (isW(el, "tbl")) {
      blkIdx++;
    }
  }
  return fixed;
}

function makeFallbackRun(doc: XDocument, color: string): XElement {
  const r = createWElement(doc, "r");
  const rPr = createWElement(doc, "rPr");
  rPr.setAttribute("w:val", "");
  rPr.setAttribute("x", "");
  const b = createWElement(doc, "b");
  const col = createWElement(doc, "color");
  col.setAttribute("w:val", color);
  rPr.appendChild(b);
  rPr.appendChild(col);
  r.appendChild(rPr);
  return r;
}

/**
 * Xử lý MỘT paragraph câu hỏi:
 *  - nếu chưa có tag → chèn run [HOCMAI] (clone rPr) sau nhãn;
 *  - nếu có tag nhưng rPr KHÁC mẫu → chuẩn hóa rPr của run tag đó.
 * Trả về 1 nếu đã sửa, 0 nếu không.
 */
function tagOneParagraph(p: XElement, doc: XDocument, canonical: XElement): number {
  const raw = paraText(p);
  const m = LABEL_RE.exec(raw);
  if (!m) return 0;
  // Điểm chèn tag: ngay SAU SỐ (trước "[Trắc nghiệm...]" nếu có) — đúng chuẩn
  // mẫu Bài 13/HSA: "Câu 1. [HOCMAI] ..." / "Câu 1 [HOCMAI] – (ID:...)".
  // Nếu nhãn là "Câu N [Trắc nghiệm ...]" → chèn [HOCMAI] TRƯỚC "[Trắc".
  let end = m.index + m[0].length;
  const trnIdx = m[0].indexOf("[Trắc");
  if (trnIdx >= 0) end = m.index + trnIdx;

  // Tìm run chứa [HOCMAI] ngay sau nhãn (trong khoảng whitespace 40 ký tự)
  const after = raw.slice(end, end + 60);
  const tagMatch = /^[\s\u00A0]*(\[HOCMAI\])/.exec(after);
  if (tagMatch) {
    // Đã có tag — kiểm tra rPr có khớp mẫu không; khớp → không làm gì
    const tagStart = end + tagMatch[1].length;
    // tìm run chứa vị trí tagStart
    const runEl = findRunAt(p, tagStart);
    if (!runEl) return 0;
    if (rPrMatches(runEl, canonical)) return 0;
    // SAI định dạng → thay rPr bằng clone mẫu
    cloneRPr(canonical, runEl, doc);
    return 1;
  }

  // CHƯA có tag → mở rộng qua whitespace, chèn run mới
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
      const t = paragraphText(c);
      if (t.length > 0) {
        spans.push({ st: pos, en: pos + t.length, el: c, isRun });
        pos += t.length;
      }
    }
  }
  if (end > pos) return 0;

  let si = -1;
  for (let i = 0; i < spans.length; i++) {
    if (end <= spans[i].en) { si = i; break; }
  }
  if (si < 0) return 0;
  const sp = spans[si];
  const el = sp.el;
  const tagRun = makeTagRunClone(canonical, doc);

  if (end === sp.st) {
    p.insertBefore(tagRun, el);
    return 1;
  }
  if (!sp.isRun) return 0;
  const r = el;
  if (end === sp.en) {
    if (r.nextSibling) p.insertBefore(tagRun, r.nextSibling as XNode);
    else p.appendChild(tagRun);
    return 1;
  }

  // end nằm giữa run → split
  const txt = runText(r);
  const splitAt = end - sp.st;
  const headRun = runWithTextClone(r, doc, txt.slice(0, splitAt));
  const tailRun = runWithTextClone(r, doc, txt.slice(splitAt));
  p.replaceChild(headRun, r);
  p.insertBefore(tailRun, headRun.nextSibling as XNode);
  p.insertBefore(tagRun, tailRun);
  return 1;
}

/** Tìm run (element w:r) chứa vị trí logical pos trong paragraph */
function findRunAt(p: XElement, pos: number): XElement | null {
  const children = childrenOf(p);
  let acc = 0;
  for (const c of children) {
    if (!isW(c, "r")) continue;
    const t = runText(c);
    if (pos <= acc + t.length && t.length > 0) return c;
    acc += t.length;
  }
  return null;
}

/** So sánh rPr của run với rPr mẫu (bỏ qua lang/rsid khác nhau) */
function rPrMatches(run: XElement, canonical: XElement): boolean {
  const a = childrenOf(run).find((c) => isW(c, "rPr"));
  const b = childrenOf(canonical).find((c) => isW(c, "rPr"));
  if (!a || !b) return false;

  const sig = (rPr: XElement): string => {
    const parts: string[] = [];
    for (const c of childrenOf(rPr)) {
      const ln = localName(c);
      if (ln === "lang" || ln === "rStyle" || ln === "rsid" || ln === "rtl" || ln === "cs" || ln === "w14" || ln === "w15") continue;
      // color: so value (theo RGB); themeColor được so riêng qua value nếu có
      let attrs = "";
      const attrNames = ["w:val", "w:color", "w:themeColor", "w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"];
      for (const an of attrNames) {
              const v = (c as XElement).getAttributeNS?.("w", an.replace("w:", ""));
              if (v) attrs += `${an}=${v};`;
            }
      parts.push(`${ln}({${attrs}})`);
    }
    return parts.sort().join("|");
  };
  return sig(a) === sig(b);
}