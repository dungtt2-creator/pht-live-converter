/**
 * Xử lý phân trang (Page-flo, Keep-together)
 *
 * QUY TẮC BẤT BIẾN (yêu cầu chuẩn hóa tài liệu):
 *  1. Heading/subheading phải ĐI CÙNG đoạn nội dung ngay sau nó:
 *     đặt w:keepNext trên paragraph heading (KHÔNG áp dụng cho nội dung).
 *  2. KHÔNG đặt keepNext hàng loạt cho paragraph nội dung — tránh chuỗi
 *     keep-next khiến cả block bị đẩy sang trang tiếp theo.
 *  3. KHÔNG xóa pageBreakBefore có sẵn của heading (tôn trọng layout template),
 *     KHÔNG chèn page break thủ công mới.
 *  4. Xóa page break KHÔNG cần thiết phát sinh sau khi xóa khối lời giải:
 *     <w:br w:type="page"/> đứng ngay sau chỗ xóa + paragraph rỗng thừa.
 *  5. Widow/orphan: bật w:widowControl nếu heading chưa có (giữ giá trị sẵn có).
 *  6. Floating/anchored: KHÔNG đụng — word tự reflow quanh chúng.
 */

import { localName, isW, childrenOf } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";

/** Đặt w:keepNext + w:widowControl cho paragraph heading (nếu chưa có), trả số đã sửa */
export function fixHeadingKeepNext(
  blocks: DocBlock[],
  idx: Array<{ el: XElement; isTable: boolean }>,
  doc: XDocument,
  stripTitleTables = false,
): number {
  let fixed = 0;
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (b.kind === "table") {
      // Heading trong BẢNG 1-cell (GV dùng table + shd để tạo nền tiêu đề):
      // text cell duy nhất khớp dạng "I. NHIỆT LƯỢNG..." / "1. Tiêu đề..." → đặt keepNext
      // trên paragraph trong cell để giữ heading với nội dung sau bảng.
      const t = b as unknown as { rows: { text: string }[][] };
      const cellText = ((t.rows[0] || [])[0] || { text: "" }).text.trim();
      if (/^[IVXlc]+\s*\.\s+[A-ZÀ-ỸĀĐ]|^\d+(\.\d+)*\s*[.):]\s+[A-ZÀ-ỸĀĐ]/.test(cellText)) {
        if (i < idx.length && idx[i].isTable) {
          if (stripTitleTables) {
            // Chuyển bảng tiêu đề 1-cell → paragraph thường (giữ shd) rồi keepNext.
            // Word KHÔNG giữ heading trong bảng với nội dung SAU bảng (keepNext không
            // xuyên khối table) → heading rơi cuối trang. MAU Live chuẩn đã dùng
            // paragraph thường cho heading — làm đúng như vậy.
            const p = stripTitleTableIntoPara(idx[i].el as XElement, doc);
            if (p) {
              fixed++;
              idx[i] = { el: p, isTable: false };
            }
          } else {
            if (setKeepNextInFirstCell(idx[i].el as XElement, doc)) fixed++;
          }
        }
      }
      continue;
    }
    const blk = b as Block;
    if (blk.kind !== "heading") continue;
    // chỉ heading "I." / "1." (cấp mục lớn) — không phải heading lồng trong câu hỏi
    if (!/^[IVX]+\s*\.|^\d+(\.\d+)*\s*[.):]/.test(blk.text.trim())) continue;
    if (i + 1 >= blocks.length || i + 1 >= idx.length || idx[i + 1].isTable) continue;
    // Nội dung tiếp theo phải là paragraph (để keep với nó); nếu kế là heading khác → bỏ
    const next = blocks[i + 1];
    if (next.kind !== "table" && (next as Block).kind === "heading") continue;

    const p = idx[i].el as XElement;
    if (setKeepNextOnPara(p, doc)) fixed++;
  }
  return fixed;
}

/** Đặt keepNext + widowControl cho paragraph đầu tiên trong cell đầu bảng */
function setKeepNextInFirstCell(tbl: XElement, doc: XDocument): boolean {
  for (const tr of childrenOf(tbl)) {
    if (!isW(tr, "tr")) continue;
    for (const tc of childrenOf(tr)) {
      if (!isW(tc, "tc")) continue;
      // paragraph ĐẦU TIÊN có text trong cell (bỏ paragraph rỗng)
      for (const c of childrenOf(tc)) {
        if (!isW(c, "p")) continue;
        const txt = collectText(c as XElement);
        if (!txt.trim()) continue;
        return setKeepNextOnPara(c as XElement, doc);
      }
    }
  }
  return false;
}

/**
 * Chuyển bảng 1-cell chứa heading thành paragraph thường (giữ toàn bộ nội dung cell,
 * kể cả shd trong pPr nếu có) — vì keepNext KHÔNG xuyên qua khối table.
 * Đồng thời xóa paragraph rỗng ngay SAU heading để keepNext trỏ thẳng vào đoạn nội
 * dung thật (nếu còn rỗng, keepNext buộc với paragraph rỗng → heading vẫn rơi cuối trang).
 * Trả paragraph mới hoặc null nếu không làm được.
 */
function stripTitleTableIntoPara(tbl: XElement, doc: XDocument): XElement | null {
  // Lấy paragraph đầu (có text) trong cell đầu — clone ra ngoài bảng
  for (const tr of childrenOf(tbl)) {
    if (!isW(tr, "tr")) continue;
    for (const tc of childrenOf(tr)) {
      if (!isW(tc, "tc")) continue;
      for (const c of childrenOf(tc)) {
        if (!isW(c, "p")) continue;
        const txt = collectText(c as XElement);
        if (!txt.trim()) continue;
        const p = c.cloneNode(true) as XElement;
        // Thêm keepNext + widowControl
        setKeepNextOnPara(p, doc);
        // Thay bảng bằng paragraph
        tbl.parentNode!.insertBefore(p as XNode, tbl as XNode);
        tbl.parentNode!.removeChild(tbl as XNode);

        // Xóa paragraph rỗng liền sau (tối đa 8) — keepNext cần trỏ vào content thật
        let removed = 0;
        let sibling: XNode | null = p.nextSibling;
        while (sibling && removed < 8) {
          const s = sibling as XElement;
          if (!isW(s, "p")) break; // dừng khi gặp bảng hoặc phần tử khác
          const t = collectText(s);
          if (t.trim().length > 0) break;
          const isVisual = hasOmmOrDrawing(s);
          if (isVisual) break;
          const nxt = sibling.nextSibling;
          s.parentNode!.removeChild(s);
          removed++;
          sibling = nxt;
        }
        return p;
      }
    }
  }
  return null;
}

function hasOmmOrDrawing(p: XElement): boolean {
  const walk = (n: XNode): boolean => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      const ln = localName(c);
      if (ln === "oMath" || ln === "oMathPara" || ln === "drawing" || ln === "pict" || ln === "object" || ln === "fldChar") return true;
      if (walk(c)) return true;
    }
    return false;
  };
  return walk(p);
}

/** Đặt keepNext + widowControl lên paragraph (đúng thứ tự schema Word) */
function setKeepNextOnPara(p: XElement, doc: XDocument): boolean {
  const pPr = ensurePPr(p, doc);
  let changed = false;
  if (!childrenOf(pPr).some((c) => isW(c, "keepNext"))) {
    const kn = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:keepNext");
    insertPPrChild(pPr, kn, doc);
    changed = true;
  }
  if (!childrenOf(pPr).some((c) => isW(c, "widowControl"))) {
    const wc = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:widowControl");
    wc.setAttribute("w:val", "1");
    insertPPrChild(pPr, wc, doc);
    changed = true;
  }
  return changed;
}

function collectText(p: XElement): string {
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
 * Dọn khoảng trống phát sinh sau khi xóa khối lời giải:
 *  - paragraph rỗng (không text, không OMML/drawing) xuất hiện LIÊN TIẾP ≥2 → giữ 1
 *  - <w:br w:type="page"/> đứng TRONG paragraph rỗng hoặc ngay sau vùng xóa → bỏ
 *    (chỉ bỏ khi paragraph có <1 ký tự text thường — an toàn: không đụng page break
 *     có chủ đích giữa các đề lớn).
 * Trả về {emptyRemoved, pageBreaksRemoved}.
 */
export function cleanupSpacers(
  blocks: DocBlock[],
  idx: Array<{ el: XElement; isTable: boolean }>,
  removedSet: Set<number> | null,
  doc: XDocument,
): { emptyRemoved: number; pageBreaksRemoved: number } {
  let emptyRemoved = 0;
  let pageBreaksRemoved = 0;

  // 1) Gom các chỉ số "paragraph rỗng" (cần xét liền kề nhau trong body)
  const isParaEmpty = (i: number): boolean => {
    if (i >= idx.length || idx[i].isTable) return false;
    const p = idx[i].el as XElement;
    const txt = textOf(p);
    if (txt.length > 0) return false;
    // có OMML/drawing/field → KHÔNG xem là rỗng
    if (containsNonTextVisual(p)) return false;
    return true;
  };

  // pass 1: xóa page break trong run KHÔNG có text (page break phát sinh do xóa lời giải)
  for (let i = 0; i < idx.length; i++) {
    if (idx[i].isTable) continue;
    const p = idx[i].el as XElement;
    const pageBrs = collectPageBreaks(p);
    for (const br of pageBrs) {
      const r = br.parentNode as XElement;
      const rText = textOf(r as XElement);
      if (rText.trim().length === 0) {
        // page break nằm trong run không có text → xóa run (an toàn)
        r.parentNode!.removeChild(r as XNode);
        pageBreaksRemoved++;
      }
    }
  }

  // pass 2: gom paragraph rỗng liên tiếp (chỉ trong vùng xóa nếu biết removedSet) → giữ 1
  let run = 0;
  let runStart = -1;
  const toRemoveEmpty: number[] = [];
  for (let i = 0; i < idx.length; i++) {
    if (idx[i].isTable) { run = 0; runStart = -1; continue; }
    const empty = isParaEmpty(i);
    if (empty) {
      if (run === 0) { runStart = i; }
      run++;
      // giữ paragraph đầu của chuỗi, xóa những cái sau nếu cũng empty
      if (run > 1 && runStart >= 0) toRemoveEmpty.push(i);
    } else {
      run = 0; runStart = -1;
    }
  }
  for (const i of toRemoveEmpty) {
    if (idx[i].el.parentNode) {
      idx[i].el.parentNode!.removeChild(idx[i].el);
      emptyRemoved++;
    }
  }

  return { emptyRemoved, pageBreaksRemoved };
}

function collectPageBreaks(p: XElement): XElement[] {
  const out: XElement[] = [];
  const walk = (n: XNode) => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      const ln = localName(c);
      if (ln === "br") {
        const t = (c as XElement).getAttributeNS?.("w", "type") || (c as XElement).getAttributeNS?.("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "type");
        if (t === "page") out.push(c as XElement);
      }
      walk(c);
    }
  };
  walk(p);
  return out;
}

function textOf(p: XElement): string {
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

function containsNonTextVisual(p: XElement): boolean {
  const walk = (n: XNode): boolean => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      const ln = localName(c);
      if (ln === "oMath" || ln === "oMathPara" || ln === "drawing" || ln === "pict" || ln === "object" || ln === "fldChar") return true;
      if (walk(c)) return true;
    }
    return false;
  };
  return walk(p);
}

function ensurePPr(p: XElement, doc: XDocument): XElement {
  let pPr = childrenOf(p).find((c) => isW(c, "pPr"));
  if (!pPr) {
    pPr = doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "w:pPr");
    p.insertBefore(pPr as XNode, p.firstChild as XNode | null);
  }
  return pPr;
}

/** Chèn child vào pPr theo thứ tự schema Word (keepNext, widowControl trước spacing) */
function insertPPrChild(pPr: XElement, child: XElement, doc: XDocument): void {
  // thứ tự: pStyle | keepNext | keepLines | pageBreakBefore | ... | widowControl | ... | spacing | ind | jc ...
  const order = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
  const ln = localName(child);
  const myIdx = order.indexOf(ln);
  const kids = childrenOf(pPr);
  let insertPos: XNode | null = null;
  for (let i = 0; i < kids.length; i++) {
    const kl = order.indexOf(localName(kids[i]));
    if (kl > myIdx) { insertPos = kids[i] as XNode; break; }
  }
  if (insertPos) pPr.insertBefore(child as XNode, insertPos);
  else pPr.appendChild(child as XNode);
}