/**
 * Chuẩn hóa CÂU DẪN NGỮ LIỆU CHUNG.
 *
 * Quy tắc (yêu cầu user 10-2026):
 *   Gốc : "Dựa vào thông tin dưới đây để trả lời các câu hỏi từ 33 đến 35"
 *   Tính : x = 35 − 33 + 1 = 3
 *   Đích : "Dựa vào thông tin dưới đây để trả lời 3 câu hỏi sau."
 *
 * BẤT BIẾN: chỉ thay ĐOẠN TEXT trong các w:t đã có, KHÔNG dựng lại paragraph
 * → giữ nguyên 100% định dạng (màu, in đậm, font, sz, vị trí, khoảng cách) của
 * câu dẫn gốc, kể cả khi câu dẫn nằm chung paragraph với nhãn "[HOCMAI] – (ID: …)".
 */

import { localName } from "./xml";
import type { XNode, XElement } from "./xml";

/**
 * Tìm câu dẫn ngữ liệu chung — các biến thể gặp trong đề thật:
 *   "…để trả lời các câu hỏi từ 33 đến 35"          (HSA)
 *   "…để trả lời các câu từ 106 đến 108"            (V-ACT)
 *   "…để trả lời câu hỏi từ 1 đến 3"                (V-ACT)
 *   "…để trả lời câu hỏi 28 đến 30"                 (V-ACT)
 *   "…để trả lời các câu sau từ câu 16 - 18"        (V-ACT)
 * Chấp nhận: NBSP, khoảng trắng lặp, "từ câu A", dấu "–"/"-" thay "đến",
 * dấu chấm cuối (bị thay luôn).
 * Group 3 = Số_A, group 4 = Số_B. Phần "Dựa vào thông tin dưới đây để trả lời"
 * GIỮ NGUYÊN — chỉ phần đuôi bị thay.
 */
export const STIMULUS_RE =
  /Dựa\s+vào\s+thông\s+tin\s+dưới\s+đây\s+để\s+trả\s+lời\s+(?:các\s+)?(?:câu\s+(?:hỏi\s+)?)?(?:sau\s+)?từ\s+(?:câu\s+)?(\d{1,3})\s*(?:đến|[-–—])\s*(?:câu\s+)?(\d{1,3})\s*\.?/g;

/**
 * Biến thể KHÔNG có chữ "từ": "…để trả lời câu hỏi 28 đến 30"
 */
export const STIMULUS_RE_NO_TU =
  /Dựa\s+vào\s+thông\s+tin\s+dưới\s+đây\s+để\s+trả\s+lời\s+(?:các\s+)?câu\s+(?:hỏi\s+)?(\d{1,3})\s*(?:đến|[-–—])\s*(?:câu\s+)?(\d{1,3})\s*\.?/g;

/** Không đi vào các nhánh đồ hoạ/OLE (w:t bên trong textbox không thuộc dòng chảy) */
const SKIP_SUBTREE = new Set(["drawing", "pict", "object"]);

type Token =
  | { kind: "t"; el: XElement; text: string; start: number; end: number }
  | { kind: "ctl"; start: number; end: number };

/**
 * Duyệt paragraph theo ĐÚNG thứ tự dòng chảy (giống paragraphText: w:t cộng text,
 * w:tab → \t, w:br/w:cr → \n) và ghi lại offset của từng w:t.
 */
function collectTokens(p: XElement): { text: string; tokens: Token[] } {
  const tokens: Token[] = [];
  let text = "";
  const walk = (n: XNode) => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      const ln = localName(c);
      if (c.nodeType === 3) {
        // text "trần" trong w:r (một số generator) — không đụng tới, chỉ đếm
        text += c.nodeValue ?? "";
        continue;
      }
      if (SKIP_SUBTREE.has(ln)) continue;
      if (ln === "tab" || ln === "br" || ln === "cr") {
        const ch = ln === "tab" ? "\t" : "\n";
        tokens.push({ kind: "ctl", start: text.length, end: text.length + 1 });
        text += ch;
        continue;
      }
      if (ln === "t") {
        const s = text.length;
        const inner = collectTextOf(c);
        tokens.push({ kind: "t", el: c as XElement, text: inner, start: s, end: s + inner.length });
        text += inner;
        continue;
      }
      walk(c);
    }
  };
  walk(p);
  return { text, tokens };
}

function collectTextOf(el: XNode): string {
  let out = "";
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i];
    if (c.nodeType === 3) out += c.nodeValue ?? "";
  }
  return out;
}

function setNodeText(el: XElement, value: string): void {
  while (el.firstChild) el.removeChild(el.firstChild);
  el.appendChild(el.ownerDocument.createTextNode(value));
  // giữ khoảng trắng đầu/cuối (nếu không Word cắt mất)
  if (/^\s|\s$/.test(value)) el.setAttribute("xml:space", "preserve");
  else {
    // xoá xml:space cũ nếu không cần
    const anyEl = el as unknown as { removeAttribute?: (n: string) => void };
    anyEl.removeAttribute?.("xml:space");
  }
}

/**
 * Dọn run đã rỗng sau khi thay text: xoá w:r nếu mọi w:t của nó đều rỗng và
 * không chứa nội dung khác (drawing/object/pict/tab/br). Không đụng run OLE/ảnh.
 */
function pruneEmptyRunAfter(t: XElement): void {
  const run = t.parentNode;
  if (!run || localName(run) !== "r") return;
  for (let i = 0; i < run.childNodes.length; i++) {
    const c = run.childNodes[i];
    const ln = localName(c);
    if (ln === "rPr") continue;
    if (ln === "t") {
      if (collectTextOf(c).length > 0) return;
      continue;
    }
    // tab/br/drawing/object/pict/fldChar… → giữ nguyên run
    return;
  }
  run.parentNode?.removeChild(run);
}

/**
 * Thay đoạn [start, end) của text paragraph bằng newText — CHỈ sửa nội dung
 * các w:t đã có, giữ nguyên rPr của từng run.
 */
function replaceRange(tokens: Token[], start: number, end: number, newText: string): boolean {
  const tNodes = tokens.filter((t): t is Extract<Token, { kind: "t" }> => t.kind === "t");
  let placed = false;
  let touched = false;
  const emptied: XElement[] = [];
  for (const t of tNodes) {
    if (t.end <= start || t.start >= end) continue;
    const text = collectTextOf(t.el);
    const localStart = Math.max(0, start - t.start);
    const localEnd = Math.min(t.end - t.start, end - t.start);
    const before = text.slice(0, localStart);
    const after = text.slice(localEnd);
    // newText chỉ ghi vào w:t ĐẦU TIÊN giao với dải cần thay
    // (giữ nguyên rPr của run đầu — đúng định dạng câu dẫn gốc)
    const next = before + (placed ? "" : newText) + after;
    placed = true;
    touched = true;
    setNodeText(t.el, next);
    if (next.length === 0) emptied.push(t.el);
  }
  // dọn các run chỉ còn rPr (rỗng) — tránh để lại run "vỏ" trong XML
  for (const el of emptied) pruneEmptyRunAfter(el);
  return touched;
}

export interface StimulusChange {
  before: string;
  after: string;
  count: number;
}

export interface StimulusResult {
  changes: StimulusChange[];
}

/** Duyệt mọi paragraph trong body — KỂ CẢ paragraph nằm trong bảng (w:tbl/w:tr/w:tc).
 *  Bỏ qua nhánh đồ hoạ/OLE (w:drawing, w:pict, w:object) vì textbox không thuộc
 *  dòng chảy văn bản. */
function collectParagraphs(body: XElement): XElement[] {
  const out: XElement[] = [];
  const walk = (n: XNode) => {
    for (let i = 0; i < n.childNodes.length; i++) {
      const c = n.childNodes[i];
      if (c.nodeType !== 1) continue;
      const ln = localName(c);
      if (SKIP_SUBTREE.has(ln)) continue;
      if (ln === "p") {
        out.push(c as XElement);
        continue; // paragraph không lồng paragraph
      }
      walk(c);
    }
  };
  walk(body);
  return out;
}

/**
 * Áp dụng cho MỌI paragraph trong tài liệu (kể cả trong bảng).
 */
export function normalizeStimulusHeadings(body: XElement): StimulusResult {
  const changes: StimulusChange[] = [];
  for (const p of collectParagraphs(body)) {
    if (!p.parentNode) continue;
    const { text, tokens } = collectTokens(p);
    if (!/Dựa\s+vào/.test(text)) continue;

    // Thu thập match của CẢ 2 biến thể, sắp theo vị trí, bỏ trùng
    type Hit = { start: number; end: number; a: number; b: number };
    const hits: Hit[] = [];
    const scan = (re: RegExp, aGroup: number, bGroup: number) => {
      re.lastIndex = 0;
      let m: RegExpExecArray | null;
      while ((m = re.exec(text)) !== null) {
        const a = parseInt(m[aGroup], 10);
        const b = parseInt(m[bGroup], 10);
        if (!Number.isFinite(a) || !Number.isFinite(b)) continue;
        // Chỉ thay phần ĐUÔI: bắt đầu ngay sau "trả lời" (+ khoảng trắng)
        // (giữ nguyên "Dựa vào thông tin dưới đây để trả lời").
        const rel = m[0].toLowerCase().indexOf("trả lời");
        const afterReply = rel >= 0 ? m[0].slice(rel + "trả lời".length) : m[0];
        const cutRel = rel + "trả lời".length + (afterReply.length - afterReply.replace(/^\s+/, "").length);
        hits.push({ start: m.index + cutRel, end: m.index + m[0].length, a, b });
      }
    };
    scan(STIMULUS_RE, 1, 2);
    scan(STIMULUS_RE_NO_TU, 1, 2);
    if (!hits.length) continue;

    // Xử lý TỪ CUỐI VỀ ĐẦU để offset phía trước không lệch; bỏ hit trùng dải
    hits.sort((x, y) => y.start - x.start);
    let lastStart = Number.POSITIVE_INFINITY;
    const applied: Hit[] = [];
    for (const h of hits) {
      if (h.end > lastStart) continue; // chồng lấn với hit đã xử lý
      lastStart = h.start;
      applied.push(h);
    }

    for (const h of applied) {
      const x = h.b - h.a + 1;
      if (x <= 0) continue;
      // Giữ nguyên dấu cách sẵn có trước cụm đuôi ("…để trả lời|các câu…")
      // → không thêm space, tránh "trả lời  3 câu".
      const ok = replaceRange(tokens, h.start, h.end, `${x} câu hỏi sau.`);
      if (!ok) continue;
      changes.push({
        before: `từ ${h.a} đến ${h.b}`,
        after: `${x} câu hỏi sau.`,
        count: x,
      });
    }
  }
  return { changes };
}
