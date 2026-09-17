/**
 * Blanker: tạo "vị trí trống" trong câu kiến thức để học sinh hoàn thiện.
 *
 * Cơ sở sư phạm (đối chiếu với PHT mẫu do người làm):
 * - Câu định nghĩa "... được gọi là X" → X thành chỗ trống.
 * - Nếu X bắt đầu bằng từ loại (phép đo, đại lượng, đơn vị...) → giữ từ loại,
 *   chỉ blank phần phân biệt ("phép đo trực tiếp" → "phép đo .........").
 * - Câu chứa công thức toán OMML → bỏ qua (an toàn).
 *
 * Thao tác trực tiếp trên DOM XML để giữ nguyên rPr/định dạng của run khác.
 */

import { localName, childrenOf, isW, setTextNodeText, createWElement, paragraphText } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";

/** Từ mở đầu X để đáng blank — học thuật, học sinh cần tự nhớ.
 *  "hệ đơn vị", "đơn vị dẫn xuất" (danh từ chung) → KHÔNG blank (khớp mẫu). */
const BLANK_HEADS = new Set([
  "phép", "đại lượng", "dụng cụ", "thiết bị", "hiện tượng", "quy tắc",
  "định luật", "khái niệm", "mô hình", "phương pháp", "công cụ", "công thức",
  "nhiệt lượng", "động năng", "thế năng", "nội năng", "cơ năng", "nhiệt độ",
  "áp suất", "khối lượng", "thể tích", "tần số", "biên độ", "tốc độ", "gia tốc",
  "lực", "công", "năng lượng", "đơn vị cơ bản", "đơn vị dẫn xuất",
]);

/** Từ 2-token được coi là một head ("dụng cụ đo", "phép đo trực tiếp"...) */
const BLANK_HEADS_2 = new Set(["dụng cụ", "phép đo", "đơn vị cơ bản", "đơn vị dẫn xuất", "nhiệt lượng"]);

/** Gạch chân đủ dài cho học sinh viết */
function blanksFor(len: number): string {
  return "_".repeat(Math.max(14, Math.min(30, len + 6)));
}

const BLANKABLE =
  /(?:được\s+gọi\s+chung\s+là|được\s+gọi\s+là|gọi\s+chung\s+là|gọi\s+là)\s+([^,.;:()\n]{1,80})/g;

/**
 * Tìm vùng [start, end) trong text cần thay bằng chỗ trống.
 * Quy tắc khớp PHT mẫu:
 * - "... gọi là X" (câu NGẮN, ≤ 4 từ) → blank TOÀN BỘ X ("dụng cụ đo" → ______).
 * - "phép đo + Y" → giữ "phép đo", blank Y ("phép đo trực tiếp" → "phép đo ____").
 * - Câu dài (>140 ký tự) → bỏ qua (an toàn, như mẫu giữ nguyên).
 * CÁC VỊ TRÍ LÀ TỌA ĐỘ TRÊN TEXT ĐÃ TRIM — caller phải shift sang raw.
 */
export function findBlankRegions(text: string): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  if (text.length > 140) return out;
  BLANKABLE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = BLANKABLE.exec(text))) {
    const raw = m[1];
        const trimmed = raw.trim();
        if (!trimmed) continue;
        const words = trimmed.split(/\s+/);
                if (words.length > 5) continue; // quá dài — không blank
                const absStart = m.index + m[0].length - raw.length; // đầu raw (có thể có space đầu)
                const lead = raw.length - raw.trimStart().length; // space thừa đầu raw
                let blankStart: number;
                // head 2-token: "dụng cụ đo" → giữ nguyên cả head, blank phần sau
                const head2 = words.slice(0, 2).join(" ");
                if (BLANK_HEADS_2.has(head2)) {
                  if (words.length === 2) continue; // chỉ head — không có từ phân biệt
                  const rest = words.slice(2).join(" ");
                  const rel = raw.slice(lead).indexOf(rest);
                  blankStart = absStart + lead + (rel >= 0 ? rel : head2.length + 1);
                }
                else if (BLANK_HEADS.has(words[0])) {
                  // giữ "phép/đại lượng/...", blank phần còn lại: "phép đo trực tiếp" → "phép đo ____"
                  if (words.length === 1) continue;
                  const rest = words.slice(1).join(" ");
                  const rel = raw.slice(lead).indexOf(rest);
                  blankStart = absStart + lead + (rel >= 0 ? rel : words[0].length + 1);
                } else if (words.length >= 2 && BLANK_HEADS.has(words[1])) {
                  // "được gọi là phép đo trực tiếp" — cùng quy tắc, blank sau "phép đo"
                  if (words.length === 2) continue;
                  const rest = words.slice(2).join(" ");
                  const rel = raw.slice(lead).indexOf(rest);
                  blankStart = absStart + lead + (rel >= 0 ? rel : words.slice(0, 2).join(" ").length + 1);
                } else if (words.length <= 3) {
                  // danh từ chung ngắn ("dụng cụ đo", "nhiệt nóng chảy") → blank TOÀN BỘ (theo mẫu)
                  blankStart = absStart + lead;
                } else {
                  continue; // an toàn: câu dài, không rõ — giữ nguyên
                }
                const end = absStart + raw.length - (raw.length - raw.trimEnd().length);
                if (end - blankStart >= 2) out.push([blankStart, end]);
      }
  return out;
}

/** Text (logical) của một run w:r */
function runText(r: XElement): string {
  let out = "";
  for (const c of childrenOf(r)) {
    const ln = localName(c);
    if (ln === "t") out += c.firstChild?.nodeValue ?? "";
    else if (ln === "tab") out += "\t";
    else if (ln === "br" || ln === "cr") out += "\n";
  }
  return out;
}

function hasMathInPara(p: XElement): boolean {
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

/** Clone run nhưng chỉ giữ rPr + các con không phải text (drawing, comment ref...) */
function cloneRunShell(src: XElement, doc: XDocument): XElement {
  const r = createWElement(doc, "r");
  for (const c of childrenOf(src)) {
    const ln = localName(c);
    if (ln === "t" || ln === "tab" || ln === "br" || ln === "cr") continue;
    r.appendChild(c.cloneNode(true));
  }
  return r;
}

function runWithText(src: XElement, doc: XDocument, text: string): XElement {
  const r = cloneRunShell(src, doc);
  const t = createWElement(doc, "t");
  t.setAttribute("xml:space", "preserve");
  setTextNodeText(t, text);
  r.appendChild(t);
  return r;
}

function blankRun(src: XElement, doc: XDocument, len: number): XElement {
  const r = cloneRunShell(src, doc);
  let rPr = childrenOf(r).find((c) => isW(c, "rPr"));
  if (!rPr) {
    rPr = createWElement(doc, "rPr");
    r.appendChild(rPr); // rPr ở đầu (run shell chỉ có rPr hoặc rỗng)
    // đưa rPr lên đầu nếu có con khác (hiếm)
    const first = r.firstChild;
    if (first && first !== rPr) {
      r.removeChild(rPr);
      r.insertBefore(rPr, first);
    }
  }
  let u = childrenOf(rPr).find((c) => isW(c, "u"));
  if (!u) {
    u = createWElement(doc, "u");
    rPr.appendChild(u);
  }
  u.setAttribute("w:val", "single");
  const t = createWElement(doc, "t");
  t.setAttribute("xml:space", "preserve");
  setTextNodeText(t, blanksFor(len));
  r.appendChild(t);
  return r;
}

/**
 * Áp blank lên paragraph: cắt regions (tọa độ logical) thành run mới.
 * Giữ pPr và toàn bộ run ngoài region. Trả về số blank; 0 nếu không làm gì.
 */
export function applyBlanks(p: XElement, doc: XDocument, regions: Array<[number, number]>): number {
  if (!regions.length) return 0;
  if (hasMathInPara(p)) return 0;

  const runs: XElement[] = childrenOf(p).filter((c) => isW(c, "r"));
  if (!runs.length) return 0;

  // Span logical của từng run
  const spans: Array<{ st: number; en: number }> = [];
  let pos = 0;
  for (const r of runs) {
    const t = runText(r);
    spans.push({ st: pos, en: pos + t.length });
    pos += t.length;
  }
  if (pos === 0) return 0;

  // Vùng hợp lệ, cắt trong [0, pos]
  const regionsSorted = [...regions].sort((a, b) => a[0] - b[0]);
  const clipped: Array<[number, number]> = [];
  let lastEnd = -1;
  for (const [s, e] of regionsSorted) {
    const cs = Math.max(0, Math.min(pos, s));
    const ce = Math.max(cs, Math.min(pos, e));
    if (ce - cs >= 2 && cs >= lastEnd) {
      clipped.push([cs, ce]);
      lastEnd = ce;
    }
  }
  if (!clipped.length) return 0;

  // Dựng lại: với mỗi run, xén ra text trước / blank / text sau theo region
  const out: XElement[] = [];
  let blankCount = 0;
  let ri = 0; // region index

  for (let i = 0; i < runs.length; i++) {
    const r = runs[i];
    const sp = spans[i];
    // bỏ qua region kết thúc trước run này
    while (ri < clipped.length && clipped[ri][1] <= sp.st) ri++;
    if (ri >= clipped.length || clipped[ri][0] >= sp.en) {
      out.push(r); // run ngoài region
      continue;
    }
    // run giao với ≥1 region
    let cursor = sp.st;
    let consumed = false;
    while (ri < clipped.length && cursor < sp.en) {
      const [rs, re] = clipped[ri];
      if (re <= cursor) { ri++; continue; }
      if (rs >= sp.en) break;
      const cutS = Math.max(cursor, rs);
      const cutE = Math.min(sp.en, re);
      if (cutS > cursor) {
        out.push(runWithText(r, doc, runText(r).slice(cursor - sp.st, cutS - sp.st)));
      }
      out.push(blankRun(r, doc, cutE - cutS));
      blankCount++;
      cursor = cutE;
      consumed = true;
      ri++;
    }
    if (cursor < sp.en) {
      if (consumed) out.push(runWithText(r, doc, runText(r).slice(cursor - sp.st)));
      else out.push(r);
    }
  }

  // Thay toàn bộ con (pPr giữ nguyên — được giữ qua `out` nếu là con trực tiếp? Không — pPr không nằm trong runs)
  const pPr = childrenOf(p).find((c) => isW(c, "pPr"));
  while (p.firstChild) p.removeChild(p.firstChild);
  if (pPr) p.appendChild(pPr);
  for (const n of out) p.appendChild(n);
  return blankCount;
}

/** Tạo chỗ trống placeholder (API ổn định, không được dùng trong pipeline) */
export function makeBlankPlaceholder(paras: unknown): string {
  void paras;
  return blanksFor(12);
}

export { paragraphText };