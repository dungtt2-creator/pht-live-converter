/**
 * Converter: chuyển DOM document.xml đã parse thành PHT hoặc Live.
 *
 * PHT (Phiếu học tập): giữ mục tiêu + lý thuyết (blank hóa định nghĩa) + câu hỏi.
 *     Xóa: khối "Đáp án/Hướng dẫn/Lời giải" + nội dung giải, bảng giải thích.
 * Live (Tài liệu Live, theo mẫu kỳ thi):
 *     1) Trang NGANG (16840x11907, margin 720) — đúng template.
 *     2) Xóa mọi highlight (đánh dấu màu đáp án của GV).
 *     3) Xóa đáp án + hướng dẫn giải (cùng logic PHT).
 *     4) Gom câu hỏi vào BẢNG 2 cột viền đỏ C00000: trái = câu hỏi, phải = dòng kẻ
 *        để giáo viên viết khi chữa bài.
 *
 * Nguyên tắc bất biến: không sửa nội dung kiến thức; chỉ xóa đáp án/đánh dấu
 * và thêm cấu trúc hỗ trợ.
 */

import { localName, childrenOf, isW, paragraphText, createWElement } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";
import { applyBlanks, findBlankRegions } from "./blanker";
import { getTemplate, type TemplateId } from "./template";

export type OutputMode = "pht" | "live";

export interface ConvertResult {
  blocks: DocBlock[];
  answersRemoved: number;
  guidelinesRemoved: number;
  blanksCreated: number;
  notes: string[];
}

const QNUM = /^[Cc]âu\s*(\d+)\s*[.:]/;
const NEW_CONTEXT =
  /^(Sử dụng|Dựa vào|Đọc|Xem xét|Căn cứ|Kết hợp)\s+(thông tin|đoạn|bài|hình|bảng|đề|đồ thị)/i;
const PRESERVED_CONTEXT =
  /^Bài đọc|^Đoạn văn sau|^Thông tin sau|^DẶN DÒ|^Dặn dò|^Nguồn\s*[:：]|^Trích\s*(SGK|SBT)/;
/** "Đáp án: A. ____" dạng mẫu điền — chỉ giữ khi có gạch chân */
const ANSWER_WITH_CHOICES = /^(Đáp án|Đ\/A)\s*[:.]?\s*[A-D]\./;
const FILL_BLANK_ONLY = /_{2,}/;

function hasMath(p: XElement): boolean {
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

/** Thu thập các w:p và w:tbl trong body theo thứ tự → mảng đánh chỉ số chung */
function indexBody(body: XElement): Array<{ el: XElement; isTable: boolean }> {
  const out: Array<{ el: XElement; isTable: boolean }> = [];
  for (const child of childrenOf(body)) {
    if (isW(child, "p")) out.push({ el: child, isTable: false });
    else if (isW(child, "tbl")) out.push({ el: child, isTable: true });
  }
  return out;
}

/**
 * Tính tập block cần XÓA (đáp án/hướng dẫn giải) — dùng chung PHT & Live.
 * Trả về Set index block.
 */
function computeRemoveSet(blocks: DocBlock[]): { toRemove: Set<number> } {
  const toRemove = new Set<number>();
  let inG = false; // đang duyệt khối đáp án/hướng dẫn

  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (b.kind === "table") {
      const t = b as unknown as { rows: { text: string }[][] };
      const header = (t.rows[0] || []).map((c) => c.text).join(" ");
      // Bảng GIẢI THÍCH đáp án (xóa độc lập):
      //  1) header "Ý | Đ/A | Giải thích"
      //  2) cột đầu dòng 0 "a) Đúng/Sai" + dòng 1 có nội dung giải
      const isExplanation =
        (/Đ\/A/.test(header) && /Giải thích/.test(header)) ||
        ((t.rows[0] || []).some((c) => /^a\s*\)?\s*(Đúng|Sai)\b/.test(c.text.trim())) &&
          (t.rows[1] || []).some((c) => c.text.trim().length > 2));
      if (isExplanation) toRemove.add(i);
      continue;
    }
    const blk = b as Block;
    const txt = blk.text.trim();
    if (inG) {
      if (
        QNUM.test(txt) ||
        blk.kind === "heading" ||
        NEW_CONTEXT.test(txt) ||
        PRESERVED_CONTEXT.test(txt)
      ) {
        inG = false;
        continue; // điểm dừng khối giải — giữ
      }
      if (ANSWER_WITH_CHOICES.test(txt) && FILL_BLANK_ONLY.test(txt)) continue; // mẫu điền
      toRemove.add(i); // nội dung giải / đáp án thật
      continue;
    }
    if (blk.kind === "guideline" || blk.kind === "answer") {
      if (ANSWER_WITH_CHOICES.test(txt) && FILL_BLANK_ONLY.test(txt)) continue; // mẫu điền
      toRemove.add(i);
      inG = true;
    }
  }
  return { toRemove };
}

/** Xóa các block đã đánh dấu khỏi body */
function removeBlocks(body: XElement, idx: Array<{ el: XElement; isTable: boolean }>, toRemove: Set<number>): number {
  let n = 0;
  for (let i = 0; i < idx.length; i++) {
    if (toRemove.has(i)) {
      idx[i].el.parentNode!.removeChild(idx[i].el);
      n++;
    }
  }
  return n;
}

/**
 * Chuyển đổi chính.
 */
export function convertDom(
  blocks: DocBlock[],
  body: XElement,
  mode: OutputMode,
  templateId: TemplateId = "tsa",
  useFileTemplate = false,
): ConvertResult {
  const notes: string[] = [];
  const result: ConvertResult = {
    blocks, answersRemoved: 0, guidelinesRemoved: 0, blanksCreated: 0, notes,
  };
  const tpl = getTemplate(templateId);
  const idx = indexBody(body);
  if (idx.length !== blocks.length) {
    notes.push(`Cảnh báo: số block model (${blocks.length}) khác số phần tử body (${idx.length}).`);
  }
  const doc = body.ownerDocument;

  // ===== 0) Banner nhận diện (chỉ khi KHÔNG dùng template file) =====
  // Khi useFileTemplate=true, banner + sectPr đến từ file template thật (merge sau).
  if (!useFileTemplate && idx.length > 0) {
    const banner = makeBannerPara(doc, tpl.header, tpl.color);
    const firstEl = idx[0].el;
    body.insertBefore(banner, firstEl);
    notes.push(`Template: ${tpl.label} — đã chèn banner "${tpl.header}".`);
  }

  // ===================== PHT =====================
  if (mode === "pht") {
    const { toRemove } = computeRemoveSet(blocks);
    result.answersRemoved = removeBlocks(body, idx, toRemove);
    notes.push(`Đã xóa ${result.answersRemoved} block đáp án/hướng dẫn/bảng giải thích.`);

    // Blank hóa lý thuyết — trong MỌI vùng kiến thức (heading mở vùng mới),
    // chỉ blank lý thuyết TRƯỚC câu hỏi đầu của vùng (khớp PHT mẫu).
    let blanks = 0;
    let questionSeenInSection = false;
    blocks.forEach((b, i) => {
      if (b.kind === "table" || i >= idx.length || idx[i].isTable) return;
      const blk = b as Block;
      if (blk.kind === "heading") questionSeenInSection = false;
      if (blk.kind === "question") questionSeenInSection = true;
      if (questionSeenInSection || blk.kind !== "theory") return;
      const p = idx[i].el as XElement;
      const raw = paragraphText(p); // RAW — giữ tab đầu
      const txt = raw.trim();
      if (!txt || txt.length > 600) return;
      if (hasMath(p)) return;
      const offset = raw.length - raw.trimStart().length;
      const regions = findBlankRegions(txt).map(([s, e]) => [s + offset, e + offset] as [number, number]);
      if (!regions.length) return;
      blanks += applyBlanks(p, doc, regions);
    });
    result.blanksCreated = blanks;
    notes.push(`Đã tạo ${blanks} chỗ trống kiến thức.`);
  }

  // ===================== LIVE =====================
  if (mode === "live") {
    // ----- 1) Template kỳ thi: trang NGANG + margin đúng mẫu -----
    const sectPr = childrenOf(body).find((c) => isW(c, "sectPr"));
    if (sectPr) {
      const pgSz = childrenOf(sectPr).find((c) => isW(c, "pgSz"));
      if (pgSz) {
        pgSz.setAttribute("w:orient", "landscape");
        pgSz.setAttribute("w:w", "16840");
        pgSz.setAttribute("w:h", "11907");
      } else {
        const sz = createW(doc, "pgSz");
        sz.setAttribute("w:orient", "landscape");
        sz.setAttribute("w:w", "16840");
        sz.setAttribute("w:h", "11907");
        sectPr.appendChild(sz);
      }
      const pgMar = childrenOf(sectPr).find((c) => isW(c, "pgMar"));
      if (pgMar) {
        pgMar.setAttribute("w:top", "1560");
        pgMar.setAttribute("w:right", "720");
        pgMar.setAttribute("w:bottom", "720");
        pgMar.setAttribute("w:left", "720");
        pgMar.setAttribute("w:header", "284");
        pgMar.setAttribute("w:footer", "284");
        pgMar.setAttribute("w:gutter", "0");
      }
    }

    // ----- 2) Xóa mọi HIGHLIGHT (đánh dấu màu đáp án) -----
    const hlNodes: XElement[] = [];
    const walkHl = (n: XNode) => {
      for (let i = 0; i < n.childNodes.length; i++) {
        const c = n.childNodes[i];
        if (localName(c) === "highlight") hlNodes.push(c as XElement);
        walkHl(c);
      }
    };
    walkHl(body);
    for (const hl of hlNodes) hl.parentNode!.removeChild(hl);

    // ----- 3) Xóa đáp án + hướng dẫn giải -----
    const { toRemove } = computeRemoveSet(blocks);
    result.answersRemoved = removeBlocks(body, idx, toRemove);

    // ----- 4) Gom câu hỏi vào BẢNG 2 cột viền MÀU TEMPLATE (form Live) -----
        //    Nhóm = câu hỏi + phương án/theory ngắn liền sau, đến câu kế/heading/instruction/table.
        const RED = tpl.color;
    const groups: Array<{ elems: XElement[] }> = [];
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.kind === "table") continue;
      if ((b as Block).kind !== "question") continue;
      const els: XElement[] = [];
      let j = i;
      while (j < blocks.length) {
        const bb = blocks[j];
        if (bb.kind === "table") break;
        const bblk = bb as Block;
        if (j > i && (bblk.kind === "question" || bblk.kind === "heading" || bblk.kind === "instruction")) break;
        if (j < idx.length && !idx[j].isTable) {
          const el = idx[j].el as XElement;
          if (el.parentNode) els.push(el); // bỏ phần tử đã bị xóa (trong khối giải)
        }
        j++;
      }
      if (els.length) groups.push({ elems: els });
      i = j - 1;
    }

    // Xây bảng 1x2 và DI CHUYỂN paragraph gốc vào cột trái
    for (const g of groups) {
      const tbl = buildQuestionTable(doc, RED);
      const first = g.elems[0];
      first.parentNode!.insertBefore(tbl, first);
      for (const el of g.elems) el.parentNode!.removeChild(el);
      const leftCell = findLeftCell(tbl);
      if (leftCell) {
        for (const el of g.elems) leftCell.appendChild(el);
      }
    }

    notes.push(
      `Live: template ngang + ${hlNodes.length} highlight đã bỏ + ${result.answersRemoved} block đáp án/hướng dẫn đã xóa + ${groups.length} bảng câu hỏi có dòng kẻ.`,
    );
  }

  return result;
}

/** Tìm ô (tc) đầu tiên của bảng */
function findLeftCell(tbl: XElement): XElement | null {
  for (const c of childrenOf(tbl)) {
    if (isW(c, "tr")) {
      for (const tc of childrenOf(c)) {
        if (isW(tc, "tc")) return tc;
      }
    }
  }
  return null;
}

/** Tạo bảng 1x2 viền đỏ: cột trái chứa câu hỏi, cột phải chứa dòng kẻ viết */
function buildQuestionTable(doc: XDocument, red: string): XElement {
  const tbl = createW(doc, "tbl");
  const tblPr = createW(doc, "tblPr");
  const tblW = createW(doc, "tblW");
  tblW.setAttribute("w:w", "0");
  tblW.setAttribute("w:type", "auto");
  tblPr.appendChild(tblW);
  const jc = createW(doc, "jc");
  jc.setAttribute("w:val", "center");
  tblPr.appendChild(jc);
  const borders = createW(doc, "tblBorders");
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    const b = createW(doc, side);
    b.setAttribute("w:val", "single");
    b.setAttribute("w:sz", "4");
    b.setAttribute("w:space", "0");
    b.setAttribute("w:color", red);
    borders.appendChild(b);
  }
  tblPr.appendChild(borders);
  tbl.appendChild(tblPr);

  const grid = createW(doc, "tblGrid");
  for (const w of [9000, 6000]) {
    const gc = createW(doc, "gridCol");
    gc.setAttribute("w:w", String(w));
    grid.appendChild(gc);
  }
  tbl.appendChild(grid);

  const tr = createW(doc, "tr");
  // ô trái — câu hỏi (paragraph di chuyển vào sau)
  const tcL = createW(doc, "tc");
  const tcPrL = createW(doc, "tcPr");
  const tcWL = createW(doc, "tcW");
  tcWL.setAttribute("w:w", "9000");
  tcWL.setAttribute("w:type", "dxa");
  tcPrL.appendChild(tcWL);
  tcL.appendChild(tcPrL);
  tcL.appendChild(createW(doc, "p")); // paragraph đệm để ô hợp lệ
  tr.appendChild(tcL);

  // ô phải — dòng kẻ để GV viết chữa bài
  const tcR = createW(doc, "tc");
  const tcPrR = createW(doc, "tcPr");
  const tcWR = createW(doc, "tcW");
  tcWR.setAttribute("w:w", "6000");
  tcWR.setAttribute("w:type", "dxa");
  tcPrR.appendChild(tcWR);
  tcR.appendChild(tcPrR);
  tcR.appendChild(makeUnderlinePara(doc, 70));
  tcR.appendChild(makeUnderlinePara(doc, 70));
  tr.appendChild(tcR);

  tbl.appendChild(tr);
  return tbl;
}

/** Paragraph dòng kẻ: underline spaces để viết lên trên */
function makeUnderlinePara(doc: XDocument, nSpaces: number): XElement {
  const p = createW(doc, "p");
  const pPr = createW(doc, "pPr");
  const ind = createW(doc, "ind");
  ind.setAttribute("w:left", "120");
  pPr.appendChild(ind);
  p.appendChild(pPr);
  const r = createW(doc, "r");
  const rPr = createW(doc, "rPr");
  const u = createW(doc, "u");
  u.setAttribute("w:val", "single");
  rPr.appendChild(u);
  r.appendChild(rPr);
  const t = createW(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(" ".repeat(nSpaces)));
  r.appendChild(t);
  p.appendChild(r);
  return p;
}

function createW(doc: XDocument, name: string): XElement {
  return createWElement(doc, name);
}

/**
 * Tạo paragraph banner nhận diện kỳ thi: in đậm, màu template, căn giữa.
 * (Chỉ dùng cho chế độ không template-file giữ nguyên ORIGINAL.
 */
function makeBannerPara(doc: XDocument, text: string, color: string): XElement {
  const p = createW(doc, "p");
  const pPr = createW(doc, "pPr");
  const jc = createW(doc, "jc");
  jc.setAttribute("w:val", "center");
  pPr.appendChild(jc);
  const spacing = createW(doc, "spacing");
  spacing.setAttribute("w:after", "160");
  pPr.appendChild(spacing);
  p.appendChild(pPr);
  const r = createW(doc, "r");
  const rPr = createW(doc, "rPr");
  const b = createW(doc, "b");
  rPr.appendChild(b);
  const col = createW(doc, "color");
  col.setAttribute("w:val", color);
  rPr.appendChild(col);
  const sz = createW(doc, "sz");
  sz.setAttribute("w:val", "32"); // 16pt
  rPr.appendChild(sz);
  r.appendChild(rPr);
  const t = createW(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  p.appendChild(r);
  return p;
}

/** API ổn định */
export function convertParsed(
  blocks: DocBlock[],
  body: XElement,
  mode: OutputMode,
  templateId: TemplateId = "tsa",
  useFileTemplate = false,
): ConvertResult {
  return convertDom(blocks, body, mode, templateId, useFileTemplate);
}