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
import { addHocmaiTags } from "./tag-hocmai";
import { getTemplate, type TemplateId } from "./template";
import { fixHeadingKeepNext, cleanupSpacers } from "./pagination";

export type OutputMode = "pht" | "live";

export interface ConvertResult {
  blocks: DocBlock[];
  answersRemoved: number;
  guidelinesRemoved: number;
  blanksCreated: number;
  /** Số tag [HOCMAI] tự động thêm cho câu chưa chuẩn hóa */
  hocmaiTagsAdded: number;
  notes: string[];
}

const QNUM = /^[Cc]âu\s*(\d{1,3})\s*(?:[:.]|\s*\[(?:HOCMAI|Trắc nghiệm)[^\]]*\]|\s*$)/;
const NEW_CONTEXT =
  /^(Sử dụng|Dựa vào|Đọc|Xem xét|Căn cứ|Kết hợp)\s+(thông tin|đoạn|bài|hình|bảng|đề|đồ thị)/i;
const PRESERVED_CONTEXT =
  /^Bài đọc|^Đoạn văn sau|^Thông tin sau|^DẶN DÒ|^Dặn dò|^Nguồn\s*[:：]|^Trích\s*(SGK|SBT)/;
/** "Đáp án: A. ____" dạng mẫu điền — chỉ giữ khi có gạch chân */
const ANSWER_WITH_CHOICES = /^(Đáp án|Đ\/A)\s*[:.]?\s*[A-D]\./;
/** "Đáp án. A. 1,3; B. 2,4; C. 5" — đáp án câu NỐI CỘT: mẫu GIỮ + blank số thành ____ */
const FILL_JOIN_ANSWER = /^Đáp\s?án\s*[:.]\s*A\.[^B]*;(?:\s*)B\./i;
const FILL_BLANK_ONLY = /_{2,}/;
/**
 * Điểm bắt đầu vùng ĐÁP ÁN/giải CUỐI FILE (đề thi thật TN THPT/HSA/QDA
 * có "ĐÁP ÁN THAM KHẢO" + bảng đáp án + "LỜI GIẢI THAM KHẢO" dưới cùng):
 * viết HOA toàn bộ + theo sau là chữ HOA → xóa TỚI HẾT FILE (kể cả bảng/heading/câu).
 * Không khớp "Đáp án: A." (chữ thường, mẫu điền câu nối cột).
 */
const ANSWER_KEY_START =
  /^(?:ĐÁP ÁN\s+[A-ZÀ-Ỹ]|BẢNG ĐÁP ÁN|LỜI GIẢI\s+[A-ZÀ-Ỹ]|HƯỚNG DẪN GIẢI\s+[A-ZÀ-Ỹ]|ĐÁP ÁN ĐỀ)/;
/** Stimulus dạng TN THPT: "Nội dung câu 1 và 2: ..." — giữ ngoài khung + ngắt nhóm */
const CONTEXT_BODY_RE = /^Nội dung\s+(?:câu|các\s+câu)/i;

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
function computeRemoveSet(blocks: DocBlock[]): { toRemove: Set<number>; joinAnswerBlocks: number[] } {
  const toRemove = new Set<number>();
  const joinAnswerBlocks: number[] = [];
  let inG = false; // đang duyệt khối đáp án/hướng dẫn
  let inAnswerKey = false; // vùng ĐÁP ÁN/giải CUỐI FILE (xóa tới hết)

  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    // Vùng đáp án/giải tham khảo cuối file (đề thi thật TN THPT/HSA/QDA):
    // từ dòng "ĐÁP ÁN THAM KHẢO"/"BẢNG ĐÁP ÁN"/"LỜI GIẢI..." → xóa TẤT CẢ
    // (kể cả bảng đáp án + heading + "Câu N. Chọn X" giả) tới hết tài liệu.
    if (b.kind !== "table" && ANSWER_KEY_START.test((b as Block).text.trim())) {
      inAnswerKey = true;
    }
    if (inAnswerKey) {
      toRemove.add(i);
      continue;
    }
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
          if (FILL_JOIN_ANSWER.test(txt)) {
            // Đáp án câu nối cột: GIỮ + blank số thành ____ (đúng mẫu Live Bài 13)
            // Không mở inG — khối "Hướng dẫn." đứng riêng sau sẽ tự mở.
            joinAnswerBlocks.push(i);
            continue;
          }
          toRemove.add(i);
          inG = true;
        }
  }
  return { toRemove, joinAnswerBlocks };
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
 * Blank dòng đáp án câu nối cột ("Đáp án. A. 1,3;B. 2,4;C. 5" → "Đáp án: A. ____; ____ B. ____; ____; C. ____").
 * Lặp trên RAW paragraphText (giữ tab đầu) — thay mọi SỐ bằng ____, giữ nhãn A./B./C. và dấu ;.
 */
function blankJoinAnswers(
  body: XElement,
  idx: Array<{ el: XElement; isTable: boolean }>,
  targets: number[],
): number {
  if (!targets.length) return 0;
  const NUM_RE = /[0-9][0-9.,]*/g;
  const visit = (n: XNode): void => {
    if (localName(n) === "t") {
      const tn = n as unknown as { textContent: string | null; nodeValue: string | null };
      const cur = tn.textContent ?? tn.nodeValue ?? "";
      if (NUM_RE.test(cur)) tn.textContent = cur.replace(NUM_RE, "____");
    }
    for (let i = 0; i < n.childNodes.length; i++) visit(n.childNodes[i]);
  };
  for (const i of targets) {
    if (i < idx.length && !idx[i].isTable) visit(idx[i].el as XNode);
  }
  return targets.length;
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
      blocks, answersRemoved: 0, guidelinesRemoved: 0, blanksCreated: 0, hocmaiTagsAdded: 0, notes,
    };
  const tpl = getTemplate(templateId);
  const idx = indexBody(body);
  if (idx.length !== blocks.length) {
    notes.push(`Cảnh báo: số block model (${blocks.length}) khác số phần tử body (${idx.length}).`);
  }
  const doc = body.ownerDocument;

    // ===== 0) KHÔNG chèn banner/template — giữ NGUYÊN phần đầu file gốc (tiêu đề bài học
    //     của GV). Yêu cầu user: "Phần tiêu đề của bài học cần giữ cố định như file gốc,
    //     không tự chỉnh sửa tiêu đề của file". Trước đây template TSA inject banner
    //     "TỔNG ÔN MÔN VẬT LÍ SỰ CHUYỂN THỂ..." vào MỌI file → tiêu đề bài sai =====

  // ===================== PHT =====================
      if (mode === "pht") {
        // 0b) Chuẩn hóa: tự thêm tag [HOCMAI] cho câu chưa có (theo mẫu Bài 13)
        result.hocmaiTagsAdded = addHocmaiTags(blocks, body, doc, tpl.color);
        if (result.hocmaiTagsAdded > 0) {
          notes.push(`Đã tự động thêm tag [HOCMAI] cho ${result.hocmaiTagsAdded} câu chưa chuẩn hóa.`);
        }

        // 0c) Xóa HIGHLIGHT đánh dấu đáp án đúng (yêu cầu: đưa phương án về đồng nhất,
        //     giống phương án không được đánh dấu — không lấy màu highlight cho tag)
        const hlNodesPht: XElement[] = [];
        const walkHlPht = (n: XNode) => {
          for (let i = 0; i < n.childNodes.length; i++) {
            const c = n.childNodes[i];
            if (localName(c) === "highlight") hlNodesPht.push(c as XElement);
            walkHlPht(c);
          }
        };
        walkHlPht(body);
        for (const hl of hlNodesPht) hl.parentNode!.removeChild(hl);

        const { toRemove, joinAnswerBlocks } = computeRemoveSet(blocks);
            result.answersRemoved = removeBlocks(body, idx, toRemove);
            const joinBlanks = blankJoinAnswers(body, idx, joinAnswerBlocks);
            if (joinBlanks > 0) notes.push(`Đã blank ${joinBlanks} dòng đáp án nối cột (giữ mẫu điền).`);
            notes.push(`Đã xóa ${result.answersRemoved} block đáp án/hướng dẫn/bảng giải thích + ${hlNodesPht.length} highlight đáp án.`);

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

      // ----- 5) Phân trang: heading đi cùng đoạn đầu + dọn khoảng trống phát sinh -----
      const hFixed = fixHeadingKeepNext(blocks, idx, doc);
      if (hFixed > 0) notes.push(`Đã bật keep-with-next + widow-control cho ${hFixed} heading (giữ heading với đoạn đầu tiên).`);
      const cleaned = cleanupSpacers(blocks, idx, toRemove, doc);
      if (cleaned.emptyRemoved > 0 || cleaned.pageBreaksRemoved > 0) {
        notes.push(`Đã dọn ${cleaned.emptyRemoved} paragraph rỗng thừa + ${cleaned.pageBreaksRemoved} page break phát sinh.`);
      }
    }

  // ===================== LIVE =====================
    if (mode === "live") {
      // ----- 0) Chuẩn hóa: tự thêm tag [HOCMAI] (chạy trên DOM — trước khi xóa) -----
      result.hocmaiTagsAdded = addHocmaiTags(blocks, body, doc, tpl.color);
      if (result.hocmaiTagsAdded > 0) {
        notes.push(`Đã tự động thêm tag [HOCMAI] cho ${result.hocmaiTagsAdded} câu chưa chuẩn hóa.`);
      }

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
        const { toRemove, joinAnswerBlocks } = computeRemoveSet(blocks);
        result.answersRemoved = removeBlocks(body, idx, toRemove);
        const joinBlanks = blankJoinAnswers(body, idx, joinAnswerBlocks);
        if (joinBlanks > 0) notes.push(`Đã blank ${joinBlanks} dòng đáp án nối cột (giữ mẫu điền).`);

    // ----- 4) Gom câu hỏi vào BẢNG 2 cột viền WAVE XANH (đúng mẫu Live Bài 13) -----
        const WAVE = "00B0F0";
        const groups: Array<{ elems: XElement[]; nQuestions: number }> = [];
        for (let i = 0; i < blocks.length; i++) {
          const b = blocks[i];
          if (b.kind === "table" || (b as Block).kind !== "question") continue;
          const els: XElement[] = [];

                    // Mẫu Bài 13: stimulus "Sử dụng thông tin..." GIỮ NGUYÊN ngoài khung
                    // (paragraph thường trước khung). Gom tối đa 3 câu hỏi liên tiếp/khung —
                    // mẫu xen 2-3 câu theo dung lượng nội dung, 3 là cân bằng đẹp.

                    let j = i;
                          let nQ = 0;
                          while (j < blocks.length) {
                            const bb = blocks[j];
                            if (bb.kind === "table") break; // bảng ngữ liệu: ngắt nhóm, để nguyên ngoài khung
                            const bblk = bb as Block;
                            if (j > i && (bblk.kind === "heading" || bblk.kind === "instruction")) break;
                            if (bblk.kind === "question") {
                              if (nQ >= 3) break; // mẫu gom 2-3 câu liên tiếp/khung
                              nQ++;
                            }
                  if (j < idx.length && !idx[j].isTable) {
                    const el = idx[j].el as XElement;
                    if (el.parentNode) els.push(el); // bỏ phần tử đã bị xóa (trong khối giải)
                  }
                  j++;
                }
          if (els.length) groups.push({ elems: els, nQuestions: nQ });
          i = j - 1;
        }

        // Xây bảng wave 1x2 và DI CHUYỂN paragraph gốc vào cột trái; cột phải = dòng chấm
        for (const g of groups) {
          const nDots = Math.max(6, 7 * g.nQuestions);
          const tbl = buildQuestionTable(doc, WAVE, nDots);
          const first = g.elems[0];
          first.parentNode!.insertBefore(tbl, first);
          for (const el of g.elems) el.parentNode!.removeChild(el);
          const leftCell = findLeftCell(tbl);
          if (leftCell) {
            for (const el of g.elems) leftCell.appendChild(el);
          }
        }

        notes.push(
                  `Live: template ngang + ${hlNodes.length} highlight đã bỏ + ${result.answersRemoved} block đáp án/hướng dẫn đã xóa + ${groups.length} khung câu hỏi viền wave (${groups.reduce((s, g) => s + g.nQuestions, 0)} câu, dòng chấm chữa bài).`,
                );

                // ----- 5) Phân trang: heading đi cùng đoạn đầu + dọn khoảng trống phát sinh -----
                        // Live: strip bảng tiêu đề 1-cell → paragraph (keepNext không xuyên table,
                        // heading trong bảng vẫn rơi cuối trang — khớp MAU dùng paragraph thường).
                        const hFixedL = fixHeadingKeepNext(blocks, idx, doc, true);
                        if (hFixedL > 0) notes.push(`Đã chuẩn hóa ${hFixedL} heading (bảng tiêu đề → paragraph + keep-with-next + widow-control).`);
                const cleanedL = cleanupSpacers(blocks, idx, toRemove, doc);
                if (cleanedL.emptyRemoved > 0 || cleanedL.pageBreaksRemoved > 0) {
                  notes.push(`Đã dọn ${cleanedL.emptyRemoved} paragraph rỗng thừa + ${cleanedL.pageBreaksRemoved} page break phát sinh.`);
                }
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

/**
 * Tạo bảng 2 cột viền WAVE xanh (đúng mẫu Live Bài 13):
 * cột trái chứa câu hỏi (9876 dxa), cột phải chứa nDots dòng chấm để GV chữa bài (5821 dxa).
 * Viền wave màu 00B0F0 sz=12; rPr dòng chấm = Montserrat italic 10pt màu 00B0F0.
 */
function buildQuestionTable(doc: XDocument, wave: string, nDots: number): XElement {
  const tbl = createW(doc, "tbl");
  const tblPr = createW(doc, "tblPr");
  const tblStyle = createW(doc, "tblStyle");
  tblStyle.setAttribute("w:val", "TableGrid");
  tblPr.appendChild(tblStyle);
  const tblW = createW(doc, "tblW");
  tblW.setAttribute("w:w", "15697");
  tblW.setAttribute("w:type", "dxa");
  tblPr.appendChild(tblW);
  const borders = createW(doc, "tblBorders");
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    const b = createW(doc, side);
    b.setAttribute("w:val", "wave");
    b.setAttribute("w:sz", "12");
    b.setAttribute("w:space", "0");
    b.setAttribute("w:color", wave);
    borders.appendChild(b);
  }
  tblPr.appendChild(borders);
  const tblLook = createW(doc, "tblLook");
  tblLook.setAttribute("w:val", "04A0");
  tblLook.setAttribute("w:firstRow", "1");
  tblLook.setAttribute("w:lastRow", "0");
  tblLook.setAttribute("w:firstColumn", "1");
  tblLook.setAttribute("w:lastColumn", "0");
  tblLook.setAttribute("w:noHBand", "0");
  tblLook.setAttribute("w:noVBand", "1");
  tblPr.appendChild(tblLook);
  tbl.appendChild(tblPr);

  const grid = createW(doc, "tblGrid");
  for (const w of [9876, 5821]) {
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
  tcWL.setAttribute("w:w", "9876");
  tcWL.setAttribute("w:type", "dxa");
  tcPrL.appendChild(tcWL);
  tcL.appendChild(tcPrL);
  tcL.appendChild(createW(doc, "p")); // paragraph đệm để ô hợp lệ
  tr.appendChild(tcL);

  // ô phải — dòng chấm Montserrat italic 10pt màu wave (GV viết chữa bài)
  const tcR = createW(doc, "tc");
  const tcPrR = createW(doc, "tcPr");
  const tcWR = createW(doc, "tcW");
  tcWR.setAttribute("w:w", "5821");
  tcWR.setAttribute("w:type", "dxa");
  tcPrR.appendChild(tcWR);
  tcR.appendChild(tcPrR);
  for (let k = 0; k < nDots; k++) tcR.appendChild(makeDotsPara(doc, wave));
  tr.appendChild(tcR);

  tbl.appendChild(tr);
  return tbl;
}

/** Paragraph dòng chấm "………..…" — đúng mẫu: Montserrat italic 10pt màu wave */
function makeDotsPara(doc: XDocument, wave: string): XElement {
  const p = createW(doc, "p");
  const pPr = createW(doc, "pPr");
  const rPrP = createW(doc, "rPr");
  const fontsP = createW(doc, "rFonts");
  fontsP.setAttribute("w:ascii", "Montserrat");
  fontsP.setAttribute("w:eastAsia", "Montserrat");
  fontsP.setAttribute("w:hAnsi", "Montserrat");
  fontsP.setAttribute("w:cs", "Montserrat");
  rPrP.appendChild(fontsP);
  rPrP.appendChild(createW(doc, "i"));
  rPrP.appendChild(createW(doc, "iCs"));
  const colorP = createW(doc, "color");
  colorP.setAttribute("w:val", wave);
  rPrP.appendChild(colorP);
  const szP = createW(doc, "sz");
  szP.setAttribute("w:val", "20"); // 10pt
  rPrP.appendChild(szP);
  const szCsP = createW(doc, "szCs");
  szCsP.setAttribute("w:val", "20");
  rPrP.appendChild(szCsP);
  pPr.appendChild(rPrP);
  p.appendChild(pPr);

  const r = createW(doc, "r");
  const rPr = createW(doc, "rPr");
  const fonts = createW(doc, "rFonts");
  fonts.setAttribute("w:ascii", "Montserrat");
  fonts.setAttribute("w:eastAsia", "Montserrat");
  fonts.setAttribute("w:hAnsi", "Montserrat");
  fonts.setAttribute("w:cs", "Montserrat");
  rPr.appendChild(fonts);
  rPr.appendChild(createW(doc, "i"));
  rPr.appendChild(createW(doc, "iCs"));
  const color = createW(doc, "color");
  color.setAttribute("w:val", wave);
  rPr.appendChild(color);
  const sz = createW(doc, "sz");
  sz.setAttribute("w:val", "20"); // 10pt
  rPr.appendChild(sz);
  const szCs = createW(doc, "szCs");
  szCs.setAttribute("w:val", "20");
  rPr.appendChild(szCs);
  r.appendChild(rPr);
  const t = createW(doc, "t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode("………..………………………………………………………………………………………..………………"));
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