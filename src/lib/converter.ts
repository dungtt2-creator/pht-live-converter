/**
 * Converter: chuyển DOM document.xml đã parse thành PHT hoặc Live.
 *
 * PHT  (Phiếu học tập):  giữ mục tiêu + lý thuyết (blank hóa định nghĩa) + câu hỏi.
 *                        Xóa: "Đáp án:", "Hướng dẫn:", "Lời giải:", bảng giải thích.
 * Live (Tài liệu Live): giữ nguyên 100% nội dung GV (không blank, không xóa kiến thức).
 *
 * Nguyên tắc bất biến: không sửa nội dung kiến thức; chỉ xóa/che phần đáp án (PHT)
 * và thêm cấu trúc hỗ trợ giảng dạy (Live).
 */

import { localName, childrenOf, isW, paragraphText, createWElement } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";
import { applyBlanks, findBlankRegions } from "./blanker";

export type OutputMode = "pht" | "live";

export interface ConvertResult {
  blocks: DocBlock[];
  answersRemoved: number;
  guidelinesRemoved: number;
  blanksCreated: number;
  notes: string[];
}

const QNUM = /^[Cc]âu\s*(\d+)\s*[.:]/;

function isQuestionPara(p: XElement): boolean {
  return QNUM.test(paragraphText(p).trim());
}

/** Bảng câu hỏi đúng/sai (cột Đúng/Sai) hoặc bảng giải thích (cột Đ/A) */
function isAnswerTable(tbl: XElement): boolean {
  const trs = childrenOf(tbl).filter((c) => isW(c, "tr"));
  if (!trs.length) return false;
  const first = childrenOf(trs[0]).filter((c) => isW(c, "tc"));
  const headerText = first.map((tc) => {
    const p = childrenOf(tc).find((c) => isW(c, "p"));
    return p ? paragraphText(p) : "";
  }).join(" ");
  return /Đ\/A|Giải thích|Nhận định/.test(headerText) && /Đúng|Sai|Đ\/A/.test(headerText);
}

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

/**
 * Thu thập các w:p và w:tbl trong body theo thứ tự → mảng đánh chỉ số chung
 */
function indexBody(body: XElement): Array<{ el: XElement; isTable: boolean }> {
  const out: Array<{ el: XElement; isTable: boolean }> = [];
  for (const child of childrenOf(body)) {
    if (isW(child, "p")) out.push({ el: child, isTable: false });
    else if (isW(child, "tbl")) out.push({ el: child, isTable: true });
  }
  return out;
}

/**
 * Chuyển đổi chính: thao tác trực tiếp trên body XML.
 */
export function convertDom(blocks: DocBlock[], body: XElement, mode: OutputMode): ConvertResult {
  const notes: string[] = [];
  const result: ConvertResult = {
    blocks, answersRemoved: 0, guidelinesRemoved: 0, blanksCreated: 0, notes,
  };

  const idx = indexBody(body);
  if (idx.length !== blocks.length) {
    notes.push(`Cảnh báo: số block model (${blocks.length}) khác số phần tử body (${idx.length}) — có thể có phần tử lạ.`);
  }

  const doc = body.ownerDocument;

  // ===================== PHT =====================
    if (mode === "pht") {
      // Nguồn tham chiếu: PHT mẫu do người biên soạn.
      // - GIỮ: mục tiêu, lý thuyết + đoạn dài (đề chung "Sử dụng thông tin...", ngữ liệu),
      //         câu hỏi, bảng đúng/sai/điền từ, "DẶN DÒ".
      // - XÓA CHỈ: khối "Đáp án/Hướng dẫn/Lời giải" + nội dung giải ngay sau.
      // Điểm dừng khối giải: câu hỏi mới, heading, "Sử dụng thông tin...", đề chung
      // điều kiện, hay hết tài liệu.
      const toRemove = new Set<number>();
            const QNUM_LOCAL = /^[Cc]âu\s*(\d+)\s*[.:]/;
            // Điểm dừng khối giải = NGỮ LIỆU/ĐỀ mới (phải GIỮ)
            const NEW_CONTEXT =
              /^(Sử dụng|Dựa vào|Đọc|Xem xét|Căn cứ|Kết hợp)\s+(thông tin|đoạn|bài|hình|bảng|đề|đồ thị)/i;
            const PRESERVED_CONTEXT = /^Bài đọc|^Đoạn văn sau|^Thông tin sau|^DẶN DÒ|^Dặn dò|^Nguồn\s*[:：]|^Trích\s*(SGK|SBT)/;
            // "Đáp án: A. ____; ____ B. ____" = PHƯƠNG ÁN ĐIỀN (câu hỏi) — GIỮ, không xóa
                        // "Đáp án: A. ____; ____ B. ____" = PHƯƠNG ÁN ĐIỀN (câu hỏi) — GIỮ, không xóa.
                                    // "Đáp án. A. 1,3; B. 2,4; C. 5" = ĐÁP ÁN THẬT — XÓA (không phải mẫu điền).
                                    const ANSWER_WITH_CHOICES = /^(Đáp án|Đ\/A)\s*[:.]?\s*[A-D]\./;
                                    const FILL_BLANK_ONLY = /_{2,}/;
      let inGuideline = false; // đang duyệt khối đáp án/hướng dẫn
      for (let i = 0; i < blocks.length; i++) {
        const b = blocks[i];
        if (b.kind === "table") {
                  const t = b as unknown as { rows: { text: string }[][] };
                  const header = (t.rows[0] || []).map((c) => c.text).join(" ");
                  // Bảng GIẢI THÍCH đáp án — NHẬN DIỆN ĐỘC LẬP (xóa luôn, không cần inGuideline):
                  //  1) header chứa "Đ/A" + "Giải thích" ("Ý | Đ/A | Giải thích")
                  //  2) cột đầu dòng 0 khớp "a) Đúng/Sai"/"a Đúng" + có nội dung giải (>2 ký tự)
                  // KHÔNG nhầm bảng câu hỏi Đ/S: header "Phát biểu | Đúng | Sai" (cột Đúng/Sai rỗng để HS tick).
                  const isExplanation =
                    (/Đ\/A/.test(header) && /Giải thích/.test(header)) ||
                    ((t.rows[0] || []).some((c) => /^a\s*\)?\s*(Đúng|Sai)\b/.test(c.text.trim())) &&
                     (t.rows[1] || []).some((c) => c.text.trim().length > 2));
                  if (isExplanation) toRemove.add(i);
                  continue;
                }
        const blk = b as Block;
        const txt = blk.text.trim();
        if (inGuideline) {
                          // Điểm dừng khối giải: câu hỏi mới / heading / ngữ liệu mới
                          if (
                            QNUM_LOCAL.test(txt) ||
                            blk.kind === "heading" ||
                            NEW_CONTEXT.test(txt) ||
                            PRESERVED_CONTEXT.test(txt)
                          ) {
                            inGuideline = false;
                            continue; // không xóa block dừng
                          }
                          if (ANSWER_WITH_CHOICES.test(txt) && FILL_BLANK_ONLY.test(txt)) {
                            // "Đáp án: A. ____" = mẫu điền — GIỮ, không kết thúc khối
                            continue;
                          }
                          // nội dung giải / đáp án thật — xóa
                          toRemove.add(i);
                          continue;
                        }
                                  if (blk.kind === "guideline" || blk.kind === "answer") {
                                              // "Đáp án: A. ____" = mẫu điền — KHÔNG mở khối xóa
                                              if (ANSWER_WITH_CHOICES.test(txt) && FILL_BLANK_ONLY.test(txt)) continue;
                                              toRemove.add(i);
                                              inGuideline = true;
                                            }
      }

      let removedCount = 0;
      for (let i = 0; i < idx.length; i++) {
        if (toRemove.has(i)) {
          idx[i].el.parentNode!.removeChild(idx[i].el);
          removedCount++;
        }
      }
      result.answersRemoved = removedCount;
      notes.push(`Đã xóa ${removedCount} block thuộc khối đáp án/hướng dẫn.`);

    // 3) Blank hóa lý thuyết — trong MỌI vùng kiến thức (mỗi heading I/II/III mở vùng mới),
        //    chỉ blank các lý thuyết TRƯỚC câu hỏi đầu của vùng đó (khớp PHT mẫu).
        //    QUAN TRỌNG: tọa độ region phải tính trên RAW paragraphText(p) (chưa trim)
        //    vì applyBlanks khớp tọa độ trên cùng text đó.
        let blanks = 0;
        let questionSeenInSection = false; // reset mỗi khi gặp heading
        blocks.forEach((b, i) => {
          if (b.kind === "table" || i >= idx.length || idx[i].isTable) return;
          const blk = b as Block;
          if (blk.kind === "heading") questionSeenInSection = false;
          if (blk.kind === "question") questionSeenInSection = true;
          if (questionSeenInSection || blk.kind !== "theory") return;
          const p = idx[i].el as XElement;
                    const raw = paragraphText(p); // RAW — giữ tab/khoảng trắng đầu
                    const txt = raw.trim();
                    if (!txt || txt.length > 600) return;
                    if (hasMath(p)) return;
                    // regions tính trên trimmed; shift sang raw (leading whitespace)
                    const offset = raw.length - raw.trimStart().length;
                    const regions = findBlankRegions(txt)
                      .map(([s, e]) => [s + offset, e + offset] as [number, number]);
                    if (!regions.length) return;
                    blanks += applyBlanks(p, doc, regions);
        });
    result.blanksCreated = blanks;
    notes.push(`Đã tạo ${blanks} chỗ trống kiến thức.`);
  }

  // ===================== LIVE =====================
    if (mode === "live") {
      // 1) Chuyển trang sang NGANG (landscape) — dễ giáo viên viết khi chữa bài.
      //    Chỉnh sectPr cuối body (w:pgSz w:orient="landscape", đổi w/h).
      const sectPr = childrenOf(body).find((c) => isW(c, "sectPr"));
      if (sectPr) {
        const pgSz = childrenOf(sectPr).find((c) => isW(c, "pgSz"));
        if (pgSz) {
          pgSz.setAttribute("w:orient", "landscape");
          pgSz.setAttribute("w:w", "16838"); // A4 dọc 11906 → ngang 16838
          pgSz.setAttribute("w:h", "11906");
        } else {
          const sz = createW(doc, "pgSz");
          sz.setAttribute("w:orient", "landscape");
          sz.setAttribute("w:w", "16838");
          sz.setAttribute("w:h", "11906");
          sectPr.appendChild(sz);
        }
      }
      // 2) Giữ nguyên nội dung GV; thêm vùng trả lời = DÒNG KẺ (underline spaces)
      //    sau mỗi câu hỏi để giáo viên ghi trong lúc chữa bài.
      const parIdx = new Map<number, XElement>();
      blocks.forEach((b, i) => {
        if (b.kind !== "table" && i < idx.length) {
          parIdx.set(i, idx[i].el as XElement);
        }
      });
      let inserted = 0;
      for (let i = blocks.length - 1; i >= 0; i--) {
        const b = blocks[i];
        if (b.kind === "table") continue;
        const blk = b as Block;
        if (blk.kind !== "question") continue;
        const p = parIdx.get(i);
        if (!p) continue;
        const ans = createWritingLine(doc);
        p.parentNode!.insertBefore(ans, p.nextSibling);
        inserted++;
      }
      result.notes.push(`Live: trang ngang + ${inserted} dòng kẻ sau câu hỏi (giữ nguyên nội dung GV).`);
    }

  return result;
}

/** Tạo paragraph dòng kẻ (underline spaces) để giáo viên viết khi chữa bài */
function createWritingLine(doc: XDocument): XElement {
  const p = createW(doc, "p");
  const pPr = createW(doc, "pPr");
  const ind = createW(doc, "ind");
  ind.setAttribute("w:left", "360");
  ind.setAttribute("w:right", "360");
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
  // 100 khoảng trắng gạch dưới → kẻ liền nét, viết được lên trên
  t.appendChild(doc.createTextNode(" ".repeat(100)));
  r.appendChild(t);
  p.appendChild(r);
  return p;
}

function createW(doc: XDocument, name: string): XElement {
  return createWElement(doc, name);
}

/** API ổn định */
export function convertParsed(blocks: DocBlock[], body: XElement, mode: OutputMode): ConvertResult {
  return convertDom(blocks, body, mode);
}