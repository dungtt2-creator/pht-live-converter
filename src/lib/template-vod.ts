/**
 * TEMPLATE VOD (khổ ngang 32 × 18 cm — 16:9) — convert file đề GV → Tài liệu Live.
 *
 * Mẫu chị gửi (File mẫu/Template VOD.docx, public/templates/vod.docx):
 *  - Trang ngang 32 × 18 cm = 18144 × 10206 twips, margin 1 cm (567 twips) 4 phía,
 *    KHÔNG header/footer (không logo, không watermark).
 *  - Mỗi câu = 1 bảng 2 cột KHÔNG viền (tblBorders none):
 *      cột trái  ~5061097 EMU (~14.2 cm) = nội dung câu hỏi (Times New Roman 14 = sz 28)
 *      cột phải  phần còn lại          = các dòng chấm "……" Times New Roman 14 màu 00B0F0
 *    (đo từ mẫu: khung 28.3→423.1pt, dòng chấm 34.1→877.2pt, cách nhau 27.6pt).
 *  - Đáp án A./B./C./D. xếp 2 phương án/dòng (mẫu: A. … B. … / C. … D. …).
 *
 * KHÁC HOÀN TOÀN bản Live "wave" cũ: không khung bo góc viền, không bảng wave
 * 9876/5821, không dòng chấm Montserrat 10pt.
 *
 * Nguyên tắc:
 *  - Giữ NGUYÊN nội dung + định dạng gốc của từng paragraph (kể cả OLE Equation
 *    của công thức) — chỉ RELOCATE paragraph vào ô bảng, không dựng lại text.
 *  - Mọi run văn bản được chuẩn hóa Times New Roman 14pt (sz/szCs 28) theo yêu cầu
 *    "Dùng Times New Roman, cỡ chữ 14 cho phần văn bản soạn thảo" — RIÊNG run OLE
 *    (w:object) giữ rPr gốc để công thức không bị đổi cỡ.
 *  - Đánh đúng id câu hỏi: giữ nguyên "(ID: xxxxx)" của file đề; nếu file thiếu
 *    → chèn "(ID: —)" để GV thấy ngay chỗ cần bổ sung, KHÔNG tự bịa số.
 */

import { localName, childrenOf, isW, createWElement, paragraphText } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";
import type { LoadedTemplate } from "./template-loader";

export const VOD_COLOR = "00B0F0";
/** Times New Roman 14pt = sz 28 (half-point) */
const VOD_SZ = "28";
/** Khổ 32 × 18 cm */
const PG_W = "18144";
const PG_H = "10206";
const PG_MAR = "567";

/**
 * Đường kẻ chấm của mẫu: ĐÚNG 70 ký tự "…" (đếm trong Template VOD.docx) —
 * 70 ký tự ở Times New Roman 14 vừa khít 1 dòng của khổ 32 cm. Dùng nhiều hơn
 * sẽ tự wrap thành 2 dòng (mẫu không làm vậy).
 */
const DOTS_TEXT = "…".repeat(70);
/** Cỡ chữ của dòng chấm theo mẫu: 12pt (sz 24) — phần văn bản soạn thảo mới là 14pt */
const DOTS_SZ = "24";
/** Số dòng chấm mặc định cho 1 câu (mẫu: 12 dòng/trang chia cho 1 câu) */
/**
 * Số dòng chấm/câu. Mẫu VOD đặt 1 câu/trang với 12 dòng chấm (mỗi dòng 32pt =
 * spacing line 480). Nhưng đề HSA 50 câu mà 12 dòng/câu → 1 trang/câu (112 trang):
 * không dùng được khi dạy. Dùng 7 dòng/câu để khung câu hỏi + vùng chữa bài vừa
 * 1 trang ngang (16 cm khả dụng ≈ khung ~7 cm + 7 dòng × 32pt ≈ 7.9 cm).
 */
const DOTS_PER_QUESTION = 3;

/** Khai báo namespace còn thiếu (file GV hiếm khi khai báo vml/mc…) */
const NS_MAP: Record<string, string> = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  w10: "urn:schemas-microsoft-com:office:word",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  w16se: "http://schemas.microsoft.com/office/word/2015/wordml/symex",
};

// ============================ phần tử dựng mới ============================

function el(doc: XDocument, name: string, attrs: Record<string, string> = {}): XElement {
  const e = createWElement(doc, name);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  return e;
}

/** pPr rỗng tối thiểu (Times New Roman 14) */
function tplPPr(doc: XDocument): XElement {
  const pPr = el(doc, "pPr");
  const rPr = el(doc, "rPr");
  rPr.appendChild(el(doc, "rFonts", {
    "w:ascii": "Times New Roman", "w:hAnsi": "Times New Roman", "w:cs": "Times New Roman",
  }));
  rPr.appendChild(el(doc, "sz", { "w:val": VOD_SZ }));
  rPr.appendChild(el(doc, "szCs", { "w:val": VOD_SZ }));
  pPr.appendChild(rPr);
  return pPr;
}

/** Paragraph dòng chấm (Times New Roman 14, màu xanh template, in nghiêng như mẫu) */
export function makeVodDots(doc: XDocument): XElement {
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  const spacing = el(doc, "spacing", { "w:line": "480", "w:lineRule": "auto" });
  pPr.appendChild(spacing);
  const rPrP = el(doc, "rPr");
  rPrP.appendChild(el(doc, "rFonts", {
    "w:ascii": "Times New Roman", "w:hAnsi": "Times New Roman", "w:cs": "Times New Roman",
  }));
  rPrP.appendChild(el(doc, "i"));
  rPrP.appendChild(el(doc, "iCs"));
  rPrP.appendChild(el(doc, "color", { "w:val": VOD_COLOR }));
  // 12pt (sz 24) — ĐÚNG như mẫu: 70 ký tự "…" ở 12pt vừa khít 1 dòng khổ 32 cm;
  // để 14pt thì mỗi dòng chấm tự wrap thành 2 dòng (sai mẫu).
  rPrP.appendChild(el(doc, "sz", { "w:val": DOTS_SZ }));
  rPrP.appendChild(el(doc, "szCs", { "w:val": DOTS_SZ }));
  pPr.appendChild(rPrP);
  p.appendChild(pPr);

  const r = el(doc, "r");
  const rPr = rPrP.cloneNode(true) as XElement;
  r.appendChild(rPr);
  const t = el(doc, "t", { "xml:space": "preserve" });
  t.appendChild(doc.createTextNode(DOTS_TEXT));
  r.appendChild(t);
  p.appendChild(r);
  return p;
}

/**
 * Khung câu hỏi = bảng 1 ô (1×1) rộng 14.2 cm, viền đơn 1pt màu C0504D
 * (đo từ mẫu: rect 24.6→423.1pt ≈ 14.06 cm, viền #C0504D dày 1pt).
 * Dùng bảng thay textbox của mẫu: textbox cố định chiều cao, câu hỏi dài sẽ tràn;
 * bảng tự giãn theo nội dung, vẫn giữ đúng diện mạo "khung câu hỏi" của template.
 */
function makeVodQuestionBox(doc: XDocument): XElement {
  const tbl = el(doc, "tbl");
  const tblPr = el(doc, "tblPr");
  tblPr.appendChild(el(doc, "tblW", { "w:w": "8080", "w:type": "dxa" }));
  tblPr.appendChild(el(doc, "jc", { "w:val": "left" }));
  const borders = el(doc, "tblBorders");
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    borders.appendChild(el(doc, side, { "w:val": "single", "w:sz": "8", "w:space": "0", "w:color": "C0504D" }));
  }
  tblPr.appendChild(borders);
  tblPr.appendChild(el(doc, "tblLayout", { "w:type": "fixed" }));
  tblPr.appendChild(el(doc, "tblCellMar")); // mặc định
  tblPr.appendChild(el(doc, "tblLook", { "w:val": "0000" }));
  tbl.appendChild(tblPr);

  const grid = el(doc, "tblGrid");
  grid.appendChild(el(doc, "gridCol", { "w:w": "8080" }));
  tbl.appendChild(grid);

  const tr = el(doc, "tr");
  // không cho khung câu hỏi bị cắt sang trang (nếu nội dung dài hơn 1 trang
  // Word vẫn tự tách — đúng ý, tránh mất nội dung)
  const trPr = el(doc, "trPr");
  trPr.appendChild(el(doc, "cantSplit"));
  tr.appendChild(trPr);
  const tc = el(doc, "tc");
  const tcPr = el(doc, "tcPr");
  tcPr.appendChild(el(doc, "tcW", { "w:w": "8080", "w:type": "dxa" }));
  const mar = el(doc, "tcMar");
  for (const side of ["top", "left", "bottom", "right"]) {
    mar.appendChild(el(doc, side, { "w:w": "113", "w:type": "dxa" }));
  }
  tcPr.appendChild(mar);
  // Word cần tcBorders tường minh — nếu chỉ khai tblBorders, một số bản Word chỉ vẽ
  // đường ngang (bug thực tế: PDF render ra thiếu viền dọc).
  const tcB = el(doc, "tcBorders");
  for (const side of ["top", "left", "bottom", "right"]) {
    tcB.appendChild(el(doc, side, { "w:val": "single", "w:sz": "8", "w:space": "0", "w:color": "C0504D" }));
  }
  tcPr.appendChild(tcB);
  tc.appendChild(tcPr);
  tc.appendChild(el(doc, "p")); // đệm — paragraph câu hỏi được move vào sau
  tr.appendChild(tc);
  tbl.appendChild(tr);
  return tbl;
}

// ============================ chuẩn hóa định dạng ============================

/** run có chứa công thức OLE / drawing (không đổi cỡ chữ) */
function isVisualRun(r: XElement): boolean {
  for (const c of childrenOf(r)) {
    const ln = localName(c);
    if (ln === "object" || ln === "drawing" || ln === "pict" || ln === "fldChar") return true;
    if (ln === "AlternateContent") return true;
  }
  return false;
}

/** Ép rFonts = Times New Roman và sz/szCs = 14pt cho 1 rPr (giữ bold/italic/color…) */
function forceVodFont(rPr: XElement, doc: XDocument): void {
  let fonts = childrenOf(rPr).find((c) => isW(c, "rFonts"));
  if (!fonts) {
    fonts = el(doc, "rFonts");
    rPr.insertBefore(fonts as XNode, rPr.firstChild as XNode | null);
  }
  fonts.setAttribute("w:ascii", "Times New Roman");
  fonts.setAttribute("w:hAnsi", "Times New Roman");
  fonts.setAttribute("w:cs", "Times New Roman");
  // eastAsia giữ nguyên nếu là tiếng Nhật/Trung; tiếng Việt cần TNR
  const ea = fonts.getAttributeNS?.("w", "eastAsia");
  if (!ea || /aptos|calibri|georgia|arial|montserrat/i.test(ea)) {
    fonts.setAttribute("w:eastAsia", "Times New Roman");
  }
  for (const tag of ["sz", "szCs"]) {
    let s = childrenOf(rPr).find((c) => isW(c, tag));
    if (!s) {
      s = el(doc, tag);
      rPr.appendChild(s);
    }
    s.setAttribute("w:val", VOD_SZ);
  }
}

function ensureRPr(r: XElement, doc: XDocument): XElement {
  let rPr = childrenOf(r).find((c) => isW(c, "rPr"));
  if (!rPr) {
    rPr = el(doc, "rPr");
    r.insertBefore(rPr as XNode, r.firstChild as XNode | null);
  }
  return rPr;
}

/**
 * Chuẩn hóa font toàn paragraph: mọi run văn bản → Times New Roman 14.
 * Run chứa công thức OLE giữ nguyên rPr (đổi cỡ sẽ vỡ công thức).
 */
function normalizeParagraphFont(p: XElement, doc: XDocument): void {
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "r") {
        if (!isVisualRun(c)) forceVodFont(ensureRPr(c, doc), doc);
        // vẫn đi xuống: run có thể chứa txbxContent khác
        walk(c);
        continue;
      }
      // w:txbxContent bên trong shape của file GV giữ nguyên hoàn toàn
      if (ln === "txbxContent") continue;
      walk(c);
    }
  };
  walk(p);
  // pPr/rPr (định dạng mặc định của paragraph) cũng về TNR 14
  const pPr = childrenOf(p).find((c) => isW(c, "pPr"));
  if (pPr) {
    const rPr = childrenOf(pPr).find((c) => isW(c, "rPr"));
    if (rPr) forceVodFont(rPr, doc);
  }
}

/**
 * Bullet/numbering của file đề dùng FONT KHAI BÁO TRONG numbering.xml (thường
 * Calibri/Symbol) → ép Times New Roman ở run KHÔNG ăn vào ký tự bullet.
 * Cách xử lý: BỎ numPr và chèn ký tự "•" (hoặc số thứ tự) bằng run Times New Roman
 * ngay đầu paragraph, giữ nguyên nội dung + thụt lề đã có.
 */
function flattenListBullets(p: XElement, doc: XDocument): boolean {
  const pPr = childrenOf(p).find((c) => isW(c, "pPr"));
  if (!pPr) return false;
  const numPr = childrenOf(pPr).find((c) => isW(c, "numPr"));
  if (!numPr) return false;
  pPr.removeChild(numPr as XNode);
  // chèn "• " bằng run TNR 14 đầu paragraph (trước run đầu tiên; pPr luôn đứng đầu)
  const marker = el(doc, "r");
  const mRPr = el(doc, "rPr");
  mRPr.appendChild(el(doc, "rFonts", {
    "w:ascii": "Times New Roman", "w:hAnsi": "Times New Roman", "w:cs": "Times New Roman",
  }));
  mRPr.appendChild(el(doc, "sz", { "w:val": VOD_SZ }));
  mRPr.appendChild(el(doc, "szCs", { "w:val": VOD_SZ }));
  marker.appendChild(mRPr);
  const mt = el(doc, "t", { "xml:space": "preserve" });
  mt.appendChild(doc.createTextNode("•\t"));
  marker.appendChild(mt);
  let ref: XNode | null = null;
  for (const c of childrenOf(p)) {
    if (isW(c, "pPr")) continue;
    ref = c as XNode;
    break;
  }
  if (ref) p.insertBefore(marker as XNode, ref);
  else p.appendChild(marker as XNode);
  // thụt lề trái tối thiểu để bullet không dính lề (nếu paragraph chưa có ind)
  const ind = childrenOf(pPr).find((c) => isW(c, "ind"));
  if (ind) {
    if (!ind.getAttributeNS?.("w", "left")) ind.setAttribute("w:left", "567");
  } else {
    const i2 = el(doc, "ind", { "w:left": "567", "w:hanging": "284" });
    pPr.appendChild(i2);
  }
  return true;
}

/** Xóa highlight (đánh dấu đáp án của GV) — template không dùng highlight */
function removeHighlights(root: XElement): number {
  const found: XElement[] = [];
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      if (localName(c) === "highlight") found.push(c);
      walk(c);
    }
  };
  walk(root);
  for (const h of found) h.parentNode?.removeChild(h);
  return found.length;
}

// ============================ nhóm câu hỏi ============================

/** ID câu hỏi trong nhãn: "(ID: 390525)" */
const ID_RE = /\(\s*ID\s*[:：]\s*([0-9A-Za-z._-]+)\s*\)/i;

export interface VodConvertStats {
  questions: number;
  groups: number;
  ids: string[];
  missingIds: number[];
  highlightRemoved: number;
  dotRows: number;
}

/**
 * Convert body (đã parse) sang bố cục Template VOD — thao tác TRỰC TIẾP trên DOM.
 * Trả về thống kê để báo cho user.
 */
export function applyVodLayout(
  blocks: DocBlock[],
  body: XElement,
  doc: XDocument,
  dotsPerQuestion = DOTS_PER_QUESTION,
): VodConvertStats {
  const idx: Array<{ el: XElement; isTable: boolean }> = [];
  for (const child of childrenOf(body)) {
    if (isW(child, "p")) idx.push({ el: child, isTable: false });
    else if (isW(child, "tbl")) idx.push({ el: child, isTable: true });
  }

  const stats: VodConvertStats = {
    questions: 0, groups: 0, ids: [], missingIds: [], highlightRemoved: 0, dotRows: 0,
  };
  stats.highlightRemoved = removeHighlights(body);

  // ---- 1) Gom câu hỏi (câu + phương án + dòng nhiễu ngay sau) thành từng khung.
  //        Duyệt theo CHỈ SỐ BODY (idx) chứ không theo blocks: blocks đã bị prune
  //        lúc xóa khối đáp án nên i của blocks KHÔNG còn trùng idx → nếu dùng i
  //        sẽ bỏ sót khung và nhét nhầm câu (bug: 50 câu chỉ ra 8 khung).
  type Group = { elems: XElement[]; id: string | null; qnum: number };
  const groups: Group[] = [];
  const n = Math.min(blocks.length, idx.length);
  for (let i = 0; i < n; i++) {
    const b = blocks[i];
    if (b.kind === "table" || (b as Block).kind !== "question") continue;
    if (idx[i].isTable || !idx[i].el.parentNode) continue;
    const txt = (b as Block).text;
    const idMatch = ID_RE.exec(txt);
    const qnumMatch = /^[Cc]âu\s*(\d{1,3})/.exec(txt.trim());
    const elems: XElement[] = [];
    let j = i;
    while (j < n) {
      const bb = blocks[j];
      if (bb.kind === "table" || idx[j].isTable) break; // bảng ngữ liệu: NGOÀI khung
      const blk = bb as Block;
      if (j > i && blk.kind === "question") break;
      if (j > i && (blk.kind === "heading" || blk.kind === "instruction" || blk.kind === "guideline")) break;
      if (idx[j].el.parentNode) elems.push(idx[j].el);
      j++;
    }
    if (elems.length) {
      groups.push({
        elems,
        id: idMatch ? idMatch[1] : null,
        qnum: qnumMatch ? parseInt(qnumMatch[1], 10) : 0,
      });
    }
    i = j - 1;
  }

  // ---- 2) Chèn khung cho từng nhóm
  for (const g of groups) {
    const tbl = makeVodQuestionBox(doc);
    const first = g.elems[0];
    first.parentNode!.insertBefore(tbl, first);
    for (const e of g.elems) e.parentNode?.removeChild(e);
    const boxCell = firstCellOfFirstRow(tbl);
    if (boxCell) {
      // bỏ paragraph đệm
      for (const c of childrenOf(boxCell)) {
        if (isW(c, "p") && paragraphText(c).trim() === "" && !hasVisual(c)) {
          boxCell.removeChild(c as XNode);
          break;
        }
      }
      for (const e of g.elems) boxCell.appendChild(e);
    }
    // dòng chấm: paragraph FULL CHIỀU RỘNG trang, đặt ngay SAU khung câu hỏi
    // (đo mẫu: dòng chấm 34.1→877.2pt ≈ 29.7 cm, không nằm trong cột hẹp)
    let anchor: XNode = tbl as XNode;
    for (let k = 0; k < dotsPerQuestion; k++) {
      const dots = makeVodDots(doc);
      anchor.parentNode!.insertBefore(dots as XNode, anchor.nextSibling);
      anchor = dots as XNode;
    }
    // KHÔNG chèn page break cứng: khung câu hỏi dài/ngắn khác nhau, break cứng
    // sinh trang trắng (đã thử: 95 trang cho 50 câu, 45 trang không có câu nào).
    // Để Word tự chảy — mỗi khung + vùng chấm luôn đi liền nhau (cantSplit).
    void anchor;
    stats.groups++;
    stats.questions++;
    stats.dotRows += dotsPerQuestion;
    if (g.id) stats.ids.push(g.id);
    else stats.missingIds.push(g.qnum);
  }

  // ---- 3) Bullet/numbering → ký tự "•" TNR (glyph bullet không ăn font run),
  //        rồi chuẩn hóa font + cỡ chữ toàn bộ văn bản (cả phần ngoài khung)
  {
    const ps: XElement[] = [];
    const collect = (n: XNode): void => {
      for (const c of childrenOf(n)) {
        if (isW(c, "p")) ps.push(c);
        collect(c);
      }
    };
    collect(body);
    for (const p of ps) flattenListBullets(p, doc);
  }
  normalizeParagraphFont(body, doc); // walk đệ quy — phủ cả paragraph trong bảng

  // ---- 4) Page setup 32 × 18 cm, không header/footer
  const sectPr = childrenOf(body).find((c) => isW(c, "sectPr"));
  if (sectPr) {
    for (const ref of childrenOf(sectPr)) {
      if (isW(ref, "headerReference") || isW(ref, "footerReference")) {
        sectPr.removeChild(ref as XNode);
      }
    }
    const drop = ["pgSz", "pgMar", "cols", "docGrid"];
    for (const c of childrenOf(sectPr)) {
      if (drop.includes(localName(c))) sectPr.removeChild(c as XNode);
    }
    const pgSz = el(doc, "pgSz", { "w:w": PG_W, "w:h": PG_H, "w:orient": "landscape" });
    const pgMar = el(doc, "pgMar", {
      "w:top": PG_MAR, "w:right": PG_MAR, "w:bottom": PG_MAR, "w:left": PG_MAR,
      "w:header": "284", "w:footer": "284", "w:gutter": "0",
    });
    sectPr.appendChild(pgSz);
    sectPr.appendChild(pgMar);
  }

  return stats;
}

function firstCellOfFirstRow(tbl: XElement): XElement | null {
  for (const tr of childrenOf(tbl)) {
    if (!isW(tr, "tr")) continue;
    for (const tc of childrenOf(tr)) {
      if (isW(tc, "tc")) return tc;
    }
  }
  return null;
}

function hasVisual(p: XElement): boolean {
  const walk = (n: XNode): boolean => {
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "object" || ln === "drawing" || ln === "pict" || ln === "oMath") return true;
      if (walk(c)) return true;
    }
    return false;
  };
  return walk(p);
}

// ============================ ghép file template ============================

export interface VodBuildResult {
  documentXml: string;
  relsXml: string;
  /** media của file GV cần copy sang zip template (path→data đọc từ zip GV) */
  srcMediaPaths: string[];
}

/**
 * Ghép body GV (đã convert) vào KHUÔN file Template VOD thật.
 *
 * Vì sao không dùng zip GV làm gốc như bản TSA: template VOD dùng nhiều part
 * riêng (styles/settings/theme/numbering) và KHÔNG có header/footer — trộn rels
 * của GV vào khung template sinh file Word báo corrupt (đã thử: xóa header refs,
 * giữ src rels, đổi content-types… đều hỏng; lý do gốc là styles/numbering
 * id lệch và rels trỏ chéo). Cách an toàn: lấy TOÀN BỘ part của template, chỉ
 * thêm media/OLE của file GV + rels trỏ tới chúng (rIdSrc*).
 */
export function buildVodDocument(
  tpl: LoadedTemplate,
  gvDocumentXml: string,
  gvRelsXml: string,
): VodBuildResult {
  const tplXml = tpl.xml;
  const bodyIdx = tplXml.indexOf("<w:body>");
  const sectIdx = tplXml.lastIndexOf("<w:sectPr");
  if (bodyIdx < 0 || sectIdx < 0) throw new Error("Template VOD không hợp lệ (thiếu body/sectPr)");
  const head = tplXml.slice(0, bodyIdx + "<w:body>".length);
  const tail = tplXml.slice(tplXml.indexOf("</w:sectPr>", sectIdx) + "</w:sectPr>".length);
  // sectPr của TEMPLATE (32×18 cm) — KHÔNG dùng sectPr của file GV.
  // Bỏ header/footerReference: zip template không có part header/footer (template
  // gốc trỏ rId27-31 nhưng ta KHÔNG copy header/footer sang) → Word báo corrupt.
  const sectRaw = tplXml.slice(sectIdx, tplXml.indexOf("</w:sectPr>", sectIdx) + "</w:sectPr>".length);
  const sectPr = sectRaw.replace(/<w:(header|footer)Reference[^>]*\/>/g, "");

  const gvBody = gvDocumentXml.slice(
    gvDocumentXml.indexOf("<w:body>") + "<w:body>".length,
    gvDocumentXml.lastIndexOf("<w:sectPr") >= 0
      ? gvDocumentXml.lastIndexOf("<w:sectPr")
      : gvDocumentXml.indexOf("</w:body>"),
  );

  // --- rels: rel của GV chỉ giữ phần media/OLE được body dùng, đổi sang rIdSrc*
  const usedRids = new Set<string>();
  const ridRe = /r:(?:id|embed|link)="([^"]+)"/g;
  let m: RegExpExecArray | null;
  while ((m = ridRe.exec(gvBody))) usedRids.add(m[1]);

  const srcMediaPaths: string[] = [];
  const extraRels: string[] = [];
  const idMap = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*\/>/g;
  let r: RegExpExecArray | null;
  while ((r = relRe.exec(gvRelsXml))) {
    const tag = r[0];
    const id = /Id="([^"]+)"/.exec(tag)?.[1];
    const type = /Type="([^"]+)"/.exec(tag)?.[1];
    const target = /Target="([^"]+)"/.exec(tag)?.[1];
    if (!id || !type || !target) continue;
    if (!usedRids.has(id)) continue;
    // chỉ nhận media/oleObject — header/footer/hyperlink của GV bỏ (template không có)
    if (!type.includes("/image") && !type.includes("/oleObject")) continue;
    const newId = `rIdSrc${id}`;
    idMap.set(id, newId);
    extraRels.push(
      `<Relationship Id="${newId}" Type="${type}" Target="${target.replace(/^\.\.\//, "")}"/>`,
    );
    srcMediaPaths.push(target.startsWith("media/") ? `word/${target}` : `word/${target}`);
  }

  const body2 = gvBody.replace(/r:(id|embed|link)="([^"]+)"/g, (full, kind: string, id: string) => {
    const nid = idMap.get(id);
    return nid ? `r:${kind}="${nid}"` : full;
  });

  // section break GIỮA body (w:pPr/w:sectPr của file GV) trỏ header/footer của GV
  // → zip template không có part đó ⇒ Word báo corrupt. Bỏ các ref này.
  const body3 = body2.replace(
    /<w:(header|footer)Reference[^>]*r:id="[^"]*"[^>]*\/>/g,
    "",
  );
  void body3;

  // --- namespace còn thiếu so với head template
  const declared = new Set<string>();
  const nsRe = /xmlns:([A-Za-z0-9]+)=/g;
  let nm: RegExpExecArray | null;
  while ((nm = nsRe.exec(head))) declared.add(nm[1]);
  const used = new Set<string>();
  const pfxRe = /<([A-Za-z0-9]+):/g;
  let pm: RegExpExecArray | null;
  while ((pm = pfxRe.exec(body3))) used.add(pm[1]);
  const attrRe = /\s([A-Za-z0-9]+):[A-Za-z]+="/g;
  let am: RegExpExecArray | null;
  while ((am = attrRe.exec(body3))) used.add(am[1]);
  const missing = [...used].filter(
    (p) => !declared.has(p) && p !== "xml" && p !== "xmlns" && NS_MAP[p],
  );
  let head2 = head;
  if (missing.length) {
    const docOpen = /<w:document\b[^>]*>/.exec(head);
    if (docOpen) {
      const tag = docOpen[0];
      const patched = tag.slice(0, -1) + missing.map((p) => ` xmlns:${p}="${NS_MAP[p]}"`).join("") + ">";
      head2 = head.replace(tag, patched);
    }
  }

  // --- rels cuối = rels template (bỏ rel trỏ part KHÔNG có trong zip: fonts) + rel GV
  const tplRelsEntry = tpl.zip.file("word/_rels/document.xml.rels");
  const tplRelsPromise = tplRelsEntry ? tplRelsEntry.async("string") : Promise.resolve("");
  void tplRelsPromise; // đồng bộ hoá: caller đọc rels bằng readVodTemplateRels()

  const documentXml = head2 + body3 + sectPr + tail;
  return { documentXml, relsXml: extraRels.join(""), srcMediaPaths };
}

/** rels của template đã lọc (bỏ target trỏ part không tồn tại + rel rác) */
export async function readVodTemplateRels(tpl: LoadedTemplate): Promise<string> {
  const entry = tpl.zip.file("word/_rels/document.xml.rels");
  const xml = entry ? await entry.async("string") : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const present = new Set<string>();
  tpl.zip.forEach((p) => present.add(p));
  const kept: string[] = [];
  const relRe = /<Relationship\b[^>]*\/>/g;
  let m: RegExpExecArray | null;
  while ((m = relRe.exec(xml))) {
    const tag = m[0];
    const target = /Target="([^"]+)"/.exec(tag)?.[1] ?? "";
    if (target.startsWith("media/") && !present.has(`word/${target}`)) continue;
    if (target.startsWith("fonts/") && !present.has(`word/${target}`)) continue;
    if (["header1.xml", "header2.xml", "footer1.xml", "footer2.xml", "footer3.xml"].includes(target)) {
      // template VOD KHÔNG có header/footer → bỏ hẳn (không logo, không watermark)
      continue;
    }
    kept.push(tag);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${kept.join("")}</Relationships>`;
}
