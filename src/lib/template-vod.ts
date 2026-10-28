/**
 * TEMPLATE VOD (khổ ngang 32 × 18 cm — 16:9) — convert file đề GV → Tài liệu Live.
 *
 * BỐ CỤC ĐẦU RA (theo yêu cầu chốt 10-2026):
 *  1. KHÔNG header/tiêu đề đầu trang: bỏ hết "MÃ ĐỀ…", "Số lượng câu hỏi…",
 *     "Thời gian làm bài…", "Ngày làm bài…" của file đề.
 *  2. KHÔNG in "Câu 1.", "Câu 2." — mỗi khung bắt đầu thẳng bằng
 *     "[HOCMAI] - (ID: xxxxxx)" in đậm, màu đỏ.
 *  3. Khung bo tròn 4 góc, viền mảnh đỏ gạch (#C0504D) — dùng shape VML roundrect
 *     đúng như file mẫu (arcsize 10923f ≈ 10923/100000 = góc bo nhỏ như mẫu;
 *     bảng Word KHÔNG bo góc được nên phải dùng shape). Shape đặt INLINE
 *     (không floating) để nội dung bên dưới tự đẩy xuống, không chồng lên nhau.
 *  4. Đáp án A/B/C/D nằm NGANG trên 1 dòng, dãn đều bằng 4 tab stop chia đều khung.
 *  5. Dưới khung: các dòng kẻ chấm màu xanh dương nhạt, dài LIỀN hết bề rộng lề
 *     trang, khoảng cách đều. Dùng tab stop leader="dot" (Word tự sinh dấu chấm
 *     đều tăm tắp) thay vì chuỗi ký tự "…" — chuỗi ký tự bị wrap sinh dòng lẻ
 *     và đoạn đứt nét khi không vừa đúng số ký tự.
 */

import { localName, childrenOf, isW, createWElement, paragraphText } from "./xml";
import type { XNode, XElement, XDocument } from "./xml";
import type { DocBlock, Block } from "./model";
import type { LoadedTemplate } from "./template-loader";

/** Màu dòng kẻ chấm (xanh dương nhạt — đúng mẫu) */
export const VOD_COLOR = "00B0F0";
/** Màu viền khung + màu tag [HOCMAI]: đỏ gạch/đỏ đô */
export const VOD_BORDER = "C0504D"; // đỏ gạch (mẫu); nhạt hơn: D99694
/** Màu nhãn phương án A./B./C./D. — xanh dương như mẫu */
export const VOD_CHOICE_LABEL = "0070C0";
export const VOD_TAG_COLOR = "C00000";
/** Times New Roman 14pt = sz 28 (half-point) — phần văn bản soạn thảo */
const VOD_SZ = "28";
/** Cỡ chữ dòng kẻ chấm (12pt) — giống mẫu, để chấm thưa và nhẹ hơn văn bản */
const DOTS_SZ = "24";
/** Khổ 32 × 18 cm = 18144 × 10206 twips */
const PG_W = "18144";
const PG_H = "10206";
/** Lề 1 cm = 567 twips */
const PG_MAR = 567;
/** Bề rộng vùng soạn thảo = 18144 - 2×567 = 17010 twips (đúng lề trang) */
const TEXT_W = 18144 - PG_MAR * 2;
/**
 * Bề rộng khung.
 *  - Câu KHÔNG có hình vẽ: khung chiếm ~1/2 bề rộng trang (yêu cầu 10-2026).
 *  - Câu CÓ hình vẽ: khung ~55% bên TRÁI, hình đặt ở cột PHẢI ngoài khung.
 * Dòng kẻ chấm luôn chạy hết bề rộng trang (17010 twips) ở dưới cùng.
 */
const BOX_W_FULL = TEXT_W;
const BOX_W_HALF = Math.round(TEXT_W * 0.5);      // câu không có hình: 1/2 trang
const BOX_W_WITH_IMG = Math.round(TEXT_W * 0.55); // câu có hình: khung trái + hình phải
/** Bề rộng khung mặc định (dùng cho hàm dựng khung khi không truyền) */
const BOX_W = BOX_W_HALF;
/**
 * Góc bo NHẸ (vuông hơn hẳn mẫu cũ 10923f): arcsize 3000f ≈ 3% cạnh ngắn —
 * bo mềm 4 góc nhưng nhìn vẫn là khung chữ nhật (yêu cầu "bo nhẹ, tránh bo tròn quá nhiều").
 */
const ARC_SIZE = "3000f";
/**
 * Số dòng kẻ chấm/câu: KHÔNG cố định — tính động để phủ kín phần còn lại của
 * trang tới mép lề dưới (yêu cầu: "phủ kín toàn bộ phần còn lại của trang").
 * DOTS_LINE là khoảng cách dòng cố định → số dòng = (chiều cao khả dụng - chiều
 * cao khung) / DOTS_LINE_PT.
 */
/**
 * Bước dòng kẻ chấm THẬT (đo từ PDF render): 27.6pt ≈ 552 twips
 * (spacing line=360 + chiều cao chữ 12pt). Dùng số này để tính số dòng vừa trang —
 * nếu tính theo 360 sẽ thừa ~35% số dòng → tràn sang trang sau.
 */
/**
 * Bước dòng kẻ chấm ĐO THỰC TẾ (calibrate bằng Word + PDF, không suy đoán):
 * spacing w:line=360 → mỗi dòng chấm chiếm 19.3pt = 386 twips trong bản in.
 * Với trang 18 cm (khả dụng 16 cm = 9072 twips) chứa được 23 dòng → phủ tới
 * cách mép lề dưới chỉ 8.9pt. Đây là con số dùng để tính số dòng cho mỗi câu.
 */
const DOTS_LINE_TWIPS = 386;
/** Bước dòng khi dùng lineRule=exact (twips) — chốt cứng để khoảng cách đều tuyệt đối */
const DOTS_STEP_EXACT = 386;
const DOTS_LINE = "360";
/** Số dòng chấm tối đa của 1 trang TRỐNG (không có khung) — dùng làm mốc trừ dần */
const DOTS_FULL_PAGE = 23;
/** Chiều cao khả dụng 1 trang (twips): 18 cm - 2×1 cm lề = 16 cm = 9072 twips */
const PAGE_BODY_H = 10206 - PG_MAR * 2;

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

function tnr(doc: XDocument, sz = VOD_SZ): XElement {
  const rPr = el(doc, "rPr");
  rPr.appendChild(el(doc, "rFonts", {
    "w:ascii": "Times New Roman", "w:eastAsia": "Times New Roman",
    "w:hAnsi": "Times New Roman", "w:cs": "Times New Roman",
  }));
  rPr.appendChild(el(doc, "sz", { "w:val": sz }));
  rPr.appendChild(el(doc, "szCs", { "w:val": sz }));
  return rPr;
}

/** Run text đơn giản (TNR, cỡ cho trước, có thể đậm/màu) */
function runText(
  doc: XDocument,
  text: string,
  opts: { sz?: string; bold?: boolean; color?: string } = {},
): XElement {
  const r = el(doc, "r");
  const rPr = tnr(doc, opts.sz ?? VOD_SZ);
  if (opts.bold) {
    rPr.appendChild(el(doc, "b"));
    rPr.appendChild(el(doc, "bCs"));
  }
  if (opts.color) rPr.appendChild(el(doc, "color", { "w:val": opts.color }));
  r.appendChild(rPr);
  const t = el(doc, "t", { "xml:space": "preserve" });
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

/**
 * MỘT dòng kẻ chấm dài liền hết bề rộng lề trang.
 * Tab stop RIGHT tại TEXT_W với leader="dot" → Word tự vẽ chấm đều từ lề trái
 * tới lề phải, không bao giờ wrap thành dòng thứ hai hay đứt đoạn.
 */
export function makeVodDots(doc: XDocument, lineTw = 360): XElement {
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  const tabs = el(doc, "tabs");
  tabs.appendChild(el(doc, "tab", { "w:val": "right", "w:leader": "dot", "w:pos": String(TEXT_W) }));
  pPr.appendChild(tabs);
  // lineRule="exact" để mọi dòng chấm cách nhau CHÍNH XÁC cùng một bước
  // (dùng "auto" thì Word cộng thêm leading font → khoảng cách dao động 22-31pt).
  pPr.appendChild(el(doc, "spacing", { "w:before": "0", "w:after": "0", "w:line": String(lineTw), "w:lineRule": "exact" }));
  const rPrP = tnr(doc, DOTS_SZ);
  rPrP.appendChild(el(doc, "color", { "w:val": VOD_COLOR }));
  pPr.appendChild(rPrP);
  p.appendChild(pPr);

  const r = el(doc, "r");
  const rPr = tnr(doc, DOTS_SZ);
  rPr.appendChild(el(doc, "color", { "w:val": VOD_COLOR }));
  r.appendChild(rPr);
  r.appendChild(el(doc, "tab"));
  p.appendChild(r);
  return p;
}

/**
 * Khung câu hỏi = shape VML roundrect, viền 1pt đỏ gạch/hồng, nền trắng, bo NHẸ.
 * - arcsize 3000f: bo nhẹ (vuông hơn mẫu cũ 10923f) theo yêu cầu.
 * - `mso-fit-shape-to-text:t` + width = full bề rộng lề: khung TỰ CAO theo nội
 *   dung nên đề bài + 4 phương án luôn NẰM GỌN trong viền, không tràn/đè ra ngoài.
 * - Shape INLINE trong 1 paragraph (không floating): mọi thứ chèn sau tự chảy
 *   xuống dưới khung, không đè lên nhau (floating đã thử → text chồng nhau).
 */
function makeVodQuestionBox(doc: XDocument, pageBreakBefore = false, boxW = BOX_W): XElement {
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  // 1 câu/1 trang: dùng pageBreakBefore trên CHÍNH paragraph của khung (không
  // dùng paragraph chứa break riêng — cách đó sinh trang trắng xen kẽ).
  if (pageBreakBefore) pPr.appendChild(el(doc, "pageBreakBefore"));
  pPr.appendChild(el(doc, "keepNext"));
  pPr.appendChild(el(doc, "spacing", { "w:before": "0", "w:after": "0", "w:line": "240", "w:lineRule": "auto" }));
  pPr.appendChild(el(doc, "ind", { "w:left": "0", "w:right": "0", "w:firstLine": "0" }));
  p.appendChild(pPr);

  const r = el(doc, "r");
  const pict = el(doc, "pict");
  const shape = elVml(doc, "roundrect", {
    arcsize: ARC_SIZE,
    style: `width:${(boxW / 1440).toFixed(2)}in`,
    fillcolor: "#FFFFFF",
    strokecolor: `#${VOD_BORDER}`,
    strokeweight: "1pt",
  });
  const textbox = elVml(doc, "textbox", {
    style: "mso-fit-shape-to-text:t",
    inset: "8pt,5pt,8pt,5pt",
  });
  const content = createWElement(doc, "txbxContent");
  content.appendChild(el(doc, "p")); // đệm — nội dung câu hỏi được chuyển vào sau
  textbox.appendChild(content);
  shape.appendChild(textbox);
  pict.appendChild(shape);
  r.appendChild(pict);
  p.appendChild(r);
  return p;
}

/** Tạo element VML (không thuộc namespace w) */
function elVml(doc: XDocument, name: string, attrs: Record<string, string>): XElement {
  const e = doc.createElementNS("urn:schemas-microsoft-com:vml", `v:${name}`);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  return e;
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
 * BỎ thụt lề trái/phải + tab thừa của file đề: nội dung nằm trong khung, thụt lề
 * của file gốc sẽ làm chữ dồn lệch (bug gặp với câu có tab đầu 283 twips).
 */
function normalizeParagraphFont(p: XElement, doc: XDocument): void {
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "r") {
        if (!isVisualRun(c)) forceVodFont(ensureRPr(c, doc), doc);
        walk(c);
        continue;
      }
      if (ln === "txbxContent") continue;
      walk(c);
    }
  };
  walk(p);
  const pPr = childrenOf(p).find((c) => isW(c, "pPr"));
  if (pPr) {
    const rPr = childrenOf(pPr).find((c) => isW(c, "rPr"));
    if (rPr) forceVodFont(rPr, doc);
    // bỏ thụt lề (khung đã lo phần lề); giữ căn đều 2 bên nếu có
    for (const c of childrenOf(pPr)) {
      if (isW(c, "ind")) pPr.removeChild(c as XNode);
    }
    pPr.appendChild(el(doc, "ind", { "w:left": "0", "w:right": "0", "w:firstLine": "0" }));
  }
}

/**
 * Bullet/numbering của file đề dùng FONT KHAI BÁO TRONG numbering.xml (thường
 * Calibri/Symbol) → ép Times New Roman ở run KHÔNG ăn vào ký tự bullet.
 * Cách xử lý: BỎ numPr và chèn ký tự "•" bằng run Times New Roman.
 */
function flattenListBullets(p: XElement, doc: XDocument): boolean {
  const pPr = childrenOf(p).find((c) => isW(c, "pPr"));
  if (!pPr) return false;
  const numPr = childrenOf(pPr).find((c) => isW(c, "numPr"));
  if (!numPr) return false;
  pPr.removeChild(numPr as XNode);
  const marker = runText(doc, "•\t");
  let ref: XNode | null = null;
  for (const c of childrenOf(p)) {
    if (isW(c, "pPr")) continue;
    ref = c as XNode;
    break;
  }
  if (ref) p.insertBefore(marker as XNode, ref);
  else p.appendChild(marker as XNode);
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

// ============================ nhãn câu hỏi ============================

/** ID câu hỏi trong nhãn: "(ID: 390525)" */
const ID_RE = /\(\s*ID\s*[:：]\s*([0-9A-Za-z._-]+)\s*\)/i;

/**
 * Nhãn "Câu N." đứng trước tag — cần bỏ hẳn theo yêu cầu "không thêm Câu 1./Câu 2.".
 * Khớp: "Câu 5." / "Câu 5:" / "Câu 5 [HOCMAI]" / tab đầu / khoảng trắng.
 */
const QLABEL_RE = /^[\s\u00A0\t]*[Cc]âu\s*\d{1,3}\s*(?:[.:]\s*)?/;

/** Tag "[HOCMAI] - (ID: x)" — tìm trong text để tách phần đầu ra khỏi run */
const TAG_ID_RE = /\[\s*HOCMAI\s*\]\s*[-\u2013\u2014:.]?\s*\(\s*ID\s*[:：]\s*[0-9A-Za-z._-]+\s*\)/i;

/**
 * Bỏ nhãn "Câu N." ở đầu paragraph câu hỏi — CẮT THEO OFFSET KÝ TỰ LOGIC
 * (đếm qua w:t + w:tab như paragraphText), không theo độ dài run.
 *
 * PITFALL đã dính: cắt theo run rồi trừ dồn độ dài làm lệch khi run chứa
 * w:tab (text logic dài hơn text thật) → xóa lẹm cả "[HOC" của tag, ra "MAI]".
 * Cách đúng: tính offset cần bỏ trên chuỗi logic, rồi rewrite từng w:t theo offset.
 */
function stripQuestionLabel(p: XElement, doc: XDocument): boolean {
  const raw = paragraphText(p);
  const m = QLABEL_RE.exec(raw);
  if (!m || !m[0] || m[0].length === 0) return false;
  return cutLogicalFront(p, m[0].length, doc);
}

/**
 * Bỏ `count` ký tự đầu của paragraph (tính trên text logic: text + tab + br).
 * Trả về true nếu bỏ được đủ.
 */
function cutLogicalFront(p: XElement, count: number, doc: XDocument): boolean {
  let left = count;
  // Duyệt tuần tự mọi node con theo đúng thứ tự paragraphText: w:t (text), w:tab, w:br
  const nodes: Array<{ kind: "t" | "tab" | "br"; el: XElement }> = [];
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "t") nodes.push({ kind: "t", el: c });
      else if (ln === "tab") nodes.push({ kind: "tab", el: c });
      else if (ln === "br" || ln === "cr") nodes.push({ kind: "br", el: c });
      else walk(c);
    }
  };
  walk(p);

  for (const nd of nodes) {
    if (left <= 0) break;
    if (nd.kind !== "t") {
      // tab/br chiếm 1 ký tự logic
      nd.el.parentNode?.removeChild(nd.el as XNode);
      left -= 1;
      continue;
    }
    const cur = textInT(nd.el);
    if (cur.length <= left) {
      left -= cur.length;
      nd.el.parentNode?.removeChild(nd.el as XNode);
      continue;
    }
    const keep = cur.slice(left);
    left = 0;
    while (nd.el.firstChild) nd.el.removeChild(nd.el.firstChild);
    nd.el.appendChild(doc.createTextNode(keep));
  }
  return left === 0;
}

/** Text logic của 1 run (text + tab + br) — khớp cách đếm của paragraphText */
function runLogicalText(r: XElement): string {
  let out = "";
  const walk = (n: XNode): void => {
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

/** Text trực tiếp trong 1 w:t */
function textInT(t: XElement): string {
  let out = "";
  for (let i = 0; i < t.childNodes.length; i++) {
    const c = t.childNodes[i];
    if (c.nodeType === 3) out += c.nodeValue ?? "";
  }
  return out;
}

/**
 * Chuẩn hóa tag "[HOCMAI] - (ID: x)": in ĐẬM + ĐỎ (C00000) đúng yêu cầu.
 * Áp rPr cho mọi run nằm trong phạm vi tag (từ ký tự đầu tới hết "(ID: x)").
 */
function styleTagRuns(p: XElement, doc: XDocument): boolean {
  const raw = paragraphText(p);
  const m = TAG_ID_RE.exec(raw);
  if (!m) return false;
  const start = m.index;
  const end = m.index + m[0].length;
  let touched = false;

  // Chỉ tô đúng phần ký tự nằm trong [start, end): run nào VƯỢT biên thì TÁCH
  // thành 3 run (trước / trong / sau) để không đậm-đỏ lấn sang nội dung câu hỏi.
  let pos = 0;
  const runs: XElement[] = [];
  const collect = (node: XNode): void => {
    for (const c of childrenOf(node)) {
      if (isW(c, "r")) runs.push(c);
      else if (localName(c) !== "txbxContent") collect(c);
    }
  };
  collect(p);

  for (const r of runs) {
    const len = runLogicalText(r).length;
    const rStart = pos;
    const rEnd = pos + len;
    pos = rEnd;
    if (len === 0) continue;
    if (rEnd <= start || rStart >= end) continue;
    if (isVisualRun(r)) continue; // công thức: không đụng

    const fromIn = Math.max(0, start - rStart);
    const toIn = Math.min(len, end - rStart);
    if (fromIn === 0 && toIn === len) {
      applyTagFormat(r, doc);
      touched = true;
      continue;
    }
    // tách: head (trước tag) / mid (trong tag) / tail (sau tag)
    const head = fromIn > 0 ? cloneRunTrimmed(r, 0, fromIn, doc) : null;
    const mid = cloneRunTrimmed(r, fromIn, toIn, doc);
    const tail = toIn < len ? cloneRunTrimmed(r, toIn, len, doc) : null;
    const parent = r.parentNode;
    if (!parent || !mid) continue;
    if (head) parent.insertBefore(head as XNode, r as XNode);
    parent.insertBefore(mid as XNode, r as XNode);
    if (tail) parent.insertBefore(tail as XNode, r as XNode);
    else parent.insertBefore(runText(doc, " "), r as XNode);
    parent.removeChild(r as XNode);
    applyTagFormat(mid, doc);
    touched = true;
  }
  return touched;
}

/** Ép rPr: Times New Roman 14, đậm, màu đỏ C00000 (giữ nguyên phần khác) */
function applyTagFormat(r: XElement, doc: XDocument): void {
  const rPr = ensureRPr(r, doc);
  forceVodFont(rPr, doc);
  if (!childrenOf(rPr).some((c) => isW(c, "b"))) rPr.appendChild(el(doc, "b"));
  if (!childrenOf(rPr).some((c) => isW(c, "bCs"))) rPr.appendChild(el(doc, "bCs"));
  let color = childrenOf(rPr).find((c) => isW(c, "color"));
  if (!color) {
    color = el(doc, "color");
    rPr.appendChild(color);
  }
  color.setAttribute("w:val", VOD_TAG_COLOR);
  for (const c of childrenOf(rPr)) {
    if (localName(c) === "highlight") rPr.removeChild(c as XNode);
  }
}

// ============================ đáp án trên 1 dòng ============================

/** Nhãn phương án: "A." "B." "C." "D." (đầu chuỗi hoặc sau khoảng trắng/tab) */
const CHOICE_RE = /(?:^|[\s\u00A0\t.．,;:)\]}])([A-D])\s*[.)]\s*/gu;

/**
 * Dồn 4 phương án A/B/C/D về MỘT dòng, dãn đều bằng tab stop trong bề rộng khung.
 *
 * CÁCH LÀM AN TOÀN (bài học 3 lần sửa):
 *  - KHÔNG tính offset ký tự rồi cắt theo offset: offset tính từ paragraphText
 *    (duyệt ĐỆ QUY mọi run, kể cả run lồng trong w:ins/w:hyperlink/w:smartTag)
 *    sẽ lệch với vòng lặp chỉ duyệt run TRỰC TIẾP → cắt lẹm chữ, mất công thức.
 *  - Thay vào đó duyệt ĐỆ QUY đúng thứ tự tài liệu, xét TỪNG RUN theo nội dung
 *    text của chính run đó: run nào mở đầu bằng "A." / "B." / "C." / "D." thì mở
 *    bucket mới và cắt đúng phần nhãn trong run ấy; run không có nhãn thì nối
 *    vào bucket hiện hành. Run rỗng (OLE Equation = 0 ký tự text) luôn thuộc
 *    bucket hiện hành → không bao giờ mất công thức của phương án.
 *  - Run chỉ có khoảng trắng/tab (phân cách của file đề) bị bỏ, thay bằng tab stop.
 */
function mergeChoicesIntoOneLine(
  choiceParas: XElement[],
  doc: XDocument,
  boxW = BOX_W,
): { merged: XElement | null; labels: number } {
  if (!choiceParas.length) return { merged: null, labels: 0 };

  const buckets = new Map<string, XElement[]>();
  let labels = 0;
  const cur: { label: string | null } = { label: null };

  const push = (label: string, r: XElement) => {
    if (!buckets.has(label)) { buckets.set(label, []); labels++; }
    buckets.get(label)!.push(r);
  };

  /** Duyệt đệ quy mọi w:r theo đúng thứ tự tài liệu */
  const walkRuns = (node: XNode, visit: (r: XElement) => void): void => {
    for (const c of childrenOf(node)) {
      if (isW(c, "r")) { visit(c); continue; }
      // không đi vào txbxContent (khung khác) / object (công thức nằm trong run)
      if (localName(c) === "txbxContent") continue;
      walkRuns(c, visit);
    }
  };

  for (const p of choiceParas) {
    // Gom run của paragraph theo thứ tự; đồng thời gỡ chúng khỏi paragraph gốc
    const runs: XElement[] = [];
    walkRuns(p as XNode, (r) => runs.push(r));

    for (const r of runs) {
      const text = runLogicalText(r);
      const visual = isVisualRun(r);

      if (!text) {
        // OLE/drawing (0 ký tự text) hoặc run rỗng → thuộc nhãn hiện hành
        if (visual && cur.label) push(cur.label, r);
        continue;
      }

      // Run có thể mở đầu bằng nhãn (kèm tab/khoảng trắng phía trước)
      CHOICE_RE.lastIndex = 0;
      const m = CHOICE_RE.exec(text);
      const leadOnly = text.trim().length === 0;

      if (leadOnly) continue; // run chỉ có tab/khoảng trắng → bỏ

      if (m && m.index === 0) {
        // Cả run bắt đầu bằng nhãn: xác định nhãn + phần nội dung sau nhãn
        const label = m[1];
        cur.label = label;
        const cut = m[0].length;
        if (cut < text.length) {
          const tail = cloneRunTrimmed(r, cut, text.length, doc);
          if (tail) push(label, tail);
        }
        continue;
      }

      // Run không chứa nhãn ở đầu → nối vào bucket hiện hành
      if (cur.label) push(cur.label, r);
    }
  }

  const present = ["A", "B", "C", "D"].filter((l) => buckets.has(l));
  if (present.length < 2) return { merged: null, labels };

  // Dựng paragraph 1 dòng, tab stop chia đều bề rộng khung
  // 4 cột đều nhau: tab stop cho B/C/D tại step, 2·step, 3·step (A nằm ở lề trái,
  // KHÔNG cần tab). Bỏ tab stop tại 0 vì Word bỏ qua tab trỏ về đúng vị trí hiện tại
  // → cả 4 đáp án bị dồn thành nhiều dòng.
  const step = Math.floor(boxW / present.length);
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  const tabs = el(doc, "tabs");
  for (let i = 1; i < present.length; i++) {
    tabs.appendChild(el(doc, "tab", { "w:val": "left", "w:pos": String(i * step) }));
  }
  pPr.appendChild(tabs);
  pPr.appendChild(el(doc, "spacing", { "w:before": "60", "w:after": "40", "w:line": "360", "w:lineRule": "auto" }));
  pPr.appendChild(el(doc, "ind", { "w:left": "0", "w:right": "0", "w:firstLine": "0" }));
  pPr.appendChild(tnr(doc));
  p.appendChild(pPr);

  present.forEach((lab, i) => {
    // Tab CHỈ đặt TRƯỚC đáp án thứ 2 trở đi. Đặt tab trước A sẽ đẩy A sang tab
    // stop đầu tiên — nhưng nếu tab stop đầu = 0 thì Word BỎ QUA tab, cả dòng dồn
    // thành nhiều dòng (bug: A/B/C/D rơi xuống 3 dòng khác nhau).
    if (i > 0) {
      const rTab = el(doc, "r");
      rTab.appendChild(tnr(doc));
      rTab.appendChild(el(doc, "tab"));
      p.appendChild(rTab);
    }
    p.appendChild(runText(doc, `${lab}. `, { bold: true, color: VOD_CHOICE_LABEL }));
    for (const r of buckets.get(lab)!) p.appendChild(r as XNode);
  });
  return { merged: p, labels };
}

/**
 * Clone run chỉ giữ text logic trong [from, to). Chỉ gọi cho run TEXT (không có
 * công thức/ảnh) — clone sâu rồi cắt đầu/đuôi trên các w:t.
 */
function cloneRunTrimmed(r: XElement, from: number, to: number, doc: XDocument): XElement | null {
  if (to <= from) return null;
  const clone = r.cloneNode(true) as XElement;
  let pos = 0;
  const ts: XElement[] = [];
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      if (localName(c) === "t") ts.push(c);
      walk(c);
    }
  };
  walk(clone);
  for (const t of ts) {
    const cur = textInT(t);
    const s0 = pos;
    const e0 = pos + cur.length;
    pos = e0;
    if (e0 <= from || s0 >= to) { t.parentNode?.removeChild(t as XNode); continue; }
    const keepFrom = Math.max(0, from - s0);
    const keepTo = Math.min(cur.length, to - s0);
    const kept = cur.slice(keepFrom, keepTo);
    while (t.firstChild) t.removeChild(t.firstChild);
    if (kept) t.appendChild(doc.createTextNode(kept));
    else t.parentNode?.removeChild(t as XNode);
  }
  return runLogicalText(clone).trim().length || isVisualRun(clone) ? clone : null;
}

// ============================ nhóm câu hỏi ============================

export interface VodConvertStats {
  questions: number;
  groups: number;
  ids: string[];
  missingIds: number[];
  highlightRemoved: number;
  dotRows: number;
  /** số paragraph tiêu đề đầu trang đã bỏ (MÃ ĐỀ / Số lượng câu hỏi / …) */
  headerRemoved: number;
  /** số câu đã dồn 4 đáp án về 1 dòng */
  choicesMerged: number;
  /** số câu đã đưa hình vẽ ra ngoài khung (cột phải) */
  picturesMovedToRight: number;
}

/**
 * Nhận diện paragraph TIÊU ĐỀ ĐẦU TRANG của file đề để xóa hoàn toàn:
 * "MÃ ĐỀ: …", "Số lượng câu hỏi: …", "Thời gian làm bài: …", "Ngày làm bài: …",
 * "Họ và tên thí sinh", "Số báo danh", "ĐỀ THI GỒM …", dòng "Môn: …".
 */
const HEADER_LINE_RE =
  /^\s*(M[ÃA]\s*(Đ[ỀE]|DE)|S[ỐO]\s*L[ƯU][ỢO]NG\s*C[ÂA]U|S[ỐO]\s*C[ÂA]U|TH[ỜO]I\s*GIAN\s*L[ÀA]M\s*B[ÀA]I|NG[ÀA]Y\s*L[ÀA]M\s*B[ÀA]I|H[ỌO]\s*(V[ÀA]\s*T[ÊE]N|T[ÊE]N)|S[ỐO]\s*B[ÁA]O\s*DANH|M[ÔO]N\s*THI|B[ÀA]I\s*THI|K[ỲY]\s*THI|TH[ỜO]I\s*L[ƯU][ỢO]NG|ĐI[ỂE]M|GHI\s*CH[ÚU])\b/iu;

function isHeaderLine(txt: string): boolean {
  const t = txt.trim();
  if (!t) return false;
  if (HEADER_LINE_RE.test(t)) return true;
  // So khớp không phụ thuộc kiểu chữ (file đề có khi viết HOA toàn bộ)
  const low = t.toLowerCase();
  const prefixes = [
    "mã đề", "ma de", "số lượng câu", "so luong cau", "số câu", "so cau",
    "thời gian làm bài", "thoi gian lam bai", "ngày làm bài", "ngay lam bai",
    "họ và tên", "ho va ten", "họ tên", "ho ten", "số báo danh", "so bao danh",
    "môn thi", "mon thi", "bài thi", "bai thi", "kỳ thi", "ky thi",
    "thời lượng", "thoi luong", "ghi chú", "ghi chu",
  ];
  if (prefixes.some((p) => low.startsWith(p))) return true;
  // "50 câu" / "HSA – BF" đứng riêng cũng là dòng tiêu đề của file đề
  if (/^\d+\s*c[âa]u[\s.]*$/iu.test(t)) return true;
  if (/^[A-Z]{2,5}\s*[–\-—]\s*[A-Z]{1,4}\s*$/u.test(t)) return true;
  // dòng chỉ có "MÃ ĐỀ: XXX" hoặc "ĐỀ SỐ 01"
  if (/^(đề|de)\s*(số|so)?\s*\d+/iu.test(t)) return true;
  return false;
}

/**
 * Convert body (đã parse) sang bố cục Template VOD — thao tác TRỰC TIẾP trên DOM.
 * Trả về thống kê để báo cho user.
 */
export function applyVodLayout(
  blocks: DocBlock[],
  body: XElement,
  doc: XDocument,
  dotsPerQuestion = 0, // 0 = TỰ ĐỘNG lấp kín trang (1 câu/trang)
): VodConvertStats {
  const idx: Array<{ el: XElement; isTable: boolean }> = [];
  for (const child of childrenOf(body)) {
    if (isW(child, "p")) idx.push({ el: child, isTable: false });
    else if (isW(child, "tbl")) idx.push({ el: child, isTable: true });
  }

  const stats: VodConvertStats = {
    questions: 0, groups: 0, ids: [], missingIds: [], highlightRemoved: 0, dotRows: 0,
    headerRemoved: 0, choicesMerged: 0, picturesMovedToRight: 0,
  };
  stats.highlightRemoved = removeHighlights(body);

  // ---- 0) XÓA TIÊU ĐỀ ĐẦU TRANG (MÃ ĐỀ / Số lượng câu hỏi / Thời gian / Ngày làm bài)
  //      Chỉ xét các paragraph TRƯỚC câu hỏi đầu tiên — không đụng nội dung đề.
  {
    const limit = Math.min(blocks.length, idx.length);
    let firstQ = limit;
    for (let i = 0; i < limit; i++) {
      const b = blocks[i];
      if (b.kind !== "table" && (b as Block).kind === "question") { firstQ = i; break; }
    }
    for (let i = 0; i < firstQ; i++) {
      if (idx[i].isTable || !idx[i].el.parentNode) continue;
      const b = blocks[i];
      if (b.kind !== "table" && (b as Block).kind === "question") break;
      const txt = b.kind === "table" ? "" : (b as Block).text;
      if (isHeaderLine(txt) || !txt.trim()) {
        idx[i].el.parentNode?.removeChild(idx[i].el as XNode);
        stats.headerRemoved++;
      }
    }
  }

  // ---- 1) Gom câu hỏi (câu + phương án + dòng nhiễu ngay sau) thành từng khung.
  //        Duyệt theo CHỈ SỐ BODY (idx) — blocks đã bị prune lúc xóa khối đáp án
  //        nên i của blocks KHÔNG còn trùng idx (bug cũ: 50 câu chỉ ra 8 khung).
  type Group = {
    qPara: XElement;
    elems: XElement[];      // toàn bộ paragraph của câu (câu hỏi + phương án)
    choiceParas: XElement[]; // riêng các paragraph được phân loại "choices"
    id: string | null;
    qnum: number;
  };
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
    const choiceParas: XElement[] = [];
    let j = i;
    while (j < n) {
      const bb = blocks[j];
      // BẢNG NGỮ LIỆU (bảng tần số/số liệu của chính câu này) → ĐƯA VÀO KHUNG.
      // Trước đây `break` tại bảng khiến phần "Tứ phân vị…" + 4 đáp án rơi ra
      // ngoài khung (bug: trang không có khung, chỉ có chữ trôi nổi).
      if (bb.kind === "table" || idx[j].isTable) {
        if (j > i && idx[j].el.parentNode) elems.push(idx[j].el);
        j++;
        continue;
      }
      const blk = bb as Block;
      if (j > i && blk.kind === "question") break;
      // "instruction" = ngữ liệu chung ("Dựa vào thông tin dưới đây…", "Cho đa giác
      // lồi 16 cạnh.") → ĐƯA VÀO KHUNG của câu này (yêu cầu: toàn bộ đề bài nằm
      // gọn trong 1 khung). Chỉ heading/guideline mới ngắt nhóm.
      if (j > i && (blk.kind === "heading" || blk.kind === "guideline")) break;
      if (idx[j].el.parentNode) {
        elems.push(idx[j].el);
        if (blk.kind === "choices") choiceParas.push(idx[j].el);
      }
      j++;
    }
    if (elems.length) {
      groups.push({
        qPara: idx[i].el,
        elems,
        choiceParas,
        id: idMatch ? idMatch[1] : null,
        qnum: qnumMatch ? parseInt(qnumMatch[1], 10) : 0,
      });
    }
    i = j - 1;
  }

  // ---- 2) Bỏ nhãn "Câu N." + ép tag [HOCMAI] - (ID: x) đậm đỏ (làm trước khi dựng khung)
  for (const g of groups) {
    stripQuestionLabel(g.qPara, doc);
    styleTagRuns(g.qPara, doc);
  }

  // ---- 3) Dựng khung bo góc + nội dung + dòng chấm cho từng câu
  let firstGroup = true;
  for (const g of groups) {
    // 3-before) PHÁT HIỆN HÌNH VẼ của câu này.
    //  - Có hình ⇒ khung hẹp (~55% trái) + hình đặt cột PHẢI ngoài khung.
    //  - Không có hình ⇒ khung 1/2 bề rộng trang.
    const picParas: XElement[] = [];
    for (const e of g.elems) {
      if (!isW(e, "p")) continue;
      if (isBigPicture(e)) {
        const txt = paragraphText(e).trim();
        if (txt.length > 0) {
          // hình nằm chung paragraph với chữ câu hỏi → tách hình ra paragraph riêng
          const picP = extractPicturePara(e, doc);
          if (picP) picParas.push(picP);
        } else {
          picParas.push(e); // paragraph chỉ có hình → đưa nguyên ra ngoài
        }
      }
    }
    const hasPicture = picParas.length > 0;
    const boxW = hasPicture ? BOX_W_WITH_IMG : BOX_W_HALF;

    // Câu đầu giữ nguyên trang 1; các câu sau bắt đầu trang mới (1 câu/trang).
    const boxP = makeVodQuestionBox(doc, !firstGroup, boxW);
    firstGroup = false;
    const first = g.elems[0];
    first.parentNode!.insertBefore(boxP as XNode, first as XNode);

    // 3a) Dồn 4 phương án về 1 dòng (nếu nhận diện được >= 2 phương án)
    let mergedChoices: XElement | null = null;
    let choiceLabels = 0;
    if (g.choiceParas.length) {
      const res = mergeChoicesIntoOneLine(g.choiceParas, doc, boxW);
      mergedChoices = res.merged;
      choiceLabels = res.labels;
    }

    // 3b) Chuyển các paragraph còn lại vào txbxContent của khung
    const content = findTxbxContent(boxP);
    const picSet = new Set<XElement>(picParas);
    const skip = new Set<XElement>(g.choiceParas);
    if (content) {
      for (const c of childrenOf(content)) {
        if (isW(c, "p") && paragraphText(c).trim() === "" && !hasVisual(c)) {
          content.removeChild(c as XNode);
          break;
        }
      }
      for (const e of g.elems) {
        if (skip.has(e) || picSet.has(e)) continue;
        e.parentNode?.removeChild(e as XNode);
        content.appendChild(e);
      }
      if (mergedChoices) {
        content.appendChild(mergedChoices);
        if (choiceLabels >= 2) stats.choicesMerged++;
      }
    }
    // dọn paragraph phương án cũ còn sót trong body (đã thay bằng bản gộp)
    for (const cp of g.choiceParas) cp.parentNode?.removeChild(cp as XNode);
    // gỡ paragraph chỉ-có-hình khỏi body (sẽ đặt lại vào cột phải nếu có hình)
    for (const pp of picParas) pp.parentNode?.removeChild(pp as XNode);

    // 3b2) CÓ HÌNH: thay paragraph chứa khung bằng bảng 2 cột KHÔNG viền
    //      (trái = khung câu hỏi + đáp án, phải = hình vẽ) → hình nằm ngoài khung.
    let boxAnchor: XElement = boxP;
    if (hasPicture) {
      // Chèn bảng 2 cột vào ĐÚNG vị trí khung cũ TRƯỚC khi nhấc khung ra khỏi body,
      // nếu không bảng không có parentNode → insertBefore(null) crash.
      const two = makeTwoColumnLayout(doc, boxP, picParas, boxW);
      boxAnchor = two;
      stats.picturesMovedToRight++;
    }

    // 3c) Dòng kẻ chấm ngay SAU khung — số dòng TÍNH ĐỘNG để phủ kín phần còn
    //     lại của trang tới mép lề dưới (mỗi câu 1 trang).
    //     Chiều cao khung = tổng chiều cao ước lượng từng paragraph (đọc cả khối
    //     hình/đồ thị nhúng) + viền/inset.
    // 3c) Vùng kẻ chấm: ĐẶT ĐỦ SỐ DÒNG để phủ tới mép lề dưới rồi để Word CẮT phần
    //     dư ở cuối trang — nhờ vậy không bao giờ tràn sang trang sau (bug cũ: tràn
    //     sinh trang trắng) mà vẫn phủ kín 100% phần còn lại của trang.
    //     Số dòng đặt = phần còn lại ước lượng + 3 dòng dư (dư bị cắt, không hiện).
    // Số dòng chấm = số dòng còn chỗ sau khung, tính từ MỐC CALIBRATE thật
    // (trang trống chứa 23 dòng chấm) để phủ kín tới mép lề dưới mà không tràn.
    // Cộng chiều cao BẢNG NGỮ LIỆU đứng ngay trước khung (bảng nằm ngoài khung
    // theo mẫu) — nếu bỏ qua, số dòng chấm tính thừa → tràn trang, sinh trang trắng
    // (bug gặp ở câu có bảng tần số).
    const leadingTableH = tableHeightBefore(body, boxP);
    // Ước lượng chiều cao khung. Hệ số 1.35 là HỆ SỐ HIỆU CHỈNH đo được: ước lượng
    // theo ký tự thường THẤP hơn chiều cao thật (do công thức OLE, nhiều dòng đáp án),
    // nên nhân lên để không cấp thừa dòng chấm → tránh tràn trang.
    // Ước lượng chiều cao khung — HIỆU CHỈNH theo số đo thật trên PDF:
    // khung câu 132 ký tự + 4 đáp án cao 106pt = 2120 twips; ước lượng ký tự cho
    // 1200 twips → cần hệ số nhân để khớp. Hệ số 0.6 (trước đây ước CAO gấp ~1.7
    // lần) để số dòng chấm phủ kín phần còn lại của trang.
    // Khung hẹp (1/2 trang) làm chữ wrap gấp ~2 lần số dòng so với khung full-width,
    // nên hệ số chiều cao phải nhân theo tỉ lệ bề rộng khung.
    const widthFactor = TEXT_W / boxW; // 2.0 khi khung 1/2 trang, 1.8 khi có hình
    const estH = g.elems.reduce(
      (acc, e) => acc + estimateParaHeightTwips(e) * widthFactor, 0,
    );
    const boxH = estH * 0.6 + 900 + leadingTableH + 900; // +900 = 1 dòng an toàn
    const availH = PAGE_BODY_H - boxH;
    const nDots = dotsPerQuestion > 0
      ? dotsPerQuestion
      : Math.max(1, Math.min(DOTS_FULL_PAGE, Math.round(availH / DOTS_LINE_TWIPS)));
    let anchor: XNode = boxAnchor as XNode;
    for (let k = 0; k < nDots; k++) {
      const dots = makeVodDots(doc, DOTS_LINE_TWIPS);
      anchor.parentNode!.insertBefore(dots as XNode, anchor.nextSibling);
      anchor = dots as XNode;
    }
    // Dòng chấm nằm trong ô bảng cao cố định = chiều cao khả dụng → Word cắt phần
    // vượt quá mép dưới thay vì đẩy sang trang mới.

    stats.groups++;
    stats.questions++;
    stats.dotRows += nDots;
    if (g.id) stats.ids.push(g.id);
    else stats.missingIds.push(g.qnum);
  }

  // ---- 4) Bullet/numbering → ký tự "•" TNR, rồi chuẩn hóa font toàn tài liệu
  {
    const ps: XElement[] = [];
    const collect = (node: XNode): void => {
      for (const c of childrenOf(node)) {
        if (isW(c, "p")) ps.push(c);
        collect(c);
      }
    };
    collect(body);
    for (const p of ps) flattenListBullets(p, doc);
  }
  normalizeParagraphFont(body, doc);

  // ---- 5) Page setup 32 × 18 cm, không header/footer
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
    sectPr.appendChild(el(doc, "pgSz", { "w:w": PG_W, "w:h": PG_H, "w:orient": "landscape" }));
    sectPr.appendChild(el(doc, "pgMar", {
      "w:top": String(PG_MAR), "w:right": String(PG_MAR),
      "w:bottom": String(PG_MAR), "w:left": String(PG_MAR),
      "w:header": "284", "w:footer": "284", "w:gutter": "0",
    }));
  }

  return stats;
}

/**
 * Ước lượng chiều cao (twips) của 1 paragraph trong khung — để tính số dòng chấm
 * vừa đúng 1 trang.
 *
 * PITFALL: câu có HÌNH VẼ/ĐỒ THỊ (w:drawing / w:object OLE lớn) không thể ước
 * lượng theo số ký tự — nếu bỏ qua, chiều cao khung bị tính thiếu → số dòng chấm
 * tính thừa → tràn sang trang sau sinh trang trắng (bug gặp ở câu có hình).
 * Đọc kích thước thật từ wp:extent cy (EMU; 1 twip = 635 EMU) và w:shape style.
 */
function estimateParaHeightTwips(p: XElement): number {
  // BẢNG (ngữ liệu) trong khung: mỗi hàng ≈ 380 twips
  if (isW(p, "tbl")) {
    const rows = childrenOf(p).filter((c) => isW(c, "tr")).length || 3;
    return rows * 380 + 200;
  }
  const chars = paragraphText(p).length;
  // Hệ số hiệu chỉnh theo số liệu ĐO trên PDF render thật (không phải lý thuyết):
  // khung ~15 dòng chữ cao 94.5pt = 1890 twips → ~126 twips/dòng; lấy 150 để
  // chừa biên an toàn cho công thức OLE (cao hơn chữ thường).
  // Hệ số hiệu chỉnh TỪ SỐ LIỆU ĐO THẬT trên PDF (10 trang mẫu, câu 114–419 ký tự):
  //   boxH_pt ≈ ceil(chars/105) * 45 + 60  → quy ra twips: dòng = 900, biên = 1200.
  // (Trước đây ước 340 twips/dòng → khung bị tính cao gấp ~2 lần thực tế, nên số
  //  dòng kẻ chấm bị cắt bớt và chừa trống 1/3 trang.)
  // Công thức hiệu chỉnh từ PDF thật: mỗi dòng chữ (chứa cả công thức OLE) cao
  // ~18pt = 360 twips, ~65 ký tự/dòng ở TNR 14 khổ 30 cm; +1 dòng cho hàng đáp án;
  // +14pt (280 twips) cho inset/viền của shape.
  const textH = Math.max(1, Math.ceil(chars / 105)) * 900;

  // chiều cao ảnh/đồ thị nhúng (drawing: wp:extent cy)
  let visualH = 0;
  let maxVisualH = 0;
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "extent") {
        const cy = attrAny(c as XElement, "cy");
        const v = cy ? parseInt(cy, 10) : 0;
        // 1 twip = 635 EMU; +8% cho khoảng dòng bao quanh ảnh
        if (v > 0) maxVisualH = Math.max(maxVisualH, Math.round((v / 635) * 1.08));
      } else if (ln === "shape") {
        const style = attrAny(c as XElement, "style") || "";
        const m = /height\s*:\s*([\d.]+)pt/i.exec(style);
        if (m) maxVisualH = Math.max(maxVisualH, Math.round(parseFloat(m[1]) * 20 * 1.08));
      }
      walk(c);
    }
  };
  walk(p);
  // Ảnh/đồ thị lớn nằm CHUNG paragraph với câu hỏi → chiều cao paragraph = ảnh,
  // KHÔNG cộng dồn (cộng dồn làm ước lượng vống → cắt bớt dòng chấm oan).
  visualH = maxVisualH;

  // Mỗi công thức OLE (Equation) hiển thị như 1 "khối" giữa dòng: cộng thêm
  // 260 twips/công thức (đo từ PDF: câu 4 có đồ thị lớn; câu nhiều phân số làm
  // khung cao hơn hẳn ước lượng theo ký tự).
  let nObj = 0;
  const countObj = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      if (localName(c) === "object" || localName(c) === "pict") nObj++;
      countObj(c);
    }
  };
  countObj(p);
  const objH = Math.min(2400, nObj * 150);

  return Math.max(textH + objH, visualH, 340) + (visualH > 0 ? 300 : 0);
}

/** Đọc attribute không phân biệt namespace (xmldom/native) */
function attrAny(e: XElement, name: string): string | null {
  const withNs = e.getAttributeNS?.(null, name);
  if (withNs) return withNs;
  const any = (e as unknown as { getAttribute?: (n: string) => string | null }).getAttribute;
  return any ? any.call(e, name) : null;
}

/**
 * Ước lượng chiều cao các BẢNG nằm ngay TRƯỚC khung câu hỏi (bảng ngữ liệu:
 * bảng tần số, bảng số liệu… nằm ngoài khung theo mẫu). Mỗi hàng ≈ 360 twips.
 */
function tableHeightBefore(body: XElement, boxP: XElement): number {
  // XNode của dự án chỉ có nextSibling → duyệt xuôi từ body, dừng tại boxP,
  // và chỉ cộng các BẢNG đứng ngay trước (cách boxP ≤ 3 phần tử).
  const kids = childrenOf(body);
  const at = kids.indexOf(boxP);
  if (at <= 0) return 0;
  let total = 0;
  for (let k = at - 1, guard = 0; k >= 0 && guard < 3; k--, guard++) {
    const el = kids[k];
    if (isW(el, "tbl")) {
      const rows = childrenOf(el).filter((c) => isW(c, "tr")).length || 3;
      total += rows * 380 + 200;
      continue;
    }
    if (isW(el, "p")) {
      const txt = paragraphText(el).trim();
      if (txt.length > 40) break; // gặp nội dung câu trước → dừng
    }
    break;
  }
  return total;
}

/** Tìm w:txbxContent bên trong shape của khung */
/**
 * Nhận diện HÌNH VẼ / ĐỒ THỊ của đề (để đưa ra ngoài khung bên phải).
 *
 * Phân biệt với CÔNG THỨC MathType (w:object + v:shape + o:OLEObject) — công thức
 * cũng là "ảnh" nhưng phải Ở TRONG khung. Dấu hiệu phân biệt:
 *  - Hình vẽ dùng `w:drawing` + `pic:pic` (DrawingML) hoặc `v:imagedata` KHÔNG có
 *    `w:object` bao ngoài, và KHÔNG có `<o:OLEObject>`.
 *  - Ngưỡng kích thước dự phòng: rộng > 60pt hoặc cao > 50pt (công thức chỉ
 *    11–56pt, và công thức luôn nằm trong w:object nên đã bị loại ở trên).
 */
function isBigPicture(p: XElement): boolean {
  let big = false;
  const walk = (n: XNode): void => {
    if (big) return;
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "object") {
        // w:object = công thức/OLE → KHÔNG phải hình vẽ cần đưa ra ngoài
        continue;
      }
      // CHỈ dựa vào KÍCH THƯỚC: hình vẽ/đồ thị của đề luôn lớn (>=80pt ở 1 chiều).
      // Không dùng "có w:drawing là hình" — ảnh công thức nhỏ cũng có thể nằm trong
      // drawing, nhận nhầm sẽ làm MỌI câu thành bố cục 2 cột (bug gặp).
      if (ln === "extent") {
        const cx = parseInt(attrAny(c as XElement, "cx") || "0", 10);
        const cy = parseInt(attrAny(c as XElement, "cy") || "0", 10);
        if (cy / 12700 >= 80 || cx / 12700 >= 120) big = true;
      } else if (ln === "shape") {
        const style = attrAny(c as XElement, "style") || "";
        const m = /height\s*:\s*([\d.]+)pt/i.exec(style);
        if (m && parseFloat(m[1]) >= 80) big = true;
      }
      walk(c);
    }
  };
  walk(p);
  return big;
}

/**
 * Tách HÌNH ra khỏi paragraph: tạo paragraph mới chỉ chứa run có ảnh (giữ
 * r:embed/r:id nên ảnh vẫn hiển thị), trả về paragraph đó. Dùng cho trường hợp
 * hình nằm chung paragraph với chữ câu hỏi (Câu 46 của đề HSA).
 */
function extractPicturePara(src: XElement, doc: XDocument): XElement | null {
  const picRuns: XElement[] = [];
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      if (localName(c) === "r" && isBigPicture(c)) {
        picRuns.push(c);
        continue;
      }
      walk(c);
    }
  };
  walk(src);
  if (!picRuns.length) return null;
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  pPr.appendChild(el(doc, "jc", { "w:val": "center" }));
  pPr.appendChild(el(doc, "spacing", { "w:before": "0", "w:after": "0", "w:line": "240", "w:lineRule": "auto" }));
  p.appendChild(pPr);
  for (const r of picRuns) {
    r.parentNode?.removeChild(r as XNode);
    p.appendChild(r as XNode);
  }
  return p;
}

/**
 * Bố cục 2 cột cho câu CÓ HÌNH: bảng 1 hàng × 2 ô, không viền.
 *  - ô TRÁI (boxW): paragraph chứa khung câu hỏi + đáp án.
 *  - ô PHẢI: các paragraph hình (căn giữa) → hình nằm NGOÀI khung, bên phải.
 * Trả về bảng để caller dùng làm mốc chèn dòng kẻ chấm.
 */
/** Bề rộng cột phải (twips) — dùng để thu nhỏ hình cho vừa */
function rightColWidth(boxW: number): number {
  return TEXT_W - boxW - 200;
}

/**
 * Thu nhỏ ảnh cho VỪA cột phải (nếu ảnh rộng hơn cột, Word sẽ tràn ra ngoài trang).
 * Giữ tỉ lệ: scale = min(1, maxW_EMU / cx). Cập nhật cả wp:extent và a:ext.
 */
function fitPictureWidth(picPara: XElement, maxWTwips: number): void {
  const maxEmu = maxWTwips * 635;
  const walk = (n: XNode): void => {
    for (const c of childrenOf(n)) {
      const ln = localName(c);
      if (ln === "extent" || ln === "ext") {
        const cxRaw = attrAny(c as XElement, "cx");
        const cyRaw = attrAny(c as XElement, "cy");
        const cx = cxRaw ? parseInt(cxRaw, 10) : 0;
        const cy = cyRaw ? parseInt(cyRaw, 10) : 0;
        if (cx > maxEmu && cx > 0) {
          const scale = maxEmu / cx;
          (c as XElement).setAttribute("cx", String(Math.round(maxEmu)));
          (c as XElement).setAttribute("cy", String(Math.round(cy * scale)));
        }
      }
      walk(c);
    }
  };
  walk(picPara);
}

function makeTwoColumnLayout(
  doc: XDocument,
  boxP: XElement,
  picParas: XElement[],
  boxW: number,
): XElement {
  const rightW = TEXT_W - boxW - 200; // trừ khe giữa 2 cột
  // Giữ vị trí chèn = ngay trước khung câu hỏi trong body
  const host = boxP.parentNode;
  const hostNext = boxP.nextSibling;
  const tbl = el(doc, "tbl");
  const tblPr = el(doc, "tblPr");
  tblPr.appendChild(el(doc, "tblW", { "w:w": String(TEXT_W), "w:type": "dxa" }));
  tblPr.appendChild(el(doc, "jc", { "w:val": "left" }));
  const borders = el(doc, "tblBorders");
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
    borders.appendChild(el(doc, side, { "w:val": "none", "w:sz": "0", "w:space": "0", "w:color": "auto" }));
  }
  tblPr.appendChild(borders);
  tblPr.appendChild(el(doc, "tblLayout", { "w:type": "fixed" }));
  tblPr.appendChild(el(doc, "tblCellMar"));
  tblPr.appendChild(el(doc, "tblLook", { "w:val": "0000" }));
  tbl.appendChild(tblPr);

  // Word cộng inset của shape (8pt = 160 twips mỗi bên) + viền vào bề rộng cột,
  // nên cột trái phải rộng hơn khung một khoản; nếu không Word tự nới cột và
  // hình bị đẩy lệch/tràn (bug gặp: cột 9356 thay vì boxW).
  const leftColW = boxW + 400;
  const realRightW = TEXT_W - leftColW;
  const grid = el(doc, "tblGrid");
  grid.appendChild(el(doc, "gridCol", { "w:w": String(leftColW) }));
  grid.appendChild(el(doc, "gridCol", { "w:w": String(realRightW) }));
  tbl.appendChild(grid);

  const tr = el(doc, "tr");
  const trPr = el(doc, "trPr");
  trPr.appendChild(el(doc, "cantSplit"));
  tr.appendChild(trPr);

  // ô TRÁI: khung câu hỏi (chuyển từ body vào)
  const tcL = el(doc, "tc");
  const tcPrL = el(doc, "tcPr");
  tcPrL.appendChild(el(doc, "tcW", { "w:w": String(leftColW), "w:type": "dxa" }));
  const marL = el(doc, "tcMar");
  for (const s of ["top", "left", "bottom", "right"]) {
    marL.appendChild(el(doc, s, { "w:w": "0", "w:type": "dxa" }));
  }
  tcPrL.appendChild(marL);
  tcL.appendChild(tcPrL);
  boxP.parentNode?.removeChild(boxP as XNode);
  tcL.appendChild(boxP as XNode);
  tr.appendChild(tcL);

  // ô PHẢI: hình vẽ (căn giữa)
  const tcR = el(doc, "tc");
  const tcPrR = el(doc, "tcPr");
  tcPrR.appendChild(el(doc, "tcW", { "w:w": String(realRightW), "w:type": "dxa" }));
  const marR = el(doc, "tcMar");
  marR.appendChild(el(doc, "left", { "w:w": "170", "w:type": "dxa" }));
  for (const s of ["top", "bottom", "right"]) {
    marR.appendChild(el(doc, s, { "w:w": "0", "w:type": "dxa" }));
  }
  tcPrR.appendChild(marR);
  tcPrR.appendChild(el(doc, "vAlign", { "w:val": "top" }));
  tcR.appendChild(tcPrR);
  if (picParas.length) {
    for (const pp of picParas) {
      // Thu nhỏ hình vừa cột phải — ảnh gốc rộng 540pt sẽ tràn ra ngoài trang 32cm
      fitPictureWidth(pp, rightColWidth(boxW) - 200);
      const pPr = childrenOf(pp).find((c) => isW(c, "pPr"));
      if (pPr) {
        for (const c of childrenOf(pPr)) {
          if (isW(c, "jc")) pPr.removeChild(c as XNode);
        }
        pPr.appendChild(el(doc, "jc", { "w:val": "center" }));
      }
      tcR.appendChild(pp as XNode);
    }
  } else {
    tcR.appendChild(el(doc, "p"));
  }
  tr.appendChild(tcR);

  tbl.appendChild(tr);
  // Chèn vào body tại đúng vị trí khung cũ (khung đã được chuyển vào ô trái nên
  // parentNode của nó là ô, không dùng được để chèn).
  if (host) {
    host.insertBefore(tbl as XNode, hostNext as XNode | null);
  }
  return tbl;
}

function findTxbxContent(boxP: XElement): XElement | null {
  let found: XElement | null = null;
  const walk = (node: XNode): void => {
    if (found) return;
    for (const c of childrenOf(node)) {
      if (localName(c) === "txbxContent") { found = c; return; }
      walk(c);
    }
  };
  walk(boxP);
  return found;
}

function hasVisual(p: XElement): boolean {
  const walk = (node: XNode): boolean => {
    for (const c of childrenOf(node)) {
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
 * Dùng TOÀN BỘ part của template làm gốc (styles/settings/theme/numbering),
 * chỉ thêm media/OLE của file đề với rels remap rIdSrc*.
 * KHÔNG trộn rels của file đề vào khung template — đã thử đủ cách, luôn ra
 * "file appears to be corrupted" (styles/numbering id lệch + rels trỏ chéo).
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
  // sectPr template (32×18 cm). Bỏ header/footerReference: template gốc trỏ
  // rId27-31 nhưng ta KHÔNG copy header/footer sang → Word báo corrupt.
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
  const body3 = body2.replace(/<w:(header|footer)Reference[^>]*r:id="[^"]*"[^>]*\/>/g, "");

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

  const documentXml = head2 + body3 + sectPr + tail;
  return { documentXml, relsXml: extraRels.join(""), srcMediaPaths };
}

/** rels của template đã lọc (bỏ target trỏ part không tồn tại + header/footer) */
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
      continue;
    }
    kept.push(tag);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${kept.join("")}</Relationships>`;
}
