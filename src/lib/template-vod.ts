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
export const VOD_BORDER = "C0504D";
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
 * Bề rộng khung câu hỏi. Mẫu: khung 398.5pt ≈ 8080 twips (14.2 cm) — giữ đúng
 * để diện mạo giống mẫu; phần còn lại của trang là lề trống như file mẫu.
 */
const BOX_W = 8080;
/** Góc bo của shape mẫu: arcsize="10923f" (f = 1/100000) */
const ARC_SIZE = "10923f";
/** Số dòng kẻ chấm cho mỗi câu */
const DOTS_PER_QUESTION = 4;
/** Khoảng cách dòng kẻ chấm: line 480 = gấp đôi dòng đơn (~32pt) — đều nhau */
const DOTS_LINE = "480";

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
export function makeVodDots(doc: XDocument): XElement {
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  const tabs = el(doc, "tabs");
  tabs.appendChild(el(doc, "tab", { "w:val": "right", "w:leader": "dot", "w:pos": String(TEXT_W) }));
  pPr.appendChild(tabs);
  pPr.appendChild(el(doc, "spacing", { "w:before": "0", "w:after": "0", "w:line": DOTS_LINE, "w:lineRule": "auto" }));
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
 * Khung câu hỏi bo tròn 4 góc = shape VML roundrect, viền 1pt đỏ gạch, nền trắng.
 * - arcsize 10923f: đúng độ bo của mẫu (nhỏ, không phải viên thuốc).
 * - `mso-fit-shape-to-text:t` + width cố định: khung TỰ CAO theo nội dung, nên
 *   câu hỏi dài không bị tràn/cắt chữ (textbox mẫu thì cố định chiều cao).
 * - Shape đặt INLINE trong 1 paragraph (không floating): mọi thứ chèn sau nó
 *   tự chảy xuống dưới khung, không đè lên nhau.
 */
function makeVodQuestionBox(doc: XDocument): XElement {
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  pPr.appendChild(el(doc, "spacing", { "w:before": "0", "w:after": "0", "w:line": "240", "w:lineRule": "auto" }));
  pPr.appendChild(el(doc, "ind", { "w:left": "0", "w:right": "0", "w:firstLine": "0" }));
  p.appendChild(pPr);

  const r = el(doc, "r");
  const pict = el(doc, "pict");
  const shape = elVml(doc, "roundrect", {
    arcsize: ARC_SIZE,
    style: `width:${(BOX_W / 1440).toFixed(2)}in`,
    fillcolor: "#FFFFFF",
    strokecolor: `#${VOD_BORDER}`,
    strokeweight: "1pt",
  });
  const textbox = elVml(doc, "textbox", {
    style: "mso-fit-shape-to-text:t",
    inset: "6pt,3pt,6pt,3pt",
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
  // Chia đều bề rộng khung: A ở lề trái (0), các nhãn sau cách đều `step`.
  // Tab stop đầu = 0 để A không bị thụt vào; nếu để 200 thì cột A hẹp hơn các cột.
  const step = Math.floor(BOX_W / present.length);
  const p = el(doc, "p");
  const pPr = el(doc, "pPr");
  const tabs = el(doc, "tabs");
  present.forEach((_, i) => {
    tabs.appendChild(el(doc, "tab", { "w:val": "left", "w:pos": String(i * step) }));
  });
  pPr.appendChild(tabs);
  pPr.appendChild(el(doc, "spacing", { "w:before": "60", "w:after": "40", "w:line": "360", "w:lineRule": "auto" }));
  pPr.appendChild(el(doc, "ind", { "w:left": "0", "w:right": "0", "w:firstLine": "0" }));
  pPr.appendChild(tnr(doc));
  p.appendChild(pPr);

  present.forEach((lab, i) => {
    if (i > 0) {
      const rTab = el(doc, "r");
      rTab.appendChild(tnr(doc));
      rTab.appendChild(el(doc, "tab"));
      p.appendChild(rTab);
    }
    p.appendChild(runText(doc, `${lab}. `, { bold: true, color: VOD_BORDER }));
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
  dotsPerQuestion = DOTS_PER_QUESTION,
): VodConvertStats {
  const idx: Array<{ el: XElement; isTable: boolean }> = [];
  for (const child of childrenOf(body)) {
    if (isW(child, "p")) idx.push({ el: child, isTable: false });
    else if (isW(child, "tbl")) idx.push({ el: child, isTable: true });
  }

  const stats: VodConvertStats = {
    questions: 0, groups: 0, ids: [], missingIds: [], highlightRemoved: 0, dotRows: 0,
    headerRemoved: 0, choicesMerged: 0,
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
      if (bb.kind === "table" || idx[j].isTable) break; // bảng ngữ liệu: NGOÀI khung
      const blk = bb as Block;
      if (j > i && blk.kind === "question") break;
      if (j > i && (blk.kind === "heading" || blk.kind === "instruction" || blk.kind === "guideline")) break;
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
  for (const g of groups) {
    const boxP = makeVodQuestionBox(doc);
    const first = g.elems[0];
    first.parentNode!.insertBefore(boxP as XNode, first as XNode);

    // 3a) Dồn 4 phương án về 1 dòng (nếu nhận diện được >= 2 phương án)
    let mergedChoices: XElement | null = null;
    let choiceLabels = 0;
    if (g.choiceParas.length) {
      const res = mergeChoicesIntoOneLine(g.choiceParas, doc);
      mergedChoices = res.merged;
      choiceLabels = res.labels;
    }

    // 3b) Chuyển các paragraph còn lại vào txbxContent của khung
    const content = findTxbxContent(boxP);
    if (content) {
      for (const c of childrenOf(content)) {
        if (isW(c, "p") && paragraphText(c).trim() === "" && !hasVisual(c)) {
          content.removeChild(c as XNode);
          break;
        }
      }
      const skip = new Set<XElement>(g.choiceParas);
      for (const e of g.elems) {
        if (skip.has(e)) continue;
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

    // 3c) Dòng kẻ chấm ngay SAU khung, dài liền hết lề trang
    let anchor: XNode = boxP as XNode;
    for (let k = 0; k < dotsPerQuestion; k++) {
      const dots = makeVodDots(doc);
      anchor.parentNode!.insertBefore(dots as XNode, anchor.nextSibling);
      anchor = dots as XNode;
    }

    stats.groups++;
    stats.questions++;
    stats.dotRows += dotsPerQuestion;
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

/** Tìm w:txbxContent bên trong shape của khung */
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
