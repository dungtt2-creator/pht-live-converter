/**
 * Pipeline: docx (ArrayBuffer) → { phtDocx, liveDocx } (ArrayBuffer).
 * Client-side hoàn toàn — không cần server.
 */

import JSZip from "jszip";
import { loadDocx, parseXmlBlocks } from "./parser";
import { convertParsed, type OutputMode, type ConvertResult } from "./converter";
import { serializeXml, stripRedundantXmlns } from "./xml";
import type { TemplateId } from "./template";
import { loadTemplate, extractSkeleton, stripBodySkeleton, buildFromTemplateKeepTitle } from "./template-file";
import { applyVodLayout, buildVodDocument, readVodTemplateRels, type VodConvertStats } from "./template-vod";

export interface ConversionOutput {
  pht: ArrayBuffer | null;
  live: ArrayBuffer | null;
  phtResult: ConvertResult | null;
  liveResult: ConvertResult | null;
  fileName: string;
  questionCount: number;
  theoryParagraphs: number;
}

export interface ConvertOptions {
  modes: OutputMode[];
  /** Tên file gốc để đặt tên output */
  fileName?: string;
  /** Template kỳ thi (màu sắc + banner); mặc định TSA */
  templateId?: TemplateId;
  /** Dùng file template THẬT (public/templates/) làm khung output */
  useFileTemplate?: boolean;
  /**
   * "vod" = bố cục Tài liệu dạy học ghi hình: khổ 32 × 18 cm, bảng 2 cột
   * (câu hỏi | dòng chấm), Times New Roman 14. Mặc định "exam" (bản cũ).
   */
  layout?: "exam" | "vod";
}

/** Pack lại docx với document.xml mới */
async function packDocx(zip: JSZip, newDocumentXml: string): Promise<ArrayBuffer> {
  const out = new JSZip();
  // copy tất cả entry (giữ cấu trúc) trừ document.xml
  const tasks: Promise<void>[] = [];
  zip.forEach((relPath, entry) => {
    if (entry.dir) return;
    if (relPath === "word/document.xml") return; // ghi đè sau
    tasks.push(
      zip.file(relPath)!.async("uint8array").then((data) => {
        out.file(relPath, data);
      }),
    );
  });
  await Promise.all(tasks);
  out.file("word/document.xml", newDocumentXml);
  return out.generateAsync({
    type: "arraybuffer",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

/** Đổi tên file output chuẩn theo quy ước: <tên>_PHT.docx / <tên>_Live.docx */
export function outputFileName(base: string, mode: OutputMode): string {
  const lower = base.toLowerCase().replace(/\.docx?$/i, "");
  return `${lower}_${mode === "pht" ? "PHT" : "Live"}.docx`;
}

/**
 * Tên file thành phẩm GIỮ NGUYÊN cấu trúc tên file đề + hậu tố "_Live"
 * (yêu cầu user 2026-10: GV nhận biết ngay bản dạy học ghi hình).
 * VD: "HSA - BF_CB.docx" → "HSA - BF_CB_Live.docx"
 */
export function outputFileNameKeepCase(base: string, suffix = "Live"): string {
  const noExt = base.replace(/\.docx?$/i, "");
  return `${noExt}_${suffix}.docx`;
}

/**
 * Chuyển đổi một file .docx → PHT + Live.
 * Khi options.useFileTemplate=true: đầu ra được bọc trong khung file template thật
 * (banner + page setup theo kỳ thi).
 */
export async function convertDocx(
  data: ArrayBuffer,
  options: ConvertOptions = { modes: ["pht", "live"] },
): Promise<ConversionOutput> {
  const { zip, xml, media } = await loadDocx(data);
  const templateId = options.templateId ?? "tsa";
  const useFile = options.useFileTemplate ?? false;
  const layout = options.layout ?? "exam";

  let pht: ArrayBuffer | null = null;
  let live: ArrayBuffer | null = null;
  let phtResult: ConvertResult | null = null;
  let liveResult: ConvertResult | null = null;

  // Mỗi mode parse XML TƯƠI để không bị đột biến DOM của mode trước
  for (const mode of options.modes) {
    const parsed = parseXmlBlocks(xml, media);
    const result = convertParsed(parsed.blocks, parsed.body, mode, templateId, useFile, layout);
    const newXml = serializeXml(parsed.doc);

    let outBuf: ArrayBuffer;
    if (layout === "vod" && mode === "live") {
      outBuf = await wrapWithVodTemplate(newXml, zip);
    } else if (useFile) {
      outBuf = await wrapWithTemplateFileKeepTitle(templateId, mode, newXml, zip);
    } else {
      outBuf = await packDocx(zip, newXml);
    }

    if (mode === "pht") {
      pht = outBuf;
      phtResult = result;
    } else {
      live = outBuf;
      liveResult = result;
    }
  }

  const parsed0 = parseXmlBlocks(xml, media);
  const questionCount = parsed0.blocks.filter((b) => b.kind !== "table" && b.kind === "question").length;
  const theoryParagraphs = parsed0.blocks.filter((b) => b.kind !== "table" && b.kind === "theory").length;

  return {
    pht,
    live,
    phtResult,
    liveResult,
    fileName: options.fileName || "document.docx",
    questionCount,
    theoryParagraphs,
  };
}

/**
 * Bọc nội dung GV (document.xml đã convert) vào khung template thật:
 *  - skeleton: banner + sectPr từ template
 *  - zip gốc = zip GV (giữ content-types/media/OLE)
 *  - media banner template copy với tên remap, thêm rels.
 *
 * KHÔNG CHÈN BANNER TEMPLATE — giữ NGUYÊN tiêu đề bài học của file gốc
 * (yêu cầu user: "Phần tiêu đề của bài học cần giữ cố định như file gốc,
 * không tự chỉnh sửa tiêu đề của file"). Chỉ lấy sectPr (page setup:
 * khổ giấy + margin) từ template để bố cục đúng kỳ thi.
 */
async function wrapWithTemplateFileKeepTitle(
  templateId: TemplateId,
  mode: OutputMode,
  convertedXml: string,
  gvZip: JSZip,
): Promise<ArrayBuffer> {
  const tpl = await loadTemplate(templateId, mode);
  const skeleton = extractSkeleton(tpl.xml);
  const gvBody = stripBodySkeleton(convertedXml);
  // KHÔNG thêm skeleton.prefixXml (banner template) — giữ tiêu đề gốc của file GV.
  // vẫn remap sectPr header/footer rId + media ảnh nếu template có.
  const merged = await buildFromTemplateKeepTitle(tpl, skeleton, gvBody, gvZip);

  // Đóng gói: copy zip GV (trừ document.xml + rels) + merged + media template
  const out = new JSZip();
  const tasks: Promise<void>[] = [];
  gvZip.forEach((p, e) => {
    if (e.dir) return;
    if (p === "word/document.xml" || p === "word/_rels/document.xml.rels") return;
    tasks.push(
      e.async("uint8array").then((c) => {
        out.file(p, c);
      }),
    );
  });
  await Promise.all(tasks);
  out.file("word/document.xml", merged.documentXml);
  out.file("word/_rels/document.xml.rels", merged.relsXml);
  for (const [p, c] of merged.extraMedia) out.file(p, c);
  return out.generateAsync({
    type: "arraybuffer",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

export { stripRedundantXmlns }; // tái xuất cho CLI debug

/**
 * Đóng gói Tài liệu VOD: zip TEMPLATE VOD làm gốc (32×18 cm, styles/settings
 * riêng, không header/footer) + body GV đã convert + media/OLE của file GV.
 */
async function wrapWithVodTemplate(convertedXml: string, gvZip: JSZip): Promise<ArrayBuffer> {
  const tpl = await loadTemplate("vod" as TemplateId, "live");
  const gvRelsEntry = gvZip.file("word/_rels/document.xml.rels");
  const gvRels = gvRelsEntry ? await gvRelsEntry.async("string") : "";
  const merged = buildVodDocument(tpl, convertedXml, gvRels);
  const tplRels = await readVodTemplateRels(tpl);

  const out = new JSZip();
  const tasks: Promise<void>[] = [];
  tpl.zip.forEach((p, e) => {
    if (e.dir) return;
    if (p === "word/document.xml" || p === "word/_rels/document.xml.rels") return;
    // bỏ phần chỉ thuộc template không dùng (font nhúng, header/footer không tồn tại)
    tasks.push(
      e.async("uint8array").then((c) => {
        out.file(p, c);
      }),
    );
  });
  await Promise.all(tasks);

  // media + OLE của file GV (giữ tên gốc; template không có trùng tên)
  const mediaTasks: Promise<void>[] = [];
  for (const p of new Set(merged.srcMediaPaths)) {
    const f = gvZip.file(p);
    if (!f) continue;
    mediaTasks.push(
      f.async("uint8array").then((c) => {
        out.file(p, c);
      }),
    );
  }
  await Promise.all(mediaTasks);

  // content-types: bổ sung Default cho phần media của GV (emf/jpeg/gif…) nếu thiếu
  const ctEntry = tpl.zip.file("[Content_Types].xml");
  if (ctEntry) {
    let ct = await ctEntry.async("string");
    const have = new Set<string>();
    const defRe = /<Default\s+Extension="([^"]+)"/g;
    let m: RegExpExecArray | null;
    while ((m = defRe.exec(ct))) have.add(m[1].toLowerCase());
    const need: Record<string, string> = {
      jpeg: "image/jpeg", jpg: "image/jpeg", gif: "image/gif", emf: "image/x-emf",
      bmp: "image/bmp", tiff: "image/tiff", bin: "application/vnd.openxmlformats-officedocument.oleObject",
    };
    let add = "";
    for (const p of merged.srcMediaPaths) {
      const ext = (p.split(".").pop() || "").toLowerCase();
      if (!ext || have.has(ext) || !need[ext]) continue;
      have.add(ext);
      add += `<Default Extension="${ext}" ContentType="${need[ext]}"/>`;
    }
    if (add) ct = ct.replace("<Types ", "<Types ").replace(/(<Types[^>]*>)/, `$1${add}`);
    out.file("[Content_Types].xml", ct);
  }

  out.file("word/document.xml", merged.documentXml);
  out.file("word/_rels/document.xml.rels", tplRels.replace("</Relationships>", merged.relsXml + "</Relationships>"));
  return out.generateAsync({
    type: "arraybuffer",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}