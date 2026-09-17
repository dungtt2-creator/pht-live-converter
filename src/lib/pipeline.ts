/**
 * Pipeline: docx (ArrayBuffer) → { phtDocx, liveDocx } (ArrayBuffer).
 * Client-side hoàn toàn — không cần server.
 */

import JSZip from "jszip";
import { loadDocx, parseXmlBlocks } from "./parser";
import { convertParsed, type OutputMode, type ConvertResult } from "./converter";
import { serializeXml, stripRedundantXmlns } from "./xml";
import type { TemplateId } from "./template";
import { loadTemplate, extractSkeleton, stripBodySkeleton, buildFromTemplate } from "./template-file";

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

  let pht: ArrayBuffer | null = null;
  let live: ArrayBuffer | null = null;
  let phtResult: ConvertResult | null = null;
  let liveResult: ConvertResult | null = null;

  // Mỗi mode parse XML TƯƠI để không bị đột biến DOM của mode trước
  for (const mode of options.modes) {
    const parsed = parseXmlBlocks(xml, media);
    const result = convertParsed(parsed.blocks, parsed.body, mode, templateId, useFile);
    const newXml = serializeXml(parsed.doc);

    let outBuf: ArrayBuffer;
    if (useFile) {
      outBuf = await wrapWithTemplateFile(templateId, mode, newXml, zip);
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
 */
async function wrapWithTemplateFile(
  templateId: TemplateId,
  mode: OutputMode,
  convertedXml: string,
  gvZip: JSZip,
): Promise<ArrayBuffer> {
  const tpl = await loadTemplate(templateId, mode);
  const skeleton = extractSkeleton(tpl.xml);
  const gvBody = stripBodySkeleton(convertedXml);
  const merged = await buildFromTemplate(tpl, skeleton, gvBody, gvZip);

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