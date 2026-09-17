/**
 * Pipeline: docx (ArrayBuffer) → { phtDocx, liveDocx } (ArrayBuffer).
 * Client-side hoàn toàn — không cần server.
 */

import JSZip from "jszip";
import { loadDocx, parseXmlBlocks } from "./parser";
import { convertParsed, type OutputMode, type ConvertResult } from "./converter";
import { serializeXml, stripRedundantXmlns } from "./xml";
import type { TemplateId } from "./template";

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
 */
export async function convertDocx(
  data: ArrayBuffer,
  options: ConvertOptions = { modes: ["pht", "live"] },
): Promise<ConversionOutput> {
  const { zip, xml, media } = await loadDocx(data);

  let pht: ArrayBuffer | null = null;
  let live: ArrayBuffer | null = null;
  let phtResult: ConvertResult | null = null;
  let liveResult: ConvertResult | null = null;

  // Mỗi mode parse XML TƯƠI để không bị đột biến DOM của mode trước
  for (const mode of options.modes) {
    const parsed = parseXmlBlocks(xml, media);
    const result = convertParsed(parsed.blocks, parsed.body, mode, options.templateId ?? "tsa");
    const newXml = serializeXml(parsed.doc);
    if (mode === "pht") {
      pht = await packDocx(zip, newXml);
      phtResult = result;
    } else {
      live = await packDocx(zip, newXml);
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

export { stripRedundantXmlns }; // tái xuất cho CLI debug