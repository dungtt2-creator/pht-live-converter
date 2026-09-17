/// <reference types="node" />
/**
 * Parser: docx (JSZip) → chuỗi Block/Table giữ nguyên thứ tự + tham chiếu XML.
 */

import JSZip from "jszip";
import { parseXml, paragraphText, localName, isW, childrenOf, asXDocument, asXElement } from "./xml";
import type { XDocument, XElement } from "./xml";
import type { Block, TableBlock, DocBlock, BlockKind } from "./model";
import { classifyParagraph } from "./model";

export interface ParsedXml {
  blocks: DocBlock[];
  /** phần tử w:body (XML) */
  body: XElement;
  /** tên file media (ảnh) có trong docx */
  media: string[];
  /** tài liệu XML gốc */
  doc: XDocument;
}

const QNUM = /^[Cc]âu\s*(\d+)\s*[.:]/;

function readParagraph(p: XElement, idx: number): Block {
  const text = paragraphText(p);
  const t = text.trim();
  const kind: BlockKind = t ? classifyParagraph(t, idx) : "text";
  const m = t.match(QNUM);
  return {
    kind,
    text: t,
    original: t,
    questionIndex: m ? parseInt(m[1], 10) : undefined,
  } as Block;
}

function readTable(tbl: XElement, idx: number, headingContext: string | undefined): TableBlock {
  const rows: { text: string; paras: string[] }[][] = [];
  const trs = childrenOf(tbl).filter((c) => isW(c, "tr"));
  for (const tr of trs) {
    const tcs = childrenOf(tr).filter((c) => isW(c, "tc"));
    const row: { text: string; paras: string[] }[] = [];
    for (const tc of tcs) {
      const paras: string[] = [];
      for (const c of childrenOf(tc)) {
        if (isW(c, "p")) {
          const t = paragraphText(c).trim();
          if (t) paras.push(t);
        }
      }
      row.push({ text: paras.join("\n"), paras });
    }
    if (row.length) rows.push(row);
  }
  return {
    kind: "table",
    rows,
    contextHeading: headingContext,
    pos: idx,
  };
}

export function parseXmlBlocks(xml: string, media: string[]): ParsedXml {
  const doc: XDocument = parseXml(xml);
  const root = doc.documentElement;
  if (!root) throw new Error("XML rỗng");
  const body = childrenOf(root).find((c) => isW(c, "body"));
  if (!body) throw new Error("Không tìm thấy w:body");
  const blocks: DocBlock[] = [];
  let headingContext: string | undefined;

  for (const child of childrenOf(body)) {
    if (isW(child, "p")) {
      const b = readParagraph(child as XElement, blocks.length);
      if (b.kind === "heading") headingContext = b.text;
      blocks.push(b);
    } else if (isW(child, "tbl")) {
      blocks.push(readTable(child as XElement, blocks.length, headingContext));
    }
    // sectPr — bỏ qua
  }

  return { blocks, body, media, doc };
}

export interface LoadedDocx {
  zip: JSZip;
  xml: string;
  media: string[];
  fileName: string;
}

export async function loadDocx(data: ArrayBuffer): Promise<LoadedDocx> {
  const zip = await JSZip.loadAsync(data);
  const docEntry = zip.file("word/document.xml");
  if (!docEntry) {
    // một số file có thể là .docx với document.xml ở vị trí khác — fallback
    const alt = zip.file(/document\.xml$/)[0];
    if (!alt) throw new Error("Không tìm thấy word/document.xml — file không phải .docx hợp lệ");
    const xml = await alt.async("string");
    const media: string[] = [];
    zip.forEach((relPath, entry) => {
      if (relPath.startsWith("word/media/") && !entry.dir) media.push(relPath);
    });
    return { zip, xml, media, fileName: alt.name };
  }
  const xml = await docEntry.async("string");
  const media: string[] = [];
  zip.forEach((relPath, entry) => {
    if (relPath.startsWith("word/media/") && !entry.dir) media.push(relPath);
  });
  return { zip, xml, media, fileName: docEntry.name };
}

/** Số câu hỏi thật sự */
export function countRealQuestions(blocks: DocBlock[]): number {
  return blocks.filter((b): b is Block => b.kind !== "table" && b.kind === "question").length;
}