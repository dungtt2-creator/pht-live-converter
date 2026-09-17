/**
 * Tải template thật từ public/templates/ — PHÂN TÁCH browser/node:
 * - Browser: fetch() (webpack-safe).
 * - Node CLI: đọc fs qua tham số loader (scripts/ truyền vào).
 */

import JSZip from "jszip";
import type { TemplateId } from "./template";
import { TEMPLATE_FILES } from "./template-file";

export interface LoadedTemplate {
  zip: JSZip;
  xml: string;
}

/** Hàm đọc file (browser dùng fetch; Node CLI truyền fs-based reader) */
export type TemplateFileReader = (file: string) => Promise<ArrayBuffer>;

const browserReader: TemplateFileReader = async (file) => {
  const base = window.location.pathname.replace(/\/[^/]*$/, "/");
  const res = await fetch(`${base}templates/${file}`);
  if (!res.ok) throw new Error(`Không tải được template ${file}`);
  return res.arrayBuffer();
};

let activeReader: TemplateFileReader | null = null;

export function setTemplateReader(reader: TemplateFileReader | null) {
  activeReader = reader;
}

export async function loadTemplate(
  templateId: TemplateId,
  mode: "pht" | "live",
  reader?: TemplateFileReader,
): Promise<LoadedTemplate> {
  const file = TEMPLATE_FILES[templateId][mode];
  const read = reader ?? activeReader ?? browserReader;
  const buf = await read(file);
  const zip = await JSZip.loadAsync(buf);
  const doc = zip.file("word/document.xml");
  if (!doc) throw new Error(`Template ${file} thiếu word/document.xml`);
  return { zip, xml: await doc.async("string") };
}