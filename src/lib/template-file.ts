/**
 * Template THẬT (public/templates/) làm khung sườn — chiến lược INJECT:
 *
 *  - ZIP GỐC là file GV (giữ nguyên content-types, media, OLE, rels gốc — an toàn tuyệt đối).
 *  - Từ template: lấy banner prefix (logo + "TỔNG ÔN...") + sectPr (trang ngang/dọc).
 *  - Inject prefix vào đầu body GV, thay sectPr GV bằng sectPr template.
 *  - Nếu prefix chứa ảnh (r:embed): copy file media của template vào zip GV
 *    với tên remap + thêm Relationship. Tránh trùng rId với GV.
 */

import JSZip from "jszip";
import type { TemplateId } from "./template";

export const TEMPLATE_FILES: Record<TemplateId, { pht: string; live: string }> = {
  tnthpt: { pht: "pht.docx", live: "live-tnthpt.docx" },
  hsa: { pht: "pht.docx", live: "live-hsa.docx" },
  qda: { pht: "pht.docx", live: "live-qda.docx" },
  tsa: { pht: "pht.docx", live: "live-tsa.docx" },
  topclass: { pht: "pht.docx", live: "live-topclass.docx" },
};

import { loadTemplate, type LoadedTemplate } from "./template-loader";
export { loadTemplate, type LoadedTemplate };

export interface TemplateSkeleton {
  /** Phần đầu tài liệu: <?xml?> + <w:document ...ns...> + <w:body> (namespace đầy đủ) */
  docHead: string;
  /** Prefix XML: logo + banner "TỔNG ÔN..." (sau <w:body>) */
  prefixXml: string;
  /** sectPr template (page setup) */
  sectPr: string;
  /** Phần cuối sau sectPr (thường "</w:body></w:document>") */
  tailXml: string;
}

/** Tách prefix (banner) + sectPr từ template */
export function extractSkeleton(xml: string): TemplateSkeleton {
  const bodyIdx = xml.indexOf("<w:body>");
  const sectIdx = xml.lastIndexOf("<w:sectPr");
  if (bodyIdx < 0 || sectIdx < 0) throw new Error("Template không hợp lệ (thiếu body/sectPr)");

  const docHead = xml.slice(0, bodyIdx);
  const bodyStart = bodyIdx + "<w:body>".length;
  const bannerIdx = xml.indexOf("TỔNG ÔN", bodyStart);
  let prefixEnd = bodyStart;
  if (bannerIdx > bodyStart) {
    const rel = xml.slice(bodyStart, bannerIdx);
    const blockStart = bodyStart + Math.max(rel.lastIndexOf("<w:p "), rel.lastIndexOf("<w:tbl>"));
    const endP = xml.indexOf("</w:p>", blockStart);
    const endT = xml.indexOf("</w:tbl>", blockStart);
    const end = endP > 0 && (endT < 0 || endP < endT) ? endP + "</w:p>".length : endT > 0 ? endT + "</w:tbl>".length : -1;

    // Cộng thêm block tiêu đề phụ ("SỰ CHUYỂN THỂ...") — tối đa 2 paragraph liền sau
    let cur = end > 0 ? end : bodyStart;
    for (let k = 0; k < 2; k++) {
      const nextP = xml.indexOf("<w:p ", cur);
      if (nextP < 0 || nextP > sectIdx) break;
      const nextEnd = xml.indexOf("</w:p>", nextP);
      if (nextEnd < 0 || nextEnd > sectIdx) break;
      const seg = xml.slice(nextP, nextEnd);
      const t = seg.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
      const txt = t ? t[1] : "";
      if (txt.length > 60 || /Câu \d/.test(txt)) break;
      cur = nextEnd + "</w:p>".length;
    }
    prefixEnd = cur;
  }

  const sectPrEnd = xml.indexOf("</w:sectPr>", sectIdx) + "</w:sectPr>".length;
  return {
    docHead,
    prefixXml: xml.slice(bodyIdx, prefixEnd),
    sectPr: xml.slice(sectIdx, sectPrEnd),
    tailXml: xml.slice(sectPrEnd),
  };
}

export interface MergeResult {
  documentXml: string;
  /** rels mới (zip GV + rels ảnh template) */
  relsXml: string;
  /** file media template cần thêm (path → content) */
  extraMedia: Map<string, Uint8Array>;
}

/** Nội dung body GV (không sectPr) từ document.xml đã convert */
export function stripBodySkeleton(gvXml: string): string {
  const bodyIdx = gvXml.indexOf("<w:body>");
  const sectIdx = gvXml.lastIndexOf("<w:sectPr");
  if (bodyIdx < 0) throw new Error("Nội dung GV không hợp lệ");
  const endIdx = sectIdx > bodyIdx ? sectIdx : gvXml.indexOf("</w:body>");
  return gvXml.slice(bodyIdx + "<w:body>".length, endIdx);
}

/**
 * Ghép: inject prefix template + sectPr template vào body GV.
 * gvZip: zip FILE GỐC (giữ nguyên — là gốc output). tplZip: zip template (đọc media).
 */
export async function buildFromTemplate(
  tpl: LoadedTemplate,
  skeleton: TemplateSkeleton,
  gvBodyXml: string,
  gvZip: JSZip,
): Promise<MergeResult> {
  // 1) Tìm r:embed trong prefix (ảnh banner template)
  const embeds = new Set<string>();
  const embedRe = /r:embed="([^"]+)"/g;
  let mm: RegExpExecArray | null;
  while ((mm = embedRe.exec(skeleton.prefixXml))) embeds.add(mm[1]);

  // 2) Đọc rels template để map rId → target
  const tplRelsEntry = tpl.zip.file("word/_rels/document.xml.rels");
  const tplRelsXml = tplRelsEntry ? await tplRelsEntry.async("string") : "";
  const relMap = new Map<string, { type: string; target: string }>();
  const relRegex = /<Relationship\s+Id="([^"]+)"\s+Type="([^"]+)"\s+Target="([^"]+)"\s*\/?>/g;
  let rr: RegExpExecArray | null;
  while ((rr = relRegex.exec(tplRelsXml))) relMap.set(rr[1], { type: rr[2], target: rr[3] });

  // 3) Remap rId ảnh banner → rIdTPL{n}; copy media vào extraMedia
  const extraMedia = new Map<string, Uint8Array>();
  let prefixXml = skeleton.prefixXml;
  let n = 0;
  for (const rid of embeds) {
    const rel = relMap.get(rid);
    if (!rel) continue;
    // target như "media/image1.png"
    const tgt = rel.target.startsWith("media/") ? `word/${rel.target}` : rel.target;
    const tplFile = tpl.zip.file(tgt);
    if (!tplFile) continue;
    n++;
    const newRid = `rIdTpl${n}`;
    const content = await tplFile.async("uint8array");
    // tên mới tránh đè media GV
    const mediaName = rel.target.includes("/") ? rel.target.split("/").pop()! : rel.target;
    const newPath = `word/media/tpl_${mediaName}`;
    extraMedia.set(newPath, content);
    prefixXml = prefixXml.split(`r:embed="${rid}"`).join(`r:embed="${newRid}"`);
    // lưu rel mới
    embedsRel.set(newRid, { type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", target: `media/tpl_${mediaName}` });
  }

  // 4) Rel GV + rel ảnh banner mới
  let gvRelsXml = "";
  const gvRelsEntry = gvZip.file("word/_rels/document.xml.rels");
  if (gvRelsEntry) gvRelsXml = await gvRelsEntry.async("string");
  let relsXml = gvRelsXml || "";
  if (relsXml.includes("</Relationships>")) {
    let add = "";
    embedsRel.forEach((rel, rid) => {
      add += `<Relationship Id="${rid}" Type="${rel.type}" Target="${rel.target}"/>`;
    });
    relsXml = relsXml.replace("</Relationships>", add + "</Relationships>");
  } else {
    // tạo mới
    const NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    let body = "";
    embedsRel.forEach((rel, rid) => {
      body += `<Relationship Id="${rid}" Type="${rel.type}" Target="${rel.target}"/>`;
    });
    relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS}">${body}</Relationships>`;
  }

  // 5) document.xml: docHead + prefix + body GV + sectPr template + tail đóng
    const documentXml =
      skeleton.docHead +
      skeleton.prefixXml +
      gvBodyXml +
      skeleton.sectPr +
      skeleton.tailXml;

    return { documentXml, relsXml, extraMedia };
  }

  const embedsRel = new Map<string, { type: string; target: string }>();