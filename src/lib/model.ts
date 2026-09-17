/**
 * Mô hình dữ liệu trung gian sau khi parse file .docx thô của Giáo viên.
 * Document là chuỗi các Block (đoạn văn) và Table (bảng), giữ nguyên thứ tự.
 */

export type BlockKind =
  | "text"          // câu hỏi, lý thuyết, hướng dẫn, đáp án... (văn bản thường)
  | "question"      // bắt đầu bằng "Câu N." hoặc "Câu N:"
  | "choices"       // dòng 4 phương án A. B. C. D.
  | "heading"       // tiêu đề cấp I. II. / 1. 2. / 1.1
  | "answer"        // dòng "Đáp án:" / "Đáp án (1)..." / "Vị trí N: ..."
  | "guideline"     // dòng "Hướng dẫn:" / "Lời giải:"
  | "theory"        // đoạn lý thuyết (bình thường, không thuộc câu hỏi)
  | "instruction"   // dòng "Sử dụng thông tin dưới đây..." (đề chung)

export interface Block {
  kind: BlockKind;
  /** Văn bản gốc (đã strip tab, trim) */
  text: string;
  /** Phần văn bản gốc theo từng dòng (cho phân tích câu) */
  original: string;
  /** Nếu là answer/guideline đi kèm câu hỏi: số thứ tự câu hỏi gần nhất */
  questionIndex?: number;
  /** Cờ: thuộc phần bài tập (sau heading lý thuyết) */
  inExercise?: boolean;
}

export interface CellLike {
  /** văn bản của cell, nối nhiều paragraph bằng \n */
  text: string;
  /** từng paragraph riêng */
  paras: string[];
  /** có highlight vàng (tùy thuộc từng implementation) */
  highlighted?: boolean;
}

export interface TableBlock {
  kind: "table";
  rows: CellLike[][];
  /** heading ngay trước bảng (nếu có) */
  contextHeading?: string;
  /** vị trí trong chuỗi block */
  pos: number;
}

export type DocBlock = Block | TableBlock;
export type ParsedDoc = DocBlock[];

/** Quy tắc phân loại 1 đoạn văn -> BlockKind */
const ROMAN = /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)\b/;
const QNUM = /^[Cc]âu\s*(\d+)\s*[.:]/;

export function classifyParagraph(text: string, idx: number): BlockKind {
  const t = text.trim();
  if (!t) return "text";
  // Heading: "I. HỆ..." hoặc "1. Các phép đo..." hoặc "3.2. Tiêu đề"
  if (/^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)[\.\s]\s*[A-ZÀ-ỸĀĐỲ]/.test(t)) return "heading";
  if (/^(\d+(\.\d+)*)[\.\)]\s+[A-ZÀ-ỸĀĐỲ]/.test(t)) return "heading";
  // Câu hỏi
  if (QNUM.test(t)) return "question";
  // Phương án trắc nghiệm: "A. ..." (có 2+ trên cùng dòng hoặc dòng đơn)
  if (/^[A-D][\.\s]\s{1,2}/.test(t) || /^A\.[^\n]*\tB\./.test(t)) return "choices";
  // Nhãn đặc biệt
  if (/^Đáp\s?án\s*[:.·)/]/i.test(t)) {
    // "Đáp án: A. ____; ____" = mẫu điền (đề) — KHÔNG phải đáp án, giữ
    if (/_{2,}/.test(t)) return "text";
    return "answer";
  }
  if (/^(Hướng\s?dẫn|Lời\s?giải)\s*[:.)·]/i.test(t)) return "guideline";
  if (/^Sử dụng thông tin|^Dựa vào thông tin|^Cho đoạn trích|^Đọc đoạn văn/.test(t)) return "instruction";
  return "theory";
}

/** Tách chuỗi phương án "A. x\tB. y\tC. z\tD. w" thành mảng */
export function splitChoices(text: string): string[] {
  const m = text.split(/\t+/);
  return m.map((s) => s.trim()).filter(Boolean);
}

/** Đếm số câu hỏi xuất hiện trong blocks */
export function countQuestions(blocks: DocBlock[]): number {
  return blocks.filter((b): b is Block => b.kind !== "table" && b.kind === "question").length;
}