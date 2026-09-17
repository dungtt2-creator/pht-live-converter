/**
 * Định nghĩa template theo từng kỳ thi / hệ thống luyện thi.
 * Mỗi template = màu chủ đạo + dòng nhận diện in đầu tài liệu + phong cách.
 */

export type TemplateId = "tnthpt" | "hsa" | "qda" | "tsa" | "topclass";

export interface ExamTemplate {
  id: TemplateId;
  /** Tên hiển thị trên UI */
  label: string;
  /** Màu chủ đạo (viền bảng Live, dòng nhận diện) — hex không dấu # */
  color: string;
  /** Dòng nhận diện in đậm đầu PHT & Live */
  header: string;
  /** Ghi chú ngắn cho UI */
  desc: string;
  /** Live luôn landscape (form ngang) — tất cả template dùng chung */
}

export const TEMPLATES: ExamTemplate[] = [
  {
    id: "tnthpt",
    label: "TN THPT",
    color: "1F4E79",
    header: "KỲ THI TỐT NGHIỆP THPT",
    desc: "Màu xanh dương trang trọng — theo chuẩn Bộ GD&ĐT",
  },
  {
    id: "hsa",
    label: "HSA",
    color: "7030A0",
    header: "ĐÁNH GIÁ NĂNG LỰC HSA – ĐHQG HÀ NỘI",
    desc: "Màu tím đặc trưng của ĐGNL ĐHQG Hà Nội",
  },
  {
    id: "qda",
    label: "QDA",
    color: "ED7D31",
    header: "ĐÁNH GIÁ NĂNG LỰC QDA – ĐHQG TP. HỒ CHÍ MINH",
    desc: "Màu cam đỏ của ĐGNL ĐHQG TP.HCM",
  },
  {
    id: "tsa",
    label: "TSA",
    color: "C00000",
    header: "ĐÁNH GIÁ NĂNG LỰC TSA – ĐẠI HỌC BÁCH KHOA HÀ NỘI",
    desc: "Màu đỏ đặc trưng của ĐHBK HN (template hiện tại)",
  },
  {
    id: "topclass",
    label: "Topclass",
    color: "00A650",
    header: "TOPCLASS – LUYỆN THI CÙNG HOCMAI",
    desc: "Màu xanh lá của hệ thống Topclass",
  },
];

export const DEFAULT_TEMPLATE: TemplateId = "tsa";

export function getTemplate(id: string): ExamTemplate {
  return TEMPLATES.find((t) => t.id === id) ?? TEMPLATES.find((t) => t.id === DEFAULT_TEMPLATE)!;
}