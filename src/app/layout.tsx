import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Chuyển Word → Phiếu học tập & Tài liệu Live",
  description:
    "Tự động chuyển tài liệu Word thô của giáo viên thành Phiếu học tập cho học sinh (kiến thức có chỗ trống) và Tài liệu Live cho giáo viên (giữ nguyên nội dung). Xử lý hoàn toàn trên trình duyệt.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}