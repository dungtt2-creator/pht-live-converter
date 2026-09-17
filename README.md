# Chuyển Word Giáo viên → Phiếu học tập & Tài liệu Live

Web app (Next.js 15, **chạy 100% client-side**) tự động chuyển tài liệu Word thô `.docx`
của giáo viên thành 2 đầu ra:

| Sản phẩm | Đối tượng | Đặc điểm |
| --- | --- | --- |
| **Phiếu học tập (PHT)** | Học sinh | Giữ mục tiêu + hệ thống kiến thức (tự động tạo **chỗ trống gạch chân** cho HS điền), câu hỏi/bài tập; **XÓA** đáp án, hướng dẫn giải, bảng giải thích |
| **Tài liệu Live** | Giáo viên | **Giữ nguyên 100%** nội dung gốc (kiến thức + đáp án + hướng dẫn), thêm vùng trả lời dòng chấm sau câu hỏi |

## Nguyên tắc thiết kế (không vi phạm)

- **Không tự ý sửa kiến thức**: mọi biến đổi là cấu trúc (giữ/xóa/thêm vùng trống), không viết lại nội dung khoa học.
- **Giữ nguyên định dạng gốc**: xử lý trên chính `document.xml` (không dựng lại từ text) → giữ trọn OMML math, ảnh, bảng, run formatting.
- **Chỗ trống có cơ sở sư phạm**: chỉ blank các câu định nghĩa ngắn (`... gọi là X`) khớp mẫu PHT thủ công; giữ "phép đo ____" dạng có từ loại đứng đầu; câu dài/chứa công thức → bỏ qua (an toàn).
- **Xóa đáp án có điểm dừng**: xóa khối "Hướng dẫn/Lời giải/Đáp án" cho tới khi gặp câu hỏi mới, heading, ngữ liệu ("Sử dụng thông tin..."), "DẶN DÒ" — không nuốt nội dung.

## Chạy local

```bash
npm install
npm run dev      # http://localhost:3000
```

## Deploy Vercel

```bash
npm run build    # tạo thư mục ./out (static)
```

- Đẩy code lên GitHub → import vào Vercel (framework: Next.js). Không cần biến env, không cần server.
- Hoặc: `vercel` CLI trong thư mục này → follow prompts (static export tự nhận diện).

## Kiến trúc

```
src/lib/
  xml.ts        # adapter DOM: native DOMParser (browser) / @xmldom (Node CLI)
  model.ts      # phân loại block (question/heading/answer/guideline/theory/table)
  parser.ts     # docx → model blocks + tham chiếu XML body
  blanker.ts    # tìm vùng "gọi là X" → chuyển thành run gạch chân (giữ rPr)
  converter.ts  # PHT: xóa khối đáp án (inGuideline + điểm dừng); Live: giữ nguyên + vùng trả lời
  pipeline.ts   # orchestrate: parse → convert → pack .docx (JSZip)
```

## Kiểm thử offline (CLI, không cần browser)

Scripts trong `scripts/` dùng cùng code path với UI:

```bash
npx tsx scripts/convert-cli.ts "input.docx" --outdir out/
```

Chạy trên 2 file thật (Bài 1 & Bài 13 kèm kho mẫu thủ công) đã xác nhận:
- PHT: giữ nguyên mạch câu hỏi (25/25), giữ ngữ liệu "Sử dụng thông tin...", giữ "Đáp án: A. ____" (mẫu điền), giữ DẶN DÒ; xóa sạch đáp án/hướng dẫn (0 rò rỉ).
- Live: giữ 100% nội dung + thêm vùng trả lời sau mỗi câu hỏi.

## Hạn chế MVP

- Chỉ nhận `.docx` (không `.doc` legacy).
- Quy tắc blank có chủ ý dè dặt (chỉ câu định nghĩa ngắn) — chưa nhận diện blank trong bảng số liệu.
- Chưa có preview trực quan trước tải; chưa batch nhiều file.