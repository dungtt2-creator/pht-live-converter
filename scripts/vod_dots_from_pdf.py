"""
Lượt 2: tính số dòng kẻ chấm VỪA KHÍT cho từng câu.

ĐIỂM QUAN TRỌNG: đo từ VỊ TRÍ DÒNG CHẤM ĐẦU TIÊN (y của dòng chấm đầu), KHÔNG đo
theo đáy khung. Đáy khung là chiều cao do code ước lượng nên lệch 150–200pt; còn
vị trí dòng chấm đầu = nơi nội dung câu hỏi thật sự kết thúc → ỔN ĐỊNH giữa các lượt.

    nDots = floor((usableBottom - firstDotY - SAFETY) / LINE_H)

Dùng: python scripts/vod_dots_from_pdf.py <luot1.pdf> <out.json>
"""
import json
import sys
import pymupdf

LINE_H = 19.3   # pt cho 1 dòng chấm (đo thực tế từ bản in)
SAFETY = 10.0   # chừa 10pt cho sai số làm tròn


def main(pdf_path: str, out_json: str) -> None:
    doc = pymupdf.open(pdf_path)
    usable_bottom = doc[0].rect.height - 28.35   # lề dưới 1 cm
    dots: list[int] = []
    detail = []
    for i, pg in enumerate(doc):
        if "[HOCMAI]" not in pg.get_text():
            continue   # trang nối tiếp (dòng chấm tràn) → không tính
        ys = [l["bbox"][1]
              for b in pg.get_text("dict")["blocks"]
              for l in b.get("lines", [])
              if "...." in "".join(s["text"] for s in l["spans"])]
        if not ys:
            dots.append(8)
            detail.append((i + 1, "no-dots", 8))
            continue
        first = min(ys)
        free = usable_bottom - first - SAFETY
        n = max(1, int(free // LINE_H))
        dots.append(n)
        detail.append((i + 1, round(first), n))
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({"dots": dots, "lineH": LINE_H}, f, ensure_ascii=False, indent=1)
    print(f"{len(dots)} câu -> {out_json} | dòng chấm {min(dots)}..{max(dots)}")
    print("  ví dụ (trang, yDòngĐầu, nDòng):", detail[:6])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
