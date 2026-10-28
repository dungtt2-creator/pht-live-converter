"""
Đo chiều cao THẬT của từng khung câu hỏi từ PDF render (chính xác tuyệt đối),
rồi xuất báo cáo JSON để lần render sau đặt số dòng kẻ chấm vừa khít.

Dùng: python scripts/vod_measure_pdf.py <file.pdf> <file.docx.json>
Ghi: {"pages":[{"page":N,"boxBottom":Y,"freeToBottom":F,"wantDots":n}], "lineH":19.3}
"""
import json
import sys
import pymupdf

LINE_H = 19.3  # pt/dòng chấm (đo từ bản in)


def main(pdf: str, out_json: str) -> None:
    doc = pymupdf.open(pdf)
    usable_bottom = doc[0].rect.height - 28.35  # lề dưới 1 cm
    rows = []
    for i, pg in enumerate(doc):
        boxes = [d["rect"] for d in pg.get_drawings()
                 if d["rect"].width > 250 and "c" in {k[0] for k in d["items"]}]
        if not boxes:
            rows.append({"page": i + 1, "boxBottom": None, "freeToBottom": None, "wantDots": None})
            continue
        box_bottom = max(b.y1 for b in boxes)
        free = usable_bottom - box_bottom
        rows.append({
            "page": i + 1,
            "boxBottom": round(box_bottom, 1),
            "freeToBottom": round(free, 1),
            "wantDots": max(0, int(free // LINE_H)),
        })
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({"lineH": LINE_H, "usableBottom": round(usable_bottom, 1), "pages": rows},
                  f, ensure_ascii=False, indent=1)
    print(f"{len(rows)} trang -> {out_json}")
    bad = [r for r in rows if r["wantDots"] is not None and r["wantDots"] < 3]
    print("trang it dong cham:", [(r["page"], r["wantDots"]) for r in bad][:8])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
