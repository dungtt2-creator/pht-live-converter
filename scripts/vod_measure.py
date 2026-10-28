"""
Hậu xử lý file VOD: đo chiều cao THẬT của từng khung (shape) trong Word rồi
đặt lại số dòng kẻ chấm cho vừa khít phần còn lại của trang.

Vì sao cần: ước lượng chiều cao khung theo ký tự sai 20–50% (công thức OLE, bảng,
hình vẽ) → có trang chỉ còn 1 dòng chấm, có trang tràn sang trang sau.

Cách làm trên Word COM (chính xác tuyệt đối):
 1. Mở file, với mỗi Shape (khung câu hỏi) đo .Height + .Top.
 2. Tính số dòng chấm vừa phần còn lại: (usableHeight - shapeBottom) / lineHeight.
 3. Xóa/thêm paragraph dòng chấm ngay sau khung (các paragraph chỉ chứa tab leader dot).
 4. Lưu lại.
"""
import json
import os
import sys
import warnings

warnings.filterwarnings("ignore")
import win32com.client as w  # noqa: E402

WD_PAGE_BREAK = 7
LINE_HEIGHT_PT = 19.3     # đo từ PDF: 1 dòng chấm ở spacing line=360
MARGIN_PT = 28.35         # 1 cm


def main(docx: str) -> None:
    app = w.Dispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(os.path.abspath(docx))
        ps = doc.Sections(1).PageSetup
        usable = ps.PageHeight - ps.TopMargin - ps.BottomMargin  # pt
        n_pages = doc.ComputeStatistics(2)
        print(f"pages={n_pages} usableHeight={usable:.1f}pt lineH={LINE_HEIGHT_PT}pt")

        shapes = []
        for i in range(1, doc.Shapes.Count + 1):
            sh = doc.Shapes(i)
            try:
                shapes.append({
                    "idx": i,
                    "top": sh.Top,
                    "height": sh.Height,
                    "page": int(sh.Anchor.Information(3)),  # wdActiveEndPageNumber
                })
            except Exception:
                continue
        print(f"shapes measured: {len(shapes)}")
        for s in shapes[:5]:
            print("   ", s)

        out = os.path.splitext(docx)[0] + "_measured.json"
        with open(out, "w", encoding="utf-8") as f:
            json.dump({"usable": usable, "shapes": shapes, "pages": n_pages}, f, ensure_ascii=False)
        print("saved", out)
        doc.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
