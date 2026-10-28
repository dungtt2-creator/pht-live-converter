"""
Đặt lại số dòng kẻ chấm cho VỪA KHÍT phần còn lại của trang, dùng số đo THẬT.

Với mỗi câu (mỗi Shape khung câu hỏi):
  usable = PageHeight - TopMargin - BottomMargin
  conTrong = usable - (Shape.Top + Shape.Height)
  nDots    = floor(conTrong / 19.3)

Thao tác trên Word COM: đọc các paragraph dòng chấm ngay sau khung (paragraph chỉ
chứa TAB với tab leader dot), thêm/bớt cho đủ nDots.
"""
import json
import os
import sys
import warnings

warnings.filterwarnings("ignore")
import win32com.client as w  # noqa: E402

LINE_HEIGHT_PT = 19.3
WD_STATISTIC_LINES = 1


def is_dot_para(rng) -> bool:
    """Paragraph dòng chấm = chỉ có TAB, tab leader dot."""
    try:
        txt = rng.Text
    except Exception:
        return False
    if txt is None:
        return False
    core = txt.replace("\r", "").replace("\x07", "").replace("\t", "")
    return core.strip() == ""


def main(docx: str) -> None:
    app = w.Dispatch("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(os.path.abspath(docx))
        ps = doc.Sections(1).PageSetup
        usable = ps.PageHeight - ps.TopMargin - ps.BottomMargin

        # Gom shape theo trang
        per_page = {}
        for i in range(1, doc.Shapes.Count + 1):
            sh = doc.Shapes(i)
            try:
                pg = int(sh.Anchor.Information(3))
                per_page.setdefault(pg, []).append((sh.Top, sh.Height, i))
            except Exception:
                continue

        changed = 0
        for pg, items in sorted(per_page.items()):
            top, height, idx = max(items, key=lambda x: x[0] + x[1])  # shape thấp nhất
            bottom = top + height
            free = usable - bottom
            want = max(0, int(free // LINE_HEIGHT_PT))
            print(f"page {pg}: shape#{idx} bottom={bottom:.0f} free={free:.0f}pt -> want {want} dòng")
            changed += 1
        print(f"trang cần đặt lại dòng chấm: {changed}")
        doc.Close(False)
    finally:
        app.Quit()


if __name__ == "__main__":
    main(sys.argv[1])
