/**
 * ĐO BỀ RỘNG CHỮ bằng canvas trình duyệt (chỉ chạy ở client).
 *
 * Vì sao cần: ước lượng chiều cao khung câu hỏi bằng cách ĐẾM KÝ TỰ sai ±75pt
 * (công thức OLE cao hơn chữ thường, chữ Việt có dấu rộng hơn chữ Latinh). Sai số
 * này làm số dòng kẻ chấm bị tính sai → trang chỉ còn 1–2 dòng kẻ, hoặc tràn trang.
 *
 * Cách làm: dùng canvas 2D measureText với đúng font Times New Roman 14pt để tính
 * số dòng THẬT của từng đoạn văn, từ đó ra chiều cao khung chính xác hơn nhiều.
 * Không có canvas (SSR/test Node) → trả null và code dùng ước lượng cũ.
 */

export interface CanvasMeasurer {
  /** Bề rộng (pt) của 1 chuỗi ở TNR 14pt */
  widthPt(text: string): number;
  /** Số dòng 1 đoạn văn chiếm khi wrap trong bề rộng `boxPt` */
  lines(text: string, boxPt: number): number;
}

let cached: CanvasMeasurer | null | undefined;

export function getCanvasMeasurer(): CanvasMeasurer | null {
  if (cached !== undefined) return cached;
  try {
    if (typeof document === "undefined") {
      cached = null;
      return cached;
    }
    const cv = document.createElement("canvas");
    const ctx = cv.getContext("2d");
    if (!ctx) {
      cached = null;
      return cached;
    }
    // 14pt = 18.667px ở 96dpi; đo theo px rồi đổi ra pt (×0.75)
    const fontPx = 14 * (96 / 72);
    ctx.font = `${fontPx.toFixed(2)}px "Times New Roman", "Liberation Serif", serif`;
    const m: CanvasMeasurer = {
      widthPt(t: string): number {
        return ctx.measureText(t).width * 0.75;
      },
      lines(t: string, boxPt: number): number {
        if (!t) return 0;
        const words = t.split(/\s+/).filter(Boolean);
        if (!words.length) return 1;
        let n = 1;
        let cur = "";
        for (const w of words) {
          const trial = cur ? `${cur} ${w}` : w;
          if (m.widthPt(trial) > boxPt && cur) {
            n++;
            cur = w;
          } else {
            cur = trial;
          }
        }
        return n;
      },
    };
    cached = m;
    return cached;
  } catch {
    cached = null;
    return cached;
  }
}
