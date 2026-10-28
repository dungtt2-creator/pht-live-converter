"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { convertDocx, outputFileName, outputFileNameKeepCase } from "@/lib/pipeline";
import type { ConvertResult } from "@/lib/converter";

type Phase = "idle" | "converting" | "done" | "error";

interface Payload {
  name: string;
  /** null ở chế độ VOD (chỉ xuất Tài liệu dạy học) */
  pht: ArrayBuffer | null;
  live: ArrayBuffer;
  phtResult: ConvertResult | null;
  liveResult: ConvertResult | null;
  questionCount: number;
  theoryCount: number;
  ms: number;
}

function download(buf: ArrayBuffer, name: string) {
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

export default function Converter() {
  const [phase, setPhase] = useState<Phase>("idle");
    const [err, setErr] = useState<string>("");
    const [payload, setPayload] = useState<Payload | null>(null);
    const inputRef = useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = useState(false);
  /**
   * Chế độ xuất bản:
   *  - "vod"  : Tài liệu dạy học ghi hình (Template VOD — khổ 32 × 18 cm, TNR 14,
   *             câu hỏi | dòng chấm). Tên file giữ nguyên tên đề + hậu tố _Live.
   *  - "exam" : Bản Live/PHT theo mẫu kỳ thi cũ (khung wave).
   */
  const [layout, setLayout] = useState<"vod" | "exam">("vod");

  const run = useCallback(async (file: File) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setErr("Vui lòng chọn file .docx (Word). File .doc cũ không hỗ trợ ở MVP.");
      setPhase("error");
      return;
    }
    setPhase("converting");
    setErr("");
    setPayload(null);
    try {
      const ab = await file.arrayBuffer();
      const t0 = performance.now();
      const out = await convertDocx(ab, {
              modes: layout === "vod" ? ["live"] : ["pht", "live"],
              fileName: file.name,
              useFileTemplate: true, // dùng file template thật (TSA default)
              layout,
            });
      const ms = Math.round(performance.now() - t0);
      const base = file.name.replace(/\.docx?$/i, "");
      setPayload({
        name: base,
        pht: out.pht,
        live: out.live!,
        phtResult: out.phtResult,
        liveResult: out.liveResult,
        questionCount: out.questionCount,
        theoryCount: out.theoryParagraphs,
        ms,
      });
      setPhase("done");
    } catch (e) {
      console.error(e);
      setErr(e instanceof Error ? e.message : "Lỗi không xác định khi xử lý file.");
      setPhase("error");
    }
  }, [layout]);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragActive(false);
      const f = e.dataTransfer.files?.[0];
      if (f) void run(f);
    },
    [run],
  );

  const stats = useMemo(() => {
    if (!payload) return null;
    return {
      q: payload.questionCount,
      blanks: payload.phtResult?.blanksCreated ?? 0,
      removed: payload.phtResult?.answersRemoved ?? 0,
      tags: payload.phtResult?.hocmaiTagsAdded ?? 0,
    };
  }, [payload]);

  return (
    <div className="mx-auto max-w-3xl px-4 py-10">
      <header className="mb-6 text-center">
              <h1 className="text-3xl font-bold text-slate-900">
                Chuyển file đề Word → Tài liệu Live cho giáo viên
              </h1>
              <p className="mt-3 text-sm leading-relaxed text-slate-500">
                Upload file <b>.docx</b> đề của giáo viên: hệ thống chuyển sang{" "}
                <b>đúng Template VOD</b> — khổ ngang 32 × 18 cm (16:9), Times New Roman 14,
                mỗi câu là một khung câu hỏi kèm dòng chấm để giáo viên chữa bài khi ghi
                hình. Xử lý hoàn toàn trên máy bạn — không gửi dữ liệu lên server.
              </p>
            </header>

      <div className="mb-5 rounded-2xl border border-slate-200 bg-white p-4">
        <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">
          Chế độ xuất bản
        </p>
        <div className="grid gap-2 sm:grid-cols-2">
          <button
            type="button"
            onClick={() => setLayout("vod")}
            className={`rounded-xl border p-3 text-left transition ${
              layout === "vod"
                ? "border-sky-500 bg-sky-50 ring-1 ring-sky-300"
                : "border-slate-200 hover:border-sky-300"
            }`}
          >
            <span className="block text-sm font-semibold text-slate-900">
              Template VOD (ghi hình)
            </span>
            <span className="mt-1 block text-xs leading-relaxed text-slate-500">
              32 × 18 cm · Times New Roman 14 · câu hỏi + dòng chấm. Tên file:
              &lt;tên đề&gt;_Live.docx
            </span>
          </button>
          <button
            type="button"
            onClick={() => setLayout("exam")}
            className={`rounded-xl border p-3 text-left transition ${
              layout === "exam"
                ? "border-violet-500 bg-violet-50 ring-1 ring-violet-300"
                : "border-slate-200 hover:border-violet-300"
            }`}
          >
            <span className="block text-sm font-semibold text-slate-900">
              Bản kỳ thi (PHT + Live cũ)
            </span>
            <span className="mt-1 block text-xs leading-relaxed text-slate-500">
              Khung wave theo mẫu kỳ thi, xuất cả Phiếu học tập và Tài liệu Live.
            </span>
          </button>
        </div>
      </div>

            <div
        onDragOver={(e) => { e.preventDefault(); setDragActive(true); }}
        onDragLeave={() => setDragActive(false)}
        onDrop={onDrop}
        onClick={() => inputRef.current?.click()}
        className={`cursor-pointer rounded-2xl border-2 border-dashed p-10 text-center transition ${
          dragActive
            ? "border-sky-500 bg-sky-50"
            : "border-slate-300 bg-white hover:border-sky-400 hover:bg-sky-50/50"
        }`}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".docx"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) void run(f);
          }}
        />
        <div className="mx-auto mb-3 flex h-14 w-14 items-center justify-center rounded-full bg-sky-100 text-2xl">
          📄
        </div>
        <p className="font-medium text-slate-700">
          {phase === "converting" ? "Đang xử lý…" : "Kéo thả file .docx vào đây hoặc bấm để chọn"}
        </p>
        <p className="mt-1 text-xs text-slate-400">
          Định dạng khuyến nghị: tài liệu ôn thi có mục tiêu bài học, phần lý thuyết đánh số
          I/II/III, câu hỏi “Câu N.” và mục Đáp án / Hướng dẫn.
        </p>
      </div>

      {phase === "converting" && (
        <div className="mt-6 rounded-xl border border-sky-200 bg-sky-50 p-4 text-sm text-sky-800">
          <div className="mb-2 h-1.5 animate-pulse rounded-full bg-sky-300" />
          Đang phân tích cấu trúc, tạo chỗ trống kiến thức và đóng gói 2 file…
        </div>
      )}

      {phase === "error" && (
        <div className="mt-6 rounded-xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-700">
          ⚠️ {err}
        </div>
      )}

      {phase === "done" && payload && stats && (
        <div className="mt-6 space-y-4">
          <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-4 text-sm text-emerald-800">
            ✅ Xử lý xong <b>{payload.name}</b> trong {payload.ms} ms —{" "}
            {stats.q} câu hỏi phát hiện
            {layout === "vod"
              ? `, ${payload.liveResult?.answersRemoved ?? 0} khối đáp án/hướng dẫn đã loại`
              : `, ${stats.blanks} chỗ trống tạo, ${stats.removed} khối đáp án/hướng dẫn đã loại khỏi phiếu`}
            {stats.tags > 0 &&
              `, đã tự động thêm tag [HOCMAI] cho ${stats.tags} câu chưa chuẩn hóa`}
            .
          </div>

          <div className={`grid gap-4 ${payload.pht ? "sm:grid-cols-2" : ""}`}>
            {payload.pht && (
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-base font-semibold text-slate-900">📝 Phiếu học tập</h3>
              <p className="mt-1 min-h-[3.5rem] text-xs leading-relaxed text-slate-500">
                Cho học sinh: giữ mục tiêu + hệ thống kiến thức (đã tạo chỗ trống bằng
                gạch chân), câu hỏi bài tập; bỏ phần đáp án &amp; hướng dẫn giải.
              </p>
              <button
                onClick={() => payload.pht && download(payload.pht, outputFileName(payload.name, "pht"))}
                className="mt-3 w-full rounded-lg bg-sky-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-sky-700"
              >
                ⬇ Tải PHT (.docx)
              </button>
            </div>
            )}

            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-base font-semibold text-slate-900">🎤 Tài liệu Live</h3>
              <p className="mt-1 min-h-[3.5rem] text-xs leading-relaxed text-slate-500">
                {layout === "vod"
                  ? "Template VOD: khổ 32 × 18 cm (16:9), Times New Roman 14, mỗi câu là khung câu hỏi kèm dòng chấm để giáo viên chữa bài khi ghi hình."
                  : "Cho giáo viên: giữ nguyên 100% kiến thức gốc, bổ sung vùng để giáo viên viết ở mỗi câu hỏi khi giảng dạy trực tiếp."}
              </p>
              <button
                onClick={() =>
                  download(
                    payload.live,
                    layout === "vod"
                      ? outputFileNameKeepCase(payload.name, "Live")
                      : outputFileName(payload.name, "live"),
                  )
                }
                className="mt-3 w-full rounded-lg bg-violet-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-violet-700"
              >
                ⬇ Tải Live (.docx)
              </button>
            </div>
          </div>
        </div>
      )}

      <footer className="mt-10 border-t border-slate-200 pt-4 text-center text-xs text-slate-400">
        MVP chạy 100% client-side (Next.js + docx/xml parsing) — sẵn sàng deploy lên Vercel.
        Không tự ý sửa nội dung kiến thức; chỉ cấu trúc lại hình thức phục vụ giảng dạy.
      </footer>
    </div>
  );
}