"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { convertDocx, outputFileName } from "@/lib/pipeline";
import type { ConvertResult } from "@/lib/converter";
import { TEMPLATES, DEFAULT_TEMPLATE } from "@/lib/template";
import type { TemplateId } from "@/lib/template";

type Phase = "idle" | "converting" | "done" | "error";

interface Payload {
  name: string;
  pht: ArrayBuffer;
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
  const [templateId, setTemplateId] = useState<TemplateId>(DEFAULT_TEMPLATE);
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = useState(false);

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
      const out = await convertDocx(ab, { modes: ["pht", "live"], fileName: file.name, templateId });
      const ms = Math.round(performance.now() - t0);
      const base = file.name.replace(/\.docx?$/i, "");
      setPayload({
        name: base,
        pht: out.pht!,
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
  }, []);

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
    };
  }, [payload]);

  return (
    <div className="mx-auto max-w-3xl px-4 py-10">
      <header className="mb-8 text-center">
        <p className="mb-2 text-xs font-semibold uppercase tracking-[0.2em] text-sky-600">
          TSA · V-ACT · GDPT 2018 · Ôn thi Đánh giá năng lực
        </p>
        <h1 className="text-3xl font-bold text-slate-900">
          Chuyển tài liệu Word → Phiếu học tập &amp; Tài liệu Live
        </h1>
        <p className="mt-3 text-sm leading-relaxed text-slate-500">
          Upload file <b>.docx</b> thô của giáo viên: hệ thống tự tách{" "}
          <b>Phiếu học tập</b> (kiến thức cô đọng + chỗ trống cho học sinh, bỏ đáp
          án/hướng dẫn) và <b>Tài liệu Live</b> (giữ nguyên 100% nội dung giáo viên,
          thêm vùng trả lời). Xử lý hoàn toàn trên máy bạn — không gửi dữ liệu lên server.
        </p>
      </header>

      {/* Chọn template kỳ thi */}
      <div className="mb-6">
        <p className="mb-2 text-center text-sm font-semibold text-slate-600">
          Chọn template kỳ thi / hệ thống:
        </p>
        <div className="grid grid-cols-2 gap-2 sm:grid-cols-5">
          {TEMPLATES.map((t) => (
            <button
              key={t.id}
              type="button"
              onClick={() => setTemplateId(t.id as TemplateId)}
              className={`rounded-xl border-2 px-3 py-2 text-center transition ${
                templateId === t.id
                  ? "border-transparent text-white shadow-md"
                  : "border-slate-200 bg-white text-slate-600 hover:border-slate-300"
              }`}
              style={
                templateId === t.id
                  ? { backgroundColor: `#${t.color}` }
                  : { borderColor: `#${t.color}` }
              }
              title={t.desc}
            >
              <span className="block text-sm font-bold">{t.label}</span>
            </button>
          ))}
        </div>
        <p className="mt-2 text-center text-xs text-slate-400">
          {TEMPLATES.find((t) => t.id === templateId)?.desc}
        </p>
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
            {stats.q} câu hỏi phát hiện, {stats.blanks} chỗ trống tạo, {stats.removed} khối
            đáp án/hướng dẫn đã loại khỏi phiếu.
          </div>

          <div className="grid gap-4 sm:grid-cols-2">
            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-base font-semibold text-slate-900">📝 Phiếu học tập</h3>
              <p className="mt-1 min-h-[3.5rem] text-xs leading-relaxed text-slate-500">
                Cho học sinh: giữ mục tiêu + hệ thống kiến thức (đã tạo chỗ trống bằng
                gạch chân), câu hỏi bài tập; bỏ phần đáp án &amp; hướng dẫn giải.
              </p>
              <button
                onClick={() => download(payload.pht, outputFileName(payload.name, "pht"))}
                className="mt-3 w-full rounded-lg bg-sky-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-sky-700"
              >
                ⬇ Tải PHT (.docx)
              </button>
            </div>

            <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <h3 className="text-base font-semibold text-slate-900">🎤 Tài liệu Live</h3>
              <p className="mt-1 min-h-[3.5rem] text-xs leading-relaxed text-slate-500">
                Cho giáo viên: giữ nguyên 100% kiến thức &amp; đáp án gốc, bổ sung vùng
                trả lời sau mỗi câu hỏi để dạy trực tiếp.
              </p>
              <button
                onClick={() => download(payload.live, outputFileName(payload.name, "live"))}
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