import React, { useCallback, useRef, useState } from "react";
import { motion } from "framer-motion";
import { Download, UploadCloud, CheckCircle2, Link as LinkIcon, Clipboard, ClipboardCheck, AlertTriangle } from "lucide-react";

type FailedItem = { sheet: string; name: string; url: string };

function classNames(...c: (string | false | null | undefined)[]) {
  return c.filter(Boolean).join(" ");
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<string | null>(null);
  const [failedList, setFailedList] = useState<FailedItem[]>([]);
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [dragOver, setDragOver] = useState(false);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) setFile(f);
  }, []);

  const onBrowse = () => inputRef.current?.click();

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setMsg(null);
    setFailedList([]);
    if (!file) {
      setMsg("Please choose an .xlsx file.");
      return;
    }
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setMsg("Only .xlsx files are supported.");
      return;
    }

    setBusy(true);
    try {
      const form = new FormData();
      form.append("file", file);

      const res = await fetch("/api/upload", { method: "POST", body: form });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || `Server error (${res.status})`);
      }

      const blob = await res.blob();

      // Failures header (base64 json)
      const failedB64 = res.headers.get("X-Failed-Json");
      if (failedB64) {
        const bin = Uint8Array.from(atob(failedB64), (c) => c.charCodeAt(0));
        const jsonStr = new TextDecoder().decode(bin);
        setFailedList(JSON.parse(jsonStr) as FailedItem[]);
      }

      // Suggested filename
      const suggested =
        (res.headers.get("Content-Disposition") || "").match(/filename\*?=(?:UTF-8''|")?([^";]*)/)?.[1] ||
        "images.zip";

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = decodeURIComponent(suggested);
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setMsg("All set! Your ZIP is downloading.");
    } catch (err: any) {
      setMsg(err.message || "Something went wrong.");
    } finally {
      setBusy(false);
    }
  }

  // group failed by sheet
  const failedBySheet = failedList.reduce<Record<string, FailedItem[]>>((acc, it) => {
    (acc[it.sheet] ||= []).push(it);
    return acc;
  }, {});

  const copySheet = async (sheet: string) => {
    const rows = failedBySheet[sheet].map((f) => `${f.name}\t${f.url}`).join("\n");
    await navigator.clipboard.writeText(rows);
    setCopiedKey(sheet);
    setTimeout(() => setCopiedKey(null), 1200);
  };

  return (
    <div className="px-6 py-10 sm:py-16">
      <div className="mx-auto max-w-5xl">
        {/* Hero */}
        <motion.div
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          className="text-center"
        >
          <h1 className="text-4xl sm:text-5xl font-extrabold tracking-tight bg-gradient-to-r from-brand-300 via-emerald-300 to-cyan-300 bg-clip-text text-transparent">
            XLSX → Images (ZIP)
          </h1>
          <p className="mt-3 text-slate-300/90 max-w-2xl mx-auto">
            Upload an <code className="font-mono">.xlsx</code>. We detect any column starting with <b>Final</b> on each
            sheet, fetch all images (in parallel), and return a beautifully organized ZIP:
            <span className="ml-1 font-mono">&lt;Workbook&gt;/&lt;Sheet&gt;/images</span>.
          </p>
        </motion.div>

        {/* Upload card */}
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.05, duration: 0.5 }}
          className="card mt-10 p-6 sm:p-8"
        >
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
          >
            <div className="drop-outer">
              <div
                className={classNames(
                  "rounded-xl p-8 sm:p-10 text-center transition",
                  dragOver ? "bg-white/15" : "bg-slate-900/60"
                )}
              >
                <UploadCloud className="mx-auto h-10 w-10 text-emerald-300" />
                <p className="mt-3 font-semibold">Drag & drop your .xlsx here</p>
                <p className="text-slate-300/80 text-sm">— or —</p>
                <div className="mt-4">
                  <button
                    type="button"
                    onClick={onBrowse}
                    className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-brand-500 to-emerald-600 px-5 py-2.5 font-semibold text-white shadow-lg shadow-emerald-900/25 hover:brightness-110 focus:outline-none focus:ring-2 focus:ring-emerald-300"
                  >
                    <Download className="h-5 w-5" />
                    Choose file
                  </button>
                </div>
                <input
                  ref={inputRef}
                  type="file"
                  accept=".xlsx"
                  className="sr-only"
                  onChange={(e) => setFile(e.target.files?.[0] || null)}
                />

                {file && (
                  <p className="mt-4 text-sm text-emerald-200/90">
                    Selected: <span className="font-medium">{file.name}</span>
                  </p>
                )}
              </div>
            </div>
          </div>

          <form onSubmit={handleSubmit} className="mt-6 flex flex-col sm:flex-row items-center gap-3">
            <button
              type="submit"
              disabled={busy || !file}
              className={classNames(
                "w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-xl px-5 py-2.5 font-semibold shadow-lg focus:outline-none focus:ring-2",
                busy || !file
                  ? "bg-slate-600 cursor-not-allowed text-slate-200"
                  : "bg-gradient-to-r from-emerald-500 to-brand-600 text-white shadow-emerald-900/25 hover:brightness-110 focus:ring-emerald-300"
              )}
            >
              {busy ? "Processing…" : "Create ZIP"}
            </button>

            {msg && (
              <span className="text-sm text-slate-200/90">
                {msg.includes("error") ? (
                  <span className="text-rose-300">{msg}</span>
                ) : (
                  <span className="inline-flex items-center gap-1 text-emerald-300">
                    <CheckCircle2 className="h-4 w-4" /> {msg}
                  </span>
                )}
              </span>
            )}
          </form>

          {/* Indeterminate progress bar */}
          {busy && (
            <div className="relative mt-6 h-2 w-full overflow-hidden rounded-full bg-slate-700/60 progress" />
          )}
        </motion.div>

        {/* Failed section */}
        {failedList.length > 0 && (
          <motion.div
            initial={{ opacity: 0, y: 16 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.4 }}
            className="card mt-10 p-6 sm:p-8"
          >
            <div className="flex items-center gap-2 mb-4">
              <AlertTriangle className="h-5 w-5 text-amber-300" />
              <h2 className="text-lg font-semibold">
                Failed to download ({failedList.length})
              </h2>
            </div>

            <div className="space-y-6">
              {Object.entries(failedBySheet).map(([sheet, items]) => (
                <div key={sheet} className="rounded-xl border border-white/10 bg-white/5 p-4">
                  <div className="flex items-center justify-between gap-3">
                    <h3 className="font-semibold text-slate-100">
                      Sheet: <span className="text-emerald-300">{sheet}</span> — {items.length} item(s)
                    </h3>
                    <button
                      onClick={() => copySheet(sheet)}
                      className="inline-flex items-center gap-1.5 rounded-md bg-white/10 px-3 py-1.5 text-sm text-slate-100 hover:bg-white/15"
                    >
                      {copiedKey === sheet ? (
                        <>
                          <ClipboardCheck className="h-4 w-4" /> Copied
                        </>
                      ) : (
                        <>
                          <Clipboard className="h-4 w-4" /> Copy rows
                        </>
                      )}
                    </button>
                  </div>

                  <ul className="mt-3 space-y-2">
                    {items.map((it, i) => (
                      <li key={i} className="flex items-center gap-2 text-sm text-slate-200">
                        <LinkIcon className="h-4 w-4 text-rose-300 shrink-0" />
                        <a
                          href={it.url}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="underline decoration-rose-300/50 hover:decoration-rose-200"
                          title={it.url}
                        >
                          {it.name}
                        </a>
                        <span className="text-slate-400">—</span>
                        <span className="text-slate-300">sheet: {it.sheet}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>
          </motion.div>
        )}

        {/* Footer */}
        <div className="mt-10 text-center text-slate-400/80 text-xs">
          Built with FastAPI + React • Parallel downloads • Smart ZIP layout
        </div>
      </div>
    </div>
  );
}
