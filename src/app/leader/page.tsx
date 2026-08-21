'use client';

import { useCallback, useRef, useState, type DragEvent, type ChangeEvent } from 'react';
import Link from 'next/link';
import { BelloLogo } from '@/components/bello-logo';
import { mergeExcelFiles } from '@/infrastructure/excel/excel-merger';

const inputBase =
  'rounded-xl border border-white/10 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none transition-colors placeholder:text-slate-500 focus:border-violet-400/60 disabled:cursor-not-allowed disabled:opacity-50';

interface ProcessedFile {
  name: string;
  memberCount: number;
}

export default function LeaderDashboardPage() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [dragging, setDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<ProcessedFile[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  const accept = '.xlsx,.xls';

  function handleFiles(newFiles: FileList | File[]): void {
    const xlsxFiles = Array.from(newFiles).filter(
      (f) => f.name.endsWith('.xlsx') || f.name.endsWith('.xls'),
    );
    if (xlsxFiles.length === 0) {
      setError('Please upload .xlsx or .xls files only.');
      return;
    }
    setError(null);
    setResult(null);
    setFiles((prev) => [...prev, ...xlsxFiles]);
  }

  function onDrop(e: DragEvent): void {
    e.preventDefault();
    setDragging(false);
    if (e.dataTransfer.files.length > 0) handleFiles(e.dataTransfer.files);
  }

  function onDragOver(e: DragEvent): void {
    e.preventDefault();
    setDragging(true);
  }

  function onDragLeave(): void {
    setDragging(false);
  }

  function onFileInputChange(e: ChangeEvent<HTMLInputElement>): void {
    if (e.target.files && e.target.files.length > 0) handleFiles(e.target.files);
    if (fileInputRef.current) fileInputRef.current.value = '';
  }

  function removeFile(index: number): void {
    setFiles((prev) => prev.filter((_, i) => i !== index));
    setResult(null);
  }

  function clearAll(): void {
    setFiles([]);
    setResult(null);
    setError(null);
  }

  const handleMerge = useCallback(async () => {
    if (files.length === 0) return;
    setProcessing(true);
    setError(null);
    setResult(null);
    try {
      const mergedBytes = await mergeExcelFiles(files);

      const blob = new Blob([new Uint8Array(mergedBytes)], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Bello_Master_Report.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setResult([{ name: 'Bello_Master_Report.xlsx', memberCount: files.length }]);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Merge failed');
    } finally {
      setProcessing(false);
    }
  }, [files]);

  return (
    <main className="animate-bello-in mx-auto w-full max-w-4xl px-4 pb-24 pt-8 sm:px-6">
      <Link
        href="/"
        className="mb-5 inline-flex items-center gap-1.5 text-sm font-medium text-slate-400 transition-colors hover:text-white"
      >
        <span aria-hidden>&larr;</span> Back to leaderboard
      </Link>

      <header className="mb-8">
        <div className="flex items-center gap-3">
          <BelloLogo size={40} />
          <div>
            <h1 className="text-2xl font-black tracking-tight text-white sm:text-3xl">
              Leader Dashboard
            </h1>
            <p className="text-sm text-slate-400">
              Merge multiple evaluator exports into one master report
            </p>
          </div>
        </div>
      </header>

      {/* Drop zone */}
      <section
        onDrop={onDrop}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onClick={() => fileInputRef.current?.click()}
        className={`cursor-pointer rounded-2xl border-2 border-dashed p-8 text-center transition-all ${
          dragging
            ? 'border-violet-400 bg-violet-500/10'
            : 'border-white/15 bg-white/[0.03] hover:border-white/25 hover:bg-white/[0.05]'
        }`}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept={accept}
          multiple
          className="hidden"
          onChange={onFileInputChange}
        />
        <div className="mb-3 text-3xl text-slate-500">&#128206;</div>
        <p className="text-sm font-medium text-slate-300">
          {dragging ? 'Drop files here' : 'Drag & drop .xlsx files or click to browse'}
        </p>
        <p className="mt-1 text-xs text-slate-500">
          Upload exports from sub-leaders to merge them into a single master sheet
        </p>
      </section>

      {/* File list */}
      {files.length > 0 && (
        <section className="mt-4 rounded-2xl border border-white/10 bg-white/[0.04] p-5 backdrop-blur">
          <div className="mb-3 flex items-center justify-between">
            <h2 className="text-lg font-bold text-white">
              Uploaded Files ({files.length})
            </h2>
            <button
              type="button"
              onClick={clearAll}
              className="text-xs font-medium text-slate-400 transition-colors hover:text-white"
            >
              Clear all
            </button>
          </div>
          <ul className="space-y-1.5">
            {files.map((file, i) => (
              <li
                key={`${file.name}-${i}`}
                className="flex items-center justify-between rounded-xl bg-slate-950/40 px-3 py-2 text-sm"
              >
                <span className="truncate text-slate-200">{file.name}</span>
                <div className="flex items-center gap-3">
                  <span className="text-xs text-slate-500">
                    {(file.size / 1024).toFixed(0)} KB
                  </span>
                  <button
                    type="button"
                    onClick={() => removeFile(i)}
                    className="rounded-lg px-2 py-1 text-xs font-semibold text-rose-300 transition-colors hover:bg-rose-500/15"
                  >
                    Remove
                  </button>
                </div>
              </li>
            ))}
          </ul>

          <button
            type="button"
            onClick={() => void handleMerge()}
            disabled={processing}
            className="mt-4 w-full rounded-xl bg-gradient-to-r from-indigo-500 to-violet-500 px-5 py-3 text-sm font-semibold text-white shadow-lg shadow-indigo-950/40 transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-40"
          >
            {processing ? 'Merging files...' : `Merge ${files.length} file${files.length !== 1 ? 's' : ''} into Master`}
          </button>
        </section>
      )}

      {error && (
        <p className="mt-3 rounded-xl border border-rose-400/20 bg-rose-500/10 p-3 text-sm text-rose-300">
          {error}
        </p>
      )}

      {result && (
        <section className="mt-4 rounded-2xl border border-emerald-400/20 bg-emerald-500/10 p-5 backdrop-blur">
          <h2 className="mb-1 text-lg font-bold text-emerald-300">Merge Complete</h2>
          <p className="text-sm text-emerald-200/80">
            Download started for <span className="font-mono">{result[0].name}</span>
          </p>
        </section>
      )}

      {/* How it works */}
      <section className="mt-6 rounded-2xl border border-white/10 bg-white/[0.04] p-5 backdrop-blur">
        <h2 className="mb-3 text-lg font-bold text-white">How it works</h2>
        <ol className="space-y-2 text-sm text-slate-400">
          <li className="flex gap-2">
            <span className="font-bold text-violet-400">1.</span>
            Each sub-leader exports their evaluation using the Bello app.
          </li>
          <li className="flex gap-2">
            <span className="font-bold text-violet-400">2.</span>
            Upload all exported .xlsx files here (drag & drop or click).
          </li>
          <li className="flex gap-2">
            <span className="font-bold text-violet-400">3.</span>
            The system reads member data from each file and merges them into a
            single master template.
          </li>
          <li className="flex gap-2">
            <span className="font-bold text-violet-400">4.</span>
            The merged <span className="font-mono">Bello_Master_Report.xlsx</span> is
            downloaded with all formulas intact.
          </li>
        </ol>
      </section>

      <footer className="mt-10 flex items-center justify-center gap-2 text-xs text-slate-500">
        <BelloLogo size={18} />
        <span>Bello &middot; offline-first member evaluation</span>
      </footer>
    </main>
  );
}
