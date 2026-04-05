"use client";

import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import { FileDropZone } from "@/components/FileDropZone";
import { DataGrid } from "@/components/DataGrid";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import {
  Download,
  RotateCcw,
  FileSpreadsheet,
  Loader2,
  Copy,
  Check,
} from "lucide-react";
import {
  validateFileName,
  processWorkbook,
  exportToXlsx,
  TransposedRow,
} from "@/lib/wayfairTransposer";

type AppState = "Ready" | "Processing" | "Reviewing";

export default function Home() {
  const [state, setState] = useState<AppState>("Ready");
  const [data, setData] = useState<TransposedRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [fileMeta, setFileMeta] = useState<{ size: number; lastModified: number } | null>(null);
  const [copied, setCopied] = useState(false);

  const handleFile = useCallback(async (file: File) => {
    setState("Processing");
    setFileName(file.name);
    setFileMeta({ size: file.size, lastModified: file.lastModified });

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const result = processWorkbook(workbook);

      setData(result);
      setState("Reviewing");

      const dupes = result.filter((r) => r.validation_status === "Duplicate_Key").length;
      const errors = result.filter((r) => r.validation_status === "Corrupted_Value").length;

      if (dupes > 0 || errors > 0) {
        toast.warning(
          `Phát hiện ${dupes} dòng trùng lặp và ${errors} giá trị lỗi Excel.`
        );
      } else {
        toast.success(`Transpose thành công! ${result.length} dòng dữ liệu.`);
      }
    } catch (err) {
      console.error("Processing error:", err);
      toast.error(err instanceof Error ? err.message : "Lỗi xử lý file.");
      setState("Ready");
    }
  }, []);

  const handleExport = useCallback(() => {
    exportToXlsx(data, fileName);
    toast.success("Đã tải file thành công!");
  }, [data, fileName]);

  const handleCopyToClipboard = useCallback(() => {
    const rows = data.map((r) => [r.sku, r.base_heading, r.attribute_heading, r.value].join("\t"));
    const tsv = rows.join("\n");

    const onSuccess = () => {
      setCopied(true);
      toast.success("Đã copy dữ liệu cho Google Sheets!");
      alert("Đã copy toàn bộ dữ liệu thành công!");
      setTimeout(() => setCopied(false), 2000);
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(tsv).then(onSuccess).catch(() => {
        toast.error("Không thể copy vào clipboard.");
        alert("Lỗi: Không thể copy vào clipboard.");
      });
    } else {
      try {
        const textArea = document.createElement("textarea");
        textArea.value = tsv;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand("copy");
        document.body.removeChild(textArea);
        onSuccess();
      } catch (err) {
        toast.error("Trình duyệt không hỗ trợ copy.");
        alert("Lỗi: Trình duyệt không hỗ trợ copy.");
      }
    }
  }, [data]);

  const handleReset = useCallback(() => {
    setData([]);
    setFileName("");
    setFileMeta(null);
    setState("Ready");
  }, []);

  return (
    <div className="h-screen overflow-hidden bg-gradient-to-br from-pink-300 via-purple-300 to-blue-300 dark:from-pink-900 dark:via-purple-900 dark:to-blue-900 text-foreground flex flex-col">
      {/* Header */}
      <header className="border-b bg-card/80 backdrop-blur-md shrink-0 z-50">
        <div className="w-full max-w-[1440px] mx-auto flex flex-col md:flex-row items-center justify-between px-4 md:px-8 py-4 gap-4">
          <div className="flex items-center w-full md:w-auto justify-center md:justify-start">
            <div className="text-center md:text-left bg-gradient-to-r from-fuchsia-600 to-purple-700 text-white px-6 py-2 rounded-2xl shadow-sm border border-fuchsia-500/30">
              <h1 className="text-lg font-extrabold tracking-tight">
                PGRB Description Transposer
              </h1>
              <p className="text-[10px] text-fuchsia-200 font-mono uppercase tracking-widest mt-0.5 opacity-90">
                PGRB Data Isolation Engine
              </p>
            </div>
          </div>

          {state === "Reviewing" && (
            <div className="flex flex-wrap items-center justify-center gap-2 w-full md:w-auto">
              <Button variant="outline" size="sm" onClick={handleReset} className="border-2 border-rose-200 text-rose-700 hover:bg-rose-200 hover:border-rose-300 transition-all font-bold rounded-xl bg-rose-100 shadow-sm">
                <RotateCcw className="h-4 w-4 mr-1.5" />
                Làm lại
              </Button>
              <Button variant="outline" size="sm" onClick={handleCopyToClipboard} className="border-2 border-indigo-200 text-indigo-700 hover:bg-indigo-200 hover:border-indigo-300 transition-all font-bold rounded-xl bg-indigo-100 shadow-sm">
                {copied ? <Check className="h-4 w-4 mr-1.5 text-emerald-600" /> : <Copy className="h-4 w-4 mr-1.5" />}
                Copy for Sheets
              </Button>
              <Button size="sm" onClick={handleExport} className="bg-gradient-to-r from-emerald-400 to-green-500 hover:from-emerald-500 hover:to-green-600 text-white font-bold border-0 shadow-lg shadow-emerald-500/30 transition-all rounded-xl px-5">
                <Download className="h-4 w-4 mr-1.5" />
                Xuất Excel
              </Button>
            </div>
          )}
        </div>
      </header>

      {/* Main */}
      <main className="w-full max-w-[1440px] px-4 md:px-8 pt-6 pb-24 flex-1 flex flex-col min-h-0 mx-auto">
        {state === "Ready" && (
          <div className="flex-1 flex flex-col items-center justify-center pb-12 w-full">
            <div className="max-w-4xl w-full mx-auto">
              <div className="flex items-center justify-center w-full">
                 <div className="w-full">
                    <FileDropZone onFile={handleFile} />
                 </div>
              </div>
              <div className="max-w-3xl mx-auto w-full bg-card/60 backdrop-blur-sm p-6 rounded-xl border shadow-sm mt-8">
            <h3 className="font-semibold mb-3">Hướng dẫn sử dụng</h3>
            <ul className="text-sm space-y-2 text-muted-foreground list-disc pl-5">
              <li>Phải là file <code className="bg-muted px-1.5 py-0.5 rounded">.xlsx</code> xuất từ Wayfair Portal.</li>
              <li>Hệ thống tự động rà quét và chỉ lấy dữ liệu từ các dòng có chứa <strong className="text-purple-700 font-bold">Wayfair SKU</strong>.</li>
              <li>Tự động bỏ qua thẻ Store, helper text và các giá trị Feature trống (trừ Marketing Copy).</li>
              <li>Cảnh báo đỏ (Lỗi) sẽ hiện lên nếu phát hiện giá trị bị hỏng trong Excel (#N/A, #REF!) hoặc trùng lặp dữ liệu.</li>
            </ul>
          </div>
            </div>
          </div>
        )}

        {state === "Processing" && (
          <div className="flex-1 flex flex-col items-center justify-center pb-20 gap-4">
            <Loader2 className="h-10 w-10 animate-spin text-primary" />
            <p className="text-muted-foreground font-medium text-center">
              Đang phân tích <span className="text-foreground font-bold">{fileName}</span>...
              <br />
              <span className="text-xs opacity-50">Đang khởi tạo trạng thái mới 100%...</span>
            </p>
          </div>
        )}

        {state === "Reviewing" && (
          <div className="w-full flex flex-col max-h-full min-h-0 space-y-4">
            
            {/* File Info Box */}
            <div className="flex flex-wrap items-center justify-between gap-4 bg-white/80 backdrop-blur-md p-4 rounded-xl border border-white/50 shadow-sm shrink-0">
              <div className="flex items-center gap-4">
                <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-blue-400 to-blue-600 shadow-inner border border-blue-300">
                  <FileSpreadsheet className="h-6 w-6 text-white" />
                </div>
                <div className="flex flex-col">
                  <span className="text-lg font-extrabold text-slate-800 leading-tight">{fileName}</span>
                  <span className="text-xs text-slate-500 font-semibold tracking-wide">
                    SIZE: {fileMeta ? (fileMeta.size / 1024).toFixed(1) : 0} KB • SAVED: {fileMeta ? new Date(fileMeta.lastModified).toLocaleString() : ""}
                  </span>
                </div>
              </div>
              <Badge variant="outline" className="bg-emerald-100 text-emerald-800 border-emerald-200 shadow-sm px-3 py-1.5 font-bold text-xs flex items-center gap-2 rounded-lg">
                <div className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse shadow-[0_0_8px_rgba(16,185,129,0.8)]" />
                Isolating Wayfair SKU...
              </Badge>
            </div>

            <div className="flex flex-col min-h-0 space-y-4 pb-10">
              <div className="flex items-center gap-3 shrink-0 mt-3 mb-1">
                <div className="h-8 w-2 bg-gradient-to-b from-fuchsia-500 to-purple-600 rounded-full shadow-sm" />
                <h2 className="text-2xl font-black uppercase tracking-widest text-slate-900 drop-shadow-sm">Kết quả trích xuất</h2>
              </div>
              <div className="min-h-0">
                <DataGrid data={data} />
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
