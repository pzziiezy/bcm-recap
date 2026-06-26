"use client";

import { useRef, useState, useEffect, DragEvent } from "react";
import {
  CloudUpload,
  CheckCircle,
  XCircle,
  ArrowLeft,
  Clock,
  RefreshCw,
} from "lucide-react";

export interface DriveFileInfo {
  id: string;
  name: string;
  createdTime: string;
}

interface Props {
  onBack: () => void;
  onFileInfoChange: (info: DriveFileInfo | null) => void;
}

export function formatDateTime(isoString: string): string {
  const d = new Date(isoString);
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

export default function SpacemanManager({ onBack, onFileInfoChange }: Props) {
  const [latestFile, setLatestFile] = useState<DriveFileInfo | null>(null);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<"idle" | "success" | "error">("idle");
  const [uploadError, setUploadError] = useState("");
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const fetchLatest = async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/spaceman/latest");
      const data = await res.json();
      const file: DriveFileInfo | null = data.file ?? null;
      setLatestFile(file);
      onFileInfoChange(file);
    } catch {
      setLatestFile(null);
      onFileInfoChange(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchLatest();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleUpload = async (file: File) => {
    setUploading(true);
    setUploadStatus("idle");
    setUploadError("");
    try {
      const formData = new FormData();
      formData.append("file", file);
      const res = await fetch("/api/spaceman/upload", {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        const data = await res.json();
        throw new Error(data.error || "อัปโหลดล้มเหลว");
      }
      setUploadStatus("success");
      await fetchLatest();
    } catch (err) {
      setUploadStatus("error");
      setUploadError(String(err));
    } finally {
      setUploading(false);
    }
  };

  const handleFiles = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    handleUpload(files[0]);
  };

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragging(false);
    if (!uploading) handleFiles(e.dataTransfer.files);
  };

  return (
    <div className="space-y-6">
      <button
        onClick={onBack}
        className="flex items-center gap-2 text-sm text-slate-600 hover:text-[#E91E8C] transition-colors font-medium"
      >
        <ArrowLeft className="w-4 h-4" />
        กลับสู่หน้าหลัก
      </button>

      {/* Status card */}
      <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />
        <div className="px-6 py-4 border-b border-pink-50 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-1 h-5 rounded-full bg-gradient-to-b from-[#E91E8C] to-[#F15A22]" />
            <h2 className="font-bold text-slate-800 text-lg">สถานะไฟล์ใน Google Drive</h2>
          </div>
          <button
            onClick={fetchLatest}
            disabled={loading}
            className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-[#E91E8C] transition-colors disabled:opacity-40"
          >
            <RefreshCw className={`w-3.5 h-3.5 ${loading ? "animate-spin" : ""}`} />
            รีเฟรช
          </button>
        </div>
        <div className="p-6">
          {loading ? (
            <div className="flex items-center gap-3 text-slate-500">
              <div className="animate-spin rounded-full h-5 w-5 border-2 border-pink-200 border-t-[#E91E8C]" />
              <span className="text-sm">กำลังตรวจสอบ...</span>
            </div>
          ) : latestFile ? (
            <div className="flex items-start gap-4">
              <div className="rounded-full bg-green-100 p-3 flex-shrink-0">
                <CheckCircle className="w-6 h-6 text-green-500" />
              </div>
              <div>
                <p className="font-semibold text-slate-800">{latestFile.name}</p>
                <div className="flex items-center gap-1.5 mt-1.5 text-sm text-slate-500">
                  <Clock className="w-3.5 h-3.5 flex-shrink-0" />
                  <span>
                    อัปโหลดล่าสุด:{" "}
                    <strong className="text-slate-700 font-semibold">
                      {formatDateTime(latestFile.createdTime)}
                    </strong>
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-1">
                  ไฟล์นี้จะถูกใช้อัตโนมัติใน Step 3
                </p>
              </div>
            </div>
          ) : (
            <div className="flex items-center gap-3 text-slate-500">
              <XCircle className="w-5 h-5 text-slate-400" />
              <span className="text-sm">ยังไม่มีไฟล์ใน Google Drive — อัปโหลดด้านล่างได้เลย</span>
            </div>
          )}
        </div>
      </div>

      {/* Upload card */}
      <div className="bg-white rounded-2xl shadow-sm border border-pink-100 overflow-hidden">
        <div className="h-1 bg-gradient-to-r from-[#E91E8C] via-[#00A6E2] via-[#FFD100] via-[#F15A22] to-[#72BF44]" />
        <div className="px-6 py-4 border-b border-pink-50 flex items-center gap-3">
          <div className="w-1 h-5 rounded-full bg-gradient-to-b from-[#E91E8C] to-[#F15A22]" />
          <h2 className="font-bold text-slate-800 text-lg">อัปโหลดไฟล์ DATA_SPACEMAN ใหม่</h2>
        </div>
        <div className="p-6 space-y-4">
          <div
            onClick={() => !uploading && inputRef.current?.click()}
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            className={`
              relative flex flex-col items-center justify-center gap-3
              border-2 border-dashed rounded-xl p-10 transition-all duration-200
              ${uploading
                ? "opacity-60 cursor-not-allowed border-pink-200 bg-pink-50/30"
                : dragging
                  ? "border-[#E91E8C] bg-pink-50 scale-[1.01] cursor-pointer"
                  : "border-pink-200 bg-pink-50/30 hover:border-[#E91E8C] hover:bg-pink-50 cursor-pointer"
              }
            `}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(e) => handleFiles(e.target.files)}
            />
            {uploading ? (
              <div className="animate-spin rounded-full h-10 w-10 border-4 border-pink-200 border-t-[#E91E8C]" />
            ) : (
              <CloudUpload
                className={`w-10 h-10 ${dragging ? "text-[#E91E8C]" : "text-pink-300"}`}
              />
            )}
            <div className="text-center">
              <p className="font-semibold text-slate-700">DATA_SPACEMAN.xlsx</p>
              <p className="text-sm text-slate-400 mt-1">
                {uploading
                  ? "กำลังอัปโหลดไปยัง Google Drive..."
                  : "คลิกเพื่อเลือกไฟล์ หรือลากไฟล์มาวางที่นี่"}
              </p>
            </div>
          </div>

          {uploadStatus === "success" && (
            <div className="flex items-center gap-2 px-4 py-3 bg-green-50 border border-green-200 rounded-xl text-green-700 text-sm">
              <CheckCircle className="w-4 h-4 flex-shrink-0" />
              อัปโหลดสำเร็จ! ไฟล์ถูกบันทึกใน Google Drive แล้ว
            </div>
          )}

          {uploadStatus === "error" && (
            <div className="flex items-start gap-2 px-4 py-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-sm">
              <XCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
              <span>อัปโหลดล้มเหลว: {uploadError}</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
