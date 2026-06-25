"use client";
import { useRef, useState, DragEvent } from "react";
import { Upload, CheckCircle, X } from "lucide-react";

interface Props {
  label: string;
  accept: string;
  multiple?: boolean;
  files: File[];
  onFiles: (files: File[]) => void;
  hint?: string;
}

export default function DropZone({
  label,
  accept,
  multiple = false,
  files,
  onFiles,
  hint,
}: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const handleFiles = (incoming: FileList | null) => {
    if (!incoming) return;
    const arr = Array.from(incoming);
    onFiles(multiple ? [...files, ...arr] : arr);
  };

  const onDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragging(false);
    handleFiles(e.dataTransfer.files);
  };

  const removeFile = (i: number) => {
    onFiles(files.filter((_, idx) => idx !== i));
  };

  return (
    <div className="space-y-3">
      <div
        onClick={() => inputRef.current?.click()}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        className={`
          relative flex flex-col items-center justify-center gap-3
          border-2 border-dashed rounded-xl p-8 cursor-pointer
          transition-all duration-200
          ${dragging
            ? "border-[#E91E8C] bg-pink-50 scale-[1.01]"
            : files.length > 0
              ? "border-[#72BF44] bg-green-50"
              : "border-pink-200 bg-pink-50/30 hover:border-[#E91E8C] hover:bg-pink-50"
          }
        `}
      >
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          multiple={multiple}
          className="hidden"
          onChange={(e) => handleFiles(e.target.files)}
        />
        {files.length > 0 ? (
          <CheckCircle className="w-10 h-10 text-[#72BF44]" />
        ) : (
          <Upload className={`w-10 h-10 ${dragging ? "text-[#E91E8C]" : "text-pink-300"}`} />
        )}
        <div className="text-center">
          <p className="font-semibold text-slate-700">{label}</p>
          {hint && <p className="text-sm text-slate-400 mt-1">{hint}</p>}
          <p className="text-xs text-slate-400 mt-2">
            คลิกเพื่อเลือกไฟล์ หรือลากไฟล์มาวางที่นี่
          </p>
        </div>
      </div>

      {files.length > 0 && (
        <ul className="space-y-1">
          {files.map((f, i) => (
            <li
              key={i}
              className="flex items-center justify-between px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm"
            >
              <span className="text-slate-700 truncate max-w-xs">{f.name}</span>
              <button
                onClick={() => removeFile(i)}
                className="ml-2 text-slate-400 hover:text-red-500 transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
