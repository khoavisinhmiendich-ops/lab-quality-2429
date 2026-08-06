"use client";

import React, { useEffect, useRef, useState } from "react";
import Image from "next/image";
import { FileText, Download, Loader2 } from "lucide-react";

interface EditableFormProps {
  docCode: string;
  docTitle: string;
  pdfPath?: string;
  isPdf?: boolean;
  isWord?: boolean;
  viewMode?: "preview" | "edit";
}

export function EditableForm({
  docCode,
  docTitle,
  pdfPath,
  isPdf = false,
  isWord = false,
  viewMode = "edit",
}: EditableFormProps) {
  const [contentHtml, setContentHtml] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(true);
  const containerRef = useRef<HTMLDivElement>(null);

  // Đổi key sang v99 để ép làm sạch localStorage cũ của người dùng
  const storageKey = `v101_clean_margins_${pdfPath || docCode}`;

  useEffect(() => {
    const fetchContent = async () => {
      setLoading(true);
      try {
        const saved = localStorage.getItem(storageKey);
        if (saved) {
          setContentHtml(saved);
          setLoading(false);
          return;
        }

        if (pdfPath) {
          const res = await fetch(`/api/read-doc?path=${encodeURIComponent(pdfPath)}`);
          const data = await res.json();
          if (data.html) {
            setContentHtml(data.html);
          }
        }
      } catch (err) {
        console.error("Lỗi khi tải nội dung:", err);
      } finally {
        setLoading(false);
      }
    };

    fetchContent();
  }, [pdfPath, storageKey]);

  const handleInput = () => {
    if (containerRef.current) {
      localStorage.setItem(storageKey, containerRef.current.innerHTML);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (e.key === "Tab") {
      e.preventDefault();
      document.execCommand("insertText", false, "\t");
    }
  };

  return (
    <div className="flex-1 bg-[#edebe9] rounded-lg border border-[#d2d0ce] shadow-inner overflow-auto p-4 flex justify-center items-start">
      {/* CSS Cưỡng chế bắt buộc bảng phải co giãn 100% theo khung A4 */}
   <style jsx global>{`
  .a4-paper-container {
    width: 210mm !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
  }
  
  /* Triệt tiêu lề thụt âm (negative margin) của file Word */
  .a4-paper-container table {
    width: 100% !important;
    max-width: 100% !important;
    table-layout: fixed !important;
    border-collapse: collapse !important;
    margin-top: 8px !important;
    margin-bottom: 8px !important;
    margin-left: 0 !important;   /* Xóa lề trái bị thụt âm do Word */
    margin-right: 0 !important;  /* Xóa lề phải bị lệch */
    transform: none !important;
  }

  .a4-paper-container td,
  .a4-paper-container th {
    word-wrap: break-word !important;
    word-break: break-all !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
    padding: 4px 3px !important;
    font-size: 11px !important;
    border: 1px solid #323130 !important;
  }

  .a4-paper-container img {
    max-width: 100% !important;
    height: auto !important;
  }
`}</style>

      {isPdf ? (
        <iframe src={`${encodeURI(pdfPath || "")}#toolbar=1`} className="w-full h-full border-none" title={docTitle} />
      ) : isWord && viewMode === "preview" ? (
        <div className="flex flex-col items-center justify-center h-full p-8 text-center bg-white m-6 rounded-lg shadow-sm border border-[#edebe9]">
          <FileText className="w-16 h-16 text-[#107c41] mb-3" />
          <h3 className="text-base font-semibold text-[#201f1e] mb-1">{docTitle}</h3>
          <p className="text-xs text-[#605e5c] mb-6 max-w-md">
            File Word đang ở chế độ xem nhanh. Chuyển sang thẻ &ldquo;Soạn thảo&rdquo; để điền thông tin trực tiếp vào bảng.
          </p>
          {pdfPath && (
            <a
              href={encodeURI(pdfPath)}
              download
              className="flex items-center gap-2 px-5 py-2 bg-[#107c41] hover:bg-[#0b5a2f] text-white rounded text-xs font-semibold shadow-sm transition-colors"
            >
              <Download className="w-4 h-4" /> Tải xuống file Word
            </a>
          )}
        </div>
      ) : loading ? (
        <div className="flex items-center justify-center h-full gap-2 text-xs text-[#605e5c] bg-white w-full">
          <Loader2 className="w-5 h-5 animate-spin text-[#107c41]" />
          <span>Đang tải biểu mẫu từ hệ thống...</span>
        </div>
      ) : (
        <div className="a4-paper-container min-h-[297mm] h-fit bg-white border border-[#c8c6c4] shadow-md p-6 sm:p-[10mm] text-[#201f1e] relative mb-10 shrink-0">
          
          {/* Header Bệnh viện */}
          <div className="flex items-center justify-between border-b-2 border-slate-800 pb-2 mb-3 w-full">
            <div className="flex items-center gap-3">
              <div className="shrink-0">
                <Image 
                  src="/hospital-logo.png" 
                  alt="Logo Bệnh viện" 
                  width={40} 
                  height={40} 
                  className="object-contain w-10 h-10"
                />
              </div>
              <div>
                <h4 className="text-[11px] font-bold uppercase tracking-wider text-slate-900 m-0">
                  BỆNH VIỆN PHONG - DA LIỄU TW QUY HÒA
                </h4>
                <p className="text-[10px] text-slate-600 m-0">QUẢN LÝ CHẤT LƯỢNG KHOA VI SINH - MIỄN DỊCH</p>
              </div>
            </div>
            <div className="text-right ml-4">
              <span className="text-[11px] font-bold text-rose-700 font-mono bg-rose-50 px-2 py-0.5 rounded border border-rose-200">
                {docCode}
              </span>
            </div>
          </div>

          {/* Vùng soạn thảo Word */}
          <div
            ref={containerRef}
            contentEditable={true}
            suppressContentEditableWarning={true}
            onInput={handleInput}
            onKeyDown={handleKeyDown}
            dangerouslySetInnerHTML={{ __html: contentHtml }}
            className="outline-none w-full text-sm text-[#201f1e] bg-white leading-relaxed font-['Times_New_Roman',Times,serif] select-text user-select-auto cursor-text"
          />

        </div>
      )}
    </div>
  );
}

export default EditableForm;