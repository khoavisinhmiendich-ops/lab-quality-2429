"use client";

import React, { useState, useEffect } from "react";
import { DocumentItem } from "@/data/documentsData";

interface EditableFormProps {
  selectedFile: DocumentItem | null;
}

export default function EditableForm({ selectedFile }: EditableFormProps) {
  const [content, setContent] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);

  useEffect(() => {
    let isMounted = true;

    const loadDocument = async () => {
      if (!selectedFile?.fileName) {
        if (isMounted) setContent("");
        return;
      }

      setLoading(true);
      try {
        const res = await fetch(
          `/api/read-doc?file=${encodeURIComponent(selectedFile.fileName)}`
        );
        const data = await res.json();
        if (isMounted) {
          if (data.success && data.type === "docx") {
            setContent(data.html);
          } else {
            setContent("");
          }
        }
      } catch {
        if (isMounted) setContent("");
      } finally {
        if (isMounted) setLoading(false);
      }
    };

    loadDocument();

    return () => {
      isMounted = false;
    };
  }, [selectedFile]);

  const execCommand = (command: string, value: string = "") => {
    document.execCommand(command, false, value);
  };

  if (!selectedFile) {
    return (
      <div className="flex-1 flex items-center justify-center text-gray-400 text-sm bg-gray-50">
        Vui lòng chọn tài liệu hoặc biểu mẫu từ danh mục bên trái
      </div>
    );
  }

  return (
    <div className="flex-1 flex flex-col h-screen bg-gray-100 overflow-hidden">
      {/* Header thanh công cụ MS Word */}
      <div className="bg-white border-b px-4 py-2 shadow-sm flex flex-wrap items-center gap-2">
        <h1 className="font-bold text-gray-800 text-sm mr-4 border-r pr-4">
          {selectedFile.title}
        </h1>

        {selectedFile.type === "docx" && (
          <div className="flex items-center gap-1 bg-gray-50 p-1 rounded border text-xs">
            <button
              onClick={() => execCommand("bold")}
              className="px-2 py-1 font-bold hover:bg-gray-200 rounded"
              title="In đậm"
            >
              B
            </button>
            <button
              onClick={() => execCommand("italic")}
              className="px-2 py-1 italic hover:bg-gray-200 rounded"
              title="In nghiêng"
            >
              I
            </button>
            <button
              onClick={() => execCommand("underline")}
              className="px-2 py-1 underline hover:bg-gray-200 rounded"
              title="Gạch chân"
            >
              U
            </button>
            <span className="border-r h-4 my-auto mx-1"></span>
            <button
              onClick={() => execCommand("justifyLeft")}
              className="px-2 py-1 hover:bg-gray-200 rounded"
            >
              Left
            </button>
            <button
              onClick={() => execCommand("justifyCenter")}
              className="px-2 py-1 hover:bg-gray-200 rounded"
            >
              Center
            </button>
            <button
              onClick={() => execCommand("justifyRight")}
              className="px-2 py-1 hover:bg-gray-200 rounded"
            >
              Right
            </button>
            <span className="border-r h-4 my-auto mx-1"></span>
            <button
              onClick={() => window.print()}
              className="px-2 py-1 bg-emerald-600 text-white rounded hover:bg-emerald-700"
            >
              🖨️ In / Export PDF
            </button>
          </div>
        )}
      </div>

      {/* Vùng xem/chỉnh sửa tài liệu */}
      <div className="flex-1 overflow-y-auto p-6 flex justify-center">
        {loading ? (
          <div className="text-gray-500 text-sm my-auto">Đang tải nội dung...</div>
        ) : selectedFile.type === "pdf" ? (
          <iframe
            src={`/docs/${selectedFile.fileName}`}
            className="w-full h-full border rounded-lg shadow bg-white"
            title="PDF Viewer"
          />
        ) : (
          <div
            contentEditable
            suppressContentEditableWarning
            className="w-[210mm] min-h-[297mm] bg-white p-[25mm] shadow-lg border outline-none text-gray-800 text-sm leading-relaxed"
            dangerouslySetInnerHTML={{ __html: content }}
          />
        )}
      </div>
    </div>
  );
}