'use client';

import React, { useState, useEffect, useRef } from 'react';
import FolderTree, { DocumentNode } from '@/components/FolderTree';

export default function HomePage() {
  const [selectedFile, setSelectedFile] = useState<DocumentNode | null>(null);
  const [htmlContent, setHtmlContent] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isSaved, setIsSaved] = useState<boolean>(true);
  const editorRef = useRef<HTMLDivElement>(null);

  // Xác định định dạng file
  const getFileType = (file: DocumentNode) => {
    if (file.type === 'pdf' || file.path?.toLowerCase().endsWith('.pdf')) {
      return 'pdf';
    }
    if (
      file.type === 'doc' ||
      file.type === 'docx' ||
      file.path?.toLowerCase().endsWith('.doc') ||
      file.path?.toLowerCase().endsWith('.docx')
    ) {
      return 'word';
    }
    return file.type || 'text';
  };

  // Load và chuyển đổi nội dung file
  useEffect(() => {
    if (!selectedFile) return;

    const fileType = getFileType(selectedFile);
    if (fileType !== 'word' || !selectedFile.path) return;

    let isSubscribed = true;
    const storageKey = `form_data_${selectedFile.path || selectedFile.id || selectedFile.title}`;

    const loadDocument = async () => {
      // 1. Kiểm tra LocalStorage
      const savedData = localStorage.getItem(storageKey);

      if (savedData) {
        if (isSubscribed) {
          setHtmlContent(savedData);
          setIsLoading(false);
          setIsSaved(true);
        }
        return;
      }

      // 2. Tải và convert file Word sang HTML
      if (isSubscribed) {
        setIsLoading(true);
      }

      try {
        const res = await fetch(selectedFile.path!);
        if (!res.ok) throw new Error('Không thể tải file');

        const arrayBuffer = await res.arrayBuffer();
        const mammoth = await import('mammoth');
        const result = await mammoth.convertToHtml({ arrayBuffer });

        if (isSubscribed) {
          setHtmlContent(result.value);
          setIsSaved(true);
        }
      } catch (err) {
        console.error('Lỗi chuyển đổi file Word:', err);
        if (isSubscribed) {
          setHtmlContent('<p style="color: red;">Không thể tải nội dung biểu mẫu.</p>');
        }
      } finally {
        if (isSubscribed) {
          setIsLoading(false);
        }
      }
    };

    loadDocument();

    return () => {
      isSubscribed = false;
    };
  }, [selectedFile]);

  // Tự động lưu vào LocalStorage khi chỉnh sửa/nhập liệu
  const handleInput = () => {
    if (!selectedFile || !editorRef.current) return;
    const storageKey = `form_data_${selectedFile.path || selectedFile.id || selectedFile.title}`;
    const newContent = editorRef.current.innerHTML;
    localStorage.setItem(storageKey, newContent);
    setIsSaved(true);
  };

  // Nút khôi phục biểu mẫu gốc
  const handleReset = async () => {
    if (!selectedFile || !selectedFile.path) return;
    const storageKey = `form_data_${selectedFile.path || selectedFile.id || selectedFile.title}`;
    localStorage.removeItem(storageKey);

    setIsLoading(true);
    try {
      const res = await fetch(selectedFile.path);
      const arrayBuffer = await res.arrayBuffer();
      const mammoth = await import('mammoth');
      const result = await mammoth.convertToHtml({ arrayBuffer });
      setHtmlContent(result.value);
      if (editorRef.current) {
        editorRef.current.innerHTML = result.value;
      }
    } catch (err) {
      console.error('Lỗi khôi phục mẫu:', err);
    } finally {
      setIsLoading(false);
      setIsSaved(true);
    }
  };

  // Nút In / Trích xuất PDF
  const handlePrint = () => {
    window.print();
  };

  const renderContent = () => {
    if (!selectedFile) {
      return (
        <div className="flex flex-col items-center justify-center h-full text-slate-400">
          <span className="text-5xl mb-3">📄</span>
          <p className="text-sm font-medium">Chọn một tài liệu hoặc biểu mẫu bên danh mục để xem nội dung</p>
        </div>
      );
    }

    const fileType = getFileType(selectedFile);

    // 1. Dạng Tra cứu Internet (Text)
    if (fileType === 'text' || selectedFile.content) {
      return (
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs max-w-4xl mx-auto my-4 overflow-y-auto max-h-full">
          <div className="flex items-center justify-between pb-4 mb-4 border-b border-slate-100">
            <span className="px-2.5 py-1 bg-blue-50 text-blue-700 font-semibold text-xs rounded-full border border-blue-200">
              🌐 Tra cứu Internet
            </span>
          </div>
          <div className="prose prose-slate max-w-none text-slate-800 leading-relaxed whitespace-pre-line text-sm">
            {selectedFile.content}
          </div>
        </div>
      );
    }

    // 2. Dạng PDF
    if (fileType === 'pdf' && selectedFile.path) {
      return (
        <iframe
          src={`${selectedFile.path}#toolbar=1`}
          className="w-full h-full border-0 rounded-lg shadow-inner"
          title={selectedFile.title}
        />
      );
    }

    // 3. Dạng Word/Biểu mẫu -> Khung A4 co giãn tự do
    if (fileType === 'word') {
      return (
        <div className="flex flex-col h-full bg-slate-100 overflow-hidden">
          {/* Thanh công cụ biểu mẫu */}
          <div className="bg-white border-b border-slate-200 px-4 py-2 flex items-center justify-between shadow-xs print:hidden shrink-0 z-10">
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold px-2.5 py-1 bg-green-50 text-green-700 rounded-md border border-green-200 flex items-center gap-1">
                ✍️ Cho phép điền trực tiếp
              </span>
              {isSaved && (
                <span className="text-xs text-slate-500 flex items-center gap-1">
                  ✓ Đã tự động lưu
                </span>
              )}
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleReset}
                className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-lg transition-all cursor-pointer"
              >
                🔄 Đặt lại mẫu gốc
              </button>
              <button
                onClick={handlePrint}
                className="px-3 py-1.5 text-xs font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-xs transition-all flex items-center gap-1 cursor-pointer"
              >
                🖨️ In / Trích xuất PDF
              </button>
            </div>
          </div>

          {/* Vùng chứa tờ A4 - Cho phép cuộn dọc hoàn toàn khi A4 giãn dài */}
          <div className="flex-1 overflow-y-auto p-8 flex justify-center bg-slate-200/70">
            {isLoading ? (
              <div className="flex items-center gap-2 text-blue-600 font-semibold text-sm self-center">
                <span className="animate-spin text-lg">⏳</span> Đang chuẩn bị biểu mẫu...
              </div>
            ) : (
              <div
                ref={editorRef}
                contentEditable
                suppressContentEditableWarning
                onInput={() => {
                  setIsSaved(false);
                  handleInput();
                }}
                dangerouslySetInnerHTML={{ __html: htmlContent }}
                className="bg-white shadow-2xl border border-slate-300 p-12 min-h-[297mm] h-auto w-[210mm] outline-none text-black prose prose-slate max-w-none focus:ring-2 focus:ring-blue-500 rounded-xs mb-12 self-start [&_table]:w-full [&_table]:border-collapse [&_table]:my-2 [&_td]:border [&_td]:border-black [&_td]:p-2 [&_th]:border [&_th]:border-black [&_th]:p-2 print:shadow-none print:border-none print:w-full print:p-0 print:m-0"
                style={{
                  boxSizing: 'border-box',
                  wordBreak: 'break-word',
                }}
              />
            )}
          </div>
        </div>
      );
    }

    return null;
  };

  return (
    <div className="flex h-screen bg-slate-100 overflow-hidden">
      {/* Cột danh mục bên trái */}
      <FolderTree onSelectFile={(file) => setSelectedFile(file)} selectedFile={selectedFile} />

      {/* Cột xem và điền biểu mẫu chính bên phải */}
      <main className="flex-1 h-full p-4 overflow-hidden print:p-0">
        <div className="h-full bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden flex flex-col print:border-none">
          {renderContent()}
        </div>
      </main>
    </div>
  );
}