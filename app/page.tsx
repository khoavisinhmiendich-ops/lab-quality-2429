'use client';

import React, { useState, useEffect, useRef } from 'react';
import FolderTree, { DocumentNode } from '@/components/FolderTree';

export default function HomePage() {
  const [selectedFile, setSelectedFile] = useState<DocumentNode | null>(null);
  const [htmlContent, setHtmlContent] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isSaved, setIsSaved] = useState<boolean>(true);
  const editorRef = useRef<HTMLDivElement>(null);
  const saveTimeoutRef = useRef<NodeJS.Timeout | null>(null);

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

  useEffect(() => {
    if (!selectedFile) return;

    const fileType = getFileType(selectedFile);
    if (fileType !== 'word' || !selectedFile.path) return;

    let isSubscribed = true;
    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;

    const loadDocument = async () => {
      if (isSubscribed) setIsLoading(true);

      try {
        // Kiểm tra dữ liệu đã lưu trên Cloud / LocalStorage trước
        const cloudRes = await fetch(`/api/document-data?key=${encodeURIComponent(docKey)}`);
        const cloudData = await cloudRes.json();

        if (cloudData && cloudData.content) {
          if (isSubscribed) {
            setHtmlContent(cloudData.content);
            setIsSaved(true);
            setIsLoading(false);
          }
          return;
        }

        const savedLocal = localStorage.getItem(docKey);
        if (savedLocal) {
          if (isSubscribed) {
            setHtmlContent(savedLocal);
            setIsSaved(true);
            setIsLoading(false);
          }
          return;
        }

        // Tải file gốc từ thư mục dự án
        const res = await fetch(selectedFile.path!);
        if (!res.ok) throw new Error('Không thể tải file gốc');

        const arrayBuffer = await res.arrayBuffer();
        const mammoth = await import('mammoth');
        const result = await mammoth.convertToHtml({ arrayBuffer });

        if (isSubscribed) {
          setHtmlContent(result.value);
          setIsSaved(true);
        }
      } catch (err) {
        console.error('Lỗi tải tài liệu:', err);
        if (isSubscribed) {
          // Xử lý phòng thủ: Cho phép soạn thảo trực tiếp nếu file gốc không đọc được cấu trúc XML
          setHtmlContent(`
            <div style="padding: 12px; color: #b45309; background: #fffbeb; border: 1px solid #fde68a; margin-bottom: 15px; border-radius: 8px; font-family: 'Times New Roman', Times, serif; font-size: 13pt;">
              ⚠️ <b>Lưu ý:</b> File gốc có cấu trúc mã nguồn cũ. Hệ thống đã mở chế độ soạn thảo trực tiếp. Bạn có thể nhập nội dung hoặc chỉnh sửa bình thường, dữ liệu sẽ tự động lưu lại.
            </div>
            <p style="font-family: 'Times New Roman', Times, serif; font-size: 13pt;">Nhập nội dung biểu mẫu tại đây...</p>
          `);
          setIsSaved(true);
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

  const handleInput = () => {
    if (!selectedFile || !editorRef.current) return;
    setIsSaved(false);

    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;
    const newContent = editorRef.current.innerHTML;

    localStorage.setItem(docKey, newContent);

    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);

    saveTimeoutRef.current = setTimeout(async () => {
      try {
        await fetch('/api/document-data', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ key: docKey, content: newContent }),
        });
        setIsSaved(true);
      } catch (err) {
        console.error('Lỗi đồng bộ Cloud:', err);
      }
    }, 800);
  };

  const handleReset = async () => {
    if (!selectedFile || !selectedFile.path) return;
    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;

    localStorage.removeItem(docKey);
    setIsLoading(true);

    try {
      await fetch('/api/document-data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: docKey, content: null }),
      });

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

  const handlePrint = () => {
    window.print();
  };

  const renderContent = () => {
    if (!selectedFile) {
      return (
        <div className="flex flex-col items-center justify-center h-full text-slate-400 animate-fadeIn transition-all duration-500">
          <span className="text-6xl mb-4 animate-bounce">📄</span>
          <p className="text-sm font-medium tracking-wide">Chọn một tài liệu hoặc biểu mẫu bên danh mục để xem nội dung</p>
        </div>
      );
    }

    const fileType = getFileType(selectedFile);

    if (fileType === 'text' || selectedFile.content) {
      return (
        <div className="bg-white p-8 rounded-2xl border border-slate-200/80 shadow-lg max-w-4xl mx-auto my-6 overflow-y-auto max-h-full transition-all duration-500 transform animate-fadeIn">
          <div className="flex items-center justify-between pb-4 mb-4 border-b border-slate-100">
            <span className="px-3 py-1 bg-blue-50 text-blue-700 font-semibold text-xs rounded-full border border-blue-200 shadow-2xs">
              🌐 Tra cứu Internet
            </span>
          </div>
          <div className="prose prose-slate max-w-none text-slate-800 leading-relaxed whitespace-pre-line text-sm" style={{ fontFamily: '"Times New Roman", Times, serif' }}>
            {selectedFile.content}
          </div>
        </div>
      );
    }

    if (fileType === 'pdf' && selectedFile.path) {
      return (
        <iframe
          src={`${selectedFile.path}#toolbar=1`}
          className="w-full h-full border-0 rounded-xl shadow-md transition-all duration-500 animate-fadeIn"
          title={selectedFile.title}
        />
      );
    }

    if (fileType === 'word') {
      return (
        <div className="flex flex-col h-full bg-slate-100 overflow-hidden transition-all duration-300">
          <div className="bg-white/90 backdrop-blur-md border-b border-slate-200/80 px-6 py-3 flex items-center justify-between shadow-sm print:hidden shrink-0 z-10 transition-all duration-300">
            <div className="flex items-center gap-3">
              <span className="text-xs font-semibold px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg border border-emerald-200 shadow-2xs flex items-center gap-1.5 transition-all duration-300 hover:scale-105">
                ✍️ Cho phép điền trực tiếp
              </span>
              <span className="text-xs text-slate-500 font-medium flex items-center gap-1.5 transition-all">
                {isSaved ? '☁️ Đã đồng bộ Cloud' : '⏳ Đang lưu dữ liệu...'}
              </span>
            </div>

            <div className="flex items-center gap-2.5">
              <button
                onClick={handleReset}
                className="px-3.5 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 rounded-xl transition-all duration-300 transform hover:scale-105 cursor-pointer shadow-2xs"
              >
                🔄 Đặt lại mẫu gốc
              </button>
              <button
                onClick={handlePrint}
                className="px-3.5 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-md transition-all duration-300 transform hover:scale-105 flex items-center gap-1.5 cursor-pointer"
              >
                🖨️ In / Trích xuất PDF
              </button>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto p-10 flex justify-center bg-slate-200/80">
            {isLoading ? (
              <div className="flex items-center gap-3 text-blue-600 font-semibold text-sm self-center animate-pulse">
                <span className="animate-spin text-xl">⏳</span> Đang đồng bộ dữ liệu từ Cloud...
              </div>
            ) : (
              <div
                ref={editorRef}
                contentEditable
                suppressContentEditableWarning
                onInput={handleInput}
                dangerouslySetInnerHTML={{ __html: htmlContent }}
                className="bg-white shadow-2xl border border-slate-300 p-16 min-h-[297mm] h-auto w-[210mm] outline-none text-black prose prose-slate max-w-none focus:ring-4 focus:ring-blue-400/30 rounded-xs mb-12 self-start animate-fadeIn transition-all duration-300 [&_table]:w-full [&_table]:table-fixed [&_table]:border-collapse [&_table]:my-3 [&_td]:border [&_td]:border-black [&_td]:p-1.5 [&_td]:overflow-hidden [&_td]:text-xs [&_th]:border [&_th]:border-black [&_th]:p-1.5 print:shadow-none print:border-none print:w-full print:p-0 print:m-0"
                style={{
                  boxSizing: 'border-box',
                  wordBreak: 'break-word',
                  fontFamily: '"Times New Roman", Times, serif', // Cố định toàn bộ font Times New Roman
                  fontSize: '13pt',
                  lineHeight: '1.4',
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
    <div className="flex h-screen bg-slate-100 overflow-hidden font-sans">
      <FolderTree onSelectFile={(file) => setSelectedFile(file)} selectedFile={selectedFile} />

      <main className="flex-1 h-full p-4 overflow-hidden print:p-0 transition-all duration-300">
        <div className="h-full bg-white rounded-2xl shadow-md border border-slate-200/80 overflow-hidden flex flex-col print:border-none transition-all duration-300 transform">
          {renderContent()}
        </div>
      </main>
    </div>
  );
}