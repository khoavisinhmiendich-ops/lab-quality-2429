'use client';

import React, { useEffect, useState, useRef, useCallback } from 'react';
import mammoth from 'mammoth';

interface DocViewerProps {
  fileUrl: string | null;
  filePath: string | null;
}

export default function DocViewer({ fileUrl, filePath }: DocViewerProps) {
  const [content, setContent] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [saveStatus, setSaveStatus] = useState<'saved' | 'saving' | 'unsaved'>('saved');
  const [lastSavedTime, setLastSavedTime] = useState<string>('');

  const editorRef = useRef<HTMLDivElement>(null);
  const saveTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  const isPdf = Boolean(fileUrl?.toLowerCase().includes('.pdf'));

  useEffect(() => {
    if (!fileUrl || isPdf) return;

    let isMounted = true;
    const timer = setTimeout(() => {
      setLoading(true);
      setContent('');
    }, 0);

    const loadDocxContent = async () => {
      try {
        const res = await fetch(fileUrl);
        const arrayBuffer = await res.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });

        if (isMounted) {
          setContent(result.value);
          setLoading(false);
          setSaveStatus('saved');
          setLastSavedTime(
            new Date().toLocaleTimeString('vi-VN', {
              hour: '2-digit',
              minute: '2-digit',
              second: '2-digit',
            })
          );
        }
      } catch (err) {
        console.error('Lỗi đọc file Word:', err);
        if (isMounted) {
          setContent('<p className="text-red-500 font-serif">Không thể đọc nội dung file này.</p>');
          setLoading(false);
        }
      }
    };

    loadDocxContent();

    return () => {
      isMounted = false;
      clearTimeout(timer);
    };
  }, [fileUrl, isPdf]);

  const saveToServer = useCallback(
    async (htmlContent: string) => {
      if (!filePath) return;
      setSaveStatus('saving');

      try {
        const res = await fetch('/api/save-doc', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filePath, htmlContent }),
        });

        if (res.ok) {
          setSaveStatus('saved');
          const now = new Date().toLocaleTimeString('vi-VN', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
          });
          setLastSavedTime(now);
        } else {
          setSaveStatus('unsaved');
        }
      } catch {
        setSaveStatus('unsaved');
      }
    },
    [filePath]
  );

  const handleInput = () => {
    setSaveStatus('unsaved');

    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    saveTimeoutRef.current = setTimeout(() => {
      if (editorRef.current) {
        saveToServer(editorRef.current.innerHTML);
      }
    }, 500);
  };

  if (!fileUrl) {
    return (
      <div className="flex items-center justify-center h-full text-slate-400 font-['Times_New_Roman',serif] text-base">
        Vui lòng chọn 1 tài liệu ở danh mục bên trái để chỉnh sửa
      </div>
    );
  }

  if (isPdf) {
    return <iframe src={fileUrl} className="w-full h-full border-0 rounded shadow-md" title="PDF Viewer" />;
  }

  if (loading) {
    return <div className="p-6 text-slate-500 font-['Times_New_Roman',serif]">Đang tải tài liệu...</div>;
  }

  return (
    <div className="relative max-w-5xl mx-auto my-2 flex flex-col h-[calc(100vh-100px)] transition-all">
      {/* Trạng thái lưu tự động */}
      <div className="flex justify-end mb-2 shrink-0">
        <span className="text-xs font-sans font-semibold px-3 py-1.5 rounded-full border bg-white shadow-sm flex items-center gap-1.5">
          {saveStatus === 'saving' && (
            <span className="text-amber-600 flex items-center gap-1">
              <span className="animate-spin inline-block">⏳</span> Đang tự động lưu...
            </span>
          )}
          {saveStatus === 'saved' && (
            <span className="text-emerald-600">
              ✓ Đã lưu tự động thành công {lastSavedTime && `(${lastSavedTime})`}
            </span>
          )}
          {saveStatus === 'unsaved' && <span className="text-rose-500">● Đang nhập...</span>}
        </span>
      </div>

      {/* Vùng trang giấy Word Times New Roman chuẩn thẩm mỹ */}
      <div className="bg-white p-10 md:p-14 shadow-xl rounded-sm border border-slate-300 flex-1 overflow-y-auto overflow-x-auto min-w-0 transition-all">
        <div
          ref={editorRef}
          contentEditable
          suppressContentEditableWarning
          onInput={handleInput}
          className="prose max-w-none text-slate-900 outline-none focus:ring-0 w-full
            font-['Times_New_Roman',serif] text-[15px] leading-relaxed tracking-normal text-justify
            [&_p]:my-2 [&_h1]:font-bold [&_h1]:text-2xl [&_h2]:font-bold [&_h2]:text-xl [&_h3]:font-bold [&_h3]:text-lg
            [&_table]:border-collapse [&_table]:w-max [&_table]:min-w-full [&_table]:my-4 [&_table]:shadow-sm
            [&_td]:border [&_td]:border-slate-800 [&_td]:p-2 [&_td]:text-center [&_td]:min-w-11.25 [&_td]:text-sm
            [&_th]:border [&_th]:border-slate-800 [&_th]:p-2 [&_th]:font-bold [&_th]:text-center [&_th]:bg-slate-50 [&_th]:text-sm"
          dangerouslySetInnerHTML={{ __html: content }}
        />
      </div>
    </div>
  );
}