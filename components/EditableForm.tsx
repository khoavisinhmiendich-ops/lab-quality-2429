'use client';

import React, { useState, useEffect, useRef } from 'react';
import Image from 'next/image';
import { Printer, CheckCircle2, Download, Loader2, RefreshCw, FileText, Edit3, Eye } from 'lucide-react';

interface Props {
  docCode: string;
  docTitle: string;
  pdfPath?: string;
}

export const EditableForm: React.FC<Props> = ({ docCode, docTitle, pdfPath }) => {
  const storageKey = `v24_ms_style_logo_content_${pdfPath || docCode}`;
  const [loading, setLoading] = useState<boolean>(false);
  const [lastSaved, setLastSaved] = useState<string>('Đã đồng bộ với đám mây');
  const [viewMode, setViewMode] = useState<'edit' | 'preview'>('edit');

  const [contentHtml, setContentHtml] = useState<string>('');
  const containerRef = useRef<HTMLDivElement>(null);

  const isPdf = pdfPath?.toLowerCase().endsWith('.pdf');
  const isWord = pdfPath?.toLowerCase().endsWith('.docx') || pdfPath?.toLowerCase().endsWith('.doc');

  useEffect(() => {
    let isMounted = true;

    const fetchContent = async () => {
      if (!pdfPath || isPdf) return;

      if (isMounted) setLoading(true);

      const savedHtml = typeof window !== 'undefined' ? localStorage.getItem(storageKey) : null;
      if (savedHtml) {
        if (isMounted) {
          setContentHtml(savedHtml);
          setLastSaved('Đã tải bản lưu gần đây');
          setLoading(false);
        }
        return;
      }

      try {
        const res = await fetch(`/api/read-doc?path=${encodeURIComponent(pdfPath)}`);
        const data = await res.json();

        if (!res.ok || !data.success || !data.html) {
          throw new Error('Không đọc được dữ liệu file');
        }

        if (isMounted) {
          const cleanedHtml = data.html
            .replace(/style="[^"]*"/gi, (match: string) => {
              return match
                .replace(/height:[^;]+(;|$)/gi, '')
                .replace(/line-height:[^;]+(;|$)/gi, '')
                .replace(/margin-top:[^;]+(;|$)/gi, '')
                .replace(/margin-bottom:[^;]+(;|$)/gi, '');
            });

          setContentHtml(cleanedHtml);
          setLastSaved('Đã đồng bộ');
        }
      } catch {
        const defaultContent = `
          <div style="font-family: 'Times New Roman', Times, serif; color: #000; line-height: 1.6;">
            <h3 style="text-align: center; font-weight: bold; margin-bottom: 15px; text-transform: uppercase;">${docTitle}</h3>
            <p><b>Mã tài liệu:</b> ${docCode}</p>
            <hr style="margin: 15px 0; border: 0; border-top: 1px solid #e1dfdd;" />
            <p><b>1. Mục đích:</b></p>
            <p style="padding-left: 20px; color: #605e5c;">[Nhập mục đích áp dụng của tài liệu tại đây...]</p>
            <p><b>2. Nội dung chi tiết:</b></p>
            <p style="padding-left: 20px; color: #605e5c;">[Nhập nội dung chi tiết hoặc dán dữ liệu biểu mẫu tại đây...]</p>
          </div>
        `;
        if (isMounted) {
          setContentHtml(defaultContent);
          setLastSaved('Đang ở chế độ soạn thảo');
        }
      } finally {
        if (isMounted) setLoading(false);
      }
    };

    fetchContent();

    return () => {
      isMounted = false;
    };
  }, [pdfPath, storageKey, docCode, docTitle, isPdf]);

  useEffect(() => {
    if (!containerRef.current) return;
    const cells = containerRef.current.querySelectorAll('td, th');
    cells.forEach((cell) => {
      const htmlCell = cell as HTMLElement;
      if (htmlCell.getAttribute('contenteditable') !== 'true') {
        htmlCell.setAttribute('contenteditable', 'true');
        htmlCell.style.outline = 'none';
      }
    });
  }, [contentHtml]);

  const handleBlur = () => {
    if (!containerRef.current) return;

    const cells = containerRef.current.querySelectorAll('td, th');
    cells.forEach((cell) => cell.removeAttribute('contenteditable'));

    const html = containerRef.current.innerHTML;

    cells.forEach((cell) => cell.setAttribute('contenteditable', 'true'));

    localStorage.setItem(storageKey, html);
    const timeStr = new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    setLastSaved(`Đã lưu lúc ${timeStr}`);
  };

  const handleReset = () => {
    if (confirm('Bạn có chắc chắn muốn khôi phục lại nội dung gốc từ file Word?')) {
      localStorage.removeItem(storageKey);
      window.location.reload();
    }
  };

  // Tailwind classes for the editable document body — kept functionally identical
  // to the original (table-fixed layout, tight cell padding, single-line cells),
  // just written as a clean list instead of a string with embedded comments.
  const editableAreaClassName = [
    "outline-none w-full text-[#201f1e] bg-white font-['Times_New_Roman',Times,serif]",
    'select-text flex flex-col items-center',
    '[&_table]:w-full [&_table]:border-collapse [&_table]:table-fixed',
    '[&_table]:my-0.5 [&_table]:mx-auto [&_table]:max-w-full [&_table]:leading-none',
    '[&_td]:border [&_td]:border-[#323130] [&_td]:py-0.5 [&_td]:px-1 [&_td]:align-middle',
    '[&_td]:text-[10px] [&_td]:overflow-hidden [&_td]:text-ellipsis [&_td]:whitespace-nowrap',
    '[&_td]:focus:bg-teal-50 [&_td]:focus:outline-2 [&_td]:focus:outline-teal-600 [&_td]:cursor-text',
    '[&_td]:transition-colors [&_td]:duration-150',
    '[&_th]:border [&_th]:border-[#323130] [&_th]:py-0.5 [&_th]:px-1 [&_th]:bg-[#f3f2f1]',
    '[&_th]:text-center [&_th]:font-bold [&_th]:text-[10px] [&_th]:whitespace-nowrap',
    '[&_p]:my-0 [&_p]:leading-none',
  ].join(' ');

  return (
    <div className="flex-1 bg-[#EEF4F3] p-5 flex flex-col h-full overflow-hidden font-['Times_New_Roman',Times,serif]">
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');
        .font-ui { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }

        @keyframes riseIn {
          from { opacity: 0; transform: translateY(-6px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-riseIn { animation: riseIn 0.45s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes paperIn {
          from { opacity: 0; transform: translateY(14px) scale(0.985); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-paperIn { animation: paperIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @media (prefers-reduced-motion: reduce) {
          .animate-riseIn, .animate-paperIn { animation: none !important; }
        }
      `}</style>

      {/* Thanh công cụ */}
      <div className="mb-4 bg-white/95 backdrop-blur border border-slate-200 rounded-xl shadow-[0_1px_2px_rgba(15,50,55,0.04),0_10px_24px_-14px_rgba(15,50,55,0.15)] px-4 py-3 flex items-center justify-between shrink-0 font-ui animate-riseIn">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-gradient-to-br from-[#0E3A41] to-teal-700 flex items-center justify-center text-white shadow-sm shrink-0">
            <FileText className="w-5 h-5" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-xs font-semibold text-slate-800 max-w-sm truncate">{docTitle}</h2>
              <span
                className="text-[10px] bg-slate-50 text-slate-500 px-1.5 py-0.5 rounded border border-slate-200"
                style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}
              >
                {docCode}
              </span>
            </div>
            <div className="flex items-center gap-1.5 text-[11px] text-teal-700 font-medium mt-0.5">
              <CheckCircle2 className="w-3.5 h-3.5" />
              <span>{lastSaved}</span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {isWord && (
            <div className="relative flex bg-slate-100 p-0.5 rounded-lg border border-slate-200 mr-2">
              <span
                className={`absolute top-0.5 bottom-0.5 w-[calc(50%-2px)] rounded-md bg-white shadow-sm transition-transform duration-300 ease-out ${
                  viewMode === 'preview' ? 'translate-x-[calc(100%+4px)]' : 'translate-x-0'
                }`}
              />
              <button
                onClick={() => setViewMode('edit')}
                className={`relative z-10 flex items-center gap-1.5 px-3 py-1 rounded text-xs font-medium transition-colors duration-200 cursor-pointer ${
                  viewMode === 'edit' ? 'text-teal-700 font-semibold' : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                <Edit3 className="w-3.5 h-3.5" /> Soạn thảo
              </button>
              <button
                onClick={() => setViewMode('preview')}
                className={`relative z-10 flex items-center gap-1.5 px-3 py-1 rounded text-xs font-medium transition-colors duration-200 cursor-pointer ${
                  viewMode === 'preview' ? 'text-teal-700 font-semibold' : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                <Eye className="w-3.5 h-3.5" /> Xem trước
              </button>
            </div>
          )}

          <button
            onClick={handleReset}
            title="Làm mới nội dung"
            className="flex items-center gap-1 px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 rounded-lg border border-slate-200 text-xs font-medium transition-all duration-200 hover:-translate-y-px shadow-sm cursor-pointer"
          >
            <RefreshCw className="w-3.5 h-3.5 text-slate-500" /> Khôi phục
          </button>

          {pdfPath && (
            <a
              href={encodeURI(pdfPath)}
              download
              className="flex items-center gap-1 px-3 py-1.5 bg-white hover:bg-slate-50 text-slate-700 rounded-lg border border-slate-200 text-xs font-medium transition-all duration-200 hover:-translate-y-px shadow-sm cursor-pointer"
            >
              <Download className="w-3.5 h-3.5 text-slate-500" /> Tải gốc
            </a>
          )}

          <button
            onClick={() => window.print()}
            className="flex items-center gap-1.5 px-4 py-1.5 bg-teal-700 hover:bg-teal-600 text-white rounded-lg text-xs font-medium transition-all duration-200 hover:-translate-y-px shadow-sm shadow-teal-900/15 cursor-pointer"
          >
            <Printer className="w-3.5 h-3.5" /> In / Xuất PDF
          </button>
        </div>
      </div>

      {/* Vùng hiển thị tài liệu */}
      <div className="flex-1 bg-[#E4EDEC] rounded-xl border border-slate-200 shadow-inner overflow-hidden flex flex-col">
        {isPdf ? (
          <iframe src={`${encodeURI(pdfPath || '')}#toolbar=1`} className="w-full h-full border-none" title={docTitle} />
        ) : isWord && viewMode === 'preview' ? (
          <div className="flex flex-col items-center justify-center h-full p-8 text-center bg-white m-6 rounded-xl shadow-sm border border-slate-200 font-ui animate-paperIn">
            <div className="w-16 h-16 rounded-2xl bg-teal-50 border border-teal-200/70 flex items-center justify-center mb-4">
              <FileText className="w-8 h-8 text-teal-700" />
            </div>
            <h3 className="text-base font-semibold text-slate-800 mb-1">{docTitle}</h3>
            <p className="text-xs text-slate-500 mb-6 max-w-md">
              File Word đang ở chế độ xem nhanh. Chuyển sang thẻ &ldquo;Soạn thảo&rdquo; để điền thông tin trực tiếp vào bảng.
            </p>
            {pdfPath && (
              <a
                href={encodeURI(pdfPath)}
                download
                className="flex items-center gap-2 px-5 py-2 bg-teal-700 hover:bg-teal-600 text-white rounded-lg text-xs font-semibold shadow-sm transition-all duration-200 hover:-translate-y-px cursor-pointer"
              >
                <Download className="w-4 h-4" /> Tải xuống file Word
              </a>
            )}
          </div>
        ) : loading ? (
          <div className="flex items-center justify-center h-full gap-2.5 text-xs text-slate-500 bg-white font-ui">
            <Loader2 className="w-5 h-5 animate-spin text-teal-700" />
            <span>Đang tải biểu mẫu từ hệ thống...</span>
          </div>
        ) : (
          <div className="flex-1 overflow-y-auto p-6 flex justify-center bg-[#EEF4F3]">
            {/* Thu nhỏ lề trang giấy từ 20mm xuống 15mm và thu gọn khoảng cách header */}
            <div className="w-[210mm] min-h-[297mm] bg-white border border-slate-300 shadow-[0_1px_1px_rgba(15,50,55,0.05),0_18px_36px_-16px_rgba(15,50,55,0.18)] p-[15mm] rounded-sm text-[#201f1e] relative animate-paperIn">

              {/* Phần tiêu đề chuẩn bệnh viện gọn gàng, chống bị đẩy nội dung */}
              <div className="flex items-center justify-between border-b-2 border-slate-800 pb-2 mb-3">
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
                    <h4 className="text-[11px] font-bold uppercase tracking-wider text-slate-900 m-0">BỆNH VIỆN PHONG - DA LIỄU TW QUY HÒA</h4>
                    <p className="text-[10px] text-slate-600 m-0">QUẢN LÝ CHẤT LƯỢNG KHOA VI SINH - MIỄN DỊCH</p>
                  </div>
                </div>
                <div className="text-right">
                  <span
                    className="text-[11px] font-bold text-rose-700 bg-rose-50 px-2 py-0.5 rounded border border-rose-200"
                    style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}
                  >
                    {docCode}
                  </span>
                </div>
              </div>

              <div className="text-[10px] text-teal-700 font-semibold mb-3 pb-1 border-b border-slate-100 flex justify-between items-center uppercase tracking-wide font-ui">
                <span>Nhấp vào các ô trong bảng để nhập dữ liệu</span>
                <span
                  className="text-slate-400"
                  style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}
                >
                  {docCode}
                </span>
              </div>

              <div
                ref={containerRef}
                onBlur={handleBlur}
                dangerouslySetInnerHTML={{ __html: contentHtml }}
                className={editableAreaClassName}
              />
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
