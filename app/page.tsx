'use client';

import React, { useState, useEffect, useRef } from 'react';
import FolderTree, { DocumentNode } from '@/components/FolderTree';

export default function HomePage() {
  // Trạng thái Pass Key
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [username, setUsername] = useState<string>('');
  const [password, setPassword] = useState<string>('');
  const [loginError, setLoginError] = useState<string>('');
  const [isLoggingOut, setIsLoggingOut] = useState<boolean>(false);

  const [selectedFile, setSelectedFile] = useState<DocumentNode | null>(null);
  const [htmlContent, setHtmlContent] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isSaved, setIsSaved] = useState<boolean>(true);
  const editorRef = useRef<HTMLDivElement>(null);
  const saveTimeoutRef = useRef<NodeJS.Timeout | null>(null);

  // Trạng thái xem file Excel (.xlsx)
  const [excelSheets, setExcelSheets] = useState<{ name: string; html: string }[]>([]);
  const [activeSheet, setActiveSheet] = useState<number>(0);
  const [isExcelLoading, setIsExcelLoading] = useState<boolean>(false);

  const getFileType = (file: DocumentNode) => {
    if (file.type === 'pdf' || file.path?.toLowerCase().endsWith('.pdf')) {
      return 'pdf';
    }
    if (
      file.type === 'xlsx' ||
      file.type === 'xls' ||
      file.path?.toLowerCase().endsWith('.xlsx') ||
      file.path?.toLowerCase().endsWith('.xls')
    ) {
      return 'excel';
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

  // Tải & chuyển đổi file Word (.docx) — giữ nguyên logic gốc
  useEffect(() => {
    if (!selectedFile) return;

    const fileType = getFileType(selectedFile);
    if (fileType !== 'word' || !selectedFile.path) return;

    let isSubscribed = true;
    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;

    const loadDocument = async () => {
      if (isSubscribed) setIsLoading(true);

      try {
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

  // Tải & chuyển đổi file Excel (.xlsx) — chỉ xem, không chỉnh sửa
  useEffect(() => {
    if (!selectedFile) return;

    const fileType = getFileType(selectedFile);
    if (fileType !== 'excel' || !selectedFile.path) return;

    let isSubscribed = true;

    const loadExcel = async () => {
      if (isSubscribed) {
        setIsExcelLoading(true);
        setActiveSheet(0);
        setExcelSheets([]);
      }

      try {
        const res = await fetch(selectedFile.path!);
        if (!res.ok) throw new Error('Không thể tải file gốc');

        const arrayBuffer = await res.arrayBuffer();
        const XLSX = await import('xlsx');
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        const sheets = workbook.SheetNames.map((name) => {
          const worksheet = workbook.Sheets[name];
          const html = XLSX.utils.sheet_to_html(worksheet, { editable: false });
          return { name, html };
        });

        if (isSubscribed) {
          setExcelSheets(sheets.length > 0 ? sheets : [{ name: 'Sheet1', html: '<p style="padding:16px;">Bảng tính trống.</p>' }]);
        }
      } catch (err) {
        console.error('Lỗi tải bảng tính:', err);
        if (isSubscribed) {
          setExcelSheets([
            {
              name: 'Lỗi',
              html: `<div style="padding: 16px; color: #b45309; background: #fffbeb; border: 1px solid #fde68a; border-radius: 8px; font-family: 'Times New Roman', Times, serif; font-size: 13pt;">⚠️ <b>Lưu ý:</b> Không thể đọc nội dung file Excel này.</div>`,
            },
          ]);
        }
      } finally {
        if (isSubscribed) setIsExcelLoading(false);
      }
    };

    loadExcel();

    return () => {
      isSubscribed = false;
    };
  }, [selectedFile]);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    // Tài khoản mặc định: admin / 654321
    if (username.trim() === 'admin' && password === '654321') {
      setIsAuthenticated(true);
      setLoginError('');
    } else {
      setLoginError('Tên đăng nhập hoặc mật khẩu Pass Key không chính xác!');
    }
  };

  const handleLogout = () => {
    setIsLoggingOut(true);
    setTimeout(() => {
      setIsAuthenticated(false);
      setUsername('');
      setPassword('');
      setIsLoggingOut(false);
    }, 500);
  };

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

  // ---- Icon set (inline SVG, stroke-based — no external icon dependency) ----
  const Icon = {
    Lock: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4.5" y="10.5" width="15" height="9.5" rx="2.2" />
        <path d="M7.5 10.5V7.8a4.5 4.5 0 0 1 9 0v2.7" />
        <circle cx="12" cy="15" r="1.4" fill="currentColor" stroke="none" />
      </svg>
    ),
    User: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="12" cy="8" r="3.4" />
        <path d="M5 19.5c1.2-3.4 4-5 7-5s5.8 1.6 7 5" />
      </svg>
    ),
    Key: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="8" cy="15.5" r="3.2" />
        <path d="M10.3 13.2 18 5.5m0 0h-3.4M18 5.5V9" />
      </svg>
    ),
    Doc: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
        <path d="M13.2 3.5V8h5" />
      </svg>
    ),
    Table: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4" y="5" width="16" height="14" rx="1.6" />
        <path d="M4 10h16M4 15h16M10 5v14M15 5v14" />
      </svg>
    ),
    Globe: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="12" cy="12" r="8" />
        <path d="M4 12h16M12 4c2.4 2.2 3.6 5 3.6 8s-1.2 5.8-3.6 8c-2.4-2.2-3.6-5-3.6-8s1.2-5.8 3.6-8Z" />
      </svg>
    ),
    Pen: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4 20l.9-3.6L15.6 5.7a1.7 1.7 0 0 1 2.4 0l.3.3a1.7 1.7 0 0 1 0 2.4L7.6 19.1 4 20Z" />
        <path d="M13.8 7.5l2.7 2.7" />
      </svg>
    ),
    Cloud: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M7 18h10a3.5 3.5 0 0 0 .4-6.98A5.5 5.5 0 0 0 7.1 9.6 4 4 0 0 0 7 18Z" />
      </svg>
    ),
    Refresh: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4.5 12a7.5 7.5 0 0 1 12.6-5.5M19.5 12a7.5 7.5 0 0 1-12.6 5.5" />
        <path d="M17.5 4.5V8h-3.4M6.5 19.5V16h3.4" />
      </svg>
    ),
    Print: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 8.5V4h11v4.5" />
        <rect x="4.5" y="8.5" width="15" height="7" rx="1.4" />
        <path d="M6.5 15h11V20h-11z" />
      </svg>
    ),
    Logout: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M14.5 8V6a1.5 1.5 0 0 0-1.5-1.5H6A1.5 1.5 0 0 0 4.5 6v12A1.5 1.5 0 0 0 6 19.5h7A1.5 1.5 0 0 0 14.5 18v-2" />
        <path d="M9.5 12h10m0 0-3-3m3 3-3 3" />
      </svg>
    ),
    Spinner: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" {...p}>
        <circle cx="12" cy="12" r="9" stroke="currentColor" strokeOpacity="0.2" strokeWidth="2.4" />
        <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" />
      </svg>
    ),
    File: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
        <path d="M9 12.5h6M9 15.5h6M9 9.5h2.5" />
      </svg>
    ),
  };

  const renderContent = () => {
    if (!selectedFile) {
      return (
        <div className="flex flex-col items-center justify-center h-full text-center px-6 animate-riseIn">
          <div className="relative mb-5">
            <div className="absolute inset-0 rounded-full bg-teal-500/10 blur-xl scale-150" />
            <div className="relative w-16 h-16 rounded-2xl bg-gradient-to-br from-teal-50 to-teal-100 border border-teal-200/70 flex items-center justify-center shadow-sm">
              <Icon.File className="w-7 h-7 text-teal-600" />
            </div>
          </div>
          <p className="text-[13px] font-medium text-slate-500 tracking-wide max-w-xs">
            Chọn một tài liệu hoặc biểu mẫu bên danh mục để xem nội dung
          </p>
        </div>
      );
    }

    const fileType = getFileType(selectedFile);
    const contentKey = selectedFile.id || selectedFile.path || selectedFile.title;

    if (fileType === 'text' || selectedFile.content) {
      return (
        <div
          key={contentKey}
          className="bg-white p-8 rounded-2xl border border-slate-200/70 shadow-[0_1px_2px_rgba(15,50,55,0.04),0_12px_28px_-12px_rgba(15,50,55,0.12)] max-w-4xl mx-auto my-6 overflow-y-auto max-h-full animate-riseIn"
        >
          <div className="flex items-center justify-between pb-4 mb-4 border-b border-slate-100">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-teal-50 text-teal-700 font-semibold text-[11px] uppercase tracking-wide rounded-full border border-teal-200/70">
              <Icon.Globe className="w-3.5 h-3.5" />
              Tra cứu Internet
            </span>
          </div>
          <div
            className="prose prose-slate max-w-none text-slate-800 leading-relaxed whitespace-pre-line text-sm"
            style={{ fontFamily: '"Times New Roman", Times, serif' }}
          >
            {selectedFile.content}
          </div>
        </div>
      );
    }

    if (fileType === 'pdf' && selectedFile.path) {
      return (
        <iframe
          key={contentKey}
          src={`${selectedFile.path}#toolbar=1`}
          className="w-full h-full border-0 rounded-xl shadow-md animate-riseIn"
          title={selectedFile.title}
        />
      );
    }

    if (fileType === 'excel') {
      const activeSheetData = excelSheets[activeSheet];
      return (
        <div key={contentKey} className="flex flex-col h-full bg-[#EEF4F3] overflow-hidden animate-riseIn">
          <div className="bg-white/85 backdrop-blur-md border-b border-slate-200/70 px-6 py-3 flex items-center justify-between gap-4 shadow-sm print:hidden shrink-0 z-10 animate-slideDown">
            <div className="flex items-center gap-2.5 overflow-x-auto no-scrollbar">
              <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold px-3 py-1.5 bg-sky-50 text-sky-700 rounded-full border border-sky-200/80 shrink-0">
                <Icon.Table className="w-3.5 h-3.5" />
                Bảng tính (chỉ xem)
              </span>

              {excelSheets.length > 1 && (
                <div className="flex items-center gap-1 shrink-0">
                  {excelSheets.map((sheet, idx) => (
                    <button
                      key={sheet.name + idx}
                      onClick={() => setActiveSheet(idx)}
                      className={`px-3 py-1.5 text-[11.5px] font-semibold rounded-lg transition-all duration-200 cursor-pointer hover:-translate-y-px active:translate-y-0 active:scale-95 ${
                        activeSheet === idx
                          ? 'bg-teal-700 text-white shadow-sm shadow-teal-900/20'
                          : 'bg-slate-50 text-slate-600 hover:bg-slate-100 border border-slate-200'
                      }`}
                    >
                      {sheet.name}
                    </button>
                  ))}
                </div>
              )}
            </div>

            <button
              onClick={handlePrint}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-white bg-teal-700 hover:bg-teal-600 rounded-xl shadow-sm shadow-teal-900/20 transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer shrink-0"
            >
              <Icon.Print className="w-3.5 h-3.5" />
              In / Trích xuất PDF
            </button>
          </div>

          <div className="flex-1 overflow-auto p-8 flex justify-center">
            {isExcelLoading ? (
              <div className="flex items-center gap-2.5 text-teal-700 font-semibold text-[13px] self-center animate-riseIn">
                <Icon.Spinner className="w-4 h-4 animate-spin" />
                Đang tải bảng tính...
              </div>
            ) : (
              <div
                key={activeSheet}
                className="bg-white shadow-[0_1px_1px_rgba(15,50,55,0.05),0_20px_40px_-16px_rgba(15,50,55,0.18)] border border-slate-200 rounded-sm p-4 min-w-full w-fit self-start animate-popIn [&_table]:border-collapse [&_table]:w-full [&_td]:border [&_td]:border-slate-300 [&_td]:p-1.5 [&_td]:text-[12px] [&_td]:align-top [&_td]:whitespace-nowrap [&_th]:border [&_th]:border-slate-300 [&_th]:p-1.5 [&_th]:bg-slate-50 [&_th]:text-[12px] [&_th]:font-semibold print:shadow-none print:border-none"
                style={{ fontFamily: '"Times New Roman", Times, serif' }}
                dangerouslySetInnerHTML={{ __html: activeSheetData?.html || '' }}
              />
            )}
          </div>
        </div>
      );
    }

    if (fileType === 'word') {
      return (
        <div key={contentKey} className="flex flex-col h-full bg-[#EEF4F3] overflow-hidden animate-riseIn">
          <div className="bg-white/85 backdrop-blur-md border-b border-slate-200/70 px-6 py-3 flex items-center justify-between shadow-sm print:hidden shrink-0 z-10 animate-slideDown">
            <div className="flex items-center gap-3">
              <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-full border border-emerald-200/80">
                <Icon.Pen className="w-3.5 h-3.5" />
                Cho phép điền trực tiếp
              </span>
              <span className="inline-flex items-center gap-1.5 text-[11px] font-medium text-slate-500" style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}>
                {isSaved ? (
                  <>
                    <span className="relative flex h-1.5 w-1.5">
                      <span className="absolute inline-flex h-full w-full rounded-full bg-teal-400 opacity-60" />
                      <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-teal-500" />
                    </span>
                    Đã đồng bộ Cloud
                  </>
                ) : (
                  <>
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse" />
                    Đang lưu dữ liệu...
                  </>
                )}
              </span>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleReset}
                className="group inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Refresh className="w-3.5 h-3.5 transition-transform duration-500 group-hover:rotate-180" />
                Đặt lại mẫu gốc
              </button>
              <button
                onClick={handlePrint}
                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-white bg-teal-700 hover:bg-teal-600 rounded-xl shadow-sm shadow-teal-900/20 transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Print className="w-3.5 h-3.5" />
                In / Trích xuất PDF
              </button>
            </div>
          </div>

          <div className="flex-1 overflow-y-auto p-10 flex justify-center">
            {isLoading ? (
              <div className="flex items-center gap-2.5 text-teal-700 font-semibold text-[13px] self-center animate-riseIn">
                <Icon.Spinner className="w-4 h-4 animate-spin" />
                Đang đồng bộ dữ liệu từ Cloud...
              </div>
            ) : (
              <div
                ref={editorRef}
                contentEditable
                suppressContentEditableWarning
                onInput={handleInput}
                dangerouslySetInnerHTML={{ __html: htmlContent }}
                className="bg-white shadow-[0_1px_1px_rgba(15,50,55,0.05),0_20px_40px_-16px_rgba(15,50,55,0.18)] border border-slate-200 p-16 min-h-[297mm] h-auto w-[210mm] outline-none text-black prose prose-slate max-w-none focus:ring-4 focus:ring-teal-500/20 focus:border-teal-300 rounded-sm mb-12 self-start transition-shadow duration-300 animate-popIn [&_table]:w-full [&_table]:table-fixed [&_table]:border-collapse [&_table]:my-3 [&_td]:border [&_td]:border-black [&_td]:p-1.5 [&_td]:overflow-hidden [&_td]:text-xs [&_th]:border [&_th]:border-black [&_th]:p-1.5 print:shadow-none print:border-none print:w-full print:p-0 print:m-0"
                style={{
                  boxSizing: 'border-box',
                  wordBreak: 'break-word',
                  fontFamily: '"Times New Roman", Times, serif',
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
    <div className="relative h-screen w-screen overflow-hidden bg-[#F4F9F8] font-sans">
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');

        .font-display { font-family: 'Fraunces', 'Times New Roman', serif; }
        .font-ui { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }

        .no-scrollbar::-webkit-scrollbar { display: none; }
        .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }

        @keyframes riseIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-riseIn { animation: riseIn 0.45s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes popIn {
          from { opacity: 0; transform: translateY(8px) scale(0.985); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-popIn { animation: popIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes slideDown {
          from { opacity: 0; transform: translateY(-10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-slideDown { animation: slideDown 0.4s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes veilFade {
          from { opacity: 0; }
          to { opacity: 1; }
        }
        .animate-veilFade { animation: veilFade 0.4s ease both; }

        @keyframes cardIn {
          from { opacity: 0; transform: translateY(18px) scale(0.97); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-cardIn { animation: cardIn 0.55s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes fieldIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-fieldIn { animation: fieldIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes ringSpin {
          to { transform: rotate(360deg); }
        }
        .animate-ringSpin { animation: ringSpin 10s linear infinite; }

        @keyframes ringSpinReverse {
          to { transform: rotate(-360deg); }
        }
        .animate-ringSpinReverse { animation: ringSpinReverse 14s linear infinite; }

        @keyframes shakeX {
          10%, 90% { transform: translateX(-1px); }
          20%, 80% { transform: translateX(2px); }
          30%, 50%, 70% { transform: translateX(-4px); }
          40%, 60% { transform: translateX(4px); }
        }
        .animate-shakeX { animation: shakeX 0.5s cubic-bezier(0.36,0.07,0.19,0.97) both; }

        @media (prefers-reduced-motion: reduce) {
          .animate-riseIn, .animate-popIn, .animate-slideDown, .animate-veilFade,
          .animate-cardIn, .animate-fieldIn, .animate-ringSpin, .animate-ringSpinReverse, .animate-shakeX {
            animation: none !important;
          }
        }
      `}</style>

      {/* MÀN HÌNH ĐĂNG NHẬP PASS KEY */}
      {!isAuthenticated && (
        <div
          className={`absolute inset-0 z-50 flex items-center justify-center p-4 transition-opacity duration-500 ${isLoggingOut ? 'opacity-0' : 'opacity-100 animate-veilFade'}`}
          style={{
            background:
              'radial-gradient(120% 120% at 50% 0%, #0E3A41 0%, #0A2A30 55%, #081F24 100%)',
          }}
        >
          {/* ambient texture */}
          <div
            className="pointer-events-none absolute inset-0 opacity-[0.07]"
            style={{
              backgroundImage:
                'radial-gradient(circle at 20% 20%, #fff 0, transparent 45%), radial-gradient(circle at 80% 70%, #14B8AA 0, transparent 40%)',
            }}
          />

          <div className="relative w-full max-w-md">
            <div className="bg-white rounded-[28px] shadow-[0_30px_80px_-20px_rgba(0,0,0,0.5)] border border-white/10 p-8 animate-cardIn">
              <div className="text-center mb-7">
                <div className="relative inline-flex items-center justify-center w-16 h-16 mb-4">
                  <span className="absolute inset-0 rounded-full border border-teal-200 animate-ringSpin" style={{ borderStyle: 'dashed' }} />
                  <span className="absolute inset-[3px] rounded-full border border-teal-300/60 animate-ringSpinReverse" style={{ borderStyle: 'dotted' }} />
                  <span className="absolute inset-1.5 rounded-full bg-teal-50" />
                  <Icon.Lock className="relative w-6 h-6 text-teal-700" />
                </div>
                <h2 className="font-display text-[21px] font-semibold text-[#0E3A41] tracking-tight leading-snug">
                  HỆ THỐNG QUẢN LÝ CHẤT LƯỢNG QĐ-2429/BYT
                </h2>
                <h2 className="font-display text-[21px] font-semibold text-[#0E3A41] tracking-tight leading-snug">
                  KHOA VI SINH - MIỄN DỊCH
                </h2>
                <p className="text-[11.5px] text-slate-400 font-medium mt-1.5 font-ui">
                  Bệnh viện Phong &ndash; Da liễu TW Quy Hòa
                </p>
              </div>

              <form onSubmit={handleLogin} className="space-y-4 font-ui">
                <div className="animate-fieldIn" style={{ animationDelay: '60ms' }}>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1.5 uppercase tracking-wider">
                    Tên đăng nhập
                  </label>
                  <div className="relative">
                    <Icon.User className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input
                      type="text"
                      value={username}
                      onChange={(e) => setUsername(e.target.value)}
                      placeholder="Nhập tên đăng nhập (VD: admin)"
                      className="w-full pl-10 pr-4 py-2.75 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:ring-4 focus:ring-teal-500/15 focus:border-teal-400 font-medium text-[13px] text-slate-800 transition-all duration-200 placeholder:text-slate-400"
                      required
                    />
                  </div>
                </div>

                <div className="animate-fieldIn" style={{ animationDelay: '140ms' }}>
                  <label className="block text-[11px] font-bold text-slate-500 mb-1.5 uppercase tracking-wider">
                    Mật khẩu (Pass Key)
                  </label>
                  <div className="relative">
                    <Icon.Key className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                    <input
                      type="password"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      placeholder="Nhập mật khẩu"
                      className="w-full pl-10 pr-4 py-2.75 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:ring-4 focus:ring-teal-500/15 focus:border-teal-400 font-medium text-[13px] text-slate-800 transition-all duration-200 placeholder:text-slate-400"
                      required
                    />
                  </div>
                </div>

                {loginError && (
                  <div className="p-3 bg-rose-50 border border-rose-200 text-rose-600 rounded-2xl text-[11.5px] font-semibold text-center animate-shakeX">
                    {loginError}
                  </div>
                )}

                <button
                  type="submit"
                  className="w-full py-3 bg-[#0E3A41] hover:bg-[#0A2C31] text-white rounded-2xl font-bold shadow-lg shadow-teal-950/20 transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-[0.98] cursor-pointer text-[12.5px] uppercase tracking-wider mt-1 animate-fieldIn"
                  style={{ animationDelay: '220ms' }}
                >
                  Mở khóa truy cập trang web 🚀
                </button>
              </form>

              <div className="mt-6 pt-4 border-t border-slate-100 text-center text-[10.5px] text-slate-400 font-medium font-ui animate-fieldIn" style={{ animationDelay: '280ms' }}>
                <span className="font-bold text-teal-700">© 2026 Khoa Vi sinh - Miễn dịch, Bệnh viện Phong - Da liễu TW Quy Hòa</span> {' '}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* GIAO DIỆN CHÍNH */}
      <div
        className={`h-full flex transition-all duration-700 ease-out ${
          isAuthenticated ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4 pointer-events-none'
        }`}
      >
        <FolderTree onSelectFile={(file) => setSelectedFile(file)} selectedFile={selectedFile} />

        <main className="flex-1 h-full p-4 overflow-hidden print:p-0 flex flex-col font-ui">
          <div className="flex justify-end mb-2.5 shrink-0">
            <button
              onClick={handleLogout}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200/80 rounded-xl font-bold transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer text-[12px]"
            >
              <Icon.Logout className="w-3.5 h-3.5" />
              Khóa lại (Đăng xuất)
            </button>
          </div>

          <div className="flex-1 bg-white rounded-2xl shadow-[0_1px_2px_rgba(15,50,55,0.04),0_16px_36px_-16px_rgba(15,50,55,0.12)] border border-slate-200/70 overflow-hidden flex flex-col print:border-none">
            {renderContent()}
          </div>
        </main>
      </div>
    </div>
  );
}