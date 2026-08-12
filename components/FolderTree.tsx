'use client';

import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Bot } from 'lucide-react';

export interface DocumentNode {
  id: string;
  title: string;
  fileName?: string;
  type?: string;
  path?: string;
  content?: string;
  children?: DocumentNode[];
}

interface FolderTreeProps {
  onSelectFile: (file: DocumentNode) => void;
  selectedFile?: DocumentNode | null;
}

interface ExternalResult {
  title: string;
  text: string;
  url?: string;
}

interface Message {
  sender: 'user' | 'bot';
  text: string;
  matchedFile?: DocumentNode;
  externalResult?: ExternalResult;
}

const normalizeRomanNumerals = (str: string): string => {
  return str
    .toLowerCase()
    .replace(/\bchương\s+i\b/g, 'chương 1')
    .replace(/\bchương\s+ii\b/g, 'chương 2')
    .replace(/\bchương\s+iii\b/g, 'chương 3')
    .replace(/\bchương\s+iv\b/g, 'chương 4')
    .replace(/\bchương\s+v\b/g, 'chương 5')
    .replace(/\bchương\s+vi\b/g, 'chương 6')
    .replace(/\bchương\s+vii\b/g, 'chương 7')
    .replace(/\bchương\s+viii\b/g, 'chương 8')
    .replace(/\bchương\s+ix\b/g, 'chương 9')
    .replace(/\bchương\s+x\b/g, 'chương 10')
    .replace(/\bchương\s+xi\b/g, 'chương 11')
    .replace(/\bchương\s+xii\b/g, 'chương 12');
};

const filterTree = (nodes: DocumentNode[], query: string): DocumentNode[] => {
  if (!query.trim()) return nodes;
  const cleanQuery = normalizeRomanNumerals(query);

  return nodes
    .map((node) => {
      const cleanTitle = normalizeRomanNumerals(node.title);
      if (node.children) {
        const filteredChildren = filterTree(node.children, query);
        if (filteredChildren.length > 0) {
          return { ...node, children: filteredChildren };
        }
      }
      if (cleanTitle.includes(cleanQuery) || node.title.toLowerCase().includes(query.toLowerCase())) {
        return node;
      }
      return null;
    })
    .filter(Boolean) as DocumentNode[];
};

// ---- Icon set (inline SVG, stroke-based — no external icon dependency) ----
const Icon = {
  Folder: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M4 6.2A1.2 1.2 0 0 1 5.2 5h4.4l1.8 2h7.4A1.2 1.2 0 0 1 20 8.2v9.6A1.2 1.2 0 0 1 18.8 19H5.2A1.2 1.2 0 0 1 4 17.8Z" />
    </svg>
  ),
  FolderOpen: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M4 8.2A1.2 1.2 0 0 1 5.2 7h4.4l1.8 2h7.4a1 1 0 0 1 .96 1.27l-1.7 6.5a1.2 1.2 0 0 1-1.16.9H5.6a1.2 1.2 0 0 1-1.2-1.2Z" />
    </svg>
  ),
  Doc: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
      <path d="M13.2 3.5V8h5M9 12.5h6M9 15.5h6" />
    </svg>
  ),
  Pdf: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
      <path d="M13.2 3.5V8h5" />
      <text x="7.3" y="17" fontSize="6.2" fontWeight="700" fill="currentColor" stroke="none" fontFamily="Inter, sans-serif">PDF</text>
    </svg>
  ),
  Search: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <circle cx="11" cy="11" r="6.5" />
      <path d="m20 20-4.3-4.3" />
    </svg>
  ),
  Clock: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <circle cx="12" cy="12" r="8" />
      <path d="M12 7.5V12l3 2" />
    </svg>
  ),
  Download: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M12 4v11m0 0-3.5-3.5M12 15l3.5-3.5" />
      <path d="M5 17.5v1.7a1.3 1.3 0 0 0 1.3 1.3h11.4a1.3 1.3 0 0 0 1.3-1.3v-1.7" />
    </svg>
  ),
  Print: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M6.5 8.5V4h11v4.5" />
      <rect x="4.5" y="8.5" width="15" height="7" rx="1.4" />
      <path d="M6.5 15h11V20h-11z" />
    </svg>
  ),
  Bot: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <rect x="5" y="9" width="14" height="10" rx="2.4" />
      <path d="M12 9V6m-3.5 0h7M9 14v.5M15 14v.5" />
      <circle cx="12" cy="4.3" r="1.1" fill="currentColor" stroke="none" />
    </svg>
  ),
  Send: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M4.5 11.8 19 4.7 13.4 19.3l-2.7-6.1-6.2-1.4Z" />
      <path d="m10.7 13.2 3.2-3.4" />
    </svg>
  ),
  X: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M6 6l12 12M18 6 6 18" />
    </svg>
  ),
  ArrowRight: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M4.5 12h15m0 0-5-5m5 5-5 5" />
    </svg>
  ),
  Link: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M9.5 14.5 14.5 9.5" />
      <path d="M11 8.2 13.1 6a3 3 0 0 1 4.2 4.2l-2.2 2.2M13 15.8 10.9 18a3 3 0 0 1-4.2-4.2l2.2-2.2" />
    </svg>
  ),
  Spinner: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" {...p}>
      <circle cx="12" cy="12" r="9" stroke="currentColor" strokeOpacity="0.2" strokeWidth="2.4" />
      <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" />
    </svg>
  ),
  BookOpen: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M12 6.5c-1.6-1.2-3.7-1.7-6-1.5v12c2.3-.2 4.4.3 6 1.5 1.6-1.2 3.7-1.7 6-1.5v-12c-2.3-.2-4.4.3-6 1.5Z" />
      <path d="M12 6.5V19" />
    </svg>
  ),
  Layers: (p: React.SVGProps<SVGSVGElement>) => (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" {...p}>
      <path d="M12 4.5 20 9l-8 4.5L4 9Z" />
      <path d="m4 13.5 8 4.5 8-4.5" />
    </svg>
  ),
};

export default function FolderTree({ onSelectFile, selectedFile }: FolderTreeProps) {
  const [treeData, setTreeData] = useState<DocumentNode[]>([]);
  const [openFolders, setOpenFolders] = useState<Record<string, boolean>>({});
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [currentTime, setCurrentTime] = useState<string>('');

  // State Chatbot AI
  const [isChatOpen, setIsChatOpen] = useState<boolean>(false);
  const [chatInput, setChatInput] = useState<string>('');
  const [isAiThinking, setIsAiThinking] = useState<boolean>(false);
  const [messages, setMessages] = useState<Message[]>([
    {
      sender: 'bot',
      text: 'Xin chào! Tôi là Trợ lý AI Khoa Vi sinh - Miễn dịch. Tôi có thể giúp gì cho bạn?',
    },
  ]);
  const chatEndRef = useRef<HTMLDivElement>(null);

  // Cuộn tự động xuống tin nhắn cuối cùng trong khung chat
  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isAiThinking]);

  // Cập nhật đồng hồ thời gian thực
  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(
        now.toLocaleDateString('vi-VN', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' }) +
          ' | ' +
          now.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      );
    };

    updateTime();
    const timer = setInterval(updateTime, 1000);
    return () => clearInterval(timer);
  }, []);

  // Lấy danh mục hồ sơ từ API nội bộ
  useEffect(() => {
    fetch('/api/docs')
      .then((res) => res.json())
      .then((data) => setTreeData(data))
      .catch((err) => console.error('Lỗi lấy cây thư mục:', err));
  }, []);

  // Thống kê số lượng Chương, PDF và Biểu mẫu
  const stats = useMemo(() => {
    let chapters = 0;
    let pdfs = 0;
    let docs = 0;

    const countNodes = (nodes: DocumentNode[]) => {
      nodes.forEach((node) => {
        if (node.children && node.children.length > 0) {
          chapters += 1;
          countNodes(node.children);
        } else {
          const isPdfFile = node.type === 'pdf' || node.title.toLowerCase().endsWith('.pdf') || node.fileName?.toLowerCase().endsWith('.pdf');
          if (isPdfFile) {
            pdfs += 1;
          } else {
            docs += 1;
          }
        }
      });
    };

    countNodes(treeData);
    return { chapters, pdfs, docs };
  }, [treeData]);

  // Danh sách tất cả các node phẳng để hỗ trợ tìm kiếm AI
  const allNodesList = useMemo(() => {
    const list: DocumentNode[] = [];
    const extractAll = (nodes: DocumentNode[]) => {
      nodes.forEach((node) => {
        list.push(node);
        if (node.children) {
          extractAll(node.children);
        }
      });
    };
    extractAll(treeData);
    return list;
  }, [treeData]);

  // Các thao tác thanh công cụ (Tải)
  const handleDownloadWord = () => {
    if (!selectedFile?.path) {
      alert('Vui lòng chọn 1 file biểu mẫu/tài liệu để tải về!');
      return;
    }
    const link = document.createElement('a');
    link.href = selectedFile.path;
    link.download = selectedFile.fileName || `${selectedFile.title}.docx`;
    link.click();
  };

  const handleDownloadPdf = () => {
    if (!selectedFile?.path) {
      alert('Vui lòng chọn 1 file để tải PDF!');
      return;
    }
    const link = document.createElement('a');
    link.href = selectedFile.path;
    link.download = selectedFile.fileName || `${selectedFile.title}.pdf`;
    link.click();
  };

  const handlePrint = () => {
    window.print();
  };

  // Mở bài viết Internet ra giao diện đọc chính bên phải
  const handleOpenExternalArticle = (ext: ExternalResult) => {
    const externalDoc: DocumentNode = {
      id: `ext-${Date.now()}`,
      title: ext.title || 'Thông tin tra cứu từ Internet',
      type: 'text',
      content: `🌐 THÔNG TIN TRA CỨU TỪ INTERNET\n-----------------------------------\n\n📌 Tiêu đề: ${ext.title}\n\n🔗 Link nguồn: ${ext.url || 'Không có URL'}\n\n📝 Nội dung chi tiết:\n${ext.text}`,
    };

    setSelectedFileId(externalDoc.id);
    onSelectFile(externalDoc);
    setIsChatOpen(false); // Đóng khung chat AI để người dùng tập trung đọc
  };

  // Xử lý gửi tin nhắn hỏi AI
  const handleSendMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!chatInput.trim() || isAiThinking) return;

    const userQuery = chatInput.trim();
    setMessages((prev) => [...prev, { sender: 'user', text: userQuery }]);
    setChatInput('');
    setIsAiThinking(true);

    const normalizedQuery = normalizeRomanNumerals(userQuery);

    // 1. Tìm trong tài liệu nội bộ
    const matched = allNodesList.find((node) => {
      const normalizedTitle = normalizeRomanNumerals(node.title);
      return (
        normalizedTitle.includes(normalizedQuery) ||
        node.title.toLowerCase().includes(userQuery.toLowerCase())
      );
    });

    if (matched) {
      setTimeout(() => {
        const isFolder = Boolean(matched.children && matched.children.length > 0);

        if (isFolder) {
          setSearchQuery(userQuery);
          setMessages((prev) => [
            ...prev,
            {
              sender: 'bot',
              text: `📂 Đã tìm thấy **${matched.title}** trong bộ hồ sơ. Tôi đã tự động lọc danh mục này bên trái cho bạn!`,
            },
          ]);
        } else {
          setMessages((prev) => [
            ...prev,
            {
              sender: 'bot',
              text: `📄 Đã tìm thấy tài liệu phù hợp: **"${matched.title}"**. Bấm vào bên dưới để mở:`,
              matchedFile: matched,
            },
          ]);
        }
        setIsAiThinking(false);
      }, 300);
      return;
    }

    // 2. Tra cứu Internet qua API Backend
    try {
      const res = await fetch(`/api/search?q=${encodeURIComponent(userQuery)}`);
      if (res.ok) {
        const data = await res.json();

        // Đảm bảo tạo kết quả tra cứu nếu có text hoặc url trả về
        if (data && (data.text || data.url)) {
          const contentText = data.text || `Nội dung tra cứu từ nguồn: ${data.url}`;

          setMessages((prev) => [
            ...prev,
            {
              sender: 'bot',
              text: `🌐 **Thông tin tra cứu từ Internet:**\n\n${data.title || userQuery}`,
              externalResult: {
                title: data.title || `Kết quả tra cứu: ${userQuery}`,
                text: contentText,
                url: data.url || '',
              },
            },
          ]);
          setIsAiThinking(false);
          return;
        }
      }
    } catch (err) {
      console.error('Lỗi tra cứu internet:', err);
    }

    // 3. Thông báo gợi ý nếu không thể kết nối hoặc không tìm thấy dữ liệu
    setMessages((prev) => [
      ...prev,
      {
        sender: 'bot',
        text: `Không tìm thấy thông tin trùng khớp cho từ khóa "${userQuery}".\n\n💡 **Gợi ý:** Bạn có thể nhập tên Chương hoặc số mã hiệu (Ví dụ: "5.12", "Chương 5" hoặc "Chương V").`,
      },
    ]);
    setIsAiThinking(false);
  };

  const toggleFolder = (folderId: string) => {
    setOpenFolders((prev) => ({ ...prev, [folderId]: !prev[folderId] }));
  };

  const handleFileClick = (file: DocumentNode) => {
    setSelectedFileId(file.id);
    onSelectFile(file);
  };

  const filteredTreeData = useMemo(() => {
    return filterTree(treeData, searchQuery);
  }, [treeData, searchQuery]);

  const renderTree = (nodes: DocumentNode[]) => {
    return (
      <ul className="pl-3 space-y-0.5 border-l border-slate-200 ml-2">
        {nodes.map((node) => {
          const isFolder = Boolean(node.children && node.children.length > 0);
          const isOpen = searchQuery.trim() ? true : (openFolders[node.id] ?? false);
          const isSelected = selectedFileId === node.id;
          const isPdfFile = node.type === 'pdf' || node.title.toLowerCase().endsWith('.pdf');

          if (isFolder) {
            return (
              <li key={node.id} className="my-0.5">
                <div
                  onClick={() => toggleFolder(node.id)}
                  className="group flex items-center gap-2 cursor-pointer font-semibold text-amber-800 hover:text-amber-900 py-1.5 px-2 rounded-lg hover:bg-amber-50 transition-colors duration-150 text-xs select-none"
                >
                  <Icon.ArrowRight
                    className={`w-3 h-3 shrink-0 text-amber-500 transition-transform duration-300 ${isOpen ? 'rotate-90' : 'rotate-0'}`}
                  />
                  {isOpen ? (
                    <Icon.FolderOpen className="w-4 h-4 shrink-0 text-amber-500" />
                  ) : (
                    <Icon.Folder className="w-4 h-4 shrink-0 text-amber-500" />
                  )}
                  <span className="truncate">{node.title}</span>
                </div>

                <div
                  className={`grid transition-all duration-300 ease-in-out ${
                    isOpen ? 'grid-rows-[1fr] opacity-100 mt-0.5' : 'grid-rows-[0fr] opacity-0'
                  }`}
                >
                  <div className="overflow-hidden">
                    {node.children && renderTree(node.children)}
                  </div>
                </div>
              </li>
            );
          }

          return (
            <li key={node.id}>
              <button
                onClick={() => handleFileClick(node)}
                className={`w-full text-left flex items-center gap-2 py-1.5 px-2 rounded-lg text-xs transition-all duration-150 ${
                  isSelected
                    ? 'bg-teal-700 text-white font-semibold shadow-sm shadow-teal-900/20'
                    : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                }`}
              >
                {isPdfFile ? (
                  <Icon.Pdf className={`w-4 h-4 shrink-0 ${isSelected ? 'text-white' : 'text-rose-500'}`} />
                ) : (
                  <Icon.Doc className={`w-4 h-4 shrink-0 ${isSelected ? 'text-white' : 'text-teal-600'}`} />
                )}
                <span className="truncate">{node.title}</span>
              </button>
            </li>
          );
        })}
      </ul>
    );
  };

  return (
    <aside className="w-80 h-full bg-white border-r border-slate-200/80 p-4 shrink-0 flex flex-col relative font-ui overflow-visible">
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');
        .font-display { font-family: 'Fraunces', 'Times New Roman', serif; }
        .font-ui { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }

        @keyframes riseIn {
          from { opacity: 0; transform: translateY(8px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-riseIn { animation: riseIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes slideUpFade {
          from { opacity: 0; transform: translateY(14px) scale(0.98); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-slideUpFade { animation: slideUpFade 0.35s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes bubbleIn {
          from { opacity: 0; transform: translateY(6px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-bubbleIn { animation: bubbleIn 0.3s ease both; }

        @keyframes softPulse {
          0%, 100% { opacity: 0.55; }
          50% { opacity: 1; }
        }
        .animate-softPulse { animation: softPulse 1.6s ease-in-out infinite; }

        @media (prefers-reduced-motion: reduce) {
          .animate-riseIn, .animate-slideUpFade, .animate-bubbleIn, .animate-softPulse {
            animation: none !important;
          }
        }
      `}</style>

      {/* Header Bệnh viện */}
      <div className="text-center pb-4 mb-3 border-b border-slate-200 shrink-0">
        <div className="inline-flex items-center justify-center w-9 h-9 rounded-xl bg-teal-50 border border-teal-200/70 mb-2">
          <Icon.Layers className="w-4.5 h-4.5 text-teal-700" />
        </div>
        <h3 className="text-[10.5px] font-bold text-slate-500 uppercase tracking-wider">
          BỆNH VIỆN PHONG - DA LIỄU TW QUY HÒA
        </h3>
        <h2 className="font-display text-[13.5px] font-semibold text-[#0E3A41] uppercase tracking-wide mt-0.5">
          Khoa Vi Sinh - Miễn Dịch
        </h2>
      </div>

      {/* Thanh công cụ */}
      <div className="grid grid-cols-3 gap-1.5 mb-3 shrink-0">
        <button
          onClick={handleDownloadWord}
          className="py-1.5 px-1.5 bg-teal-50 hover:bg-teal-100 text-teal-700 border border-teal-200 rounded-lg text-[10.5px] font-semibold flex flex-col items-center justify-center gap-0.5 transition-all duration-150 hover:-translate-y-px cursor-pointer"
          title="Tải file Word hiện tại"
        >
          <Icon.Download className="w-3.5 h-3.5" />
          Word
        </button>
        <button
          onClick={handleDownloadPdf}
          className="py-1.5 px-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-lg text-[10.5px] font-semibold flex flex-col items-center justify-center gap-0.5 transition-all duration-150 hover:-translate-y-px cursor-pointer"
          title="Tải / Xuất file PDF"
        >
          <Icon.Pdf className="w-3.5 h-3.5" />
          PDF
        </button>
        <button
          onClick={handlePrint}
          className="py-1.5 px-1.5 bg-slate-50 hover:bg-slate-100 text-slate-600 border border-slate-200 rounded-lg text-[10.5px] font-semibold flex flex-col items-center justify-center gap-0.5 transition-all duration-150 hover:-translate-y-px cursor-pointer"
          title="In tài liệu hiện tại"
        >
          <Icon.Print className="w-3.5 h-3.5" />
          In File
        </button>
      </div>

      {/* Đồng hồ hiển thị thời gian */}
      <div className="mb-3 px-2.5 py-1.5 bg-slate-50 rounded-lg border border-slate-200 text-center shrink-0">
        <span
          className="text-[10.5px] font-medium text-slate-500 flex items-center justify-center gap-1.5"
          style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}
        >
          <Icon.Clock className="w-3.5 h-3.5 text-teal-600" />
          {currentTime || 'Đang tải thời gian...'}
        </span>
      </div>

      {/* Tiêu đề Bộ hồ sơ */}
      <div className="mb-2 shrink-0">
        <h1 className="font-bold text-slate-800 text-[13px] flex items-center gap-2">
          <Icon.BookOpen className="w-4 h-4 text-teal-700" />
          Hồ sơ quản lý chất lượng 2429
        </h1>
      </div>

      {/* Bảng thống kê */}
      <div className="grid grid-cols-3 gap-1.5 mb-3 p-2.5 bg-gradient-to-br from-teal-50 to-teal-50/40 rounded-xl border border-teal-100 shrink-0 text-center">
        <div className="flex flex-col">
          <span className="text-[9.5px] text-slate-500 font-semibold uppercase tracking-wide">Chương</span>
          <span className="text-sm font-bold text-amber-700">{stats.chapters}</span>
        </div>
        <div className="flex flex-col border-x border-teal-100/80">
          <span className="text-[9.5px] text-slate-500 font-semibold uppercase tracking-wide">PDF</span>
          <span className="text-sm font-bold text-rose-600">{stats.pdfs}</span>
        </div>
        <div className="flex flex-col">
          <span className="text-[9.5px] text-slate-500 font-semibold uppercase tracking-wide">Biểu mẫu</span>
          <span className="text-sm font-bold text-teal-700">{stats.docs}</span>
        </div>
      </div>

      {/* Ô tìm kiếm danh mục */}
      <div className="mb-3 shrink-0 relative">
        <Icon.Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-400" />
        <input
          type="text"
          placeholder="Tìm kiếm tài liệu, biểu mẫu..."
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className="w-full pl-9 pr-8 py-1.75 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-teal-500/12 focus:border-teal-400 focus:bg-white transition-all duration-200 placeholder:text-slate-400"
        />
        {searchQuery && (
          <button
            onClick={() => setSearchQuery('')}
            className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
          >
            <Icon.X className="w-3.5 h-3.5" />
          </button>
        )}
      </div>

      {/* Danh mục Cây thư mục */}
      <div className="flex-1 overflow-y-auto pr-1 animate-riseIn">
        {filteredTreeData.length > 0 ? (
          renderTree(filteredTreeData)
        ) : (
          <div className="text-[11px] text-slate-400 p-3 text-center font-medium">
            {searchQuery ? 'Không tìm thấy tài liệu phù hợp' : 'Đang tải danh mục...'}
          </div>
        )}
      </div>

      {/* Nút bật/tắt Trợ Lý AI */}
      <button
  onClick={() => setIsChatOpen(!isChatOpen)}
  className="mt-3 w-full py-2.5 px-3 bg-gradient-to-r from-[#0E3A41] to-teal-700 hover:to-teal-600 text-white font-semibold rounded-xl shadow-md shadow-teal-950/15 hover:shadow-lg transition-all duration-200 flex items-center justify-center gap-2 text-xs shrink-0 active:scale-[0.98] cursor-pointer"
>
  {/* eslint-disable-next-line @next/next/no-img-element */}
  <img src="/icons/ai-assistant.gif" alt="Trợ lý AI" className="w-5 h-5 shrink-0 object-contain" />
  <span>{isChatOpen ? 'Đóng Trợ Lý AI' : 'Hỏi Trợ Lý AI (Tra cứu 2429 và các thông tin khác)'}</span>
</button>

      {/* Khung chat AI Popup */}
      {isChatOpen && (
        <div className="absolute top-20 bottom-16 left-4 right-4 bg-white border border-slate-200 rounded-2xl shadow-[0_20px_50px_-16px_rgba(15,50,55,0.35)] z-50 flex flex-col overflow-hidden animate-slideUpFade origin-bottom">
          <div className="bg-gradient-to-r from-[#0E3A41] to-teal-700 text-white px-3.5 py-3 flex items-center justify-between text-xs font-bold shrink-0">
            <div className="flex items-center gap-2">
              <span className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-white/15">
                <Icon.Bot className="w-3.5 h-3.5" />
              </span>
              <span className="font-ui">Trợ Lý AI Khoa Vi Sinh - Miễn dịch</span>
            </div>
            <button
              onClick={() => setIsChatOpen(false)}
              className="text-white/70 hover:text-white transition-colors cursor-pointer"
            >
              <Icon.X className="w-4 h-4" />
            </button>
          </div>

          <div className="flex-1 p-3 overflow-y-auto space-y-2.5 text-xs bg-slate-50/60">
            {messages.map((msg, index) => (
              <div
                key={index}
                className={`flex flex-col animate-bubbleIn ${msg.sender === 'user' ? 'items-end' : 'items-start'}`}
              >
                <div
                  className={`max-w-[85%] px-3 py-2 rounded-2xl text-xs leading-relaxed ${
                    msg.sender === 'user'
                      ? 'bg-teal-700 text-white rounded-br-md'
                      : 'bg-white text-slate-800 border border-slate-200 shadow-sm rounded-bl-md'
                  }`}
                >
                  <p className="whitespace-pre-line">{msg.text}</p>

                  {/* Nút bấm nếu tìm thấy tài liệu nội bộ */}
                  {msg.matchedFile && (
                    <button
                      onClick={() => handleFileClick(msg.matchedFile!)}
                      className="mt-2 w-full text-left px-2.5 py-1.5 bg-teal-50 hover:bg-teal-100 text-teal-700 rounded-lg border border-teal-200 font-semibold flex items-center gap-1.5 transition-colors duration-150 cursor-pointer"
                    >
                      <Icon.Folder className="w-3.5 h-3.5 shrink-0" />
                      <span className="truncate">{msg.matchedFile.title}</span>
                    </button>
                  )}

                  {/* Nút mở bài viết trên màn hình chính */}
                  {msg.externalResult && (
                    <div className="mt-2 pt-2 border-t border-slate-100 flex flex-col gap-1.5">
                      <button
                        onClick={() => handleOpenExternalArticle(msg.externalResult!)}
                        className="w-full text-left px-2.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-lg border border-emerald-200 font-semibold flex items-center justify-between transition-colors duration-150 cursor-pointer"
                      >
                        <span className="truncate flex items-center gap-1.5">
                          <Icon.BookOpen className="w-3.5 h-3.5 shrink-0" />
                          Đọc bài viết trên màn hình...
                        </span>
                        <Icon.ArrowRight className="w-3.5 h-3.5 shrink-0" />
                      </button>

                      {msg.externalResult.url && (
                        <a
                          href={msg.externalResult.url}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-[10px] text-teal-600 hover:underline truncate flex items-center gap-1"
                        >
                          <Icon.Link className="w-3 h-3 shrink-0" />
                          Nguồn: {msg.externalResult.url}
                        </a>
                      )}
                    </div>
                  )}
                </div>
              </div>
            ))}
            {isAiThinking && (
              <div className="flex items-start animate-bubbleIn">
                <div className="bg-white text-slate-500 border border-slate-200 px-3 py-2 rounded-2xl rounded-bl-md text-xs flex items-center gap-1.5">
                  <Icon.Spinner className="w-3.5 h-3.5 animate-spin text-teal-600" />
                  <span className="animate-softPulse">Đang tra cứu dữ liệu...</span>
                </div>
              </div>
            )}
            <div ref={chatEndRef} />
          </div>

          <form onSubmit={handleSendMessage} className="p-2 bg-white border-t border-slate-200 flex gap-1.5 shrink-0">
            <input
              type="text"
              placeholder="Hỏi về Chương, File hoặc tra cứu Internet..."
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              className="flex-1 px-3 py-1.75 text-xs bg-slate-100 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-teal-500/12 focus:border-teal-400 focus:bg-white transition-all duration-200"
            />
            <button
              type="submit"
              disabled={isAiThinking}
              className="px-3 py-1.75 bg-teal-700 hover:bg-teal-600 text-white rounded-xl font-semibold text-xs transition-all duration-150 disabled:opacity-40 cursor-pointer flex items-center gap-1.5"
            >
              <Icon.Send className="w-3.5 h-3.5" />
              Gửi
            </button>
          </form>
        </div>
      )}
    </aside>
  );
}
