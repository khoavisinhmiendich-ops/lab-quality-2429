'use client';

import React, { useState, useEffect, useMemo, useRef } from 'react';

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

  // Các thao tác thanh công cụ (Tải / In)
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
      <ul className="pl-3 space-y-1 border-l border-slate-200 ml-2">
        {nodes.map((node) => {
          const isFolder = Boolean(node.children && node.children.length > 0);
          const isOpen = searchQuery.trim() ? true : (openFolders[node.id] ?? false);
          const isSelected = selectedFileId === node.id;
          const isPdfFile = node.type === 'pdf' || node.title.toLowerCase().endsWith('.pdf');

          if (isFolder) {
            return (
              <li key={node.id} className="my-1">
                <div
                  onClick={() => toggleFolder(node.id)}
                  className="flex items-center gap-2 cursor-pointer font-medium text-amber-700 hover:text-amber-900 py-1.5 px-2 rounded-md hover:bg-amber-50 transition-colors text-xs select-none"
                >
                  <span className={`transform transition-transform duration-200 text-[10px] ${isOpen ? 'rotate-90' : 'rotate-0'}`}>
                    ▶
                  </span>
                  <span>{isOpen ? '📂' : '📁'}</span>
                  <span className="truncate">{node.title}</span>
                </div>

                <div
                  className={`grid transition-all duration-300 ease-in-out ${
                    isOpen ? 'grid-rows-[1fr] opacity-100 mt-1' : 'grid-rows-[0fr] opacity-0'
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
                className={`w-full text-left flex items-center gap-2 py-1.5 px-2 rounded-md text-xs transition-all duration-150 ${
                  isSelected
                    ? 'bg-blue-600 text-white font-semibold shadow-sm'
                    : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900'
                }`}
              >
                <span className="shrink-0">{isPdfFile ? '📕' : '📝'}</span>
                <span className="truncate">{node.title}</span>
              </button>
            </li>
          );
        })}
      </ul>
    );
  };

  return (
    <aside className="w-80 h-full bg-white border-r border-slate-200 overflow-y-auto p-4 shrink-0 shadow-sm flex flex-col relative">
      {/* Header Bệnh viện */}
      <div className="text-center pb-3 mb-3 border-b border-slate-200 shrink-0">
        <h3 className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
          BỆNH VIỆN PHONG - DA LIỄU TW QUY HÒA
        </h3>
        <h2 className="text-xs font-black text-blue-800 uppercase tracking-wide mt-0.5">
          KHOA VI SINH - MIỄN DỊCH
        </h2>
      </div>

      {/* Thanh công cụ */}
      <div className="grid grid-cols-3 gap-1 mb-3 shrink-0">
        <button
          onClick={handleDownloadWord}
          className="py-1 px-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-300 rounded text-[11px] font-semibold flex items-center justify-center gap-1 transition-colors cursor-pointer"
          title="Tải file Word hiện tại"
        >
          📥 Word
        </button>
        <button
          onClick={handleDownloadPdf}
          className="py-1 px-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-300 rounded text-[11px] font-semibold flex items-center justify-center gap-1 transition-colors cursor-pointer"
          title="Tải / Xuất file PDF"
        >
          📄 PDF
        </button>
        <button
          onClick={handlePrint}
          className="py-1 px-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 rounded text-[11px] font-semibold flex items-center justify-center gap-1 transition-colors cursor-pointer"
          title="In tài liệu hiện tại"
        >
          🖨️ In File
        </button>
      </div>

      {/* Đồng hồ hiển thị thời gian */}
      <div className="mb-3 px-2 py-1.5 bg-slate-50 rounded-md border border-slate-200 text-center shrink-0">
        <span className="text-[11px] font-medium text-slate-600 font-mono flex items-center justify-center gap-1.5">
          <span>⏰</span> {currentTime || 'Đang tải thời gian...'}
        </span>
      </div>

      {/* Tiêu đề Bộ hồ sơ */}
      <div className="mb-2 shrink-0">
        <h1 className="font-bold text-slate-800 text-sm flex items-center gap-2">
          📑 HỒ SƠ QUẢN LÝ CHẤT LƯỢNG 2429
        </h1>
      </div>

      {/* Bảng thống kê */}
      <div className="grid grid-cols-3 gap-1.5 mb-3 p-2 bg-blue-50/60 rounded-lg border border-blue-100 shrink-0 text-center">
        <div className="flex flex-col">
          <span className="text-[10px] text-slate-500 font-medium">Chương</span>
          <span className="text-xs font-bold text-amber-700">{stats.chapters}</span>
        </div>
        <div className="flex flex-col border-x border-blue-100">
          <span className="text-[10px] text-slate-500 font-medium">PDF</span>
          <span className="text-xs font-bold text-red-600">{stats.pdfs}</span>
        </div>
        <div className="flex flex-col">
          <span className="text-[10px] text-slate-500 font-medium">Biểu mẫu</span>
          <span className="text-xs font-bold text-blue-700">{stats.docs}</span>
        </div>
      </div>

      {/* Ô tìm kiếm danh mục */}
      <div className="mb-3 shrink-0 relative">
        <input
          type="text"
          placeholder="🔍 Tìm kiếm tài liệu, biểu mẫu..."
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          className="w-full pl-3 pr-8 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:bg-white transition-all placeholder:text-slate-400"
        />
        {searchQuery && (
          <button
            onClick={() => setSearchQuery('')}
            className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs font-bold cursor-pointer"
          >
            ✕
          </button>
        )}
      </div>

      {/* Danh mục Cây thư mục */}
      <div className="flex-1 overflow-y-auto pr-1">
        {filteredTreeData.length > 0 ? (
          renderTree(filteredTreeData)
        ) : (
          <div className="text-xs text-slate-400 p-2 text-center">
            {searchQuery ? 'Không tìm thấy tài liệu phù hợp' : 'Đang tải danh mục...'}
          </div>
        )}
      </div>

      {/* Nút bật/tắt Trợ Lý AI */}
      <button
        onClick={() => setIsChatOpen(!isChatOpen)}
        className="mt-3 w-full py-2 px-3 bg-linear-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-semibold rounded-lg shadow-md hover:shadow-lg transition-all duration-200 flex items-center justify-center gap-2 text-xs shrink-0 active:scale-95 cursor-pointer"
      >
        <span className="animate-bounce">🤖</span>
        <span>{isChatOpen ? 'Đóng Trợ Lý AI' : 'Hỏi Trợ Lý AI (Tra cứu 2429)'}</span>
      </button>

      {/* Khung chat AI Popup */}
      {isChatOpen && (
        <div className="absolute bottom-16 left-4 right-4 bg-white border border-slate-300 rounded-xl shadow-2xl z-50 flex flex-col h-96 overflow-hidden animate-fade-in transition-all">
          <div className="bg-linear-to-r from-blue-700 to-indigo-800 text-white px-3 py-2.5 flex items-center justify-between text-xs font-bold">
            <div className="flex items-center gap-1.5">
              <span>🤖</span>
              <span>Trợ Lý AI Khoa Vi Sinh - Miễn dịch</span>
            </div>
            <button
              onClick={() => setIsChatOpen(false)}
              className="text-white/80 hover:text-white text-sm font-bold cursor-pointer"
            >
              ✕
            </button>
          </div>

          <div className="flex-1 p-3 overflow-y-auto space-y-2.5 text-xs bg-slate-50/50">
            {messages.map((msg, index) => (
              <div
                key={index}
                className={`flex flex-col ${
                  msg.sender === 'user' ? 'items-end' : 'items-start'
                }`}
              >
                <div
                  className={`max-w-[85%] px-3 py-2 rounded-xl text-xs leading-relaxed ${
                    msg.sender === 'user'
                      ? 'bg-blue-600 text-white rounded-br-none'
                      : 'bg-white text-slate-800 border border-slate-200 shadow-sm rounded-bl-none'
                  }`}
                >
                  <p className="whitespace-pre-line">{msg.text}</p>

                  {/* Nút bấm nếu tìm thấy tài liệu nội bộ */}
                  {msg.matchedFile && (
                    <button
                      onClick={() => handleFileClick(msg.matchedFile!)}
                      className="mt-2 w-full text-left px-2 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded border border-blue-200 font-semibold flex items-center gap-1.5 transition-colors cursor-pointer"
                    >
                      <span>📂</span>
                      <span className="truncate">{msg.matchedFile.title}</span>
                    </button>
                  )}

                  {/* Nút MÀU XANH nhảy qua bài viết màn hình chính */}
                  {msg.externalResult && (
                    <div className="mt-2 pt-2 border-t border-slate-100 flex flex-col gap-1.5">
                      <button
                        onClick={() => handleOpenExternalArticle(msg.externalResult!)}
                        className="w-full text-left px-2 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded border border-emerald-300 font-semibold flex items-center justify-between transition-colors cursor-pointer shadow-2xs"
                      >
                        <span className="truncate flex items-center gap-1">
                          📖 Đọc bài viết trên màn hình...
                        </span>
                        <span>➔</span>
                      </button>

                      {msg.externalResult.url && (
                        <a
                          href={msg.externalResult.url}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-[10px] text-blue-500 hover:underline truncate block"
                        >
                          🔗 Nguồn: {msg.externalResult.url}
                        </a>
                      )}
                    </div>
                  )}
                </div>
              </div>
            ))}
            {isAiThinking && (
              <div className="flex items-start">
                <div className="bg-white text-slate-500 border border-slate-200 px-3 py-2 rounded-xl rounded-bl-none text-xs flex items-center gap-1.5">
                  <span className="animate-spin">⏳</span> Đang tra cứu dữ liệu...
                </div>
              </div>
            )}
            <div ref={chatEndRef} />
          </div>

          <form onSubmit={handleSendMessage} className="p-2 bg-white border-t border-slate-200 flex gap-1.5">
            <input
              type="text"
              placeholder="Hỏi về Chương, File hoặc tra cứu Internet..."
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              className="flex-1 px-3 py-1.5 text-xs bg-slate-100 border border-slate-300 rounded-lg focus:outline-none focus:ring-1 focus:ring-blue-500 focus:bg-white"
            />
            <button
              type="submit"
              disabled={isAiThinking}
              className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-semibold text-xs transition-colors disabled:opacity-50 cursor-pointer"
            >
              Gửi
            </button>
          </form>
        </div>
      )}
    </aside>
  );
}