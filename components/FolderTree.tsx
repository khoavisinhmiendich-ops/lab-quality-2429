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
  searchQuery: string;
  onSearchQueryChange: (query: string) => void;
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

// ---- Phân loại tài liệu theo 12 Chương → 4 nhóm (SOP/Quy trình, Biểu mẫu, PDF, Tài liệu) ----
// Quy tắc nhận diện dựa trên tiền tố tên file thật trong 2429.2026:
//   - Đuôi .pdf                → PDF
//   - Tiền tố "XN-QTQL"        → SOP / Quy trình
//   - Tiền tố "XN-BM"          → Biểu mẫu
//   - Tiền tố "STCL" hoặc chứa "sổ tay" → Sổ tay (Sổ tay chất lượng)
//   - Còn lại                  → Tài liệu
type FileCategoryKey = 'sop' | 'form' | 'pdf' | 'manual' | 'doc';

interface CategoryGroup {
  key: FileCategoryKey;
  label: string;
  files: DocumentNode[];
}

interface ChapterGroup {
  id: string;
  title: string;
  categories: CategoryGroup[];
  totalCount: number;
}

const CATEGORY_DEFS: { key: FileCategoryKey; label: string }[] = [
  { key: 'sop', label: 'SOP / Quy trình' },
  { key: 'form', label: 'Biểu mẫu' },
  { key: 'pdf', label: 'PDF' },
  { key: 'manual', label: 'Sổ tay' },
  { key: 'doc', label: 'Tài liệu' },
];

/** Bỏ dấu tiếng Việt để so khớp không phụ thuộc dấu (vd: "Chương" ~ "chuong") */
const stripDiacritics = (str: string): string =>
  str
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, 'd')
    .toLowerCase();

/** Chỉ những thư mục cấp 1 thật sự là "Chương" (vd: "1. Chương I Tổ chức quản lý") mới được tính là 1 trong 12 Chương chính */
const isChapterFolder = (node: DocumentNode): boolean => stripDiacritics(node.title).includes('chuong');

const classifyFile = (node: DocumentNode): FileCategoryKey => {
  const ext = (node.type || '').toLowerCase();
  const name = (node.fileName || node.title || '').trim().toLowerCase();
  // Chuẩn hoá dấu gạch ngang/gạch dưới thành khoảng trắng để so khớp "so tay"/"so-tay"/"so_tay" như nhau
  const cleanName = stripDiacritics(node.fileName || node.title || '').replace(/[-_]/g, ' ');

  if (ext === 'pdf' || name.endsWith('.pdf')) return 'pdf';
  if (name.startsWith('xn-qtql')) return 'sop';
  if (name.startsWith('xn-bm')) return 'form';
  // Tiền tố thật của 5 loại sổ tay: XN-STAT, XN-STBĐ, XN-STCL, XN-STDV, XN-STTC
  if (name.startsWith('xn-st') || cleanName.includes('so tay')) return 'manual';
  return 'doc';
};

/** Duyệt đệ quy, gom mọi file lá (bỏ qua độ sâu thư mục con thật) vào một mảng phẳng */
const collectLeafFiles = (nodes: DocumentNode[], acc: DocumentNode[] = []): DocumentNode[] => {
  nodes.forEach((node) => {
    if (node.children && node.children.length > 0) {
      collectLeafFiles(node.children, acc);
    } else if (node.path) {
      acc.push(node);
    }
  });
  return acc;
};

/** Phân loại 1 danh sách file phẳng vào các nhóm cố định (SOP/Biểu mẫu/PDF/Sổ tay/Tài liệu) */
const bucketFilesByCategory = (files: DocumentNode[]): CategoryGroup[] => {
  const buckets: Record<FileCategoryKey, DocumentNode[]> = { sop: [], form: [], pdf: [], manual: [], doc: [] };
  files.forEach((file) => {
    buckets[classifyFile(file)].push(file);
  });
  return CATEGORY_DEFS.map((def) => ({ ...def, files: buckets[def.key] })).filter((cat) => cat.files.length > 0);
};

/** Dựng đúng 12 Chương (lọc bỏ các thư mục cấp 1 không phải "Chương") → mỗi chương chia phẳng thành các nhóm cố định */
const buildChapterGroups = (treeData: DocumentNode[]): ChapterGroup[] => {
  return treeData.filter(isChapterFolder).map((chapterNode) => {
    const allFiles = chapterNode.children ? collectLeafFiles(chapterNode.children) : [];
    return {
      id: chapterNode.id,
      title: chapterNode.title,
      categories: bucketFilesByCategory(allFiles),
      totalCount: allFiles.length,
    };
  });
};

/**
 * Gom mọi thứ ở cấp gốc KHÔNG PHẢI 1 trong 12 Chương (vd: thư mục "Sổ tay (5 Sổ)",
 * các file PDF rời như quyết định, tiêu chí đánh giá...) vào 1 nhóm riêng "Sổ tay & Tài liệu khác"
 * để không bị ẩn mất hoàn toàn khỏi giao diện, nhưng KHÔNG tính vào danh sách "12 Chương".
 */
const buildOtherGroup = (treeData: DocumentNode[]): ChapterGroup | null => {
  const nonChapterNodes = treeData.filter((node) => !isChapterFolder(node));
  const allFiles: DocumentNode[] = [];
  nonChapterNodes.forEach((node) => {
    if (node.children && node.children.length > 0) {
      collectLeafFiles(node.children, allFiles);
    } else if (node.path) {
      allFiles.push(node);
    }
  });

  if (allFiles.length === 0) return null;

  return {
    id: '__other__',
    title: 'Sổ tay & Tài liệu khác',
    categories: bucketFilesByCategory(allFiles),
    totalCount: allFiles.length,
  };
};

/** Lọc cây Chương/Danh mục/File theo từ khoá tìm kiếm, ẩn nhóm/chương rỗng */
const filterChapterGroups = (chapters: ChapterGroup[], query: string): ChapterGroup[] => {
  if (!query.trim()) return chapters;
  const q = query.toLowerCase();
  const cleanQ = normalizeRomanNumerals(query);

  return chapters
    .map((chapter) => {
      const categories = chapter.categories
        .map((cat) => ({
          ...cat,
          files: cat.files.filter(
            (f) => f.title.toLowerCase().includes(q) || normalizeRomanNumerals(f.title).includes(cleanQ)
          ),
        }))
        .filter((cat) => cat.files.length > 0);
      return { ...chapter, categories };
    })
    .filter((chapter) => chapter.categories.length > 0);
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

export default function FolderTree({ onSelectFile, selectedFile, searchQuery, onSearchQueryChange }: FolderTreeProps) {
  const [treeData, setTreeData] = useState<DocumentNode[]>([]);
  const [openChapters, setOpenChapters] = useState<Record<string, boolean>>({});
  const [openCategories, setOpenCategories] = useState<Record<string, boolean>>({});
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
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

  // Dựng đúng 12 Chương → 5 nhóm SOP/Biểu mẫu/PDF/Sổ tay/Tài liệu từ cây thư mục thật
  const chapterGroups = useMemo(() => buildChapterGroups(treeData), [treeData]);
  // Mọi thứ ở cấp gốc KHÔNG thuộc 1 trong 12 Chương (vd: thư mục "Sổ tay (5 Sổ)", các PDF rời)
  // — gom vào 1 nhóm riêng để không bị ẩn mất, nhưng không tính là 1 trong 12 Chương.
  const otherGroup = useMemo(() => buildOtherGroup(treeData), [treeData]);

  const allDisplayGroups = useMemo(
    () => (otherGroup ? [...chapterGroups, otherGroup] : chapterGroups),
    [chapterGroups, otherGroup]
  );

  const filteredChapterGroups = useMemo(
    () => filterChapterGroups(allDisplayGroups, searchQuery),
    [allDisplayGroups, searchQuery]
  );

  // Thống kê số lượng Chương, PDF, Biểu mẫu và Sổ tay — tính trên TOÀN BỘ tài liệu (kể cả nhóm "Sổ tay & Tài liệu khác")
  // để phản ánh đúng số liệu thật, dù chỉ 12 Chương được hiển thị là "Chương" chính thức.
  const stats = useMemo(() => {
    let pdfs = 0;
    let forms = 0;
    let manuals = 0;
    allDisplayGroups.forEach((chapter) => {
      chapter.categories.forEach((cat) => {
        if (cat.key === 'pdf') pdfs += cat.files.length;
        if (cat.key === 'form') forms += cat.files.length;
        if (cat.key === 'manual') manuals += cat.files.length;
      });
    });
    return { chapters: chapterGroups.length, pdfs, docs: forms, manuals };
  }, [chapterGroups, allDisplayGroups]);

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
          onSearchQueryChange(userQuery);
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

  const CHAPTER_COLORS = [
    { text: 'text-teal-700', bg: 'bg-teal-50', border: 'border-teal-200/80', dot: 'bg-teal-500' },
    { text: 'text-amber-700', bg: 'bg-amber-50', border: 'border-amber-200/80', dot: 'bg-amber-500' },
    { text: 'text-rose-700', bg: 'bg-rose-50', border: 'border-rose-200/80', dot: 'bg-rose-500' },
    { text: 'text-violet-700', bg: 'bg-violet-50', border: 'border-violet-200/80', dot: 'bg-violet-500' },
    { text: 'text-sky-700', bg: 'bg-sky-50', border: 'border-sky-200/80', dot: 'bg-sky-500' },
    { text: 'text-emerald-700', bg: 'bg-emerald-50', border: 'border-emerald-200/80', dot: 'bg-emerald-500' },
  ];

  const CATEGORY_ICON: Record<FileCategoryKey, (p: React.SVGProps<SVGSVGElement>) => React.ReactElement> = {
    sop: Icon.BookOpen,
    form: Icon.Doc,
    pdf: Icon.Pdf,
    manual: Icon.Layers,
    doc: Icon.Doc,
  };

  const toggleChapter = (chapterId: string) => {
    setOpenChapters((prev) => ({ ...prev, [chapterId]: !prev[chapterId] }));
  };

  const toggleCategory = (key: string) => {
    setOpenCategories((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const handleFileClick = (file: DocumentNode) => {
    setSelectedFileId(file.id);
    onSelectFile(file);
  };

  /** Cây 2 cấp: Chương (12 thư mục gốc thật) → Danh mục (SOP/Biểu mẫu/PDF/Tài liệu) → danh sách file phẳng */
  const OTHER_GROUP_COLOR = { text: 'text-slate-600', bg: 'bg-slate-50', border: 'border-slate-200/80', dot: 'bg-slate-400' };

  const renderChapters = (chapters: ChapterGroup[]) => {
    const isSearching = Boolean(searchQuery.trim());

    return (
      <ul className="space-y-1">
        {chapters.map((chapter, chapterIndex) => {
          const isChapterOpen = isSearching ? true : (openChapters[chapter.id] ?? false);
          const isOtherGroup = chapter.id === '__other__';
          const color = isOtherGroup ? OTHER_GROUP_COLOR : CHAPTER_COLORS[chapterIndex % CHAPTER_COLORS.length];

          return (
            <li key={chapter.id}>
              {isOtherGroup && (
                <div className="flex items-center gap-2 px-2 pt-2 pb-1">
                  <span className="text-[9px] font-bold uppercase tracking-wider text-slate-300">Ngoài 12 Chương</span>
                  <div className="flex-1 h-px bg-slate-100" />
                </div>
              )}
              <div
                onClick={() => toggleChapter(chapter.id)}
                className="group flex items-start gap-2 cursor-pointer font-bold py-1.5 px-2 rounded-lg hover:bg-slate-50 transition-colors duration-150 text-xs select-none"
              >
                <Icon.ArrowRight
                  className={`w-3 h-3 shrink-0 mt-0.5 text-slate-400 transition-transform duration-300 ${isChapterOpen ? 'rotate-90' : 'rotate-0'}`}
                />
                <span className={`w-2 h-2 rounded-full shrink-0 mt-1 ${color.dot}`} />
                <span className={`flex-1 break-words leading-snug ${isOtherGroup ? 'text-slate-500 italic' : 'text-slate-800'}`}>
                  {chapter.title}
                </span>
                <span className="shrink-0 text-[10px] font-semibold text-slate-400 mt-0.5">{chapter.totalCount}</span>
              </div>

              <div
                className={`grid transition-all duration-300 ease-in-out ${
                  isChapterOpen ? 'grid-rows-[1fr] opacity-100 mt-0.5' : 'grid-rows-[0fr] opacity-0'
                }`}
              >
                <div className="overflow-hidden pl-3 border-l border-slate-200 ml-3">
                  <ul className="space-y-0.5 pl-1">
                    {chapter.categories.map((cat) => {
                      const catKey = `${chapter.id}:${cat.key}`;
                      const isCatOpen = isSearching ? true : (openCategories[catKey] ?? false);
                      const CategoryIcon = CATEGORY_ICON[cat.key];

                      return (
                        <li key={catKey} className="my-0.5">
                          <div
                            onClick={() => toggleCategory(catKey)}
                            className={`flex items-center gap-2 cursor-pointer font-semibold py-1.5 px-2 rounded-lg transition-colors duration-150 text-[11.5px] select-none hover:bg-slate-50 ${color.text}`}
                          >
                            <Icon.ArrowRight
                              className={`w-3 h-3 shrink-0 text-slate-400 transition-transform duration-300 ${isCatOpen ? 'rotate-90' : 'rotate-0'}`}
                            />
                            <CategoryIcon className="w-3.5 h-3.5 shrink-0" />
                            <span className="flex-1 truncate">{cat.label}</span>
                            <span className="shrink-0 text-[10px] font-semibold text-slate-400">{cat.files.length}</span>
                          </div>

                          <div
                            className={`grid transition-all duration-300 ease-in-out ${
                              isCatOpen ? 'grid-rows-[1fr] opacity-100 mt-0.5' : 'grid-rows-[0fr] opacity-0'
                            }`}
                          >
                            <div className="overflow-hidden pl-3 border-l border-slate-200 ml-2.5">
                              <ul className="space-y-0.5 py-0.5">
                                {cat.files.map((file) => {
                                  const isSelected = selectedFileId === file.id;
                                  const isPdfFile = cat.key === 'pdf';

                                  return (
                                    <li key={file.id}>
                                      <button
                                        onClick={() => handleFileClick(file)}
                                        title={file.title}
                                        className={`w-full text-left flex items-start gap-2 py-1.5 px-2 rounded-lg text-xs transition-all duration-150 ${
                                          isSelected
                                            ? 'bg-teal-700 text-white font-semibold shadow-sm shadow-teal-900/20'
                                            : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                                        }`}
                                      >
                                        {isPdfFile ? (
                                          <Icon.Pdf className={`w-4 h-4 shrink-0 mt-0.5 ${isSelected ? 'text-white' : 'text-rose-500'}`} />
                                        ) : (
                                          <Icon.Doc className={`w-4 h-4 shrink-0 mt-0.5 ${isSelected ? 'text-white' : 'text-teal-600'}`} />
                                        )}
                                        {/* Hiển thị đầy đủ tên biểu mẫu/tài liệu, tự xuống dòng thay vì cắt "..." */}
                                        <span className="break-words leading-snug">{file.title}</span>
                                      </button>
                                    </li>
                                  );
                                })}
                              </ul>
                            </div>
                          </div>
                        </li>
                      );
                    })}
                  </ul>
                </div>
              </div>
            </li>
          );
        })}
      </ul>
    );
  };

  return (
    <div className="w-full h-full flex flex-col relative font-ui overflow-visible p-4">
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

      {/* Tiêu đề khu vực tài liệu */}
      <div className="mb-3 shrink-0">
        <h1 className="font-bold text-slate-800 text-[13px] flex items-center gap-2">
          <Icon.BookOpen className="w-4 h-4 text-teal-700" />
          Hồ sơ quản lý chất lượng 2429
        </h1>
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

      {/* Bảng thống kê */}
      <div className="grid grid-cols-4 gap-1 mb-3 p-2.5 bg-gradient-to-br from-teal-50 to-teal-50/40 rounded-xl border border-teal-100 shrink-0 text-center">
        <div className="flex flex-col">
          <span className="text-[8.5px] text-slate-500 font-semibold uppercase tracking-wide">Chương</span>
          <span className="text-sm font-bold text-amber-700">{stats.chapters}</span>
        </div>
        <div className="flex flex-col border-x border-teal-100/80">
          <span className="text-[8.5px] text-slate-500 font-semibold uppercase tracking-wide">PDF</span>
          <span className="text-sm font-bold text-rose-600">{stats.pdfs}</span>
        </div>
        <div className="flex flex-col border-r border-teal-100/80">
          <span className="text-[8.5px] text-slate-500 font-semibold uppercase tracking-wide">Biểu mẫu</span>
          <span className="text-sm font-bold text-teal-700">{stats.docs}</span>
        </div>
        <div className="flex flex-col">
          <span className="text-[8.5px] text-slate-500 font-semibold uppercase tracking-wide">Sổ tay</span>
          <span className="text-sm font-bold text-violet-700">{stats.manuals}</span>
        </div>
      </div>

      {/* Danh mục 12 Chương */}
      <div className="flex-1 overflow-y-auto pr-1 animate-riseIn">
        {filteredChapterGroups.length > 0 ? (
          renderChapters(filteredChapterGroups)
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
        <img
          src="/icons/ai-assistant.gif"
          alt="Trợ lý AI"
          className="w-5 h-5 shrink-0 object-contain rounded-full"
        />
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
    </div>
  );
}