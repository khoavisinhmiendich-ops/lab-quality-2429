'use client';

import React, { useState, useEffect } from 'react';
import { Header } from '@/components/Header';
import { EditableForm } from '@/components/EditableForm';
import { Folder, FileText, ChevronRight, ChevronDown, ShieldCheck, Loader2, Search } from 'lucide-react';

interface DocNode {
  id: string;
  title: string;
  isFolder: boolean;
  path?: string;
  code?: string;
  children?: DocNode[];
}

export default function Home() {
  const [treeData, setTreeData] = useState<DocNode[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [openNodes, setOpenNodes] = useState<Record<string, boolean>>({});
  const [selectedDoc, setSelectedDoc] = useState<{ code: string; title: string; path: string } | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');

  useEffect(() => {
    fetch('/api/docs')
      .then((res) => res.json())
      .then((data) => {
        if (data.success && data.tree) {
          setTreeData(data.tree);

          // Tự động mở tất cả thư mục cấp 1 & cấp 2 (12 Chương & Thư mục con)
          const initialOpen: Record<string, boolean> = {};
          const openAllFolders = (nodes: DocNode[]) => {
            nodes.forEach((node) => {
              if (node.isFolder) {
                initialOpen[node.id] = true;
                if (node.children) openAllFolders(node.children);
              }
            });
          };
          openAllFolders(data.tree);
          setOpenNodes(initialOpen);

          // Chọn file đầu tiên tìm thấy làm mặc định
          const findFirstFile = (nodes: DocNode[]): DocNode | null => {
            for (const n of nodes) {
              if (!n.isFolder && n.path) return n;
              if (n.children) {
                const found = findFirstFile(n.children);
                if (found) return found;
              }
            }
            return null;
          };

          const firstFile = findFirstFile(data.tree);
          if (firstFile) {
            setSelectedDoc({
              code: firstFile.code || '2429',
              title: firstFile.title,
              path: firstFile.path!,
            });
          }
        }
      })
      .finally(() => setLoading(false));
  }, []);

  const toggleNode = (id: string) => {
    setOpenNodes((prev) => ({ ...prev, [id]: !prev[id] }));
  };

  // Render cây thư mục đệ quy đa cấp
  const renderTree = (nodes: DocNode[]) => {
    return nodes.map((node) => {
      // Lọc theo từ khóa tìm kiếm nếu có
      if (searchQuery.trim() !== '') {
        const matchesSearch = node.title.toLowerCase().includes(searchQuery.toLowerCase());
        if (!node.isFolder && !matchesSearch) return null;
      }

      if (node.isFolder) {
        const isOpen = !!openNodes[node.id];
        const hasChildren = node.children && node.children.length > 0;

        if (!hasChildren) return null;

        return (
          <div key={node.id} className="space-y-0.5">
            <button
              onClick={() => toggleNode(node.id)}
              className="w-full flex items-center gap-1.5 p-1.5 rounded hover:bg-slate-100 text-xs font-bold text-slate-800 transition-all text-left"
            >
              {isOpen ? (
                <ChevronDown className="w-3.5 h-3.5 text-slate-400 shrink-0" />
              ) : (
                <ChevronRight className="w-3.5 h-3.5 text-slate-400 shrink-0" />
              )}
              <Folder className="w-4 h-4 text-amber-500 fill-amber-500/20 shrink-0" />
              <span className="truncate">{node.title}</span>
            </button>

            {isOpen && node.children && (
              <div className="ml-3 pl-1.5 border-l border-slate-200 space-y-0.5">
                {renderTree(node.children)}
              </div>
            )}
          </div>
        );
      }

      const isSelected = selectedDoc?.path === node.path;
      return (
        <button
          key={node.id}
          onClick={() =>
            setSelectedDoc({
              code: node.code || '2429',
              title: node.title,
              path: node.path!,
            })
          }
          className={`w-full text-left p-1.5 rounded text-xs transition-all flex items-center gap-2 ${
            isSelected
              ? 'bg-rose-50 text-rose-700 font-semibold border border-rose-200 shadow-sm'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <FileText className={`w-3.5 h-3.5 shrink-0 ${isSelected ? 'text-rose-600' : 'text-slate-400'}`} />
          <span className="truncate" title={node.title}>{node.title}</span>
        </button>
      );
    });
  };

  return (
    <div className="flex flex-col h-screen overflow-hidden bg-slate-100">
      <Header />

      <div className="flex flex-1 overflow-hidden">
        {/* Sidebar hiển thị trọn bộ cây thư mục */}
        <aside className="w-80 bg-white border-r border-slate-200 flex flex-col shrink-0 shadow-sm">
          <div className="p-3 bg-slate-50 border-b border-slate-200 space-y-2">
            <h3 className="text-xs font-bold uppercase text-slate-800 tracking-wide flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-emerald-600" />
              <span>HỒ SƠ QUẢN LÝ CHẤT LƯỢNG 2429</span>
            </h3>

            {/* Ô tìm kiếm nhanh biểu mẫu */}
            <div className="relative">
              <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                placeholder="Tìm tên file, biểu mẫu..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full bg-white border border-slate-200 rounded-lg pl-8 pr-3 py-1 text-xs outline-none focus:border-rose-500 transition-all"
              />
            </div>
          </div>

          <div className="flex-1 overflow-y-auto p-2 space-y-1">
            {loading ? (
              <div className="flex items-center justify-center p-8 gap-2 text-xs text-slate-400">
                <Loader2 className="w-4 h-4 animate-spin text-rose-600" />
                <span>Đang tải toàn bộ 12 chương...</span>
              </div>
            ) : (
              renderTree(treeData)
            )}
          </div>
        </aside>

        {/* Khung xem và điền văn bản */}
        <main className="flex-1 overflow-hidden bg-slate-100">
          {selectedDoc ? (
            <EditableForm docCode={selectedDoc.code} docTitle={selectedDoc.title} pdfPath={selectedDoc.path} />
          ) : (
            <div className="flex items-center justify-center h-full text-xs text-slate-400">
              Vui lòng chọn tài liệu từ danh mục bên trái
            </div>
          )}
        </main>
      </div>
    </div>
  );
}