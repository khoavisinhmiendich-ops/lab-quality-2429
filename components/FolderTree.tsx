'use client';

import React, { useEffect, useState } from 'react';
import { Folder, FileText, ChevronRight, ChevronDown, Loader2 } from 'lucide-react';

interface DirectoryItem {
  name: string;
  type: 'folder' | 'file';
  path?: string;
  children?: DirectoryItem[];
}

interface Props {
  onSelectFile: (filePath: string, fileName: string) => void;
}

export const FolderTree: React.FC<Props> = ({ onSelectFile }) => {
  const [treeData, setTreeData] = useState<DirectoryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [openFolders, setOpenFolders] = useState<Record<string, boolean>>({});

  // Gọi API lấy 100% tài liệu từ folder public/2429.2026
  useEffect(() => {
    fetch('/api/documents')
      .then((res) => res.json())
      .then((data) => {
        if (data.files) {
          setTreeData(data.files);
        }
      })
      .catch(() => console.error('Lỗi kết nối API danh mục tài liệu'))
      .finally(() => setLoading(false));
  }, []);

  const toggleFolder = (folderPath: string) => {
    setOpenFolders((prev) => ({
      ...prev,
      [folderPath]: !prev[folderPath],
    }));
  };

  const renderTree = (items: DirectoryItem[], parentPath = '') => {
    return (
      <ul className="pl-3 space-y-1 text-xs font-medium">
        {items.map((item, index) => {
          const currentPath = `${parentPath}/${item.name}`;

          if (item.type === 'folder') {
            const isOpen = openFolders[currentPath] ?? true; // Mặc định mở tất cả folder

            return (
              <li key={index} className="my-1">
                <div
                  onClick={() => toggleFolder(currentPath)}
                  className="flex items-center gap-1.5 p-1.5 hover:bg-slate-200/80 rounded cursor-pointer text-slate-800 font-bold transition-colors select-none"
                >
                  {isOpen ? (
                    <ChevronDown className="w-3.5 h-3.5 text-slate-500" />
                  ) : (
                    <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                  )}
                  <Folder className="w-4 h-4 text-amber-500 fill-amber-500/20" />
                  <span className="truncate">{item.name}</span>
                </div>

                {isOpen && item.children && item.children.length > 0 && (
                  <div className="border-l-2 border-slate-200 ml-2.5 pl-1">
                    {renderTree(item.children, currentPath)}
                  </div>
                )}
              </li>
            );
          }

          // Trường hợp là File (PDF, Biểu mẫu, Word...)
          return (
            <li key={index}>
              <div
                onClick={() => onSelectFile(item.path || '', item.name)}
                className="flex items-center gap-2 p-1.5 hover:bg-rose-50 text-slate-700 hover:text-rose-700 rounded cursor-pointer transition-colors"
              >
               <FileText className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                <span className="truncate">{item.name}</span>
              </div>
            </li>
          );
        })}
      </ul>
    );
  };

  if (loading) {
    return (
      <div className="p-4 flex items-center justify-center gap-2 text-xs text-slate-500">
        <Loader2 className="w-4 h-4 animate-spin text-rose-600" />
        <span>Đang quét 100% tài liệu 2429.2026...</span>
      </div>
    );
  }

  return (
    <div className="w-80 bg-slate-50 border-r border-slate-200 h-full overflow-y-auto p-3">
      <h2 className="text-xs font-bold uppercase text-slate-500 mb-3 tracking-wider px-1">
        DANH MỤC TÀI LIỆU 2429.2026
      </h2>
      {treeData.length === 0 ? (
        <p className="text-xs text-slate-400 p-2">Không tìm thấy file trong thư mục public/2429.2026</p>
      ) : (
        renderTree(treeData)
      )}
    </div>
  );
};