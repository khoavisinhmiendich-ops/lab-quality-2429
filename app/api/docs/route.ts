import { NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

export interface DocNode {
  id: string;
  title: string;
  isFolder: boolean;
  path?: string;
  code?: string;
  children?: DocNode[];
}

function scanDirectoryRecursively(dirPath: string, relativePath = ''): DocNode[] {
  if (!fs.existsSync(dirPath)) return [];

  const items = fs.readdirSync(dirPath);
  const nodes: DocNode[] = [];

  // Sắp xếp thư mục lên trước, file theo sau
  items.sort((a, b) => {
    const aPath = path.join(dirPath, a);
    const bPath = path.join(dirPath, b);
    const aIsDir = fs.statSync(aPath).isDirectory();
    const bIsDir = fs.statSync(bPath).isDirectory();
    if (aIsDir && !bIsDir) return -1;
    if (!aIsDir && bIsDir) return 1;
    return a.localeCompare(b, 'vi', { numeric: true });
  });

  for (const item of items) {
    // Bỏ qua file ẩn hệ thống (ví dụ: .DS_Store, ~$Word)
    if (item.startsWith('.') || item.startsWith('~$')) continue;

    const fullPath = path.join(dirPath, item);
    const relPath = path.join(relativePath, item);
    const stat = fs.statSync(fullPath);

    if (stat.isDirectory()) {
      const subChildren = scanDirectoryRecursively(fullPath, relPath);
      // Chỉ thêm thư mục nếu nó chứa nội dung bên trong
      nodes.push({
        id: relPath,
        title: item,
        isFolder: true,
        children: subChildren,
      });
    } else {
      const ext = path.extname(item).toLowerCase();
      // Nhận toàn bộ định dạng văn bản & biểu mẫu có trong kho 2429
      if (['.docx', '.doc', '.pdf', '.xls', '.xlsx', '.txt'].includes(ext)) {
        nodes.push({
          id: relPath,
          title: item,
          code: item.split(' ')[0] || '2429',
          isFolder: false,
          path: `/2429.2026/${relPath.replace(/\\/g, '/')}`,
        });
      }
    }
  }

  return nodes;
}

export async function GET() {
  try {
    const docsDir = path.join(process.cwd(), 'public', '2429.2026');
    const tree = scanDirectoryRecursively(docsDir);
    return NextResponse.json({ success: true, tree });
  } catch (error) {
    return NextResponse.json({ success: false, error: (error as Error).message }, { status: 500 });
  }
}