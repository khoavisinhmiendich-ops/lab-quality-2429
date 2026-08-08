import { NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

export interface DocumentNode {
  id: string;
  title: string;
  fileName?: string;
  type?: string;
  path?: string;
  children?: DocumentNode[];
}

function getTree(dirPath: string, relativePath = ''): DocumentNode[] {
  if (!fs.existsSync(dirPath)) return [];

  // Lọc bỏ ngay các file tạm bắt đầu bằng ~$
  const items = fs.readdirSync(dirPath, { withFileTypes: true }).filter((item) => !item.name.startsWith('~$'));

  items.sort((a, b) =>
    a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' })
  );

  return items.map((item, index) => {
    const itemRelativePath = path.join(relativePath, item.name).replace(/\\/g, '/');
    const fullPath = path.join(dirPath, item.name);
    const id = `${relativePath}-${index}-${item.name}`.replace(/[^a-zA-Z0-9]/g, '_');

    if (item.isDirectory()) {
      return {
        id,
        title: item.name,
        children: getTree(fullPath, itemRelativePath),
      };
    } else {
      const ext = path.extname(item.name).toLowerCase().replace('.', '');
      return {
        id,
        title: item.name.replace(/\.[^/.]+$/, ''),
        fileName: item.name,
        type: ext,
        path: itemRelativePath,
      };
    }
  });
}

export async function GET() {
  const rootDir = path.join(process.cwd(), 'public', '2429.2026');
  const tree = getTree(rootDir, '2429.2026');
  return NextResponse.json(tree);
}