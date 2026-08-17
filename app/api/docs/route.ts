import fs from 'fs';
import path from 'path';
import { apiSuccess } from '@/lib/api-response';
import type { DocumentNode } from '@/lib/document-types';

export const runtime = 'nodejs';

/** Thư mục gốc chứa toàn bộ tài liệu QĐ-2429/BYT, nằm trong /public để phục vụ tĩnh */
const DOCUMENT_ROOT_DIR = path.join(process.cwd(), 'public', '2429.2026');
const DOCUMENT_ROOT_RELATIVE = '2429.2026';

/** Bỏ qua các file tạm của Microsoft Office khi đang mở (dạng `~$Tên file.docx`) */
function isTemporaryOfficeFile(fileName: string): boolean {
  return fileName.startsWith('~$');
}

/**
 * Dựng cây thư mục/tài liệu đệ quy từ hệ thống file.
 * Thư mục → node có `children`; File → node có `fileName`, `type`, `path`.
 * Sắp xếp tự nhiên (numeric) theo tên để "2." đứng trước "10.".
 */
function buildDocumentTree(dirPath: string, relativePath = ''): DocumentNode[] {
  if (!fs.existsSync(dirPath)) return [];

  const items = fs
    .readdirSync(dirPath, { withFileTypes: true })
    .filter((item) => !isTemporaryOfficeFile(item.name))
    .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));

  return items.map((item, index) => {
    const itemRelativePath = path.join(relativePath, item.name).replace(/\\/g, '/');
    const fullPath = path.join(dirPath, item.name);
    const id = `${relativePath}-${index}-${item.name}`.replace(/[^a-zA-Z0-9]/g, '_');

    if (item.isDirectory()) {
      return {
        id,
        title: item.name,
        children: buildDocumentTree(fullPath, itemRelativePath),
      };
    }

    const ext = path.extname(item.name).toLowerCase().replace('.', '');
    return {
      id,
      title: item.name.replace(/\.[^/.]+$/, ''),
      fileName: item.name,
      type: ext,
      path: itemRelativePath,
    };
  });
}

/** Trả về cây thư mục/tài liệu hiện có trên đĩa, dùng để dựng FolderTree phía client */
export async function GET() {
  const tree = buildDocumentTree(DOCUMENT_ROOT_DIR, DOCUMENT_ROOT_RELATIVE);
  return apiSuccess<DocumentNode[]>(tree);
}