import fs from 'fs';
import path from 'path';
import crypto from 'crypto';

import { apiSuccess } from '@/lib/api-response';
import type { DocumentNode } from '@/lib/document-types';

export const runtime = 'nodejs';

/**
 * Thư mục gốc chứa toàn bộ tài liệu QĐ-2429/BYT,
 * nằm trong /public để phục vụ tĩnh.
 */
const DOCUMENT_ROOT_DIR = path.join(
  process.cwd(),
  'public',
  '2429.2026'
);

const DOCUMENT_ROOT_RELATIVE = '2429.2026';

/**
 * Tạo ID ổn định từ đường dẫn tương đối của tài liệu.
 *
 * QUAN TRỌNG:
 * - Không dùng index.
 * - Không dùng vị trí của file trong danh sách.
 * - Không phụ thuộc thứ tự sắp xếp.
 *
 * Vì vậy khi thêm file mới, ID của các file cũ
 * vẫn giữ nguyên.
 */
function createStableId(relativePath: string): string {
  const normalizedPath = relativePath
    .replace(/\\/g, '/')
    .trim()
    .toLowerCase();

  return crypto
    .createHash('sha256')
    .update(normalizedPath, 'utf8')
    .digest('hex')
    .slice(0, 32);
}

/**
 * Bỏ qua các file tạm của Microsoft Office
 * dạng ~$Tên file.docx
 */
function isTemporaryOfficeFile(fileName: string): boolean {
  return fileName.startsWith('~$');
}

/**
 * Dựng cây thư mục/tài liệu đệ quy từ hệ thống file.
 *
 * Thư mục:
 *   node có children
 *
 * File:
 *   node có fileName, type, path
 *
 * Sắp xếp tự nhiên để:
 *   2. đứng trước 10.
 */
function buildDocumentTree(
  dirPath: string,
  relativePath = ''
): DocumentNode[] {
  if (!fs.existsSync(dirPath)) {
    return [];
  }

  const items = fs
    .readdirSync(dirPath, {
      withFileTypes: true,
    })
    .filter(
      (item) => !isTemporaryOfficeFile(item.name)
    )
    .sort((a, b) =>
      a.name.localeCompare(
        b.name,
        undefined,
        {
          numeric: true,
          sensitivity: 'base',
        }
      )
    );

  return items.map((item) => {
    const itemRelativePath = path
      .join(relativePath, item.name)
      .replace(/\\/g, '/');

    const fullPath = path.join(
      dirPath,
      item.name
    );

    /**
     * ID ỔN ĐỊNH
     *
     * Không sử dụng index.
     */
    const id = createStableId(
      itemRelativePath
    );

    /**
     * Thư mục
     */
    if (item.isDirectory()) {
      return {
        id,
        title: item.name,
        children: buildDocumentTree(
          fullPath,
          itemRelativePath
        ),
      };
    }

    /**
     * File
     */
    const ext = path
      .extname(item.name)
      .toLowerCase()
      .replace('.', '');

    return {
      id,
      title: item.name.replace(
        /\.[^/.]+$/,
        ''
      ),
      fileName: item.name,
      type: ext,
      path: itemRelativePath,
    };
  });
}

/**
 * Trả về cây thư mục/tài liệu hiện có trên đĩa,
 * dùng để dựng FolderTree phía client.
 */
export async function GET() {
  const tree = buildDocumentTree(
    DOCUMENT_ROOT_DIR,
    DOCUMENT_ROOT_RELATIVE
  );

  return apiSuccess<DocumentNode[]>(
    tree
  );
}