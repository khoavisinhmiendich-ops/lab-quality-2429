import fs from 'fs';
import path from 'path';

/**
 * Tiện ích dùng chung để thao tác an toàn với hệ thống file trong thư mục `public`.
 * Mọi API route đọc/ghi file tĩnh (Word, PDF, Excel...) đều phải đi qua các hàm ở
 * đây để chống path traversal (vd: `../../etc/passwd`) — không tự viết lại logic
 * kiểm tra đường dẫn riêng lẻ ở từng route.
 */

/** Thư mục gốc duy nhất được phép đọc/ghi file thông qua các API này */
export const PUBLIC_DIR = path.join(process.cwd(), 'public');

/** Kiểm tra một đường dẫn tuyệt đối có thực sự nằm bên trong baseDir hay không */
export function isPathInsideBase(baseDir: string, targetPath: string): boolean {
  const relative = path.relative(baseDir, targetPath);
  return relative !== '' && !relative.startsWith('..') && !path.isAbsolute(relative);
}

/**
 * Đệ quy tìm file theo tên (không phân biệt hoa/thường) trong một thư mục gốc.
 * Trả về đường dẫn tuyệt đối nếu tìm thấy, ngược lại `null`.
 */
export function findFileRecursive(dir: string, targetFileName: string): string | null {
  let items: fs.Dirent[];
  try {
    items = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return null;
  }

  const lowerTarget = targetFileName.toLowerCase();

  // Duyệt file trước, thư mục sau để tránh đệ quy sâu không cần thiết
  for (const item of items) {
    if (!item.isDirectory() && item.name.toLowerCase() === lowerTarget) {
      return path.join(dir, item.name);
    }
  }
  for (const item of items) {
    if (item.isDirectory()) {
      const found = findFileRecursive(path.join(dir, item.name), targetFileName);
      if (found) return found;
    }
  }
  return null;
}

/**
 * Chuẩn hoá + xác thực một đường dẫn tương đối do client cung cấp, đảm bảo nó
 * luôn phân giải vào bên trong `PUBLIC_DIR`. Dùng cho MỌI thao tác đọc hoặc GHI
 * file dựa trên input người dùng. Trả về đường dẫn tuyệt đối hợp lệ, hoặc `null`
 * nếu phát hiện path traversal.
 */
export function resolveSafePath(rawParam: string): string | null {
  const decodedPath = decodeURIComponent(rawParam).replace(/^[/\\]+/, '');
  const directPath = path.resolve(PUBLIC_DIR, decodedPath);

  if (!isPathInsideBase(PUBLIC_DIR, directPath)) {
    return null;
  }

  return directPath;
}

/**
 * Giống `resolveSafePath` nhưng dùng cho mục đích ĐỌC file: nếu không tìm thấy
 * theo đường dẫn trực tiếp, sẽ quét toàn bộ `PUBLIC_DIR` theo tên file làm
 * phương án dự phòng (hữu ích khi cấu trúc thư mục phía client không khớp 100%
 * với ổ đĩa). Chỉ trả về đường dẫn khi file thực sự tồn tại và hợp lệ.
 */
export function resolveExistingFilePath(rawParam: string): string | null {
  const directPath = resolveSafePath(rawParam);

  if (directPath && fs.existsSync(directPath) && fs.statSync(directPath).isFile()) {
    return directPath;
  }

  const fileNameOnly = path.basename(decodeURIComponent(rawParam));
  const foundPath = findFileRecursive(PUBLIC_DIR, fileNameOnly);

  if (foundPath && isPathInsideBase(PUBLIC_DIR, foundPath)) {
    return foundPath;
  }

  return null;
}