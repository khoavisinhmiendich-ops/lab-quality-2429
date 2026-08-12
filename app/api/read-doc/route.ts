import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

// Danh sách MIME type được phép, tránh trả về file tuỳ ý dạng octet-stream không kiểm soát
const MIME_TYPES: Record<string, string> = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.doc': 'application/msword',
  '.pdf': 'application/pdf',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
};

const PUBLIC_DIR = path.join(process.cwd(), 'public');

/**
 * Đệ quy tìm file theo tên (không phân biệt hoa/thường) trong một thư mục gốc.
 * Trả về đường dẫn tuyệt đối nếu tìm thấy, ngược lại null.
 */
function findFileRecursive(dir: string, targetFileName: string): string | null {
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
 * Kiểm tra một đường dẫn tuyệt đối có thực sự nằm bên trong baseDir hay không.
 * Chống path traversal (../../etc/passwd, v.v.)
 */
function isPathInsideBase(baseDir: string, targetPath: string): boolean {
  const relative = path.relative(baseDir, targetPath);
  return (
    relative !== '' &&
    !relative.startsWith('..') &&
    !path.isAbsolute(relative)
  );
}

function resolveFilePath(rawParam: string): string | null {
  // 1. Giải mã URL và loại bỏ ký tự "/" hoặc "\" ở đầu
  const decodedPath = decodeURIComponent(rawParam).replace(/^[/\\]+/, '');

  // 2. Thử đường dẫn trực tiếp bên trong thư mục public
  const directPath = path.resolve(PUBLIC_DIR, decodedPath);

  // Chặn path traversal ngay từ bước đầu
  if (!isPathInsideBase(PUBLIC_DIR, directPath)) {
    return null;
  }

  if (fs.existsSync(directPath) && fs.statSync(directPath).isFile()) {
    return directPath;
  }

  // 3. Không thấy theo path trực tiếp -> quét toàn bộ public theo tên file
  const fileNameOnly = path.basename(decodedPath);
  const foundPath = findFileRecursive(PUBLIC_DIR, fileNameOnly);

  if (foundPath && isPathInsideBase(PUBLIC_DIR, foundPath)) {
    return foundPath;
  }

  return null;
}

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const filePathParam = searchParams.get('file');

  if (!filePathParam) {
    return NextResponse.json({ error: 'Missing file parameter' }, { status: 400 });
  }

  let absolutePath: string | null;
  try {
    absolutePath = resolveFilePath(filePathParam);
  } catch (error) {
    console.error('[read-doc API] Invalid file parameter:', error);
    return NextResponse.json({ error: 'Invalid file parameter' }, { status: 400 });
  }

  if (!absolutePath) {
    console.error(`[read-doc API] File not found for param: ${filePathParam}`);
    return NextResponse.json({ error: 'File not found' }, { status: 404 });
  }

  const ext = path.extname(absolutePath).toLowerCase();
  const contentType = MIME_TYPES[ext];

  // Chỉ cho phép các định dạng đã khai báo, tránh phục vụ file hệ thống khác
  if (!contentType) {
    return NextResponse.json({ error: 'Unsupported file type' }, { status: 415 });
  }

  try {
    const fileBuffer = fs.readFileSync(absolutePath);
    const fileName = path.basename(absolutePath);

    return new NextResponse(fileBuffer, {
      status: 200,
      headers: {
        'Content-Type': contentType,
        'Content-Length': String(fileBuffer.length),
        'Content-Disposition': `inline; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        'Cache-Control': 'private, max-age=0, must-revalidate',
      },
    });
  } catch (error) {
    console.error('[read-doc API] Error reading file:', error);
    return NextResponse.json({ error: 'Error reading file' }, { status: 500 });
  }
}