import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';

// Hàm đệ quy quét và tìm file trong thư mục public nếu truyền thiếu path
function findFileRecursive(dir: string, targetFileName: string): string | null {
  if (!fs.existsSync(dir)) return null;

  const items = fs.readdirSync(dir, { withFileTypes: true });

  for (const item of items) {
    const fullPath = path.join(dir, item.name);
    if (item.isDirectory()) {
      const found = findFileRecursive(fullPath, targetFileName);
      if (found) return found;
    } else if (item.name.toLowerCase() === targetFileName.toLowerCase()) {
      return fullPath;
    }
  }

  return null;
}

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const filePathParam = searchParams.get('file');

  if (!filePathParam) {
    return NextResponse.json({ error: 'Missing file parameter' }, { status: 400 });
  }

  // 1. Giải mã URL (xử lý %20, dấu tiếng Việt) và xóa dấu / ở đầu
  const decodedPath = decodeURIComponent(filePathParam).replace(/^[/\\]+/, '');
  const publicDir = path.join(process.cwd(), 'public');

  // 2. Thử đường dẫn trực tiếp (Ví dụ: public/2429.2026/1. Chuong I.../file.docx)
  let absolutePath = path.join(publicDir, decodedPath);

  // 3. Nếu không tìm thấy, tự động quét toàn bộ thư mục public để tìm tên file
  if (!fs.existsSync(absolutePath)) {
    const fileNameOnly = path.basename(decodedPath);
    const foundPath = findFileRecursive(publicDir, fileNameOnly);

    if (foundPath) {
      absolutePath = foundPath;
    } else {
      console.error(`[read-doc API] File not found: ${absolutePath}`);
      return NextResponse.json({ error: 'File not found' }, { status: 404 });
    }
  }

  try {
    const fileBuffer = fs.readFileSync(absolutePath);
    const ext = path.extname(absolutePath).toLowerCase();

    // Xác định Header loại file
    let contentType = 'application/octet-stream';
    if (ext === '.docx') {
      contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    } else if (ext === '.pdf') {
      contentType = 'application/pdf';
    } else if (ext === '.xlsx') {
      contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }

    return new NextResponse(fileBuffer, {
      status: 200,
      headers: {
        'Content-Type': contentType,
        'Content-Disposition': `inline; filename="${encodeURIComponent(path.basename(absolutePath))}"`,
      },
    });
  } catch (error) {
    console.error('[read-doc API] Error reading file:', error);
    return NextResponse.json({ error: 'Error reading file' }, { status: 500 });
  }
}