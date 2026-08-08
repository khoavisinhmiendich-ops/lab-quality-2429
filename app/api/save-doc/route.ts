import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import HTMLtoDOCX from 'html-to-docx';

export async function POST(request: NextRequest) {
  try {
    const { filePath, htmlContent } = await request.json();

    if (!filePath || !htmlContent) {
      return NextResponse.json({ error: 'Thiếu tham số dữ liệu' }, { status: 400 });
    }

    const decodedPath = decodeURIComponent(filePath).replace(/^[/\\]+/, '');
    const absolutePath = path.join(process.cwd(), 'public', decodedPath);

    if (!fs.existsSync(absolutePath)) {
      return NextResponse.json({ error: 'Không tìm thấy file để ghi đè' }, { status: 404 });
    }

    // Chuyển đổi HTML sang Buffer
    const docxBuffer = (await HTMLtoDOCX(htmlContent, null, {
      table: { row: { cantSplit: true } },
      footer: true,
      pageNumber: true,
    })) as Buffer;

    // Ghi đè file lên ổ đĩa
    fs.writeFileSync(absolutePath, docxBuffer);

    return NextResponse.json({ success: true, message: 'Đã lưu tự động thành công' });
  } catch (error) {
    console.error('Lỗi khi lưu file docx:', error);
    return NextResponse.json({ error: 'Lỗi server khi lưu file' }, { status: 500 });
  }
}