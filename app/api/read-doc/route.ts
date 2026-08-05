import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import mammoth from 'mammoth';

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const rawPath = searchParams.get('path');

  if (!rawPath) {
    return NextResponse.json({ error: 'Thiếu đường dẫn file' }, { status: 400 });
  }

  try {
    let cleanPath = decodeURIComponent(rawPath).replace(/\\/g, '/');
    cleanPath = cleanPath.replace(/^\/+/, '');
    while (cleanPath.startsWith('2429.2026/')) {
      cleanPath = cleanPath.replace('2429.2026/', '');
    }

    const fullPath = path.join(process.cwd(), 'public', '2429.2026', cleanPath);
    console.log('>>> Đường dẫn sau khi làm sạch:', fullPath);

    if (!fs.existsSync(fullPath)) {
      return NextResponse.json(
        { error: `Không tìm thấy file tại: ${fullPath}` },
        { status: 404 }
      );
    }

    const ext = path.extname(fullPath).toLowerCase();
    const fileName = path.basename(fullPath);

    if (ext === '.docx' || ext === '.doc') {
      let htmlContent = '';

      try {
        // Thử convert sang HTML trước
        // @ts-expect-error mammoth types mismatch fallback
        const result = await mammoth.convertToHtml({ path: fullPath });
        if (result && result.value && result.value.trim() !== '') {
          htmlContent = result.value;
        } else {
          // Thử trích xuất text thô nếu HTML trống
          // @ts-expect-error mammoth types mismatch fallback
          const rawResult = await mammoth.extractRawText({ path: fullPath });
          if (rawResult && rawResult.value && rawResult.value.trim() !== '') {
            const paragraphs = rawResult.value.split('\n').filter((p: string) => p.trim() !== '');
            htmlContent = paragraphs.map((p: string) => `<p style="margin-bottom: 10px; text-align: justify;">${p}</p>`).join('');
          }
        }
      } catch (err) {
        console.error('Lỗi khi đọc file bằng mammoth:', err);
      }

      // Nếu file trống hoặc cấu trúc không đọc được chữ, hiển thị khung chuẩn kèm tên file để nhập liệu ngay
      if (!htmlContent || htmlContent.trim() === '') {
        htmlContent = `
          <h3 style="text-align: center; font-weight: bold; margin-bottom: 15px; text-transform: uppercase;">BIỂU MẪU QUẢN LÝ CHẤT LƯỢNG</h3>
          <p><b>Tên tài liệu:</b> ${fileName}</p>
          <hr style="margin: 15px 0; border: 0; border-top: 1px solid #ccc;" />
          <p><b>1. Mục đích:</b></p>
          <p style="padding-left: 20px; color: #555;">[Nhập mục đích áp dụng của tài liệu tại đây...]</p>
          <p><b>2. Nội dung chi tiết:</b></p>
          <p style="padding-left: 20px; color: #555;">[Nhập nội dung chi tiết hoặc dán dữ liệu biểu mẫu tại đây...]</p>
        `;
      }

      const wrappedHtml = `
        <div style="color: #000000 !important; font-family: 'Times New Roman', Times, serif !important; font-size: 15px !important; line-height: 1.6; min-height: 500px; display: block !important; visibility: visible !important; opacity: 1 !important;">
          ${htmlContent}
        </div>
      `;

      return NextResponse.json({ success: true, html: wrappedHtml });
    }

    return NextResponse.json({ error: 'Định dạng file không hỗ trợ' }, { status: 400 });
  } catch (error) {
    console.error('Lỗi API read-doc:', error);
    return NextResponse.json(
      {
        error: 'Không thể xử lý file này.',
        details: (error as Error).message,
      },
      { status: 500 }
    );
  }
}