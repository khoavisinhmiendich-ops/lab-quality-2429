import { NextRequest } from 'next/server';
import fs from 'fs';
import HTMLtoDOCX from 'html-to-docx';
import { apiError, apiSuccess } from '@/lib/api-response';
import { resolveSafePath } from '@/lib/fs-security';

export const runtime = 'nodejs';

interface SaveDocxPayload {
  filePath?: string;
  htmlContent?: string;
}

/**
 * Ghi đè nội dung HTML đã chỉnh sửa lên file `.docx` gốc trong `/public`.
 * Dùng cho lưu vĩnh viễn xuống đĩa (khác với `/api/document-data`, vốn chỉ
 * lưu bản nháp HTML trên Firestore).
 *
 * LƯU Ý BẢO MẬT: `filePath` do client cung cấp — bắt buộc phải đi qua
 * `resolveSafePath` để chặn path traversal (vd: `filePath: "../../../etc/passwd"`)
 * trước khi ghi file. Trước đây route này ghép đường dẫn trực tiếp bằng
 * `path.join` mà không kiểm tra, cho phép ghi đè file ngoài `/public` — lỗi
 * đã được vá khi hợp nhất logic bảo mật dùng chung với `read-doc`.
 */
export async function POST(request: NextRequest) {
  let payload: SaveDocxPayload;

  try {
    payload = await request.json();
  } catch {
    return apiError('Invalid JSON body', 400);
  }

  const { filePath, htmlContent } = payload;

  if (!filePath || !htmlContent) {
    return apiError('Thiếu tham số dữ liệu', 400);
  }

  const absolutePath = resolveSafePath(filePath);

  if (!absolutePath) {
    return apiError('Đường dẫn file không hợp lệ', 400);
  }

  if (!fs.existsSync(absolutePath)) {
    return apiError('Không tìm thấy file để ghi đè', 404);
  }

  try {
    const docxBuffer = (await HTMLtoDOCX(htmlContent, null, {
      table: { row: { cantSplit: true } },
      footer: true,
      pageNumber: true,
    })) as Buffer;

    fs.writeFileSync(absolutePath, docxBuffer);

    return apiSuccess({ success: true, message: 'Đã lưu tự động thành công' });
  } catch (error) {
    console.error('[save-docx API] Lỗi khi lưu file docx:', error);
    return apiError('Lỗi server khi lưu file', 500);
  }
}