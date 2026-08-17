import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import { apiError } from '@/lib/api-response';
import { resolveExistingFilePath } from '@/lib/fs-security';
import { DOCUMENT_MIME_TYPES } from '@/lib/document-types';

export const runtime = 'nodejs';

/** Phục vụ file Word/PDF/Excel tĩnh từ `/public` theo đường dẫn hoặc tên file, có kiểm soát bảo mật */
export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const filePathParam = searchParams.get('file');

  if (!filePathParam) {
    return apiError('Missing file parameter', 400);
  }

  let absolutePath: string | null;
  try {
    absolutePath = resolveExistingFilePath(filePathParam);
  } catch (error) {
    console.error('[read-doc API] Invalid file parameter:', error);
    return apiError('Invalid file parameter', 400);
  }

  if (!absolutePath) {
    console.error(`[read-doc API] File not found for param: ${filePathParam}`);
    return apiError('File not found', 404);
  }

  const ext = path.extname(absolutePath).toLowerCase();
  const contentType = DOCUMENT_MIME_TYPES[ext];

  // Chỉ cho phép các định dạng đã khai báo, tránh phục vụ file hệ thống khác
  if (!contentType) {
    return apiError('Unsupported file type', 415);
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
    return apiError('Error reading file', 500);
  }
}