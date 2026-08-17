import { NextResponse } from 'next/server';

/**
 * Các hàm dựng response chuẩn hoá dùng chung cho toàn bộ API route.
 * Đảm bảo mọi endpoint trả về cùng một cấu trúc lỗi / thành công,
 * giúp phía client xử lý nhất quán thay vì mỗi route một kiểu.
 */

export interface ApiErrorBody {
  error: string;
}

interface ApiSuccessInit {
  status?: number;
  headers?: HeadersInit;
}

/** Trả về response lỗi JSON đồng nhất: `{ error: string }` */
export function apiError(message: string, status = 400): NextResponse<ApiErrorBody> {
  return NextResponse.json<ApiErrorBody>({ error: message }, { status });
}

/** Trả về response thành công JSON, giữ nguyên kiểu dữ liệu truyền vào */
export function apiSuccess<T>(data: T, init?: ApiSuccessInit): NextResponse<T> {
  return NextResponse.json<T>(data, {
    status: init?.status ?? 200,
    headers: init?.headers,
  });
}