/**
 * Kiểu dữ liệu và hằng số dùng chung liên quan đến tài liệu (Word/PDF/Excel)
 * được các route `document-tree` và `read-doc` cùng sử dụng.
 */

export interface DocumentNode {
  id: string;
  title: string;
  fileName?: string;
  type?: string;
  path?: string;
  children?: DocumentNode[];
}

/**
 * Danh sách MIME type được phép phục vụ qua API đọc file.
 * Chỉ những phần mở rộng khai báo ở đây mới được trả về — tránh phục vụ
 * file hệ thống hoặc định dạng tuỳ ý dưới dạng octet-stream không kiểm soát.
 */
export const DOCUMENT_MIME_TYPES: Record<string, string> = {
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.doc': 'application/msword',
  '.pdf': 'application/pdf',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
};