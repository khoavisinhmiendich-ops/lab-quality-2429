/**
 * Tiện ích gọi HTTP ra bên ngoài dùng chung — hiện chỉ phục vụ route `search`,
 * tách riêng để dễ tái sử dụng nếu có thêm route cần gọi nguồn ngoài sau này.
 */

/** User-Agent trình duyệt, tránh bị một số nguồn chặn request không giống trình duyệt thật */
export const BROWSER_USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

const DEFAULT_TIMEOUT_MS = 8000;

/** fetch có timeout, tránh route bị treo khi nguồn ngoài phản hồi chậm hoặc không phản hồi */
export async function fetchWithTimeout(url: string, timeoutMs = DEFAULT_TIMEOUT_MS): Promise<Response | null> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    return await fetch(url, {
      headers: { 'User-Agent': BROWSER_USER_AGENT },
      signal: controller.signal,
    });
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
  }
}