import { NextResponse } from 'next/server';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

// ---------- Types ----------

interface SearchResult {
  title: string;
  text: string;
  url: string;
  source: 'article' | 'snippet' | 'fallback';
}

// ---------- Constants ----------

const USER_AGENT =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

const FETCH_TIMEOUT_MS = 8000;
const STOP_WORDS_PATTERN = /\b(tra cứu|tìm kiếm|cho tôi biết|trên internet)\b/gi;
const MIN_PARAGRAPH_LENGTH = 25;
const MIN_ARTICLE_LENGTH = 80;
const MAX_ARTICLE_LENGTH = 6000;

// ---------- Helpers ----------

/** Loại bỏ các cụm từ dẫn dắt không cần thiết nhưng giữ nguyên ý định tìm kiếm của người dùng */
function normalizeQuery(rawQuery: string): string {
  const cleaned = rawQuery.replace(STOP_WORDS_PATTERN, '').trim();
  return cleaned || rawQuery.trim();
}

/** fetch có timeout, tránh route bị treo khi nguồn ngoài phản hồi chậm */
async function fetchWithTimeout(url: string, timeoutMs = FETCH_TIMEOUT_MS): Promise<Response | null> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      headers: { 'User-Agent': USER_AGENT },
      signal: controller.signal,
    });
    return res;
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
  }
}

function stripTags(html: string): string {
  return html.replace(/<[^>]+>/g, '').trim();
}

function decodeHtmlEntities(text: string): string {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}

interface DdgResult {
  url: string;
  title: string;
  snippet: string;
}

/** Lấy kết quả tự nhiên đầu tiên (bỏ qua quảng cáo) từ trang HTML của DuckDuckGo */
function parseFirstDdgResult(html: string): DdgResult | null {
  const linkPattern = /<a class="result__a"[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi;
  const snippetPattern = /<a class="result__snippet"[^>]*>([\s\S]*?)<\/a>/gi;

  const links = [...html.matchAll(linkPattern)];
  const snippets = [...html.matchAll(snippetPattern)];

  if (links.length === 0) return null;

  const [, rawHref, rawTitle] = links[0];
  const uddgMatch = rawHref.match(/uddg=([^&]+)/);
  const url = uddgMatch ? decodeURIComponent(uddgMatch[1]) : rawHref;

  return {
    url,
    title: decodeHtmlEntities(stripTags(rawTitle)),
    snippet: snippets[0] ? decodeHtmlEntities(stripTags(snippets[0][1])) : '',
  };
}

/** Loại bỏ các khối không chứa nội dung chính (script, style, header, footer, nav) */
function stripNonContentBlocks(html: string): string {
  const blockTags = ['script', 'style', 'header', 'footer', 'nav', 'noscript'];
  return blockTags.reduce((acc, tag) => {
    const pattern = new RegExp(`<${tag}\\b[^<]*(?:(?!<\\/${tag}>)<[^<]*)*<\\/${tag}>`, 'gi');
    return acc.replace(pattern, '');
  }, html);
}

/** Trích các đoạn <p> đủ dài để tạo thành nội dung bài viết */
function extractArticleText(html: string): string | null {
  const cleaned = stripNonContentBlocks(html);
  const paragraphs = cleaned.match(/<p[^>]*>([\s\S]*?)<\/p>/gi);
  if (!paragraphs || paragraphs.length === 0) return null;

  const text = paragraphs
    .map((p) => decodeHtmlEntities(stripTags(p)))
    .filter((t) => t.length > MIN_PARAGRAPH_LENGTH)
    .join('\n\n')
    .slice(0, MAX_ARTICLE_LENGTH);

  return text.length > MIN_ARTICLE_LENGTH ? text : null;
}

/** Cố gắng lấy toàn văn bài viết gốc; trả null nếu thất bại ở bất kỳ bước nào */
async function tryFetchFullArticle(url: string): Promise<string | null> {
  if (!url.startsWith('http')) return null;

  const res = await fetchWithTimeout(url);
  if (!res || !res.ok) return null;

  try {
    const html = await res.text();
    return extractArticleText(html);
  } catch {
    return null;
  }
}

function fallbackResult(rawQuery: string): SearchResult {
  return {
    title: `Kết quả tra cứu: ${rawQuery}`,
    text: `Đã hoàn tất tìm kiếm thông tin cho từ khóa "${rawQuery}". Bạn có thể xem kết quả trực tiếp từ cổng thông tin tìm kiếm.`,
    url: `https://www.google.com/search?q=${encodeURIComponent(rawQuery)}`,
    source: 'fallback',
  };
}

// ---------- Route Handler ----------

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const rawQuery = searchParams.get('q')?.trim();

  if (!rawQuery) {
    return NextResponse.json({ error: 'Missing query parameter "q"' }, { status: 400 });
  }

  const searchQuery = normalizeQuery(rawQuery);

  try {
    const searchUrl = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(searchQuery)}`;
    const searchRes = await fetchWithTimeout(searchUrl);

    if (searchRes?.ok) {
      const html = await searchRes.text();
      const ddgResult = parseFirstDdgResult(html);

      if (ddgResult) {
        const fullText = await tryFetchFullArticle(ddgResult.url);

        const result: SearchResult = fullText
          ? { title: ddgResult.title, text: fullText, url: ddgResult.url, source: 'article' }
          : {
              title: ddgResult.title,
              text:
                ddgResult.snippet ||
                `Kết quả tìm kiếm cho từ khóa "${searchQuery}". Vui lòng bấm vào nguồn bên dưới để xem đầy đủ bài viết.`,
              url: ddgResult.url,
              source: 'snippet',
            };

        return NextResponse.json(result, {
          headers: { 'Cache-Control': 'private, max-age=300' },
        });
      }
    }
  } catch (err) {
    console.error('[search API] Unexpected error:', err);
  }

  // Không lấy được kết quả từ DuckDuckGo -> trả về dự phòng động
  return NextResponse.json(fallbackResult(rawQuery), {
    headers: { 'Cache-Control': 'no-store' },
  });
}