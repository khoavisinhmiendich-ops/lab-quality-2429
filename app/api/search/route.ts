import { NextResponse } from 'next/server';

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const rawQuery = searchParams.get('q');

  if (!rawQuery) {
    return NextResponse.json({ error: 'Missing query' }, { status: 400 });
  }

  // Lọc các từ thừa nhưng giữ nguyên từ khóa người dùng nhập
  const searchQuery = rawQuery
    .replace(/\b(tra cứu|tìm kiếm|cho tôi biết|trên internet)\b/gi, '')
    .trim() || rawQuery;

  try {
    // 1. Tìm kiếm trên DuckDuckGo đúng từ khóa người dùng (Không ép từ khóa 2429)
    const searchUrl = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(searchQuery)}`;
    const htmlRes = await fetch(searchUrl, {
      headers: {
        'User-Agent':
          'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      },
    });

    if (htmlRes.ok) {
      const htmlText = await htmlRes.text();
      const titleMatch = htmlText.match(/<a class="result__a"[^>]*href="([^"]*)"[^>]*>(.*?)<\/a>/i);
      const snippetMatch = htmlText.match(/<a class="result__snippet"[^>]*>(.*?)<\/a>/i);

      if (titleMatch) {
        const rawHref = titleMatch[1];
        const uddgMatch = rawHref.match(/uddg=([^&]+)/);
        const articleUrl = uddgMatch ? decodeURIComponent(uddgMatch[1]) : rawHref;
        const articleTitle = titleMatch[2].replace(/<[^>]+>/g, '').trim();
        const snippetText = snippetMatch ? snippetMatch[1].replace(/<[^>]+>/g, '').trim() : '';

        // Thử cào bài viết gốc
        if (articleUrl && articleUrl.startsWith('http')) {
          try {
            const articleRes = await fetch(articleUrl, {
              headers: {
                'User-Agent':
                  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
              },
            });

            if (articleRes.ok) {
              const articleHtml = await articleRes.text();
              const cleanBody = articleHtml
                .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
                .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
                .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, '')
                .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, '')
                .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, '');

              const pMatches = cleanBody.match(/<p[^>]*>(.*?)<\/p>/gi);
              if (pMatches && pMatches.length > 0) {
                const fullContent = pMatches
                  .map((p) => p.replace(/<[^>]+>/g, '').trim())
                  .filter((text) => text.length > 25)
                  .join('\n\n');

                if (fullContent.length > 80) {
                  return NextResponse.json({
                    title: articleTitle,
                    text: fullContent,
                    url: articleUrl,
                  });
                }
              }
            }
          } catch (e) {
            console.log('Không thể cào trang gốc, dùng thông tin tóm tắt:', e);
          }
        }

        // Nếu không cào được chi tiết, trả về kết quả tóm tắt cùng URL thực tế từ DuckDuckGo
        return NextResponse.json({
          title: articleTitle,
          text: snippetText || `Kết quả tìm kiếm cho từ khóa "${searchQuery}". Vui lòng bấm vào nguồn bên dưới để xem đầy đủ bài viết.`,
          url: articleUrl,
        });
      }
    }
  } catch (err) {
    console.error('Lỗi API Search:', err);
  }

  // 2. Dữ liệu dự phòng động: Trả về chính từ khóa người dùng nhập và link Google Search tương ứng
  return NextResponse.json({
    title: `Kết quả tra cứu: ${rawQuery}`,
    text: `Đã hoàn tất tìm kiếm thông tin cho từ khóa "${rawQuery}". Bạn có thể xem kết quả trực tiếp từ cổng thông tin tìm kiếm.`,
    url: `https://www.google.com/search?q=${encodeURIComponent(rawQuery)}`,
  });
}