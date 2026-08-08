import { NextResponse } from 'next/server';

const globalStore: Record<string, string> = {};

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const key = searchParams.get('key');

  if (!key) {
    return NextResponse.json({ error: 'Missing key' }, { status: 400 });
  }

  const content = globalStore[key] || null;
  return NextResponse.json({ content });
}

export async function POST(request: Request) {
  try {
    const { key, content } = await request.json();

    if (!key) {
      return NextResponse.json({ error: 'Missing key' }, { status: 400 });
    }

    globalStore[key] = content;
    return NextResponse.json({ success: true });
  } catch (err) {
    console.error('Lỗi lưu trữ:', err);
    return NextResponse.json({ error: 'Invalid request' }, { status: 500 });
  }
}