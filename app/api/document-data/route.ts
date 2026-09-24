import { NextResponse } from 'next/server';
import { db } from '@/lib/firebase';
import {
  doc,
  getDoc,
  setDoc,
} from 'firebase/firestore';

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);
  const key = searchParams.get('key');

  if (!key) {
    return NextResponse.json(
      { error: 'Missing key' },
      { status: 400 }
    );
  }

  try {
    const docRef = doc(db, 'documents', key);
    const docSnap = await getDoc(docRef);

    if (docSnap.exists()) {
      return NextResponse.json({
        content: docSnap.data().content || null,
      });
    }

    return NextResponse.json({
      content: null,
    });
  } catch (err) {
    console.error('Lỗi đọc Firestore:', err);

    return NextResponse.json({
      content: null,
    });
  }
}

export async function POST(request: Request) {
  try {
    const {
      key,
      content,
      path: documentPath,
      fileName,
    } = await request.json();

    if (!key) {
      return NextResponse.json(
        { error: 'Missing key' },
        { status: 400 }
      );
    }

    /*
     * Không xóa dữ liệu Firestore chỉ vì content = null.
     * Nếu chưa có nội dung thì bỏ qua thao tác lưu.
     */
    if (content === null || content === undefined) {
      return NextResponse.json({
        success: true,
        skipped: true,
      });
    }

    const docRef = doc(db, 'documents', key);

    await setDoc(
      docRef,
      {
        content,
        path: documentPath || null,
        fileName: fileName || null,
        updatedAt: new Date().toISOString(),
      },
      {
        merge: true,
      }
    );

    return NextResponse.json({
      success: true,
    });
  } catch (err) {
    console.error('Lỗi ghi Firestore:', err);

    return NextResponse.json(
      {
        error: 'Failed to save to Firestore',
      },
      {
        status: 500,
      }
    );
  }
}