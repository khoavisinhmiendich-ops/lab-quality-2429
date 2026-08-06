import { NextResponse } from "next/server";
import mammoth from "mammoth";
import fs from "fs";
import path from "path";

export async function GET(request: Request) {
  try {
    const { searchParams } = new URL(request.url);
    const docPath = searchParams.get("path");

    if (!docPath) {
      return NextResponse.json({ error: "Missing path parameter" }, { status: 400 });
    }

    // Xác định đường dẫn file trong thư mục public
    const filePath = path.join(process.cwd(), "public", docPath);

    if (!fs.existsSync(filePath)) {
      return NextResponse.json({ error: "File not found" }, { status: 404 });
    }

    // Đọc file từ ổ đĩa
    const buffer = fs.readFileSync(filePath);
    
    // Chuyển đổi Buffer sang ArrayBuffer chuẩn cho Mammoth
    const arrayBuffer = buffer.buffer.slice(
      buffer.byteOffset,
      buffer.byteOffset + buffer.byteLength
    );

    // Chuyển đổi file Word (.docx) sang HTML
    const result = await mammoth.convertToHtml({ arrayBuffer });
    let html = result.value;

    // ⚡ XỬ LÝ TRIỆT ĐỂ 100%: Xóa bỏ toàn bộ thuộc tính gây xô lệch lề bảng A4
    html = html
      // 1. Xóa toàn bộ thẻ <col> quy định chiều rộng cột cố định
      .replace(/<col[^>]*>/gi, "")
      // 2. Xóa các thuộc tính width và height trực tiếp trên các thẻ HTML
      .replace(/\s*width="[^"]*"/gi, "")
      .replace(/\s*height="[^"]*"/gi, "")
      // 3. Làm sạch thuộc tính style="..." inline
      .replace(/style="([^"]*)"/gi, (_, styleContent: string) => {
        const cleanedStyle = styleContent
          .split(";")
          .filter((rule: string) => {
            const prop = rule.split(":")[0]?.trim().toLowerCase();
            // Loại bỏ các thuộc tính định dạng kích thước & lề bị sai lệch của Word
            return ![
              "width",
              "min-width",
              "max-width",
              "height",
              "min-height",
              "max-height",
              "margin",
              "margin-left",
              "margin-right",
              "margin-top",
              "margin-bottom",
              "text-indent",
              "white-space",
            ].includes(prop);
          })
          .join(";");
        return cleanedStyle ? `style="${cleanedStyle}"` : "";
      });

    return NextResponse.json({ html });
  } catch (error) {
    console.error("Read doc error:", error);
    return NextResponse.json(
      { error: "Error reading document from server" },
      { status: 500 }
    );
  }
}