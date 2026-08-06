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

    const filePath = path.join(process.cwd(), "public", docPath);

    if (!fs.existsSync(filePath)) {
      return NextResponse.json({ error: "File not found" }, { status: 404 });
    }

    const buffer = fs.readFileSync(filePath);
    
    // Chuyển đổi Buffer sang ArrayBuffer chuẩn cho Mammoth
    const arrayBuffer = buffer.buffer.slice(
      buffer.byteOffset,
      buffer.byteOffset + buffer.byteLength
    );

    const result = await mammoth.convertToHtml({ arrayBuffer });
    let html = result.value;

    // Lọc bỏ triệt để các style cố định và lề âm của Word
    html = html
      .replace(/<col[^>]*>/gi, "")
      .replace(/\s*width="[^"]*"/gi, "")
      .replace(/\s*height="[^"]*"/gi, "")
      .replace(/style="([^"]*)"/gi, (match: string, styleContent: string) => {
        if (!styleContent) return "";
        const cleanedStyle = styleContent
          .split(";")
          .filter((rule: string) => {
            const prop = rule.split(":")[0]?.trim().toLowerCase();
            return (
              prop &&
              ![
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
              ].includes(prop)
            );
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