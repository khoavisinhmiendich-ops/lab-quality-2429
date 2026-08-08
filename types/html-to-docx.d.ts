declare module 'html-to-docx' {
  interface HTMLToDocxOptions {
    table?: {
      row?: {
        cantSplit?: boolean;
      };
    };
    footer?: boolean;
    pageNumber?: boolean;
    header?: boolean;
    orientation?: 'portrait' | 'landscape';
    margins?: {
      top?: number;
      right?: number;
      bottom?: number;
      left?: number;
    };
  }

  function HTMLtoDOCX(
    htmlString: string,
    headerHTMLString?: string | null,
    documentOptions?: HTMLToDocxOptions,
    footerHTMLString?: string | null
  ): Promise<Buffer>; // Đổi kiểu trả về thành Promise<Buffer>

  export default HTMLtoDOCX;
}