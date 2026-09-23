'use client';

import React, { useState, useEffect, useRef } from 'react';
import FolderTree, { DocumentNode } from '@/components/FolderTree';
import type { WorkBook, CellObject, utils as XLSXUtilsNamespace } from 'xlsx';

type XLSXUtils = typeof XLSXUtilsNamespace;


type ExcelCellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: 'left' | 'center' | 'right';
  color?: string;
  bg?: string;
  fontSize?: number;
  border?: { top?: string; right?: string; bottom?: string; left?: string };
};
type ExcelCellData = { r: number; c: number; text: string; style?: ExcelCellStyle };
type ExcelMergeData = { r: number; c: number; rowSpan: number; colSpan: number };
type ExcelSheetData = {
  name: string;
  rows: ExcelCellData[][];
  merges: Record<string, ExcelMergeData>;
  skip: Set<string>;
  startRow: number;
  startCol: number;
  endCol: number;
  error?: string;
};

const EXCEL_PAGE_HEIGHT_PX = 1122.52;
const EXCEL_HEADER_HEIGHT_PX = 26;
const EXCEL_ROW_DEFAULT_HEIGHT_PX = 30;
const EXCEL_PAGE_GAP_PX = 56;

function getExcelPageRangesStable(
  sheet: ExcelSheetData,
  sheetIndex: number,
  rowHeights: Record<string, number>,
): Array<{ start: number; end: number }> {
  if (!sheet.rows.length) return [];

  const getRowHeight = (row: number) =>
    rowHeights[`${sheetIndex}-${row}`] ?? EXCEL_ROW_DEFAULT_HEIGHT_PX;

  const ranges: Array<{ start: number; end: number }> = [];
  let start = 0;

  while (start < sheet.rows.length) {
    let end = start;
    let used = EXCEL_HEADER_HEIGHT_PX;

    while (end < sheet.rows.length) {
      const absRow = sheet.rows[end]?.[0]?.r ?? end;
      const rowHeight = getRowHeight(absRow);

      if (end > start && used + rowHeight > EXCEL_PAGE_HEIGHT_PX) break;
      used += rowHeight;
      end += 1;

      // Giữ nguyên merge dọc: không cắt merge giữa hai trang.
      let extended = true;
      while (extended) {
        extended = false;
        const lastAbsRow = sheet.rows[end - 1]?.[0]?.r ?? (end - 1);

        for (const merge of Object.values(sheet.merges)) {
          const mergeStart = merge.r;
          const mergeEnd = merge.r + merge.rowSpan - 1;
          if (mergeStart >= start && mergeStart <= lastAbsRow && mergeEnd >= end) {
            const targetEnd = Math.min(sheet.rows.length, mergeEnd + 1);
            while (end < targetEnd) {
              const extraAbsRow = sheet.rows[end]?.[0]?.r ?? end;
              used += getRowHeight(extraAbsRow);
              end += 1;
            }
            extended = true;
          }
        }
      }
    }

    if (end <= start) end = Math.min(start + 1, sheet.rows.length);
    ranges.push({ start, end });
    start = end;
  }

  return ranges;
}
type CellStyleXLSX = {
  font?: { bold?: boolean; italic?: boolean; underline?: boolean; sz?: number; color?: { rgb?: string } };
  fgColor?: { rgb?: string };
  bgColor?: { rgb?: string };
  alignment?: { horizontal?: string };
};

export default function HomePage() {
  // Trạng thái Pass Key
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [username, setUsername] = useState<string>('');
  const [password, setPassword] = useState<string>('');
  const [loginError, setLoginError] = useState<string>('');
  const [isLoggingOut, setIsLoggingOut] = useState<boolean>(false);
  // --- UI-only state (không ảnh hưởng logic nghiệp vụ) ---
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);
  const [showPassword, setShowPassword] = useState<boolean>(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);
  const [isMobileSidebarOpen, setIsMobileSidebarOpen] = useState<boolean>(false);
  // --- Kéo giãn / thu nhỏ chiều rộng sidebar bằng chuột (chỉ áp dụng desktop) ---
  const SIDEBAR_MIN_WIDTH = 220;
  const SIDEBAR_MAX_WIDTH = 480;
  const SIDEBAR_DEFAULT_WIDTH = 280;
  const [sidebarWidth, setSidebarWidth] = useState<number>(SIDEBAR_DEFAULT_WIDTH);
  const [isResizingSidebar, setIsResizingSidebar] = useState<boolean>(false);
  const sidebarRef = useRef<HTMLElement>(null);
  // --- Tìm kiếm tài liệu (điều khiển từ ô tìm kiếm ở Header, dùng chung với FolderTree) ---
  const [docSearchQuery, setDocSearchQuery] = useState<string>('');
  const headerSearchRef = useRef<HTMLInputElement>(null);
  // --- Đếm số từ + zoom cho tài liệu Word (chỉ hiển thị, không ảnh hưởng logic lưu/định dạng) ---
  const [wordCount, setWordCount] = useState<number>(0);
  const WORD_ZOOM_MIN = 60;
  const WORD_ZOOM_MAX = 150;
  const WORD_ZOOM_STEP = 10;
  const [wordZoom, setWordZoom] = useState<number>(100);
  // --- Tự động tách trang thật cho Word (chèn khoảng trống thật vào DOM, không phải overlay) ---
  // A4 CSS ở 96dpi: 210 x 297mm ≈ 793.7 x 1122.5px. Dùng giá trị gần đúng để
  // đồng bộ với min-h-[297mm] của vùng soạn thảo và tránh lỗi lệch trang theo từng lần render.
  const WORD_PAGE_HEIGHT_PX = 1122.52;
  const WORD_PAGE_GAP_PX = 56; // chiều cao khoảng trống thật giữa 2 trang
  const WORD_PAGE_GAP_ATTR = 'data-page-gap'; // đánh dấu khối ngăn trang để luôn loại bỏ trước khi lưu
  const WORD_PAGE_BADGE_ATTR = 'data-page-badge'; // nhãn đầu trang, chỉ hiển thị, không thuộc nội dung Word
  // --- Zoom cho trình xem ảnh (.jpg/.jpeg/.png/.gif/.webp) ---
  const IMAGE_ZOOM_MIN = 25;
  const IMAGE_ZOOM_MAX = 300;
  const IMAGE_ZOOM_STEP = 25;
  const [imageZoom, setImageZoom] = useState<number>(100);

  const [selectedFile, setSelectedFile] = useState<DocumentNode | null>(null);
  const [htmlContent, setHtmlContent] = useState<string>('');

  // Đếm số từ từ nội dung HTML — hàm thuần, gọi trực tiếp ở mọi nơi setHtmlContent()
  // thay vì dùng useEffect riêng (tránh setState trực tiếp trong effect).
  const computeWordCount = (html: string): number => {
    const plainText = html.replace(/<[^>]*>/g, ' ');
    return plainText.trim().length > 0 ? plainText.trim().split(/\s+/).length : 0;
  };

  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isSaved, setIsSaved] = useState<boolean>(true);
  const editorRef = useRef<HTMLDivElement>(null);
  const saveTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const savedRangeRef = useRef<Range | null>(null);
  // Bộ đếm ổn định cho Shape; không dùng Date.now()/Math.random() trong render/component code.
  const shapeIdRef = useRef(0);

  // Shapes trong trình soạn thảo Word: menu + kéo thả các khối đã chèn.
  const [isShapesMenuOpen, setIsShapesMenuOpen] = useState<boolean>(false);
  // Màu viền Shape: dùng cho Shape mới và có thể áp dụng ngay cho Shape đang chọn.
  const [shapeBorderColor, setShapeBorderColor] = useState<string>('#2563eb');
  const selectedShapeRef = useRef<HTMLElement | null>(null);
  const shapeDragRef = useRef<{
    el: HTMLElement;
    startX: number;
    startY: number;
    startLeft: number;
    startTop: number;
  } | null>(null);
  const shapeResizeRef = useRef<{
    el: HTMLElement;
    startX: number;
    startY: number;
    startWidth: number;
    startHeight: number;
    startLeft: number;
    startTop: number;
    direction: string;
  } | null>(null);
  const shapeRotateRef = useRef<{
    el: HTMLElement;
    centerX: number;
    centerY: number;
    startAngle: number;
    startRotation: number;
  } | null>(null);

  // Trạng thái thanh công cụ định dạng (ribbon) kiểu Word
  const [activeFormats, setActiveFormats] = useState<Record<string, boolean>>({});
  const [currentBlock, setCurrentBlock] = useState<string>('p');

  // Trạng thái xem file Excel (.xlsx)
  type CellStyle = ExcelCellStyle;
  type ExcelCell = ExcelCellData;
  type ExcelMerge = ExcelMergeData;
  type ExcelSheet = ExcelSheetData;
  type ExcelEditEntry = { text: string; style?: CellStyle };

  const [isExcelBorderMenuOpen, setIsExcelBorderMenuOpen] = useState(false);
  const [isExcelMergeMenuOpen, setIsExcelMergeMenuOpen] = useState(false);
  const [isWordBorderMenuOpen, setIsWordBorderMenuOpen] = useState(false);
  const [isWordMergeMenuOpen, setIsWordMergeMenuOpen] = useState(false);
  const [excelSheets, setExcelSheets] = useState<ExcelSheet[]>([]);
  const [activeSheet, setActiveSheet] = useState<number>(0);
  const [isExcelLoading, setIsExcelLoading] = useState<boolean>(false);
  const [selectedCell, setSelectedCell] = useState<ExcelCell | null>(null);
  const [isExcelSaved, setIsExcelSaved] = useState<boolean>(true);
  const [excelFormulaValue, setExcelFormulaValue] = useState<string>('');
  const excelHistoryRef = useRef<ExcelSheet[][]>([]);
  const excelRedoRef = useRef<ExcelSheet[][]>([]);
  const [excelSelection, setExcelSelection] = useState<{ sheetIdx: number; r1: number; c1: number; r2: number; c2: number } | null>(null);
  const excelSaveTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  const excelWorkbookRef = useRef<WorkBook | null>(null);
  const excelUtilsRef = useRef<XLSXUtils | null>(null);
  // --- Kéo giãn/thu nhỏ độ rộng cột & chiều cao dòng trong Excel bằng chuột (hoặc chạm) ---
  const EXCEL_COL_DEFAULT_WIDTH = 92;
  const EXCEL_COL_MIN_WIDTH = 56;
  const EXCEL_COL_MAX_WIDTH = 420;
  const EXCEL_ROW_DEFAULT_HEIGHT = 30;
  const EXCEL_ROW_MIN_HEIGHT = 22;
  const EXCEL_ROW_MAX_HEIGHT = 140;
  const [excelColWidths, setExcelColWidths] = useState<Record<string, number>>({});
  const [excelRowHeights, setExcelRowHeights] = useState<Record<string, number>>({});
  const [isResizingExcelCell, setIsResizingExcelCell] = useState<boolean>(false);
  const excelResizeRef = useRef<{
    type: 'col' | 'row';
    index: number;
    sheetIdx: number;
    startPos: number;
    startSize: number;
  } | null>(null);

  const getExcelColWidth = (sheetIdx: number, col: number): number =>
    excelColWidths[`${sheetIdx}-${col}`] ?? EXCEL_COL_DEFAULT_WIDTH;

  const getExcelRowHeight = (sheetIdx: number, row: number): number =>
    excelRowHeights[`${sheetIdx}-${row}`] ?? EXCEL_ROW_DEFAULT_HEIGHT;

  const startExcelColResize = (e: React.MouseEvent | React.TouchEvent, col: number) => {
    e.preventDefault();
    e.stopPropagation();
    const clientX = 'touches' in e ? e.touches[0]?.clientX ?? 0 : e.clientX;
    excelResizeRef.current = {
      type: 'col',
      index: col,
      sheetIdx: activeSheet,
      startPos: clientX,
      startSize: getExcelColWidth(activeSheet, col),
    };
    setIsResizingExcelCell(true);
  };

  const startExcelRowResize = (e: React.MouseEvent | React.TouchEvent, row: number) => {
    e.preventDefault();
    e.stopPropagation();
    const clientY = 'touches' in e ? e.touches[0]?.clientY ?? 0 : e.clientY;
    excelResizeRef.current = {
      type: 'row',
      index: row,
      sheetIdx: activeSheet,
      startPos: clientY,
      startSize: getExcelRowHeight(activeSheet, row),
    };
    setIsResizingExcelCell(true);
  };

  useEffect(() => {
    if (!isResizingExcelCell) return;

    const applyDelta = (clientPos: number) => {
      const info = excelResizeRef.current;
      if (!info) return;
      const delta = clientPos - info.startPos;
      if (info.type === 'col') {
        const next = Math.min(EXCEL_COL_MAX_WIDTH, Math.max(EXCEL_COL_MIN_WIDTH, Math.round(info.startSize + delta)));
        setExcelColWidths((prev) => ({ ...prev, [`${info.sheetIdx}-${info.index}`]: next }));
      } else {
        const next = Math.min(EXCEL_ROW_MAX_HEIGHT, Math.max(EXCEL_ROW_MIN_HEIGHT, Math.round(info.startSize + delta)));
        setExcelRowHeights((prev) => ({ ...prev, [`${info.sheetIdx}-${info.index}`]: next }));
      }
    };

    const handleMouseMove = (e: MouseEvent) => {
      applyDelta(excelResizeRef.current?.type === 'row' ? e.clientY : e.clientX);
    };
    const handleTouchMove = (e: TouchEvent) => {
      if (e.touches.length === 0) return;
      e.preventDefault();
      applyDelta(excelResizeRef.current?.type === 'row' ? e.touches[0].clientY : e.touches[0].clientX);
    };
    const stopResizing = () => {
      excelResizeRef.current = null;
      setIsResizingExcelCell(false);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', stopResizing);
    document.addEventListener('touchmove', handleTouchMove, { passive: false });
    document.addEventListener('touchend', stopResizing);
    document.addEventListener('touchcancel', stopResizing);
    if (editorRef.current) editorRef.current.style.cursor = excelResizeRef.current?.type === 'row' ? 'row-resize' : 'col-resize';
    if (editorRef.current) editorRef.current.style.userSelect = 'none';

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', stopResizing);
      document.removeEventListener('touchmove', handleTouchMove);
      document.removeEventListener('touchend', stopResizing);
      document.removeEventListener('touchcancel', stopResizing);
      if (editorRef.current) editorRef.current.style.cursor = '';
      if (editorRef.current) editorRef.current.style.userSelect = '';
    };
  }, [isResizingExcelCell]);

  // Chuyển chỉ số cột (0-based) thành chữ cái kiểu Excel: 0->A, 25->Z, 26->AA...
  const colLetter = (n: number): string => {
    let s = '';
    let num = n;
    while (num >= 0) {
      s = String.fromCharCode((num % 26) + 65) + s;
      num = Math.floor(num / 26) - 1;
    }
    return s;
  };

  const getFileType = (file: DocumentNode) => {
    if (file.type === 'pdf' || file.path?.toLowerCase().endsWith('.pdf')) {
      return 'pdf';
    }
    if (
      file.type === 'xlsx' ||
      file.type === 'xls' ||
      file.path?.toLowerCase().endsWith('.xlsx') ||
      file.path?.toLowerCase().endsWith('.xls')
    ) {
      return 'excel';
    }
    if (
      file.type === 'doc' ||
      file.type === 'docx' ||
      file.path?.toLowerCase().endsWith('.doc') ||
      file.path?.toLowerCase().endsWith('.docx')
    ) {
      return 'word';
    }
    if (
      ['jpg', 'jpeg', 'png', 'gif', 'webp'].includes((file.type || '').toLowerCase()) ||
      /\.(jpe?g|png|gif|webp)$/i.test(file.path || '')
    ) {
      return 'image';
    }
    return file.type || 'text';
  };

  // ---- Phát hiện định dạng file thật từ các byte đầu (chữ ký file) ----
  // .doc cũ (OLE2) bắt đầu bằng chữ ký cố định — dùng để phân biệt lỗi "không phải zip"
  // là do file .doc cũ bị đổi tên đuôi thành .docx, hay do lỗi khác.
  const isOle2Signature = (buf: ArrayBuffer): boolean => {
    const bytes = new Uint8Array(buf.slice(0, 8));
    const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    return sig.every((b, i) => bytes[i] === b);
  };

  // Tải & chuyển đổi file Word (.docx) — giữ nguyên logic gốc
  useEffect(() => {
    if (!selectedFile) return;

    const fileType = getFileType(selectedFile);
    if (fileType !== 'word' || !selectedFile.path) return;

    let isSubscribed = true;
    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;

    const loadDocument = async () => {
      if (isSubscribed) {
        setIsLoading(true);
        setWordZoom(100);
      }

      // Tải file gốc 1 lần duy nhất, tái sử dụng luôn cho mammoth nếu chưa có
      // nội dung thân bài đã lưu (tránh fetch lại 2 lần).
      let originalArrayBuffer: ArrayBuffer | null = null;
      try {
        const res = await fetch(selectedFile.path!);
        if (res.ok) {
          originalArrayBuffer = await res.arrayBuffer();
        }
      } catch (err) {
        console.error('Không thể tải file gốc:', err);
      }

      try {
        const cloudRes = await fetch(`/api/document-data?key=${encodeURIComponent(docKey)}`);
        const cloudData = await cloudRes.json();

        if (cloudData && cloudData.content) {
          if (isSubscribed) {
            setHtmlContent(cloudData.content);
            setWordCount(computeWordCount(cloudData.content));
            setIsSaved(true);
            setIsLoading(false);
          }
          return;
        }

        const savedLocal = localStorage.getItem(docKey);
        if (savedLocal) {
          if (isSubscribed) {
            setHtmlContent(savedLocal);
            setWordCount(computeWordCount(savedLocal));
            setIsSaved(true);
            setIsLoading(false);
          }
          return;
        }

        if (!originalArrayBuffer) throw new Error('Không thể tải file gốc');

        const mammoth = await import('mammoth');
        const result = await mammoth.convertToHtml({ arrayBuffer: originalArrayBuffer });

        if (isSubscribed) {
          setHtmlContent(result.value);
          setWordCount(computeWordCount(result.value));
          setIsSaved(true);
        }
      } catch (err) {
        console.error('Lỗi tải tài liệu:', err);
        if (isSubscribed) {
          // Phân biệt rõ nguyên nhân: file .doc cũ bị đổi tên đuôi thành .docx (không phải file zip thật)
          // so với các lỗi khác — để đưa ra hướng khắc phục chính xác thay vì thông báo chung chung.
          const isOldDocRenamed = originalArrayBuffer ? isOle2Signature(originalArrayBuffer) : false;

          const fallbackHtml = isOldDocRenamed
            ? `
              <div style="padding: 12px; color: #b91c1c; background: #fef2f2; border: 1px solid #fecaca; margin-bottom: 15px; border-radius: 8px; font-family: 'Times New Roman', Times, serif; font-size: 13pt;">
                ⚠️ <b>Không đọc được file:</b> Đây là file Word định dạng cũ (.doc, trước Word 2007) đã được đổi tên đuôi thành ".docx", nhưng nội dung bên trong CHƯA thực sự được chuyển đổi định dạng. Hệ thống chỉ đọc được file .docx thật (Word 2007 trở lên, bản chất là file .zip).
                <br/><br/>
                <b>Cách khắc phục:</b> Mở file gốc bằng Microsoft Word → menu "Tệp" → "Lưu dưới dạng khác" (Save As) → chọn định dạng "Word Document (*.docx)" → Lưu lại, sau đó tải file .docx thật này lên hệ thống thay cho file cũ.
              </div>
              <p style="font-family: 'Times New Roman', Times, serif; font-size: 13pt;">Nhập nội dung biểu mẫu tại đây...</p>
            `
            : `
              <div style="padding: 12px; color: #b45309; background: #fffbeb; border: 1px solid #fde68a; margin-bottom: 15px; border-radius: 8px; font-family: 'Times New Roman', Times, serif; font-size: 13pt;">
                ⚠️ <b>Lưu ý:</b> File gốc có cấu trúc mã nguồn cũ. Hệ thống đã mở chế độ soạn thảo trực tiếp. Bạn có thể nhập nội dung hoặc chỉnh sửa bình thường, dữ liệu sẽ tự động lưu lại.
              </div>
              <p style="font-family: 'Times New Roman', Times, serif; font-size: 13pt;">Nhập nội dung biểu mẫu tại đây...</p>
            `;
          setHtmlContent(fallbackHtml);
          setWordCount(computeWordCount(fallbackHtml));
          setIsSaved(true);
        }
      } finally {
        if (isSubscribed) {
          setIsLoading(false);
        }
      }
    };

    loadDocument();

    return () => {
      isSubscribed = false;
    };
  }, [selectedFile]);

  // Tải & chuyển đổi file Excel (.xlsx) — chỉ xem, không chỉnh sửa
  useEffect(() => {
    if (!selectedFile) return;

    const fileType = getFileType(selectedFile);
    if (fileType !== 'excel' || !selectedFile.path) return;

    let isSubscribed = true;
    const excelDocKey = `xlsx_${selectedFile.id || selectedFile.title || selectedFile.path}`;

    const loadExcel = async () => {
      if (isSubscribed) {
        setIsExcelLoading(true);
        setActiveSheet(0);
        setExcelSheets([]);
        setSelectedCell(null);
        setExcelSelection(null);
        excelHistoryRef.current = [];
        excelRedoRef.current = [];
        setIsExcelSaved(true);
        setExcelColWidths({});
        setExcelRowHeights({});
      }

      try {
        const res = await fetch(selectedFile.path!);
        if (!res.ok) throw new Error('Không thể tải file gốc');

        const arrayBuffer = await res.arrayBuffer();
        const XLSX = await import('xlsx');
        const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
        if (isSubscribed) excelWorkbookRef.current = workbook;
        excelUtilsRef.current = XLSX.utils;

        // Tải lại các chỉnh sửa đã lưu trước đó (nếu có) để áp dụng đè lên dữ liệu gốc
        let savedEdits: Record<string, Record<string, ExcelEditEntry>> | null = null;
        try {
          const cloudRes = await fetch(`/api/document-data?key=${encodeURIComponent(excelDocKey)}`);
          const cloudData = await cloudRes.json();
          if (cloudData && cloudData.content) {
            savedEdits = JSON.parse(cloudData.content);
          }
        } catch {
          // bỏ qua, sẽ thử localStorage
        }
        if (!savedEdits) {
          const savedLocal = localStorage.getItem(excelDocKey);
          if (savedLocal) {
            try {
              savedEdits = JSON.parse(savedLocal);
            } catch {
              savedEdits = null;
            }
          }
        }

        const sheets: ExcelSheet[] = workbook.SheetNames.map((name) => {
          const worksheet = workbook.Sheets[name];
          const ref = worksheet['!ref'] || 'A1:A1';
          const range = XLSX.utils.decode_range(ref);
          const rawMerges = worksheet['!merges'] || [];

          const merges: Record<string, ExcelMerge> = {};
          const skip = new Set<string>();
          rawMerges.forEach((m) => {
            merges[`${m.s.r}-${m.s.c}`] = {
              r: m.s.r,
              c: m.s.c,
              rowSpan: m.e.r - m.s.r + 1,
              colSpan: m.e.c - m.s.c + 1,
            };
            for (let r = m.s.r; r <= m.e.r; r++) {
              for (let c = m.s.c; c <= m.e.c; c++) {
                if (r === m.s.r && c === m.s.c) continue;
                skip.add(`${r}-${c}`);
              }
            }
          });

          const readCellStyle = (cellRaw: CellObject | undefined): CellStyle | undefined => {
            const s = cellRaw?.s as CellStyleXLSX | undefined;
            if (!s) return undefined;
            const style: CellStyle = {};
            if (s.font?.bold) style.bold = true;
            if (s.font?.italic) style.italic = true;
            if (s.font?.underline) style.underline = true;
            if (s.font?.sz) style.fontSize = s.font.sz;
            if (s.font?.color?.rgb) style.color = `#${String(s.font.color.rgb).slice(-6)}`;
            if (s.fgColor?.rgb || s.bgColor?.rgb) {
              const rgb = s.fgColor?.rgb || s.bgColor?.rgb;
              if (rgb && rgb !== 'FFFFFFFF') style.bg = `#${String(rgb).slice(-6)}`;
            }
            if (s.alignment?.horizontal === 'center') style.align = 'center';
            else if (s.alignment?.horizontal === 'right') style.align = 'right';
            return Object.keys(style).length > 0 ? style : undefined;
          };

          const sheetEdits = savedEdits ? savedEdits[name] : undefined;

          const rows: ExcelCell[][] = [];
          for (let r = range.s.r; r <= range.e.r; r++) {
            const rowCells: ExcelCell[] = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
              const addr = XLSX.utils.encode_cell({ r, c });
              const cell = worksheet[addr];
              const originalText = cell ? String(cell.f ? `=${cell.f}` : (cell.w ?? cell.v ?? '')) : '';
              const originalStyle = readCellStyle(cell);
              const editKey = `${r}-${c}`;
              const edit = sheetEdits ? sheetEdits[editKey] : undefined;
              const text = edit ? edit.text : originalText;
              const style = edit && edit.style !== undefined ? edit.style : originalStyle;
              rowCells.push({ r, c, text, style });
            }
            rows.push(rowCells);
          }

          return {
            name,
            rows,
            merges,
            skip,
            startRow: range.s.r,
            startCol: range.s.c,
            endCol: range.e.c,
          };
        });

        if (isSubscribed) {
          setExcelSheets(
            sheets.length > 0
              ? sheets
              : [{ name: 'Sheet1', rows: [], merges: {}, skip: new Set(), startRow: 0, startCol: 0, endCol: 0 }]
          );
        }
      } catch (err) {
        console.error('Lỗi tải bảng tính:', err);
        if (isSubscribed) {
          setExcelSheets([
            {
              name: 'Lỗi',
              rows: [],
              merges: {},
              skip: new Set(),
              startRow: 0,
              startCol: 0,
              endCol: 0,
              error: 'Không thể đọc nội dung file Excel này.',
            },
          ]);
        }
      } finally {
        if (isSubscribed) setIsExcelLoading(false);
      }
    };

    loadExcel();

    return () => {
      isSubscribed = false;
      if (excelSaveTimeoutRef.current) clearTimeout(excelSaveTimeoutRef.current);
    };
  }, [selectedFile]);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoggingIn(true);
    // Tài khoản mặc định: admin / 654321
    window.setTimeout(() => {
      if (username.trim() === 'admin' && password === '654321') {
        setIsAuthenticated(true);
        setLoginError('');
      } else {
        setLoginError('Tên đăng nhập hoặc mật khẩu Pass Key không chính xác!');
      }
      setIsLoggingIn(false);
    }, 350);
  };

  const handleLogout = () => {
    setIsLoggingOut(true);
    setTimeout(() => {
      setIsAuthenticated(false);
      setUsername('');
      setPassword('');
      setIsLoggingOut(false);
    }, 500);
  };

  // --- Kéo giãn/thu nhỏ độ rộng cột bảng trong tài liệu Word bằng chuột (hoặc chạm) ---
  const WORD_TABLE_COL_MIN_WIDTH = 30;
  const WORD_TABLE_RESIZE_EDGE_PX = 6; // khoảng cách tới viền phải ô để bắt đầu kéo
  const [isResizingWordTableCol, setIsResizingWordTableCol] = useState<boolean>(false);
  const wordTableResizeRef = useRef<{
    table: HTMLTableElement;
    colIndex: number;
    startX: number;
    startWidth: number;
  } | null>(null);

  /** Tạo tay nắm xoay cho Shape đang chọn; tay nắm chỉ phục vụ chỉnh sửa và không được lưu vào tài liệu. */
  const ensureShapeRotateHandle = (shape: HTMLElement) => {
    if (!editorRef.current?.contains(shape)) return;
    editorRef.current.querySelectorAll('[data-smart-shape-rotate]').forEach((el) => {
      if (el.parentElement !== shape) el.remove();
    });
    let handle = shape.querySelector(':scope > [data-smart-shape-rotate]') as HTMLElement | null;
    if (!handle) {
      handle = document.createElement('span');
      handle.setAttribute('data-smart-shape-rotate', 'true');
      handle.setAttribute('contenteditable', 'false');
      handle.setAttribute('title', 'Kéo để xoay Shape');
      handle.setAttribute('aria-label', 'Kéo để xoay Shape');
      handle.style.cssText = 'position:absolute;left:50%;top:-24px;transform:translateX(-50%);width:16px;height:16px;border:2px solid #2563eb;border-radius:999px;background:#fff;box-shadow:0 1px 4px rgba(15,23,42,.18);cursor:grab;z-index:20;box-sizing:border-box;';
      shape.appendChild(handle);
    }
    return handle;
  };

  /** Nhấn gần viền phải 1 ô trong bảng (td/th) để bắt đầu kéo giãn cột đó */
  const handleEditorMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;

    // Shape có khung ngoài contenteditable=false để có thể kéo bằng viền/khoảng trống,
    // còn phần chữ bên trong vẫn contenteditable để người dùng sửa trực tiếp.
    const rotateHandle = target.closest('[data-smart-shape-rotate]') as HTMLElement | null;
    if (rotateHandle) {
      const shape = rotateHandle.closest('[data-smart-shape]') as HTMLElement | null;
      if (shape && editorRef.current?.contains(shape)) {
        e.preventDefault();
        e.stopPropagation();
        selectedShapeRef.current = shape;
        const rect = shape.getBoundingClientRect();
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;
        const startAngle = Math.atan2(e.clientY - centerY, e.clientX - centerX) * 180 / Math.PI;
        const startRotation = parseFloat(shape.dataset.shapeRotation || (shape.dataset.smartShape === 'diamond' ? '45' : '0')) || 0;
        shapeRotateRef.current = { el: shape, centerX, centerY, startAngle, startRotation };
        rotateHandle.style.cursor = 'grabbing';
        if (editorRef.current) {
          editorRef.current.style.cursor = 'grabbing';
          editorRef.current.style.userSelect = 'none';
        }
        return;
      }
    }

    const shape = target.closest('[data-smart-shape]') as HTMLElement | null;
    if (shape && editorRef.current?.contains(shape)) {
      // Chọn Shape ngay cả khi bấm vào vùng chữ, để có thể đổi màu viền sau đó.
      selectedShapeRef.current = shape;
      ensureShapeRotateHandle(shape);
      const textEditor = target.closest('[data-smart-shape-text]');
      if (!textEditor) {
        e.preventDefault();
        e.stopPropagation();
        const rect = shape.getBoundingClientRect();
        // Resize tự do từ mọi cạnh/góc; vùng giữa dùng để kéo di chuyển.
        const edge = 12;
        const nearLeft = e.clientX >= rect.left - edge && e.clientX <= rect.left + edge;
        const nearRight = e.clientX >= rect.right - edge && e.clientX <= rect.right + edge;
        const nearTop = e.clientY >= rect.top - edge && e.clientY <= rect.top + edge;
        const nearBottom = e.clientY >= rect.bottom - edge && e.clientY <= rect.bottom + edge;
        let direction = '';
        if (nearTop && nearLeft) direction = 'nw';
        else if (nearTop && nearRight) direction = 'ne';
        else if (nearBottom && nearLeft) direction = 'sw';
        else if (nearBottom && nearRight) direction = 'se';
        else if (nearLeft) direction = 'w';
        else if (nearRight) direction = 'e';
        else if (nearTop) direction = 'n';
        else if (nearBottom) direction = 's';

        const currentLeft = parseFloat(shape.dataset.shapeLeft || '0') || 0;
        const currentTop = parseFloat(shape.dataset.shapeTop || '0') || 0;
        if (direction) {
          shapeResizeRef.current = {
            el: shape,
            startX: e.clientX,
            startY: e.clientY,
            startWidth: rect.width,
            startHeight: rect.height,
            startLeft: currentLeft,
            startTop: currentTop,
            direction,
          };
          shape.style.cursor = `${direction}-resize`;
          if (editorRef.current) {
            editorRef.current.style.cursor = `${direction}-resize`;
            editorRef.current.style.userSelect = 'none';
          }
          return;
        }
        shapeDragRef.current = {
          el: shape,
          startX: e.clientX,
          startY: e.clientY,
          startLeft: currentLeft,
          startTop: currentTop,
        };
        if (editorRef.current) {
          editorRef.current.style.cursor = 'move';
          editorRef.current.style.userSelect = 'none';
        }
        return;
      }
    }

    const cell = target.closest('td, th') as HTMLTableCellElement | null;
    if (!cell) return;
    const table = cell.closest('table');
    if (!table) return;

    const rect = cell.getBoundingClientRect();
    const distanceFromRightEdge = rect.right - e.clientX;
    if (distanceFromRightEdge > WORD_TABLE_RESIZE_EDGE_PX || distanceFromRightEdge < -2) return;

    e.preventDefault();
    wordTableResizeRef.current = {
      table,
      colIndex: cell.cellIndex,
      startX: e.clientX,
      startWidth: rect.width,
    };
    setIsResizingWordTableCol(true);
  };

  /** Đổi con trỏ chuột thành col-resize khi rê gần viền phải 1 ô trong bảng, để người dùng biết có thể kéo */
  const handleEditorMouseMoveHover = (e: React.MouseEvent<HTMLDivElement>) => {
    if (isResizingWordTableCol || !editorRef.current) return;
    const target = e.target as HTMLElement;
    const cell = target.closest('td, th') as HTMLTableCellElement | null;
    if (!cell) {
      editorRef.current.style.cursor = '';
      return;
    }
    const rect = cell.getBoundingClientRect();
    const distanceFromRightEdge = rect.right - e.clientX;
    editorRef.current.style.cursor =
      distanceFromRightEdge <= WORD_TABLE_RESIZE_EDGE_PX && distanceFromRightEdge >= -2 ? 'col-resize' : '';
  };

  // ---- Ribbon: lưu / khôi phục vùng bôi đen khi bấm nút hoặc mở dropdown ----
  // (Dời lên trước vì insertRealPageGaps() bên dưới cần dùng đến 2 hàm này)
  const saveSelection = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && editorRef.current?.contains(sel.anchorNode)) {
      savedRangeRef.current = sel.getRangeAt(0).cloneRange();
    }
  };

  const restoreSelection = () => {
    const sel = window.getSelection();
    if (sel && savedRangeRef.current) {
      sel.removeAllRanges();
      sel.addRange(savedRangeRef.current);
    }
  };

  /**
   * Thu thập các điểm ngắt an toàn ở cấp khối của tài liệu.
   *
   * QUAN TRỌNG: Không chèn một <div> vào trong <table>/<tbody>/<tr>. DOM của HTML
   * không cho phép cấu trúc đó và trình duyệt có thể tự di chuyển node, làm hỏng bố cục.
   * Vì vậy bảng luôn được xem là một khối duy nhất; trang chỉ ngắt giữa các khối cấp cao nhất.
   */
  const collectWordBreakCandidates = (container: HTMLElement): HTMLElement[] => {
    return Array.from(container.children).filter((child) => {
      const el = child as HTMLElement;
      return !el.hasAttribute(WORD_PAGE_GAP_ATTR);
    }) as HTMLElement[];
  };

  /** Dựng khối ngăn giữa các trang Word. Không hiển thị/đánh số trang. */
  const buildPageGapElement = (): HTMLDivElement => {
    const gap = document.createElement('div');
    gap.setAttribute(WORD_PAGE_GAP_ATTR, 'true');
    gap.setAttribute('contenteditable', 'false');
    gap.setAttribute('aria-hidden', 'true');
    gap.style.cssText = [
      `height:${WORD_PAGE_GAP_PX}px`,
      `min-height:${WORD_PAGE_GAP_PX}px`,
      'width:calc(100% + 128px)',
      'margin:0 -64px',
      'background:#F1F3F1',
      'display:block',
      'user-select:none',
      'pointer-events:none',
      'box-sizing:border-box',
      'flex:0 0 auto',
      'box-shadow: inset 0 8px 10px -8px rgba(15,50,55,0.12), inset 0 -8px 10px -8px rgba(15,50,55,0.12)',
    ].join(';');
    return gap;
  };

  /**
   * Tách trang ổn định theo khổ A4.
   *
   * Cách cũ chèn khoảng trống theo thứ tự từ trang 1 → cuối tài liệu. Khi khoảng trống
   * đầu tiên được thêm vào, toàn bộ vị trí phía sau bị đẩy xuống, khiến mốc trang 2, 3...
   * bị lệch tiếp tục. Bản sửa này đo toàn bộ tài liệu khi CHƯA có gap, chọn điểm ngắt theo
   * đúng mốc A4 rồi chèn các gap từ cuối về đầu. Vì vậy việc chèn gap không làm thay đổi
   * các tọa độ đã đo và các trang sau không bị trôi.
   */
  const insertRealPageGaps = () => {
    const container = editorRef.current;
    if (!container || isLoading) return;

    // Chỉ giữ lại nội dung thật để đo. Các gap/nhãn cũ không được tham gia vào phép đo.
    container.querySelectorAll(`[${WORD_PAGE_GAP_ATTR}], [${WORD_PAGE_BADGE_ATTR}]`).forEach((el) => el.remove());

    const zoomFactor = Math.max(0.01, wordZoom / 100);
    const containerRect = container.getBoundingClientRect();

    const naturalCandidates = collectWordBreakCandidates(container);
    const meaningfulCandidates = naturalCandidates.filter((el) => {
      const text = (el.textContent || '').replace(/\u00a0/g, ' ').trim();
      return text.length > 0 || !!el.querySelector('img, table, svg, canvas');
    });

    // Đo chiều cao nội dung thật, KHÔNG dùng min-height/scrollHeight của editor
    // để tránh sinh thêm một trang ảo.
    const lastMeaningful = meaningfulCandidates[meaningfulCandidates.length - 1];
    const lastRect = lastMeaningful?.getBoundingClientRect();
    const naturalContentBottom = lastRect
      ? Math.max(1, (lastRect.bottom - containerRect.top) / zoomFactor)
      : 1;

    const totalPages = Math.max(
      1,
      Math.ceil(Math.max(1, naturalContentBottom - 0.5) / WORD_PAGE_HEIGHT_PX)
    );

    if (totalPages <= 1 || naturalCandidates.length === 0) return;

    const candidates = naturalCandidates
      .map((el, index) => {
        const rect = el.getBoundingClientRect();
        return {
          el,
          index,
          top: (rect.top - containerRect.top) / zoomFactor,
          bottom: (rect.bottom - containerRect.top) / zoomFactor,
        };
      })
      .filter((entry) =>
        Number.isFinite(entry.top) &&
        Number.isFinite(entry.bottom) &&
        entry.bottom > 0
      );

    const breaks: { el: HTMLElement; pageIndex: number }[] = [];
    let lastChosenIndex = -1;

    /*
     * Mỗi gap cao WORD_PAGE_GAP_PX được chèn giữa hai trang.
     * Vì vậy mốc của trang 3 trở đi phải tính cả các gap đã xuất hiện phía trước:
     *   trang 2 -> 1 * PAGE_HEIGHT
     *   trang 3 -> 2 * PAGE_HEIGHT + 1 * PAGE_GAP
     *   trang 4 -> 3 * PAGE_HEIGHT + 2 * PAGE_GAP
     * Nếu không cộng phần gap, các trang sau sẽ bị lệch và nhãn Trang N/N xuất hiện
     * sai vị trí.
     */
    for (let pageIndex = 2; pageIndex <= totalPages; pageIndex++) {
      const target =
        (pageIndex - 1) * WORD_PAGE_HEIGHT_PX +
        (pageIndex - 2) * WORD_PAGE_GAP_PX;

      const available = candidates.filter(
        (entry) => entry.index > lastChosenIndex
      );
      if (available.length === 0) break;

      // Ưu tiên block cuối cùng nằm trước mốc trang.
      let chosen = available
        .filter((entry) => entry.bottom <= target + 1)
        .at(-1);

      // Nếu mốc rơi giữa một block lớn, không được bỏ mất trang kế tiếp.
      // Chọn block đầu tiên vượt mốc làm biên an toàn.
      if (!chosen) {
        chosen = available.find((entry) => entry.bottom > target + 1);
      }

      // Fallback cuối cùng: luôn có một biên DOM nếu tài liệu còn block phía sau.
      if (!chosen) chosen = available[available.length - 1];
      if (!chosen) break;

      breaks.push({ el: chosen.el, pageIndex });
      lastChosenIndex = chosen.index;
    }

    // Chèn từ cuối về đầu để các tọa độ đã đo không bị thay đổi trong lúc chèn.
    for (let i = breaks.length - 1; i >= 0; i--) {
      const item = breaks[i];
      if (item.el.parentElement !== container) continue;
      item.el.insertAdjacentElement(
        'afterend',
        buildPageGapElement()
      );
    }
  };

  /** Lên lịch phân trang sau khi browser đã hoàn tất layout; có thêm 1 nhịp dự phòng cho font/ảnh. */
  // Đã tắt hoàn toàn chức năng tách trang Word theo yêu cầu.
  // Giữ hàm để không ảnh hưởng các luồng autosave/resize hiện có.
  const scheduleWordPagination = () => {};

  /** Lấy nội dung HTML thật (đã loại bỏ mọi khối ngăn trang) để lưu/đếm từ — không bao giờ lưu khối trang trí này */
  const getCleanEditorSnapshot = (): { html: string; text: string } => {
    if (!editorRef.current) return { html: '', text: '' };
    const clone = editorRef.current.cloneNode(true) as HTMLElement;
    clone.querySelectorAll(`[${WORD_PAGE_GAP_ATTR}], [${WORD_PAGE_BADGE_ATTR}], [data-smart-shape-rotate]`).forEach((el) => el.remove());
    return { html: clone.innerHTML, text: clone.textContent || '' };
  };

  const handleInput = () => {
    if (!selectedFile || !editorRef.current) return;
    setIsSaved(false);

    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;
    // Luôn lấy nội dung/đếm từ từ bản sao đã loại bỏ khối ngăn trang — đảm bảo không bao giờ
    // lưu nhầm hoặc đếm nhầm phần trang trí (khối ngăn trang) vào tài liệu thật.
    const { html: newContent, text: plainText } = getCleanEditorSnapshot();
    const words = plainText.trim().length > 0 ? plainText.trim().split(/\s+/).length : 0;
    setWordCount(words);

    localStorage.setItem(docKey, newContent);

    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);

    saveTimeoutRef.current = setTimeout(async () => {
      try {
        await fetch('/api/document-data', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ key: docKey, content: newContent }),
        });
        setIsSaved(true);
      } catch (err) {
        console.error('Lỗi đồng bộ Cloud:', err);
      }
      // Chỉ dựng lại khoảng trống trang SAU KHI người dùng đã tạm dừng gõ (đủ 800ms) và đã lưu xong,
      // tránh thay đổi cấu trúc DOM ngay giữa lúc đang gõ (có thể làm nhảy con trỏ).
    }, 800);
  };

  // Kéo thả Shapes bằng chuột: chỉ cập nhật vị trí của đúng shape đang chọn,
  // sau đó dùng handleInput() để lưu HTML hiện tại theo luồng autosave sẵn có.
  useEffect(() => {
    const moveShape = (clientX: number, clientY: number, shiftKey = false) => {
      const rotate = shapeRotateRef.current;
      if (rotate) {
        const angle = Math.atan2(clientY - rotate.centerY, clientX - rotate.centerX) * 180 / Math.PI;
        let rotation = rotate.startRotation + (angle - rotate.startAngle);
        if (shiftKey) rotation = Math.round(rotation / 15) * 15;
        rotation = Math.round(rotation * 10) / 10;
        rotate.el.dataset.shapeRotation = String(rotation);
        const left = parseFloat(rotate.el.dataset.shapeLeft || '0') || 0;
        const top = parseFloat(rotate.el.dataset.shapeTop || '0') || 0;
        rotate.el.style.transform = `translate(${left}px, ${top}px) rotate(${rotation}deg)`;
        return;
      }
      const resize = shapeResizeRef.current;
      if (resize) {
        const dx = clientX - resize.startX;
        const dy = clientY - resize.startY;
        const minWidth = resize.el.dataset.smartShape === 'diamond' ? 50 : 24;
        const minHeight = resize.el.dataset.smartShape === 'line' ? 8 : (resize.el.dataset.smartShape === 'diamond' ? 36 : 24);
        const dir = resize.direction;
        let width = resize.startWidth;
        let height = resize.startHeight;
        let left = resize.startLeft;
        let top = resize.startTop;

        if (dir.includes('e')) width = Math.max(minWidth, Math.round(resize.startWidth + dx));
        if (dir.includes('s')) height = Math.max(minHeight, Math.round(resize.startHeight + dy));
        if (dir.includes('w')) {
          width = Math.max(minWidth, Math.round(resize.startWidth - dx));
          left = Math.round(resize.startLeft + dx);
          if (width === minWidth) left = Math.round(resize.startLeft + resize.startWidth - minWidth);
        }
        if (dir.includes('n')) {
          height = Math.max(minHeight, Math.round(resize.startHeight - dy));
          top = Math.round(resize.startTop + dy);
          if (height === minHeight) top = Math.round(resize.startTop + resize.startHeight - minHeight);
        }

        resize.el.style.width = `${width}px`;
        resize.el.style.height = `${height}px`;
        if (resize.el.dataset.smartShape !== 'line') resize.el.style.minHeight = `${height}px`;
        const rotation = parseFloat(resize.el.dataset.shapeRotation || (resize.el.dataset.smartShape === 'diamond' ? '45' : '0')) || 0;
        resize.el.style.transform = `translate(${left}px, ${top}px) rotate(${rotation}deg)`;
        resize.el.dataset.shapeWidth = String(width);
        resize.el.dataset.shapeHeight = String(height);
        resize.el.dataset.shapeLeft = String(left);
        resize.el.dataset.shapeTop = String(top);
        return;
      }
      const info = shapeDragRef.current;
      if (!info || !editorRef.current) return;
      const dx = clientX - info.startX;
      const dy = clientY - info.startY;
      const left = Math.round(info.startLeft + dx);
      const top = Math.round(info.startTop + dy);
      const rotation = parseFloat(info.el.dataset.shapeRotation || (info.el.dataset.smartShape === 'diamond' ? '45' : '0')) || 0;
      info.el.style.transform = `translate(${left}px, ${top}px) rotate(${rotation}deg)`;
      info.el.dataset.shapeLeft = String(left);
      info.el.dataset.shapeTop = String(top);
    };

    const stopShapeDrag = () => {
      if (!shapeDragRef.current && !shapeResizeRef.current && !shapeRotateRef.current) return;
      const resizeEl = shapeResizeRef.current?.el;
      const rotateEl = shapeRotateRef.current?.el;
      if (resizeEl) resizeEl.style.cursor = 'move';
      const rotateHandle = rotateEl?.querySelector(':scope > [data-smart-shape-rotate]') as HTMLElement | null;
      if (rotateHandle) rotateHandle.style.cursor = 'grab';
      shapeDragRef.current = null;
      shapeResizeRef.current = null;
      shapeRotateRef.current = null;
      if (editorRef.current) {
        editorRef.current.style.cursor = '';
        editorRef.current.style.userSelect = '';
      }
      handleInput();
    };

    const onMouseMove = (e: MouseEvent) => moveShape(e.clientX, e.clientY, e.shiftKey);
    const onTouchMove = (e: TouchEvent) => {
      if (e.touches.length) {
        e.preventDefault();
        moveShape(e.touches[0].clientX, e.touches[0].clientY);
      }
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', stopShapeDrag);
    document.addEventListener('touchmove', onTouchMove, { passive: false });
    document.addEventListener('touchend', stopShapeDrag);
    document.addEventListener('touchcancel', stopShapeDrag);
    return () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', stopShapeDrag);
      document.removeEventListener('touchmove', onTouchMove);
      document.removeEventListener('touchend', stopShapeDrag);
      document.removeEventListener('touchcancel', stopShapeDrag);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Cho phép xóa Shape đã chọn bằng Delete/Backspace, không ảnh hưởng thao tác nhập văn bản.
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const shape = selectedShapeRef.current;
      if (!shape || !editorRef.current?.contains(shape)) return;
      const target = e.target as HTMLElement | null;
      if (target?.closest('[data-smart-shape-text]')) return;
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        e.stopPropagation();
        deleteSelectedWordShape();
      }
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Hiệu ứng kéo giãn cột bảng Word — đặt SAU handleInput để tham chiếu đúng thứ tự khai báo
  useEffect(() => {
    if (!isResizingWordTableCol) return;

    const applyWidth = (clientX: number) => {
      const info = wordTableResizeRef.current;
      if (!info) return;
      const delta = clientX - info.startX;
      const newWidth = Math.max(WORD_TABLE_COL_MIN_WIDTH, Math.round(info.startWidth + delta));
      const rows = info.table.rows;
      for (let i = 0; i < rows.length; i++) {
        const c = rows[i].cells[info.colIndex];
        if (c) c.style.width = `${newWidth}px`;
      }
    };

    const handleMouseMove = (e: MouseEvent) => applyWidth(e.clientX);
    const handleTouchMove = (e: TouchEvent) => {
      if (e.touches.length === 0) return;
      e.preventDefault();
      applyWidth(e.touches[0].clientX);
    };
    const stopResizing = () => {
      wordTableResizeRef.current = null;
      setIsResizingWordTableCol(false);
      // Lưu lại độ rộng cột vừa đổi bằng đúng luồng autosave hiện có (local + debounce cloud)
      handleInput();
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', stopResizing);
    document.addEventListener('touchmove', handleTouchMove, { passive: false });
    document.addEventListener('touchend', stopResizing);
    document.addEventListener('touchcancel', stopResizing);
    if (editorRef.current) editorRef.current.style.cursor = 'col-resize';
    if (editorRef.current) editorRef.current.style.userSelect = 'none';

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', stopResizing);
      document.removeEventListener('touchmove', handleTouchMove);
      document.removeEventListener('touchend', stopResizing);
      document.removeEventListener('touchcancel', stopResizing);
      if (editorRef.current) editorRef.current.style.cursor = '';
      if (editorRef.current) editorRef.current.style.userSelect = '';
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isResizingWordTableCol]);

  // Tách trang Word đã được tắt hoàn toàn; không tạo gap/nhãn/observer phân trang.

  const refreshActiveFormats = () => {
    try {
      setActiveFormats({
        bold: document.queryCommandState('bold'),
        italic: document.queryCommandState('italic'),
        underline: document.queryCommandState('underline'),
        strikeThrough: document.queryCommandState('strikeThrough'),
        justifyLeft: document.queryCommandState('justifyLeft'),
        justifyCenter: document.queryCommandState('justifyCenter'),
        justifyRight: document.queryCommandState('justifyRight'),
        justifyFull: document.queryCommandState('justifyFull'),
        insertUnorderedList: document.queryCommandState('insertUnorderedList'),
        insertOrderedList: document.queryCommandState('insertOrderedList'),
      });
      const block = document.queryCommandValue('formatBlock');
      if (block) setCurrentBlock(block.toString().toLowerCase());
    } catch {
      // queryCommandState có thể ném lỗi ngoài vùng contentEditable — bỏ qua an toàn
    }
  };

  useEffect(() => {
    const handleSelectionChange = () => {
      // Luôn cập nhật vùng chọn mới nhất bên trong editor.
      // Nhờ đó khi bấm các nút/combobox trên ribbon, lệnh định dạng
      // luôn tác động đúng phần người dùng vừa bôi đen, không dùng
      // lại một vùng chọn cũ của tài liệu.
      saveSelection();
      refreshActiveFormats();
    };

    document.addEventListener('selectionchange', handleSelectionChange);
    return () => document.removeEventListener('selectionchange', handleSelectionChange);
  }, []);

  // Xử lý kéo thanh chia để thay đổi chiều rộng sidebar (chuột trên desktop, chạm trên di động/tablet)
  const handleSidebarResizeStart = (e: React.MouseEvent | React.TouchEvent) => {
    e.preventDefault();
    setIsResizingSidebar(true);
  };

  useEffect(() => {
    if (!isResizingSidebar) return;

    const updateWidthFromClientX = (clientX: number) => {
      if (!sidebarRef.current) return;
      const rect = sidebarRef.current.getBoundingClientRect();
      const nextWidth = clientX - rect.left;
      // Trên màn hình nhỏ, không cho kéo rộng hơn khoảng trống thực tế còn lại
      const viewportCap = typeof window !== 'undefined' ? window.innerWidth - 24 : SIDEBAR_MAX_WIDTH;
      const clamped = Math.min(SIDEBAR_MAX_WIDTH, viewportCap, Math.max(SIDEBAR_MIN_WIDTH, nextWidth));
      setSidebarWidth(clamped);
    };

    const handleMouseMove = (e: MouseEvent) => updateWidthFromClientX(e.clientX);
    const handleTouchMove = (e: TouchEvent) => {
      if (e.touches.length === 0) return;
      // Chặn cuộn trang khi đang kéo bằng ngón tay
      e.preventDefault();
      updateWidthFromClientX(e.touches[0].clientX);
    };
    const stopResizing = () => setIsResizingSidebar(false);

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', stopResizing);
    document.addEventListener('touchmove', handleTouchMove, { passive: false });
    document.addEventListener('touchend', stopResizing);
    document.addEventListener('touchcancel', stopResizing);
    if (editorRef.current) editorRef.current.style.cursor = 'col-resize';
    if (editorRef.current) editorRef.current.style.userSelect = 'none';

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', stopResizing);
      document.removeEventListener('touchmove', handleTouchMove);
      document.removeEventListener('touchend', stopResizing);
      document.removeEventListener('touchcancel', stopResizing);
      if (editorRef.current) editorRef.current.style.cursor = '';
      if (editorRef.current) editorRef.current.style.userSelect = '';
    };
  }, [isResizingSidebar]);


  // Chạy 1 lệnh định dạng chuẩn của trình duyệt (execCommand) trên vùng đang chọn
  const execFormat = (command: string, value?: string) => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    restoreSelection();
    document.execCommand(command, false, value);
    handleInput();
    refreshActiveFormats();
  };

  // Cỡ chữ theo pt thật (Word dùng pt, execCommand mặc định chỉ hỗ trợ thang 1-7 nên cần "vá" lại)
  const applyFontSize = (pt: string) => {
    const editor = editorRef.current;
    if (!editor) return;

    editor.focus();
    restoreSelection();

    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return;

    const sourceRange = selection.getRangeAt(0).cloneRange();
    if (!editor.contains(sourceRange.commonAncestorContainer)) return;

    // Chỉ áp dụng cỡ chữ cho đúng phần đang bôi đen.
    // Không dùng querySelectorAll('font[size="7"]') trên toàn editor vì cách đó
    // sẽ đổi cả những vùng khác trong tài liệu có cùng thẻ font.
    if (!sourceRange.collapsed) {
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
      const textNodes: Text[] = [];
      let node = walker.nextNode();
      while (node) {
        const textNode = node as Text;
        if (textNode.nodeValue && sourceRange.intersectsNode(textNode)) {
          textNodes.push(textNode);
        }
        node = walker.nextNode();
      }

      const insertedSpans: HTMLSpanElement[] = [];

      // Làm từ cuối về đầu để việc thay đổi DOM không làm lệch các node
      // đã thu thập của vùng chọn.
      for (let i = textNodes.length - 1; i >= 0; i--) {
        const textNode = textNodes[i];
        const length = textNode.nodeValue?.length ?? 0;
        if (!length) continue;

        let startOffset = 0;
        let endOffset = length;

        if (sourceRange.startContainer === textNode) {
          startOffset = sourceRange.startOffset;
        }
        if (sourceRange.endContainer === textNode) {
          endOffset = sourceRange.endOffset;
        }

        startOffset = Math.max(0, Math.min(startOffset, length));
        endOffset = Math.max(0, Math.min(endOffset, length));
        if (endOffset <= startOffset) continue;

        const partRange = document.createRange();
        partRange.setStart(textNode, startOffset);
        partRange.setEnd(textNode, endOffset);

        const spanElement = document.createElement('span');
        spanElement.style.fontSize = `${pt}pt`;
        spanElement.appendChild(partRange.extractContents());
        partRange.insertNode(spanElement);
        insertedSpans.unshift(spanElement);
      }

      // Giữ lại đúng vùng vừa bôi đen để người dùng có thể tiếp tục
      // định dạng độc lập mà không bị chuyển sang toàn bộ tài liệu.
      if (insertedSpans.length > 0) {
        const nextSelection = document.createRange();
        nextSelection.setStartBefore(insertedSpans[0]);
        nextSelection.setEndAfter(insertedSpans[insertedSpans.length - 1]);
        selection.removeAllRanges();
        selection.addRange(nextSelection);
        savedRangeRef.current = nextSelection.cloneRange();
      }
    } else {
      // Khi không bôi đen, cỡ chữ mới chỉ áp dụng cho vị trí con trỏ
      // để nội dung nhập tiếp theo dùng đúng cỡ chữ đã chọn.
      document.execCommand('fontSize', false, '7');
      const active = window.getSelection();
      if (active && active.rangeCount > 0) {
        const range = active.getRangeAt(0);
        const fonts = Array.from(editor.querySelectorAll('font[size="7"]'))
          .filter((el) => range.intersectsNode(el));

        fonts.forEach((el) => {
          const spanElement = document.createElement('span');
          spanElement.style.fontSize = `${pt}pt`;
          while (el.firstChild) spanElement.appendChild(el.firstChild);
          el.replaceWith(spanElement);
        });
      }
    }

    handleInput();
    refreshActiveFormats();
  };

  const applyHeading = (tag: string) => {
    execFormat('formatBlock', tag);
  };

  const applyColor = (type: 'foreColor' | 'hiliteColor', value: string) => {
    execFormat(type, value);
  };

  // ===== CÁC CHỨC NĂNG CƠ BẢN CHO TẤT CẢ FILE WORD =====
  // Các thao tác dưới đây chỉ tác động lên vùng soạn thảo Word hiện tại,
  // không thay đổi giao diện hay dữ liệu của các module khác.
  const getWordSelectedTableCells = (): HTMLTableCellElement[] => {
    if (!editorRef.current) return [];
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return [];
    const range = sel.getRangeAt(0);
    return Array.from(editorRef.current.querySelectorAll('td,th')).filter((el) => {
      try { return range.intersectsNode(el); } catch { return false; }
    }) as HTMLTableCellElement[];
  };

  const applyWordBorder = (side: 'all'|'top'|'bottom'|'left'|'right'|'none'|'outside'|'thickOutside'|'doubleBottom'|'thickBottom'|'topBottom'|'topThickBottom'|'topDoubleBottom') => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    restoreSelection();
    const cells = getWordSelectedTableCells();
    if (cells.length) {
      const rows = new Map<HTMLTableRowElement, HTMLTableCellElement[]>();
      cells.forEach(c => { const r=c.parentElement as HTMLTableRowElement; if(r) rows.set(r,[...(rows.get(r)||[]),c]); });
      const minRow = Math.min(...cells.map(c => c.parentElement ? (c.parentElement as HTMLTableRowElement).rowIndex : 0));
      const maxRow = Math.max(...cells.map(c => c.parentElement ? (c.parentElement as HTMLTableRowElement).rowIndex : 0));
      const minCol = Math.min(...cells.map(c => c.cellIndex));
      const maxCol = Math.max(...cells.map(c => c.cellIndex));
      cells.forEach((cell) => {
        if (side === 'none') cell.style.border = 'none';
        else if (side === 'all') cell.style.border = '1px solid #000';
        else if (side === 'top') cell.style.borderTop = '1px solid #000';
        else if (side === 'bottom') cell.style.borderBottom = '1px solid #000';
        else if (side === 'left') cell.style.borderLeft = '1px solid #000';
        else if (side === 'right') cell.style.borderRight = '1px solid #000';
        else if (side === 'outside' || side === 'thickOutside') {
          const v = side === 'thickOutside' ? '2px solid #000' : '1px solid #000';
          const r = cell.parentElement ? (cell.parentElement as HTMLTableRowElement).rowIndex : 0;
          if (r === minRow) cell.style.borderTop = v;
          if (r === maxRow) cell.style.borderBottom = v;
          if (cell.cellIndex === minCol) cell.style.borderLeft = v;
          if (cell.cellIndex === maxCol) cell.style.borderRight = v;
        } else if (side === 'doubleBottom' && cell.parentElement && (cell.parentElement as HTMLTableRowElement).rowIndex === maxRow) cell.style.borderBottom = '3px double #000';
        else if (side === 'thickBottom' && cell.parentElement && (cell.parentElement as HTMLTableRowElement).rowIndex === maxRow) cell.style.borderBottom = '2px solid #000';
        else if (side === 'topBottom') { if ((cell.parentElement as HTMLTableRowElement).rowIndex === minRow) cell.style.borderTop='1px solid #000'; if ((cell.parentElement as HTMLTableRowElement).rowIndex === maxRow) cell.style.borderBottom='1px solid #000'; }
        else if (side === 'topThickBottom') { if ((cell.parentElement as HTMLTableRowElement).rowIndex === minRow) cell.style.borderTop='1px solid #000'; if ((cell.parentElement as HTMLTableRowElement).rowIndex === maxRow) cell.style.borderBottom='2px solid #000'; }
        else if (side === 'topDoubleBottom') { if ((cell.parentElement as HTMLTableRowElement).rowIndex === minRow) cell.style.borderTop='1px solid #000'; if ((cell.parentElement as HTMLTableRowElement).rowIndex === maxRow) cell.style.borderBottom='3px double #000'; }
      });
      handleInput();
      return;
    }
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
    const range = sel.getRangeAt(0);
    const spanElement = document.createElement('span');
    if (side === 'none') spanElement.style.border = 'none';
    else if (side === 'all') spanElement.style.border = '1px solid #000';
    else if (side === 'thickOutside') spanElement.style.border = '2px solid #000';
    else if (side === 'doubleBottom') spanElement.style.borderBottom = '3px double #000';
    else if (side === 'thickBottom') spanElement.style.borderBottom = '2px solid #000';
    else if (side === 'topBottom') { spanElement.style.borderTop='1px solid #000'; spanElement.style.borderBottom='1px solid #000'; }
    else if (side === 'topThickBottom') { spanElement.style.borderTop='1px solid #000'; spanElement.style.borderBottom='2px solid #000'; }
    else if (side === 'topDoubleBottom') { spanElement.style.borderTop='1px solid #000'; spanElement.style.borderBottom='3px double #000'; }
    else if (side === 'outside') spanElement.style.border = '1px solid #000';
    else if (side === 'top') spanElement.style.borderTop = '1px solid #000';
    else if (side === 'bottom') spanElement.style.borderBottom = '1px solid #000';
    else if (side === 'left') spanElement.style.borderLeft = '1px solid #000';
    else if (side === 'right') spanElement.style.borderRight = '1px solid #000';
    spanElement.appendChild(range.extractContents());
    range.insertNode(spanElement);
    const nr = document.createRange(); nr.selectNodeContents(spanElement); sel.removeAllRanges(); sel.addRange(nr); savedRangeRef.current = nr.cloneRange();
    handleInput();
  };

  const mergeWordCells = (mode: 'center'|'across'|'cells'|'unmerge') => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    restoreSelection();
    const cells = getWordSelectedTableCells();
    if (!cells.length) { alert('Hãy chọn các ô trong cùng một bảng Word trước.'); return; }
    if (mode === 'unmerge') {
      cells.forEach((cell) => {
        const colSpan = Math.max(1, cell.colSpan || 1);
        const rowSpan = Math.max(1, cell.rowSpan || 1);
        if (colSpan > 1) {
          cell.colSpan = 1;
          for (let i=1;i<colSpan;i++) { const n=cell.cloneNode(false) as HTMLTableCellElement; n.innerHTML='&nbsp;'; cell.parentElement?.insertBefore(n,cell.nextSibling); }
        }
        if (rowSpan > 1) {
          const row = cell.parentElement as HTMLTableRowElement | null;
          const index = row ? cell.cellIndex : 0;
          cell.rowSpan = 1;
          for (let rr=1; rr<rowSpan; rr++) {
            const targetRow = row?.parentElement?.children[row.sectionRowIndex + rr] as HTMLTableRowElement | undefined;
            if (targetRow) {
              const n=cell.cloneNode(false) as HTMLTableCellElement; n.innerHTML='&nbsp;'; targetRow.insertBefore(n,targetRow.children[index] || null);
            }
          }
        }
      });
      handleInput(); return;
    }
    const rows = new Map<HTMLTableRowElement, HTMLTableCellElement[]>();
    cells.forEach((cell) => { const row=cell.parentElement as HTMLTableRowElement; if (row) rows.set(row,[...(rows.get(row)||[]),cell]); });
    const rowEntries = Array.from(rows.entries()).sort((a,b)=>a[0].rowIndex-b[0].rowIndex);
    const mergeRow = (rowCells: HTMLTableCellElement[], center:boolean) => {
      rowCells.sort((a,b)=>a.cellIndex-b.cellIndex);
      const first=rowCells[0];
      const totalColSpan=rowCells.reduce((n,c)=>n+Math.max(1,c.colSpan||1),0);
      rowCells.slice(1).forEach(c=>c.remove());
      first.colSpan=totalColSpan;
      if (center) first.style.textAlign='center';
    };
    if (mode === 'across' || rowEntries.length === 1) {
      rowEntries.forEach(([,rowCells])=>mergeRow(rowCells,mode==='center'));
    } else {
      // Merge Cells / Merge & Center theo vùng chữ nhật nhiều dòng: giữ ô đầu và mở rộng cả colSpan + rowSpan.
      const firstRowCells = rowEntries[0][1].sort((a,b)=>a.cellIndex-b.cellIndex);
      const first = firstRowCells[0];
      const colSpan = firstRowCells.reduce((n,c)=>n+Math.max(1,c.colSpan||1),0);
      rowEntries.slice(1).forEach(([,rowCells])=>rowCells.forEach(c=>c.remove()));
      firstRowCells.slice(1).forEach(c=>c.remove());
      first.colSpan = colSpan;
      first.rowSpan = rowEntries.length;
      if (mode === 'center') first.style.textAlign='center';
    }
    handleInput();
  };

  const selectAllWord = () => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(editorRef.current);
    selection?.removeAllRanges();
    selection?.addRange(range);
    refreshActiveFormats();
  };

  const copyWordSelection = () => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    try {
      document.execCommand('copy');
    } catch {
      // Trình duyệt có thể chặn clipboard; không làm ảnh hưởng tài liệu.
    }
  };

  const cutWordSelection = () => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    try {
      const ok = document.execCommand('cut');
      if (ok) handleInput();
    } catch {
      // Bỏ qua nếu trình duyệt không cho phép thao tác clipboard.
    }
  };

  const pasteWordText = async () => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    restoreSelection();

    try {
      if (navigator.clipboard?.readText) {
        const text = await navigator.clipboard.readText();
        if (text) {
          document.execCommand('insertText', false, text);
          handleInput();
        }
        return;
      }
    } catch {
      // Fallback bên dưới cho trình duyệt không cấp quyền Clipboard API.
    }

    try {
      document.execCommand('paste');
      handleInput();
    } catch {
      alert('Trình duyệt không cho phép dán tự động. Vui lòng dùng Ctrl+V.');
    }
  };

  const clearWordFormatting = () => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    restoreSelection();
    document.execCommand('removeFormat', false);
    document.execCommand('unlink', false);
    handleInput();
    refreshActiveFormats();
  };

  const findAndReplaceWordText = () => {
    if (!editorRef.current) return;

    const findText = window.prompt('Tìm nội dung:', '');
    if (findText === null || findText === '') return;
    const replaceText = window.prompt('Thay thế bằng:', '');
    if (replaceText === null) return;

    const walker = document.createTreeWalker(editorRef.current, NodeFilter.SHOW_TEXT);
    const textNodes: Text[] = [];
    let currentNode: Node | null = walker.nextNode();
    while (currentNode) {
      textNodes.push(currentNode as Text);
      currentNode = walker.nextNode();
    }

    let count = 0;
    textNodes.forEach((node) => {
      const value = node.nodeValue || '';
      if (!value.includes(findText)) return;
      const occurrences = value.split(findText).length - 1;
      if (occurrences > 0) {
        node.nodeValue = value.split(findText).join(replaceText);
        count += occurrences;
      }
    });

    if (count > 0) {
      handleInput();
      alert(`Đã thay thế ${count} vị trí.`);
    } else {
      alert('Không tìm thấy nội dung cần thay thế.');
    }
  };

  // Phím tắt Ctrl+K (hoặc Cmd+K trên macOS) — focus vào ô tìm kiếm tài liệu ở Header
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const editor = editorRef.current;
      const target = e.target as HTMLElement | null;
      const mod = e.ctrlKey || e.metaKey;

      if (mod && e.key.toLowerCase() === 'k' && !editor?.contains(target)) {
        e.preventDefault();
        headerSearchRef.current?.focus();
        return;
      }

      // Phím tắt Word chỉ chạy khi con trỏ đang ở vùng soạn thảo Word.
      if (!editor || !editor.contains(target)) return;

      if (mod && e.key.toLowerCase() === 'a') {
        e.preventDefault();
        selectAllWord();
      } else if (mod && e.key.toLowerCase() === 's') {
        e.preventDefault();
        handleInput();
      } else if (mod && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        findAndReplaceWordText();
      }
    };
    document.addEventListener('keydown', handleGlobalKeyDown);
    return () => document.removeEventListener('keydown', handleGlobalKeyDown);
  }, []);

  const insertWordLink = () => {
    if (!editorRef.current) return;
    editorRef.current.focus();
    restoreSelection();
    const url = window.prompt('Nhập liên kết (URL):', 'https://');
    if (url === null || url.trim() === '') return;
    document.execCommand('createLink', false, url.trim());
    handleInput();
  };

  /** Xóa đúng Shape đang được chọn. Có thể gọi từ menu hoặc phím Delete/Backspace. */
  /** Xóa đúng Shape đang được chọn. Có thể gọi từ menu hoặc phím Delete/Backspace.
   * Dùng function declaration để có thể được gọi từ các effect khai báo phía trên
   * mà không vi phạm temporal dead zone của const/let.
   */
  function deleteSelectedWordShape() {
    const shape = selectedShapeRef.current;
    if (!shape || !editorRef.current?.contains(shape)) return;
    if (shape.nextElementSibling?.textContent === '\u00a0') {
      shape.nextElementSibling.remove();
    }
    shape.remove();
    selectedShapeRef.current = null;
    shapeDragRef.current = null;
    shapeResizeRef.current = null;
    shapeRotateRef.current = null;
    if (editorRef.current) {
      editorRef.current.style.cursor = '';
      editorRef.current.style.userSelect = '';
    }
    handleInput();
  }

  /**
   * Insert → Shapes cho tài liệu Word. Các shape được lưu trực tiếp trong HTML của tài liệu,
   * nên vẫn đi qua đúng luồng autosave hiện tại. Khung ngoài có thể kéo; chữ bên trong có thể sửa.
   */
  type WordShapeKind =
    | 'textBox' | 'rect' | 'round' | 'ellipse' | 'triangle' | 'diamond'
    | 'parallelogram' | 'trapezoid' | 'pentagon' | 'hexagon' | 'octagon'
    | 'rightChevron' | 'rightArrow' | 'leftArrow' | 'downArrow'
    | 'upArrow' | 'star' | 'braceLeft' | 'braceRight' | 'arc' | 'curve'
    | 'line' | 'arrowLine' | 'elbow' | 'elbowArrow' | 'curveArrow'
    | 'uLine' | 'doubleBrace' | 'dashedLine';

  type ExcelBorderSide = 'all'|'top'|'bottom'|'left'|'right'|'none'|'outside'|'thickOutside'|'doubleBottom'|'thickBottom'|'topBottom'|'topThickBottom'|'topDoubleBottom';
  type ExcelMergeMode = 'center'|'across'|'cells'|'unmerge';
  type WordBorderSide = ExcelBorderSide;
  type WordMergeMode = ExcelMergeMode;
  const excelBorderOptions: ReadonlyArray<readonly [ExcelBorderSide, string]> = [
    ['bottom','Bottom Border'],['top','Top Border'],['left','Left Border'],['right','Right Border'],['none','No Border'],
    ['all','All Borders'],['outside','Outside Borders'],['thickOutside','Thick Outside Borders'],['doubleBottom','Bottom Double Border'],
    ['thickBottom','Thick Bottom Border'],['topBottom','Top and Bottom Border'],['topThickBottom','Top and Thick Bottom Border'],
    ['topDoubleBottom','Top and Double Bottom Border']
  ];
  const excelMergeOptions: ReadonlyArray<readonly [ExcelMergeMode, string]> = [
    ['center','Gộp & căn giữa'],['across','Gộp theo hàng'],['cells','Gộp ô'],['unmerge','Bỏ gộp ô']
  ];
  const wordBorderOptions: ReadonlyArray<readonly [WordBorderSide, string]> = excelBorderOptions;
  const wordMergeOptions: ReadonlyArray<readonly [WordMergeMode, string]> = excelMergeOptions;


  const insertWordShape = (kind: WordShapeKind) => {
    if (!editorRef.current) return;

    editorRef.current.focus();
    restoreSelection();

    const svgShapeKinds = new Set<WordShapeKind>([
      'triangle','parallelogram','trapezoid','pentagon','hexagon','octagon','rightChevron',
      'leftArrow','upArrow','star','braceLeft','braceRight','arc','curve','arrowLine','elbow',
      'elbowArrow','curveArrow','uLine','doubleBrace'
    ]);

    const styles: Record<string, string> = {
      textBox: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:210px;min-height:56px;margin:10px 12px 10px 0;padding:8px 12px;border:1px solid ${shapeBorderColor};background:#fff;position:relative;box-sizing:border-box;cursor:move;transform:translate(0px,0px);`,
      rect: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:210px;min-height:64px;margin:10px 12px 10px 0;padding:10px 14px;border:2px solid ${shapeBorderColor};background:#fff;position:relative;box-sizing:border-box;cursor:move;transform:translate(0px,0px);`,
      round: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:210px;min-height:64px;margin:10px 12px 10px 0;padding:10px 14px;border:2px solid ${shapeBorderColor};border-radius:12px;background:#fff;position:relative;box-sizing:border-box;cursor:move;transform:translate(0px,0px);`,
      ellipse: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:210px;min-height:64px;margin:10px 12px 10px 0;padding:10px 20px;border:2px solid ${shapeBorderColor};border-radius:999px;background:#fff;position:relative;box-sizing:border-box;cursor:move;transform:translate(0px,0px);`,
      diamond: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:150px;height:100px;margin:10px 24px;padding:12px;transform:translate(0px,0px) rotate(45deg);border:2px solid ${shapeBorderColor};background:#fff;position:relative;box-sizing:border-box;cursor:move;`,
      downArrow: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:72px;height:76px;margin:8px 14px;color:${shapeBorderColor};font-size:54px;line-height:1;position:relative;box-sizing:border-box;cursor:move;`,
      rightArrow: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:120px;height:64px;margin:8px 14px;color:${shapeBorderColor};font-size:54px;line-height:1;position:relative;box-sizing:border-box;cursor:move;`,
      leftArrow: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:120px;height:64px;margin:8px 14px;color:${shapeBorderColor};position:relative;box-sizing:border-box;cursor:move;`,
      upArrow: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:72px;height:76px;margin:8px 14px;color:${shapeBorderColor};position:relative;box-sizing:border-box;cursor:move;`,
      triangle: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:120px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      parallelogram: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:170px;height:76px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      trapezoid: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:170px;height:76px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      pentagon: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:150px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      hexagon: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:160px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      octagon: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:150px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      rightChevron: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:150px;height:80px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      star: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:100px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      braceLeft: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:60px;height:100px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      braceRight: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:60px;height:100px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      doubleBrace: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:110px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      arc: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:150px;height:90px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      curve: `display:inline-flex;align-items:center;justify-content:center;vertical-align:middle;width:170px;height:80px;margin:10px 18px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      line: 'display:inline-block;vertical-align:middle;width:180px;height:28px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;',
      dashedLine: 'display:inline-block;vertical-align:middle;width:180px;height:28px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;',
      arrowLine: `display:inline-block;vertical-align:middle;width:180px;height:42px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      elbow: `display:inline-block;vertical-align:middle;width:180px;height:70px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      elbowArrow: `display:inline-block;vertical-align:middle;width:180px;height:70px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      curveArrow: `display:inline-block;vertical-align:middle;width:180px;height:70px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
      uLine: `display:inline-block;vertical-align:middle;width:150px;height:70px;margin:8px 14px;position:relative;box-sizing:border-box;cursor:move;color:${shapeBorderColor};`,
    };
    const labels: Record<string, string> = {
      textBox: 'Nội dung', rect: 'Nội dung', round: 'Nội dung', ellipse: 'Nội dung', diamond: 'Nội dung',
      downArrow: '↓', rightArrow: '→', leftArrow: '←', upArrow: '↑', triangle: 'Nội dung',
      parallelogram: 'Nội dung', trapezoid: 'Nội dung', pentagon: 'Nội dung', hexagon: 'Nội dung',
      octagon: 'Nội dung', rightChevron: 'Nội dung', star: 'Nội dung', line: '', dashedLine: '',
      arrowLine: '', elbow: '', elbowArrow: '', curveArrow: '', uLine: '', doubleBrace: '',
      braceLeft: '', braceRight: '', arc: '', curve: ''
    };

    const svg = (kind: WordShapeKind) => {
      const common = `xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 200 100" preserveAspectRatio="none" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"`;
      const body: Record<string,string> = {
        triangle: '<polygon points="100,8 190,92 10,92"/>',
        parallelogram: '<polygon points="38,8 190,8 162,92 10,92"/>',
        trapezoid: '<polygon points="38,8 162,8 190,92 10,92"/>',
        pentagon: '<polygon points="100,6 190,38 155,94 45,94 10,38"/>',
        hexagon: '<polygon points="35,8 165,8 195,50 165,92 35,92 5,50"/>',
        octagon: '<polygon points="40,7 160,7 193,40 193,60 160,93 40,93 7,60 7,40"/>',
        rightChevron: '<polyline points="10,10 82,10 142,50 82,90 10,90 70,50 10,10"/>',
        star: '<polygon points="100,5 122,35 157,36 130,57 140,91 100,70 60,91 70,57 43,36 78,35"/>',
        braceLeft: '<path d="M145 5 C105 5 115 25 82 25 C52 25 55 42 55 50 C55 58 52 75 82 75 C115 75 105 95 145 95"/>',
        braceRight: '<path d="M55 5 C95 5 85 25 118 25 C148 25 145 42 145 50 C145 58 148 75 118 75 C85 75 95 95 55 95"/>',
        doubleBrace: '<path d="M65 8 C35 8 45 24 25 24 C8 24 12 38 12 50 C12 62 8 76 25 76 C45 76 35 92 65 92 M135 8 C165 8 155 24 175 24 C192 24 188 38 188 50 C188 62 192 76 175 76 C155 76 165 92 135 92"/>',
        arc: '<path d="M20 82 A80 80 0 0 1 180 82"/>',
        curve: '<path d="M10 75 C55 5 95 95 140 25 C155 3 172 15 190 30"/>',
        arrowLine: '<line x1="12" y1="50" x2="168" y2="50"/><polygon points="188,50 168,40 168,60" fill="currentColor" stroke="none"/>',
        elbow: '<polyline points="12,82 82,82 82,18 188,18"/>',
        elbowArrow: '<polyline points="12,82 82,82 82,18 168,18"/><polygon points="188,18 168,8 168,28" fill="currentColor" stroke="none"/>',
        curveArrow: '<path d="M10 78 C55 8 115 92 168 25"/><polygon points="188,25 168,15 168,35" fill="currentColor" stroke="none"/>',
        uLine: '<path d="M20 15 V68 Q20 85 38 85 H162 Q180 85 180 68 V15"/>',
      };
      return `<svg ${common}>${body[kind] || ''}</svg>`;
    };

    let id = '';
    do {
      const seq = shapeIdRef.current++;
      id = `shape-${seq}`;
    } while (editorRef.current.querySelector(`#${id}`));
    let html = '';
    const initialRotation = kind === 'diamond' ? 45 : 0;
    const textKinds = new Set<WordShapeKind>(['textBox','rect','round','ellipse','diamond','triangle','parallelogram','trapezoid','pentagon','hexagon','octagon','rightChevron']);
    if (kind === 'line' || kind === 'dashedLine') {
      const lineStyle = kind === 'dashedLine' ? 'dashed' : 'solid';
      html = `<span data-smart-shape="${kind}" data-shape-left="0" data-shape-top="0" data-shape-rotation="${initialRotation}" data-shape-border-color="${shapeBorderColor}" contenteditable="false" id="${id}" style="${styles[kind]}"><span style="display:block;width:100%;border-top:2px ${lineStyle} ${shapeBorderColor};"></span></span><span>&nbsp;</span>`;
    } else if (svgShapeKinds.has(kind)) {
      html = `<span data-smart-shape="${kind}" data-shape-left="0" data-shape-top="0" data-shape-rotation="${initialRotation}" data-shape-border-color="${shapeBorderColor}" contenteditable="false" id="${id}" style="${styles[kind]}">${svg(kind)}</span><span>&nbsp;</span>`;
    } else if (kind === 'downArrow' || kind === 'rightArrow' || kind === 'leftArrow' || kind === 'upArrow') {
      html = `<span data-smart-shape="${kind}" data-shape-left="0" data-shape-top="0" data-shape-rotation="${initialRotation}" data-shape-border-color="${shapeBorderColor}" contenteditable="false" id="${id}" style="${styles[kind]}"><span data-smart-shape-text="true" contenteditable="true" style="display:inline-block;min-width:1em;outline:none;">${labels[kind]}</span></span><span>&nbsp;</span>`;
    } else if (textKinds.has(kind)) {
      const innerStyle = kind === 'diamond' ? 'display:block;transform:rotate(-45deg);width:100%;text-align:center;outline:none;' : 'display:block;width:100%;text-align:center;outline:none;';
      html = `<span data-smart-shape="${kind}" data-shape-left="0" data-shape-top="0" data-shape-rotation="${initialRotation}" data-shape-border-color="${shapeBorderColor}" contenteditable="false" id="${id}" style="${styles[kind]}"><span data-smart-shape-text="true" contenteditable="true" style="${innerStyle}">${labels[kind]}</span></span><span>&nbsp;</span>`;
    }
    document.execCommand('insertHTML', false, html);
    handleInput();
    setIsShapesMenuOpen(false);

    // Đặt con trỏ vào phần chữ của shape vừa chèn để nhập nội dung ngay.
    setTimeout(() => {
      const shape = editorRef.current?.querySelector(`#${id}`) as HTMLElement | null;
      const text = shape?.querySelector('[data-smart-shape-text]') as HTMLElement | null;
      if (text) {
        text.focus();
        const range = document.createRange();
        range.selectNodeContents(text);
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
      }
    }, 0);
  };

  /** Chèn 1 bảng mới vào vị trí con trỏ trong tài liệu Word (hỏi số dòng/cột trước) */
  const insertWordTable = () => {
    if (!editorRef.current) return;

    const rowsInput = window.prompt('Số dòng của bảng:', '3');
    if (rowsInput === null) return;
    const colsInput = window.prompt('Số cột của bảng:', '3');
    if (colsInput === null) return;

    const rows = Math.max(1, Math.min(50, parseInt(rowsInput, 10) || 3));
    const cols = Math.max(1, Math.min(20, parseInt(colsInput, 10) || 3));

    let tableHtml = '<table style="width:100%;border-collapse:collapse;margin:12px 0;">';
    for (let r = 0; r < rows; r++) {
      tableHtml += '<tr>';
      for (let c = 0; c < cols; c++) {
        tableHtml += '<td style="border:1px solid #000;padding:6px 8px;min-width:60px;">&nbsp;</td>';
      }
      tableHtml += '</tr>';
    }
    tableHtml += '</table><p><br></p>';

    editorRef.current.focus();
    restoreSelection();
    document.execCommand('insertHTML', false, tableHtml);
    handleInput();
  };

  /** Xóa ô bảng Word tại vị trí con trỏ. Không ảnh hưởng các phần tử ngoài bảng. */
  const deleteWordTableCell = () => {
    if (!editorRef.current) return;

    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      alert('Vui lòng đặt con trỏ vào ô cần xóa trong bảng Word.');
      return;
    }

    const anchor = selection.anchorNode;
    const cell = (anchor instanceof HTMLElement
      ? anchor.closest('td, th')
      : anchor?.parentElement?.closest('td, th')) as HTMLTableCellElement | null;

    if (!cell || !editorRef.current.contains(cell)) {
      alert('Vui lòng đặt con trỏ vào ô cần xóa trong bảng Word.');
      return;
    }

    const row = cell.parentElement as HTMLTableRowElement | null;
    const table = row?.closest('table') as HTMLTableElement | null;
    if (!row || !table) return;

    cell.remove();

    // Nếu hàng không còn ô nào, xóa hàng. Nếu bảng không còn hàng nào, xóa bảng.
    if (row.cells.length === 0) row.remove();
    if (table.rows.length === 0) table.remove();

    setTimeout(() => {
      handleInput();
    }, 0);
  };


  /** Xóa toàn bộ bảng Word tại vị trí con trỏ. Không ảnh hưởng nội dung ngoài bảng. */
  const deleteWordTable = () => {
    if (!editorRef.current) return;

    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      alert('Vui lòng đặt con trỏ vào bảng cần xóa.');
      return;
    }

    const anchor = selection.anchorNode;
    const table = (anchor instanceof HTMLElement
      ? anchor.closest('table')
      : anchor?.parentElement?.closest('table')) as HTMLTableElement | null;

    if (!table || !editorRef.current.contains(table)) {
      alert('Vui lòng đặt con trỏ vào bảng cần xóa.');
      return;
    }

    table.remove();

    setTimeout(() => {
      handleInput();
    }, 0);
  };

  const handleReset = async () => {
    if (!selectedFile || !selectedFile.path) return;
    const docKey = `doc_${selectedFile.id || selectedFile.title || selectedFile.path}`;

    localStorage.removeItem(docKey);
    setIsLoading(true);

    try {
      await fetch('/api/document-data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: docKey, content: null }),
      });

      const res = await fetch(selectedFile.path);
      const arrayBuffer = await res.arrayBuffer();

      const mammoth = await import('mammoth');
      const result = await mammoth.convertToHtml({ arrayBuffer });

      setHtmlContent(result.value);
      setWordCount(computeWordCount(result.value));
      if (editorRef.current) {
        editorRef.current.innerHTML = result.value;
      }
    } catch (err) {
      console.error('Lỗi khôi phục mẫu:', err);
    } finally {
      setIsLoading(false);
      setIsSaved(true);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  // Tải file gốc (đưa từ FolderTree lên thanh công cụ tài liệu Word) — giữ nguyên logic gốc
  const handleDownloadWord = () => {
    if (!selectedFile?.path) {
      alert('Vui lòng chọn 1 file biểu mẫu/tài liệu để tải về!');
      return;
    }
    const link = document.createElement('a');
    link.href = selectedFile.path;
    link.download = selectedFile.fileName || `${selectedFile.title}.docx`;
    link.click();
  };

  const handleDownloadPdf = () => {
    if (!selectedFile?.path) {
      alert('Vui lòng chọn 1 file để tải PDF!');
      return;
    }
    const link = document.createElement('a');
    link.href = selectedFile.path;
    link.download = selectedFile.fileName || `${selectedFile.title}.pdf`;
    link.click();
  };

  // ---- Chỉnh sửa trực tiếp bảng tính Excel (kiểu giống Excel) + tự động lưu ----
  const getExcelDocKey = () => {
    if (!selectedFile) return '';
    return `xlsx_${selectedFile.id || selectedFile.title || selectedFile.path}`;
  };

  // Đọc toàn bộ chỉnh sửa hiện tại (theo từng sheet) từ excelSheets đang hiển thị,
  // so với dữ liệu gốc trong workbook, để chỉ lưu lại phần khác biệt (nhẹ, gọn).
  const buildExcelEditsPayload = (sheets: ExcelSheet[]): Record<string, Record<string, ExcelEditEntry>> => {
    const XLSXUtils = excelUtilsRef.current;
    const payload: Record<string, Record<string, ExcelEditEntry>> = {};
    const workbook = excelWorkbookRef.current;

    sheets.forEach((sheet) => {
      const worksheet = workbook?.Sheets?.[sheet.name];
      const sheetEdits: Record<string, ExcelEditEntry> = {};
      sheet.rows.forEach((rowCells) => {
        rowCells.forEach((cell) => {
          let originalText = '';
          if (worksheet && XLSXUtils) {
            const addr = XLSXUtils.encode_cell({ r: cell.r, c: cell.c });
            const original = worksheet[addr];
            originalText = original ? String(original.w ?? original.v ?? '') : '';
          }
          const textChanged = cell.text !== originalText;
          const hasStyle = !!cell.style && Object.keys(cell.style).length > 0;
          if (textChanged || hasStyle) {
            sheetEdits[`${cell.r}-${cell.c}`] = { text: cell.text, style: cell.style };
          }
        });
      });
      if (Object.keys(sheetEdits).length > 0) {
        payload[sheet.name] = sheetEdits;
      }
    });
    return payload;
  };

  const persistExcelEdits = async (sheets: ExcelSheet[]) => {
    const docKey = getExcelDocKey();
    if (!docKey) return;

    try {
      const XLSX = await import('xlsx');
      excelUtilsRef.current = XLSX.utils;
      const payload = buildExcelEditsPayload(sheets);
      const serialized = JSON.stringify(payload);

      localStorage.setItem(docKey, serialized);

      await fetch('/api/document-data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: docKey, content: serialized }),
      });
      setIsExcelSaved(true);
    } catch (err) {
      console.error('Lỗi lưu bảng tính:', err);
    }
  };

  // Lưu tạm local ngay + debounce đồng bộ lên cloud, dùng chung cho mọi loại chỉnh sửa
  const scheduleExcelSave = (sheets: ExcelSheet[]) => {
    setIsExcelSaved(false);
    const docKey = getExcelDocKey();
    if (docKey) {
      try {
        const payload = buildExcelEditsPayload(sheets);
        localStorage.setItem(docKey, JSON.stringify(payload));
      } catch {
        // bỏ qua nếu chưa có cache XLSX utils, sẽ được lưu ở lần debounce kế tiếp
      }
    }
    if (excelSaveTimeoutRef.current) clearTimeout(excelSaveTimeoutRef.current);
    excelSaveTimeoutRef.current = setTimeout(() => {
      persistExcelEdits(sheets);
    }, 800);
  };

  // Lưu trạng thái hiện tại vào lịch sử để phục vụ Undo/Redo (tối đa 40 bước).
  const pushExcelHistory = (sheets: ExcelSheet[]) => {
    excelHistoryRef.current.push(sheets);
    if (excelHistoryRef.current.length > 40) excelHistoryRef.current.shift();
    // Mỗi thay đổi mới tạo một nhánh chỉnh sửa mới, vì vậy Redo cũ không còn phù hợp.
    excelRedoRef.current = [];
  };

  const handleExcelUndo = () => {
    const prevState = excelHistoryRef.current.pop();
    if (!prevState) return;
    setExcelSheets((current) => {
      excelRedoRef.current.push(current);
      scheduleExcelSave(prevState);
      return prevState;
    });
    setSelectedCell(null);
    setExcelSelection(null);
  };

  const handleExcelRedo = () => {
    const nextState = excelRedoRef.current.pop();
    if (!nextState) return;
    setExcelSheets((current) => {
      excelHistoryRef.current.push(current);
      if (excelHistoryRef.current.length > 40) excelHistoryRef.current.shift();
      scheduleExcelSave(nextState);
      return nextState;
    });
    setSelectedCell(null);
    setExcelSelection(null);
  };

  const getExcelSelectionBounds = () => {
    if (!selectedCell) return null;
    const sel = excelSelection && excelSelection.sheetIdx === activeSheet ? excelSelection : null;
    if (!sel) return { r1: selectedCell.r, r2: selectedCell.r, c1: selectedCell.c, c2: selectedCell.c };
    return {
      r1: Math.min(sel.r1, sel.r2),
      r2: Math.max(sel.r1, sel.r2),
      c1: Math.min(sel.c1, sel.c2),
      c2: Math.max(sel.c1, sel.c2),
    };
  };

  const isExcelCellSelected = (r: number, c: number) => {
    const b = getExcelSelectionBounds();
    return !!b && r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2;
  };

  const selectAllExcelPage = () => {
    const sheet = excelSheets[activeSheet];
    if (!sheet || !sheet.rows.length) return;
    const r1 = sheet.startRow;
    const r2 = sheet.startRow + sheet.rows.length - 1;
    const c1 = sheet.startCol;
    const c2 = sheet.endCol;
    const first = sheet.rows[0]?.find((cell) => cell.c === c1) || sheet.rows[0]?.[0];
    if (!first) return;
    setExcelSelection({ sheetIdx: activeSheet, r1, c1, r2, c2 });
    setSelectedCell(first);
    setExcelFormulaValue(first.text);
  };

  const applyExcelBorderToSelection = (side: 'all'|'top'|'bottom'|'left'|'right'|'none'|'outside'|'thickOutside'|'doubleBottom'|'thickBottom'|'topBottom'|'topThickBottom'|'topDoubleBottom') => {
    const b = getExcelSelectionBounds();
    if (!b) return;
    const line = '1px solid #000000';
    const thick = '2px solid #000000';
    const dbl = '3px double #000000';
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        return { ...sheet, rows: sheet.rows.map((rowCells) => rowCells.map((cell) => {
          if (cell.r < b.r1 || cell.r > b.r2 || cell.c < b.c1 || cell.c > b.c2) return cell;
          const border = { ...(cell.style?.border || {}) };
          if (side === 'none') return { ...cell, style: { ...cell.style, border: {} } };
          if (side === 'all') border.top = border.right = border.bottom = border.left = line;
          if (side === 'top') border.top = line;
          if (side === 'bottom') border.bottom = line;
          if (side === 'left') border.left = line;
          if (side === 'right') border.right = line;
          if (side === 'outside' || side === 'thickOutside') {
            const v = side === 'thickOutside' ? thick : line;
            if (cell.r === b.r1) border.top = v;
            if (cell.r === b.r2) border.bottom = v;
            if (cell.c === b.c1) border.left = v;
            if (cell.c === b.c2) border.right = v;
          }
          if (side === 'doubleBottom' && cell.r === b.r2) border.bottom = dbl;
          if (side === 'thickBottom' && cell.r === b.r2) border.bottom = thick;
          if (side === 'topBottom') { if (cell.r === b.r1) border.top = line; if (cell.r === b.r2) border.bottom = line; }
          if (side === 'topThickBottom') { if (cell.r === b.r1) border.top = line; if (cell.r === b.r2) border.bottom = thick; }
          if (side === 'topDoubleBottom') { if (cell.r === b.r1) border.top = line; if (cell.r === b.r2) border.bottom = dbl; }
          return { ...cell, style: { ...cell.style, border } };
        })) };
      });
      scheduleExcelSave(next);
      return next;
    });
  };

  const mergeExcelCells = (mode: 'center'|'across'|'cells'|'unmerge') => {
    const b = getExcelSelectionBounds();
    if (!b) return;
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const merges = { ...sheet.merges };
        if (mode === 'unmerge') {
          Object.entries(merges).forEach(([key, m]) => {
            const mm = m as ExcelMerge;
            const overlap = !(mm.r + mm.rowSpan - 1 < b.r1 || mm.r > b.r2 || mm.c + mm.colSpan - 1 < b.c1 || mm.c > b.c2);
            if (overlap) delete merges[key];
          });
        } else {
          const ranges = mode === 'across'
            ? Array.from({length: b.r2-b.r1+1}, (_,i) => ({r1:b.r1+i,r2:b.r1+i,c1:b.c1,c2:b.c2}))
            : [{r1:b.r1,r2:b.r2,c1:b.c1,c2:b.c2}];
          ranges.forEach(({r1,r2,c1,c2}) => {
            const key = `${r1}-${c1}`;
            const overlap = Object.entries(merges).some(([k,m]) => { const mm=m as ExcelMerge; return k !== key && !(mm.r+mm.rowSpan-1<r1 || mm.r>r2 || mm.c+mm.colSpan-1<c1 || mm.c>c2); });
            if (!overlap) merges[key] = { r:r1,c:c1,rowSpan:r2-r1+1,colSpan:c2-c1+1 };
          });
        }
        let rows = sheet.rows;
        if (mode === 'center' || mode === 'across') {
          rows = rows.map(row => row.map(cell => {
            const inside = cell.r>=b.r1 && cell.r<=b.r2 && cell.c>=b.c1 && cell.c<=b.c2;
            const shouldCenter = mode === 'center' ? inside : inside && cell.c===b.c1;
            return shouldCenter ? { ...cell, style:{...cell.style,align:'center'} } : cell;
          }));
        }
        return { ...sheet, rows, merges, skip: rebuildExcelSkip(merges) };
      });
      scheduleExcelSave(next);
      return next;
    });
  };

  const selectExcelCell = (cell: ExcelCell, extend = false) => {
    if (extend && excelSelection && excelSelection.sheetIdx === activeSheet) {
      setExcelSelection((prev) => prev ? { ...prev, r2: cell.r, c2: cell.c } : prev);
    } else {
      setExcelSelection({ sheetIdx: activeSheet, r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c });
    }
    setSelectedCell(cell);
    setExcelFormulaValue(cell.text);
  };

  // Cập nhật nội dung 1 ô (từ việc gõ trực tiếp trong ô hoặc từ thanh công thức)
  const handleExcelCellEdit = (sheetIdx: number, r: number, c: number, newText: string) => {
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== sheetIdx) return sheet;
        return {
          ...sheet,
          rows: sheet.rows.map((rowCells) =>
            rowCells.map((cell) => (cell.r === r && cell.c === c ? { ...cell, text: newText } : cell))
          ),
        };
      });
      scheduleExcelSave(next);
      return next;
    });

    setSelectedCell((prev) => (prev && prev.r === r && prev.c === c ? { ...prev, text: newText } : prev));
    setExcelFormulaValue(newText);
  };

  // Áp style cho toàn bộ vùng đang chọn; nếu chỉ chọn một ô thì chỉ tác động ô đó.
  const applyStyleToSelectedCell = (patch: CellStyle) => {
    if (!selectedCell) return;
    const bounds = getExcelSelectionBounds();
    if (!bounds) return;
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        return {
          ...sheet,
          rows: sheet.rows.map((rowCells) =>
            rowCells.map((cell) =>
              cell.r >= bounds.r1 && cell.r <= bounds.r2 && cell.c >= bounds.c1 && cell.c <= bounds.c2
                ? { ...cell, style: { ...cell.style, ...patch } }
                : cell
            )
          ),
        };
      });
      scheduleExcelSave(next);
      const found = next[activeSheet]?.rows.flat().find((cell) => cell.r === selectedCell.r && cell.c === selectedCell.c);
      if (found) setSelectedCell(found);
      return next;
    });
  };

  const toggleStyleOnSelectedCell = (key: 'bold' | 'italic' | 'underline') => {
    if (!selectedCell) return;
    const current = selectedCell.style || {};
    if (key === 'bold') applyStyleToSelectedCell({ bold: !current.bold });
    else if (key === 'italic') applyStyleToSelectedCell({ italic: !current.italic });
    else applyStyleToSelectedCell({ underline: !current.underline });
  };

  // --- Gộp ô / Bỏ gộp ô Excel: hỗ trợ vùng chọn tự do bằng Shift + nhấp chuột. ---
  const rebuildExcelSkip = (merges: Record<string, ExcelMerge>) => {
    const skip = new Set<string>();
    Object.values(merges).forEach((m) => {
      for (let rr = m.r; rr < m.r + m.rowSpan; rr++) {
        for (let cc = m.c; cc < m.c + m.colSpan; cc++) {
          if (rr === m.r && cc === m.c) continue;
          skip.add(`${rr}-${cc}`);
        }
      }
    });
    return skip;
  };

  const mergeExcelSelection = (expand: 'right' | 'down') => {
    if (!selectedCell) return;
    const b = getExcelSelectionBounds();
    if (!b) return;
    const r1 = b.r1;
    const c1 = b.c1;
    let r2 = b.r2;
    let c2 = b.c2;
    // Giữ hành vi cũ khi chỉ chọn 1 ô: mỗi lần bấm mở rộng thêm đúng 1 hàng/cột.
    if (r1 === r2 && c1 === c2) {
      if (expand === 'right') c2 += 1;
      else r2 += 1;
    }

    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        if (r2 > sheet.endCol + sheet.rows.length || c2 > sheet.endCol) return sheet;
        const maxRow = sheet.startRow + sheet.rows.length - 1;
        if (r2 > maxRow || c2 > sheet.endCol) return sheet;
        const merges = { ...sheet.merges };
        const targetKey = `${r1}-${c1}`;
        // Không cho gộp chồng lên merge khác.
        for (const [key, m] of Object.entries(merges)) {
          if (key === targetKey) continue;
          const overlap = !(m.r + m.rowSpan - 1 < r1 || m.r > r2 || m.c + m.colSpan - 1 < c1 || m.c > c2);
          if (overlap) return sheet;
        }
        // Nếu ô đầu đã là merge, mở rộng merge hiện tại trong đúng hướng.
        const existing = merges[targetKey];
        const mergedR2 = existing ? Math.max(r2, existing.r + existing.rowSpan - 1) : r2;
        const mergedC2 = existing ? Math.max(c2, existing.c + existing.colSpan - 1) : c2;
        merges[targetKey] = { r: r1, c: c1, rowSpan: mergedR2 - r1 + 1, colSpan: mergedC2 - c1 + 1 };
        return { ...sheet, merges, skip: rebuildExcelSkip(merges) };
      });
      scheduleExcelSave(next);
      return next;
    });
  };

  const mergeSelectedCellRight = () => mergeExcelSelection('right');
  const mergeSelectedCellDown = () => mergeExcelSelection('down');

  const unmergeSelectedCell = () => {
    if (!selectedCell) return;
    const b = getExcelSelectionBounds();
    if (!b) return;
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const merges = { ...sheet.merges };
        const target = Object.entries(merges).find(([, m]) =>
          selectedCell.r >= m.r && selectedCell.r < m.r + m.rowSpan && selectedCell.c >= m.c && selectedCell.c < m.c + m.colSpan
        );
        if (!target) return sheet;
        delete merges[target[0]];
        return { ...sheet, merges, skip: rebuildExcelSkip(merges) };
      });
      scheduleExcelSave(next);
      return next;
    });
  };

  // Chèn / xoá dòng hoặc cột — hoạt động theo đúng dòng/cột người dùng đang chọn, đồng thời giữ merge.
  const handleInsertRow = (afterR: number) => {
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const colCount = sheet.rows[0]?.length ?? sheet.endCol - sheet.startCol + 1;
        const newRow: ExcelCell[] = Array.from({ length: colCount }, (_, i) => ({ r: afterR + 1, c: sheet.startCol + i, text: '' }));
        const rows = sheet.rows.map((rowCells) => rowCells.map((cell) => cell.r > afterR ? { ...cell, r: cell.r + 1 } : cell));
        const insertAt = rows.findIndex((rowCells) => rowCells[0]?.r === afterR + 2);
        rows.splice(insertAt === -1 ? rows.length : insertAt, 0, newRow);
        const merges: Record<string, ExcelMerge> = {};
        Object.values(sheet.merges).forEach((m) => {
          const crosses = m.r <= afterR && afterR < m.r + m.rowSpan;
          const nm = crosses ? { ...m, rowSpan: m.rowSpan + 1 } : m.r > afterR ? { ...m, r: m.r + 1 } : { ...m };
          merges[`${nm.r}-${nm.c}`] = nm;
        });
        return { ...sheet, rows, merges, skip: rebuildExcelSkip(merges) };
      });
      scheduleExcelSave(next);
      return next;
    });
    setSelectedCell(null);
    setExcelSelection(null);
  };

  const handleDeleteRow = (targetR: number) => {
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const rows = sheet.rows
          .filter((rowCells) => rowCells[0]?.r !== targetR)
          .map((rowCells) => rowCells.map((cell) => cell.r > targetR ? { ...cell, r: cell.r - 1 } : cell));
        const merges: Record<string, ExcelMerge> = {};
        Object.values(sheet.merges).forEach((m) => {
          const start = m.r;
          const end = m.r + m.rowSpan - 1;
          if (targetR < start) {
            const nm = { ...m, r: start - 1 };
            merges[`${nm.r}-${nm.c}`] = nm;
          } else if (targetR >= start && targetR <= end) {
            if (m.rowSpan > 1) {
              const nm = { ...m, r: start === targetR ? start : start, rowSpan: m.rowSpan - 1 };
              if (start === targetR) nm.r = start;
              merges[`${nm.r}-${nm.c}`] = nm;
            }
          } else {
            merges[`${m.r}-${m.c}`] = { ...m };
          }
        });
        return { ...sheet, rows, merges, skip: rebuildExcelSkip(merges) };
      });
      scheduleExcelSave(next);
      return next;
    });
    setSelectedCell(null);
    setExcelSelection(null);
  };

  const handleInsertColumn = (afterC: number) => {
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const rows = sheet.rows.map((rowCells) => {
          const shifted = rowCells.map((cell) => cell.c > afterC ? { ...cell, c: cell.c + 1 } : cell);
          const insertAt = shifted.findIndex((cell) => cell.c === afterC + 2);
          shifted.splice(insertAt === -1 ? shifted.length : insertAt, 0, { r: rowCells[0]?.r ?? 0, c: afterC + 1, text: '' });
          return shifted;
        });
        const merges: Record<string, ExcelMerge> = {};
        Object.values(sheet.merges).forEach((m) => {
          const crosses = m.c <= afterC && afterC < m.c + m.colSpan;
          const nm = crosses ? { ...m, colSpan: m.colSpan + 1 } : m.c > afterC ? { ...m, c: m.c + 1 } : { ...m };
          merges[`${nm.r}-${nm.c}`] = nm;
        });
        return { ...sheet, rows, merges, skip: rebuildExcelSkip(merges), endCol: sheet.endCol + 1 };
      });
      scheduleExcelSave(next);
      return next;
    });
    setSelectedCell(null);
    setExcelSelection(null);
  };

  const handleDeleteColumn = (targetC: number) => {
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const rows = sheet.rows.map((rowCells) => rowCells.filter((cell) => cell.c !== targetC).map((cell) => cell.c > targetC ? { ...cell, c: cell.c - 1 } : cell));
        const merges: Record<string, ExcelMerge> = {};
        Object.values(sheet.merges).forEach((m) => {
          const start = m.c;
          const end = m.c + m.colSpan - 1;
          if (targetC < start) {
            const nm = { ...m, c: start - 1 };
            merges[`${nm.r}-${nm.c}`] = nm;
          } else if (targetC >= start && targetC <= end) {
            if (m.colSpan > 1) {
              const nm = { ...m, colSpan: m.colSpan - 1 };
              merges[`${nm.r}-${nm.c}`] = nm;
            }
          } else {
            merges[`${m.r}-${m.c}`] = { ...m };
          }
        });
        return { ...sheet, rows, merges, skip: rebuildExcelSkip(merges), endCol: Math.max(sheet.startCol, sheet.endCol - 1) };
      });
      scheduleExcelSave(next);
      return next;
    });
    setSelectedCell(null);
    setExcelSelection(null);
  };

  const handleExcelPaste = (e: React.ClipboardEvent<HTMLTableCellElement>, cell: ExcelCell) => {
    const plain = e.clipboardData.getData('text/plain');
    if (!plain || !plain.includes('\t') && !plain.includes('\n')) return;
    e.preventDefault();
    const matrix = plain.replace(/\r/g, '').split('\n').filter((row, idx, arr) => !(idx === arr.length - 1 && row === '')).map((row) => row.split('\t'));
    if (!matrix.length) return;
    setExcelSheets((prev) => {
      pushExcelHistory(prev);
      const next = prev.map((sheet, idx) => {
        if (idx !== activeSheet) return sheet;
        const map = new Map(sheet.rows.flat().map((c) => [`${c.r}-${c.c}`, c]));
        matrix.forEach((row, ri) => row.forEach((text, ci) => {
          const key = `${cell.r + ri}-${cell.c + ci}`;
          const old = map.get(key);
          if (old) map.set(key, { ...old, text });
        }));
        return { ...sheet, rows: sheet.rows.map((rowCells) => rowCells.map((c) => map.get(`${c.r}-${c.c}`) || c)) };
      });
      scheduleExcelSave(next);
      return next;
    });
  };

  // Phím tắt Excel: Ctrl/Cmd+Z, Ctrl/Cmd+Y hoặc Ctrl/Cmd+Shift+Z.
  useEffect(() => {
    const onExcelKeyDown = (e: KeyboardEvent) => {
      if (getFileType(selectedFile || ({} as DocumentNode)) !== 'excel') return;
      const target = e.target as HTMLElement | null;
      if (!target?.closest('[data-excel-editor]')) return;
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      if (e.key.toLowerCase() === 'z' && !e.shiftKey) { e.preventDefault(); handleExcelUndo(); }
      else if (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey)) { e.preventDefault(); handleExcelRedo(); }
    };
    document.addEventListener('keydown', onExcelKeyDown);
    return () => document.removeEventListener('keydown', onExcelKeyDown);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedFile]);

  // Xuất bảng tính hiện tại (đã chỉnh sửa) thành file .xlsx để tải xuống.
  // Lưu ý: thư viện xlsx bản miễn phí chỉ ghi được nội dung + ô gộp,
  // định dạng (đậm/nghiêng/màu) chỉ hiển thị trên màn hình, chưa xuất ra được file .xlsx.
  const handleDownloadExcel = async () => {
    if (!selectedFile || excelSheets.length === 0) return;
    try {
      const XLSX = await import('xlsx');
      const newWorkbook = XLSX.utils.book_new();
      excelSheets.forEach((sheet) => {
        const aoa = sheet.rows.map((rowCells) => rowCells.map((cell) => cell.text));
        const worksheet = XLSX.utils.aoa_to_sheet(aoa);
        // Ghi lại công thức thực thay vì biến công thức thành chuỗi khi xuất.
        sheet.rows.forEach((rowCells) => rowCells.forEach((cell) => {
          if (typeof cell.text === 'string' && cell.text.startsWith('=')) {
            const addr = XLSX.utils.encode_cell({ r: cell.r - sheet.startRow, c: cell.c - sheet.startCol });
            if (worksheet[addr]) worksheet[addr] = { ...worksheet[addr], f: cell.text.slice(1) };
          }
        }));
        const merges = Object.values(sheet.merges).map((m) => ({
          s: { r: m.r - sheet.startRow, c: m.c - sheet.startCol },
          e: { r: m.r - sheet.startRow + m.rowSpan - 1, c: m.c - sheet.startCol + m.colSpan - 1 },
        }));
        if (merges.length > 0) worksheet['!merges'] = merges;
        worksheet['!cols'] = Array.from({ length: sheet.endCol - sheet.startCol + 1 }, (_, i) => ({ wpx: getExcelColWidth(excelSheets.indexOf(sheet), sheet.startCol + i) }));
        worksheet['!rows'] = sheet.rows.map((rowCells) => ({ hpx: getExcelRowHeight(excelSheets.indexOf(sheet), rowCells[0]?.r ?? 0) }));
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheet.name);
      });
      const fileName = (selectedFile.title || 'bang-tinh').replace(/\.[^/.]+$/, '') + '.xlsx';
      XLSX.writeFile(newWorkbook, fileName);
    } catch (err) {
      console.error('Lỗi xuất file Excel:', err);
      alert('Không thể xuất Excel. Vui lòng thử lại.');
    }
  };

  // ---- Icon set (inline SVG, stroke-based — no external icon dependency) ----
  const Icon = {
    Lock: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4.5" y="10.5" width="15" height="9.5" rx="2.2" />
        <path d="M7.5 10.5V7.8a4.5 4.5 0 0 1 9 0v2.7" />
        <circle cx="12" cy="15" r="1.4" fill="currentColor" stroke="none" />
      </svg>
    ),
    User: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="12" cy="8" r="3.4" />
        <path d="M5 19.5c1.2-3.4 4-5 7-5s5.8 1.6 7 5" />
      </svg>
    ),
    Key: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="8" cy="15.5" r="3.2" />
        <path d="M10.3 13.2 18 5.5m0 0h-3.4M18 5.5V9" />
      </svg>
    ),
    Doc: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
        <path d="M13.2 3.5V8h5" />
      </svg>
    ),
    Table: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4" y="5" width="16" height="14" rx="1.6" />
        <path d="M4 10h16M4 15h16M10 5v14M15 5v14" />
      </svg>
    ),
    DeleteCell: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4" y="5" width="16" height="14" rx="1.6" />
        <path d="M4 10h16M4 15h16M10 5v14M15 5v5" />
        <path d="m15.5 14.5 4 4m0-4-4 4" />
      </svg>
    ),
    DeleteTable: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4" y="5" width="16" height="14" rx="1.6" />
        <path d="M4 10h16M4 15h16M10 5v14M15 5v5" />
        <path d="m15 13.5 5 5m0-5-5 5" />
      </svg>
    ),
    Shapes: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="4" y="4" width="8" height="8" rx="1.2" />
        <circle cx="16.5" cy="15.5" r="4" />
        <path d="m12.5 19.5 3-3 3 3" />
      </svg>
    ),
    Globe: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="12" cy="12" r="8" />
        <path d="M4 12h16M12 4c2.4 2.2 3.6 5 3.6 8s-1.2 5.8-3.6 8c-2.4-2.2-3.6-5-3.6-8s1.2-5.8 3.6-8Z" />
      </svg>
    ),
    Pen: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4 20l.9-3.6L15.6 5.7a1.7 1.7 0 0 1 2.4 0l.3.3a1.7 1.7 0 0 1 0 2.4L7.6 19.1 4 20Z" />
        <path d="M13.8 7.5l2.7 2.7" />
      </svg>
    ),
    Cloud: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M7 18h10a3.5 3.5 0 0 0 .4-6.98A5.5 5.5 0 0 0 7.1 9.6 4 4 0 0 0 7 18Z" />
      </svg>
    ),
    Refresh: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4.5 12a7.5 7.5 0 0 1 12.6-5.5M19.5 12a7.5 7.5 0 0 1-12.6 5.5" />
        <path d="M17.5 4.5V8h-3.4M6.5 19.5V16h3.4" />
      </svg>
    ),
    Print: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 8.5V4h11v4.5" />
        <rect x="4.5" y="8.5" width="15" height="7" rx="1.4" />
        <path d="M6.5 15h11V20h-11z" />
      </svg>
    ),
    Fill: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="m5.5 12.5 6-6 6 6-6 6z" />
        <path d="M9 6.5 13.5 2" />
        <path d="M4 17c0 1.4 1.1 2.5 2.5 2.5S9 18.4 9 17c0-1-1-2-2.5-3.5C5 15 4 16 4 17z" fill="currentColor" stroke="none" />
      </svg>
    ),
    AlignIcon: ({ align, ...p }: React.SVGProps<SVGSVGElement> & { align: 'left' | 'center' | 'right' }) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4.5 6.5h15" />
        {align === 'left' && <path d="M4.5 12h9M4.5 17.5h12" />}
        {align === 'center' && <path d="M7 12h10M6 17.5h12" />}
        {align === 'right' && <path d="M10.5 12h9M7.5 17.5h12" />}
      </svg>
    ),
    Logout: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M14.5 8V6a1.5 1.5 0 0 0-1.5-1.5H6A1.5 1.5 0 0 0 4.5 6v12A1.5 1.5 0 0 0 6 19.5h7A1.5 1.5 0 0 0 14.5 18v-2" />
        <path d="M9.5 12h10m0 0-3-3m3 3-3 3" />
      </svg>
    ),
    Spinner: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" {...p}>
        <circle cx="12" cy="12" r="9" stroke="currentColor" strokeOpacity="0.2" strokeWidth="2.4" />
        <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" />
      </svg>
    ),
    File: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
        <path d="M9 12.5h6M9 15.5h6M9 9.5h2.5" />
      </svg>
    ),
    Bold: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M7 5.5h5.3a3.4 3.4 0 0 1 0 6.8H7z" />
        <path d="M7 12.3h6a3.6 3.6 0 0 1 0 7.2H7z" />
      </svg>
    ),
    Italic: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M11.5 5.5h6M6.5 18.5h6M14.5 5.5l-4 13" />
      </svg>
    ),
    Underline: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 4.5v6.5a5.5 5.5 0 0 0 11 0V4.5" />
        <path d="M5 19.5h14" />
      </svg>
    ),
    Strike: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M5 12h14" />
        <path d="M8 6.5c0-1.4 1.7-2.2 4-2.2s4 .9 4 2.4" />
        <path d="M8 17.3c0 1.5 1.7 2.4 4 2.4s4-.9 4-2.4c0-1.2-.7-2-2.2-2.6" />
      </svg>
    ),
    AlignLeft: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" {...p}>
        <path d="M4.5 6h15M4.5 11h10M4.5 16h15M4.5 21h10" transform="translate(0 -3)" />
      </svg>
    ),
    AlignCenter: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" {...p}>
        <path d="M4.5 6h15M7 11h10M4.5 16h15M7 21h10" transform="translate(0 -3)" />
      </svg>
    ),
    AlignRight: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" {...p}>
        <path d="M4.5 6h15M9.5 11h10M4.5 16h15M9.5 21h10" transform="translate(0 -3)" />
      </svg>
    ),
    AlignJustify: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" {...p}>
        <path d="M4.5 6h15M4.5 11h15M4.5 16h15M4.5 21h15" transform="translate(0 -3)" />
      </svg>
    ),
    ListBullet: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="5.5" cy="7" r="1.1" fill="currentColor" stroke="none" />
        <circle cx="5.5" cy="12" r="1.1" fill="currentColor" stroke="none" />
        <circle cx="5.5" cy="17" r="1.1" fill="currentColor" stroke="none" />
        <path d="M9.5 7h9M9.5 12h9M9.5 17h9" />
      </svg>
    ),
    ListNumber: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M9.5 7h9M9.5 12h9M9.5 17h9" />
        <text x="2.5" y="8.3" fontSize="5.5" fill="currentColor" stroke="none">1</text>
        <text x="2.5" y="13.3" fontSize="5.5" fill="currentColor" stroke="none">2</text>
        <text x="2.5" y="18.3" fontSize="5.5" fill="currentColor" stroke="none">3</text>
      </svg>
    ),
    Indent: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4.5 6h15M10.5 11h9M10.5 16h9M4.5 21h15" transform="translate(0 -3)" />
        <path d="M5 9.5l3 2.5-3 2.5" transform="translate(0 -3)" />
      </svg>
    ),
    Outdent: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4.5 6h15M10.5 11h9M10.5 16h9M4.5 21h15" transform="translate(0 -3)" />
        <path d="M8 9.5l-3 2.5 3 2.5" transform="translate(0 -3)" />
      </svg>
    ),
    Undo: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M7.5 8.5H16a4.3 4.3 0 0 1 0 8.6h-6" />
        <path d="M10 5.2 6.5 8.5 10 11.8" />
      </svg>
    ),
    Redo: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M16.5 8.5H8a4.3 4.3 0 0 0 0 8.6h6" />
        <path d="M14 5.2l3.5 3.3L14 11.8" />
      </svg>
    ),
    TextColor: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M9 15.5 12 6l3 9.5M10 12.5h4" />
        <path d="M5 20.5h14" strokeWidth="3" />
      </svg>
    ),
    Highlight: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M13.5 4.5 19 10 9.5 19.5H4V14z" />
        <path d="M4 20.5h16" />
      </svg>
    ),
    Menu: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M4 6.5h16M4 12h16M4 17.5h16" />
      </svg>
    ),
    Close: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6 6l12 12M18 6L6 18" />
      </svg>
    ),
    ChevronsLeft: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M13 6l-6 6 6 6M19 6l-6 6 6 6" />
      </svg>
    ),
    ChevronsRight: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M11 6l6 6-6 6M5 6l6 6-6 6" />
      </svg>
    ),
    Eye: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M2.5 12s3.5-6.5 9.5-6.5S21.5 12 21.5 12s-3.5 6.5-9.5 6.5S2.5 12 2.5 12Z" />
        <circle cx="12" cy="12" r="2.6" />
      </svg>
    ),
    EyeOff: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M3.5 3.5l17 17" />
        <path d="M10.6 5.7A9.9 9.9 0 0 1 12 5.6c6 0 9.5 6.5 9.5 6.5a15.6 15.6 0 0 1-3.3 4.1M6.9 6.9C4.3 8.5 2.5 11.6 2.5 12S6 18.5 12 18.5c1.1 0 2.1-.2 3-.5" />
        <path d="M9.9 10a2.6 2.6 0 0 0 3.7 3.6" />
      </svg>
    ),
    Building: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <rect x="5" y="3.5" width="14" height="17" rx="1.4" />
        <path d="M9 8h.01M12 8h.01M15 8h.01M9 12h.01M12 12h.01M15 12h.01M9 16h.01M12 16h.01M15 16h.01" strokeWidth="2.2" />
      </svg>
    ),
    Shield: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M12 3.5l7 2.6v5.4c0 4.7-3 7.9-7 9-4-1.1-7-4.3-7-9V6.1z" />
        <path d="M9.2 12.2l1.9 1.9 3.7-3.9" />
      </svg>
    ),
    Search: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="11" cy="11" r="6.5" />
        <path d="m20 20-4.3-4.3" />
      </svg>
    ),
    Download: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M12 4v11m0 0-3.5-3.5M12 15l3.5-3.5" />
        <path d="M5 17.5v1.7a1.3 1.3 0 0 0 1.3 1.3h11.4a1.3 1.3 0 0 0 1.3-1.3v-1.7" />
      </svg>
    ),
    Pdf: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <path d="M6.5 3.5h7L18.5 8v12.5a1 1 0 0 1-1 1h-11a1 1 0 0 1-1-1V4.5a1 1 0 0 1 1-1Z" />
        <path d="M13.2 3.5V8h5" />
        <text x="7.3" y="17" fontSize="6.2" fontWeight="700" fill="currentColor" stroke="none" fontFamily="Inter, sans-serif">PDF</text>
      </svg>
    ),
    ZoomOut: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="10.5" cy="10.5" r="6.5" />
        <path d="m20 20-4.35-4.35M8 10.5h5" />
      </svg>
    ),
    ZoomIn: (p: React.SVGProps<SVGSVGElement>) => (
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}>
        <circle cx="10.5" cy="10.5" r="6.5" />
        <path d="m20 20-4.35-4.35M10.5 8v5M8 10.5h5" />
      </svg>
    ),
  };

  const RibbonBtn = ({
    onClick,
    active,
    title,
    children,
  }: {
    onClick: () => void;
    active?: boolean;
    title: string;
    children: React.ReactNode;
  }) => (
    <button
      type="button"
      title={title}
      aria-label={title}
      aria-pressed={!!active}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
      className={`inline-flex items-center justify-center w-8 h-8 rounded-lg transition-all duration-150 cursor-pointer shrink-0 ${
        active
          ? 'bg-teal-700 text-white shadow-sm shadow-teal-900/20'
          : 'text-slate-600 hover:bg-slate-100 active:bg-slate-200'
      }`}
    >
      {children}
    </button>
  );

  // Nhãn/mô tả loại tệp cho thanh trạng thái & badge
  const fileTypeMeta: Record<string, { label: string; className: string }> = {
    word: { label: 'Word', className: 'bg-blue-50 text-blue-700 border-blue-200/80' },
    excel: { label: 'Excel', className: 'bg-emerald-50 text-emerald-700 border-emerald-200/80' },
    pdf: { label: 'PDF', className: 'bg-rose-50 text-rose-700 border-rose-200/80' },
    image: { label: 'Ảnh', className: 'bg-violet-50 text-violet-700 border-violet-200/80' },
    text: { label: 'Tài liệu', className: 'bg-slate-100 text-slate-600 border-slate-200/80' },
  };

  const currentFileType = selectedFile ? getFileType(selectedFile) : null;
  const currentSyncOk = currentFileType === 'excel' ? isExcelSaved : isSaved;

  const openWordTextColorPicker = () => {
    saveSelection();
    document.getElementById('word-text-color-picker')?.click();
  };

  const openWordHighlightPicker = () => {
    saveSelection();
    document.getElementById('word-highlight-picker')?.click();
  };

  const renderContent = () => {
    if (!selectedFile) {
      return (
        <div className="flex flex-col items-center justify-center h-full text-center px-6 animate-riseIn">
          <div className="relative mb-5">
            <div className="absolute inset-0 rounded-full bg-teal-500/10 blur-xl scale-150" />
            <div className="relative w-16 h-16 rounded-2xl bg-gradient-to-br from-teal-50 to-teal-100 border border-teal-200/70 flex items-center justify-center shadow-sm">
              <Icon.File className="w-7 h-7 text-teal-600" />
            </div>
          </div>
          <h3 className="text-[15px] font-bold text-slate-700 mb-1.5">Chưa chọn tài liệu</h3>
          <p className="text-[13px] font-medium text-slate-500 tracking-wide max-w-xs leading-relaxed">
            Chọn một tài liệu từ danh mục bên trái để bắt đầu làm việc.
          </p>
          <p className="text-[12px] text-slate-400 mt-1.5 max-w-xs">Bạn có thể mở Word, PDF hoặc Excel.</p>
        </div>
      );
    }

    const fileType = getFileType(selectedFile);
    const contentKey = selectedFile.id || selectedFile.path || selectedFile.title;

    if (fileType === 'text' || selectedFile.content) {
      return (
        <div
          key={contentKey}
          className="bg-white p-8 rounded-2xl border border-slate-200/70 shadow-[0_1px_2px_rgba(15,50,55,0.04),0_12px_28px_-12px_rgba(15,50,55,0.12)] max-w-4xl mx-auto my-6 overflow-y-auto max-h-full animate-riseIn"
        >
          <div className="flex items-center justify-between pb-4 mb-4 border-b border-slate-100">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-teal-50 text-teal-700 font-semibold text-[11px] uppercase tracking-wide rounded-full border border-teal-200/70">
              <Icon.Globe className="w-3.5 h-3.5" />
              Tra cứu Internet
            </span>
          </div>
          <div
            className="prose prose-slate max-w-none text-slate-800 leading-relaxed whitespace-pre-line text-sm"
            style={{ fontFamily: '"Times New Roman", Times, serif' }}
          >
            {selectedFile.content}
          </div>
        </div>
      );
    }

    if (fileType === 'pdf' && selectedFile.path) {
      return (
        <div key={contentKey} className="flex flex-col h-full w-full min-w-0 bg-[#F1F3F1] overflow-hidden animate-riseIn">
          <div className="bg-white/90 backdrop-blur-md border-b border-slate-200/70 px-4 py-2 flex items-center justify-between gap-3 shadow-sm print:hidden shrink-0 z-10 animate-slideDown">
            <div className="flex items-center gap-2 min-w-0">
              <span className="inline-flex items-center justify-center w-7 h-7 rounded-lg bg-rose-50 border border-rose-200/80 text-rose-600 shrink-0">
                <Icon.Doc className="w-4 h-4" />
              </span>
              <span className="text-[13px] font-semibold text-slate-700 truncate">{selectedFile.title}</span>
              <span className="inline-flex items-center px-2 py-0.5 text-[10.5px] font-bold uppercase tracking-wide rounded-full border shrink-0 bg-rose-50 text-rose-700 border-rose-200/80">
                PDF
              </span>
            </div>
            <button
              onClick={handlePrint}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-white bg-teal-700 hover:bg-teal-600 rounded-xl shadow-sm shadow-teal-900/20 transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer shrink-0"
            >
              <Icon.Print className="w-3.5 h-3.5" />
              In / Trích xuất PDF
            </button>
          </div>
          <div className="flex-1 min-h-0 p-3">
            <iframe
              src={`${selectedFile.path}#toolbar=1`}
              className="w-full h-full border-0 rounded-xl shadow-md bg-white"
              title={selectedFile.title}
            />
          </div>
        </div>
      );
    }

    if (fileType === 'image' && selectedFile.path) {
      return (
        <div key={contentKey} className="flex flex-col h-full w-full min-w-0 bg-[#F1F3F1] overflow-hidden animate-riseIn">
          <div className="bg-white/90 backdrop-blur-md border-b border-slate-200/70 px-4 py-2 flex items-center justify-between gap-3 shadow-sm print:hidden shrink-0 z-10 animate-slideDown">
            <div className="flex items-center gap-2 min-w-0">
              <span className="inline-flex items-center justify-center w-7 h-7 rounded-lg bg-violet-50 border border-violet-200/80 text-violet-600 shrink-0">
                <Icon.Doc className="w-4 h-4" />
              </span>
              <span className="text-[13px] font-semibold text-slate-700 truncate">{selectedFile.title}</span>
              <span className="inline-flex items-center px-2 py-0.5 text-[10.5px] font-bold uppercase tracking-wide rounded-full border shrink-0 bg-violet-50 text-violet-700 border-violet-200/80">
                Ảnh
              </span>
            </div>

            <div className="flex items-center gap-1.5 shrink-0">
              <button
                onClick={() => setImageZoom((z) => Math.max(IMAGE_ZOOM_MIN, z - IMAGE_ZOOM_STEP))}
                disabled={imageZoom <= IMAGE_ZOOM_MIN}
                aria-label="Thu nhỏ"
                title="Thu nhỏ"
                className="w-7 h-7 flex items-center justify-center rounded-md text-slate-500 hover:text-teal-700 hover:bg-teal-50 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
              >
                <Icon.ZoomOut className="w-4 h-4" />
              </button>
              <button
                onClick={() => setImageZoom(100)}
                title="Đặt lại mức thu phóng 100%"
                className="w-12 text-center font-mono text-[11px] text-slate-500 hover:text-teal-700 cursor-pointer"
              >
                {imageZoom}%
              </button>
              <button
                onClick={() => setImageZoom((z) => Math.min(IMAGE_ZOOM_MAX, z + IMAGE_ZOOM_STEP))}
                disabled={imageZoom >= IMAGE_ZOOM_MAX}
                aria-label="Phóng to"
                title="Phóng to"
                className="w-7 h-7 flex items-center justify-center rounded-md text-slate-500 hover:text-teal-700 hover:bg-teal-50 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
              >
                <Icon.ZoomIn className="w-4 h-4" />
              </button>

              <div className="w-px h-5 bg-slate-200 mx-1" />

              <a
                href={selectedFile.path}
                download={selectedFile.fileName || selectedFile.title}
                title="Tải ảnh xuống"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-teal-700 bg-teal-50 hover:bg-teal-100 border border-teal-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Download className="w-3.5 h-3.5" />
                Tải xuống
              </a>
              <button
                onClick={handlePrint}
                title="In ảnh"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-slate-600 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Print className="w-3.5 h-3.5" />
                In
              </button>
            </div>
          </div>

          <div className="flex-1 min-h-0 overflow-auto p-6 flex items-start justify-center">
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img
              src={selectedFile.path}
              alt={selectedFile.title}
              style={{ width: `${imageZoom}%`, maxWidth: 'none' }}
              className="rounded-xl shadow-[0_1px_1px_rgba(15,50,55,0.05),0_20px_40px_-16px_rgba(15,50,55,0.18)] border border-slate-200 bg-white print:shadow-none print:border-none"
            />
          </div>
        </div>
      );
    }

    if (fileType === 'excel') {
      const activeSheetData = excelSheets[activeSheet];
      const hasError = !!activeSheetData?.error;

      // Đã tắt hoàn toàn chức năng tách trang Excel theo yêu cầu.
      // Toàn bộ dữ liệu của sheet được hiển thị liên tục trong một bảng duy nhất.

      return (
        <div key={contentKey} className="flex flex-col h-full w-full min-w-0 bg-[#F1F3F1] overflow-hidden animate-riseIn">
          {/* Thanh trên: nhãn loại file + nút in */}
          <div className="bg-white/90 backdrop-blur-md border-b border-slate-200/70 px-4 py-2 flex items-center justify-between gap-4 shadow-sm print:hidden shrink-0 z-20 animate-slideDown">
            <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-full border border-emerald-200/80">
              <Icon.Pen className="w-3.5 h-3.5" />
              Bảng tính &middot; chỉnh sửa trực tiếp
            </span>

            <span className="inline-flex items-center gap-1.5 text-[11px] font-medium text-slate-400 shrink-0">
              {isExcelSaved ? (
                <>
                  <Icon.Cloud className="w-3.5 h-3.5 text-teal-500" />
                  Đã lưu tự động
                </>
              ) : (
                <>
                  <Icon.Spinner className="w-3.5 h-3.5 animate-spin text-amber-500" />
                  Đang lưu...
                </>
              )}
            </span>

            <div className="flex-1" />

            <button
              onClick={handleExcelUndo}
              title="Hoàn tác (Ctrl+Z)"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-xl transition-all duration-200 cursor-pointer shrink-0"
            >
              <Icon.Undo className="w-3.5 h-3.5" />
              Hoàn tác
            </button>
            <button
              onClick={handleExcelRedo}
              title="Làm lại (Ctrl+Y)"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-xl transition-all duration-200 cursor-pointer shrink-0"
            >
              <Icon.Redo className="w-3.5 h-3.5" />
              Làm lại
            </button>
            <button
              onClick={handleDownloadExcel}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-teal-700 bg-teal-50 hover:bg-teal-100 border border-teal-200/80 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer shrink-0"
            >
              <Icon.Cloud className="w-3.5 h-3.5" />
              Tải xuống .xlsx
            </button>
            <button
              onClick={handlePrint}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-white bg-teal-700 hover:bg-teal-600 rounded-xl shadow-sm shadow-teal-900/20 transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer shrink-0"
            >
              <Icon.Print className="w-3.5 h-3.5" />
              In / Trích xuất PDF
            </button>
          </div>

          {/* Thanh công cụ định dạng kiểu Excel: đậm/nghiêng/gạch chân/căn lề/màu/chèn-xoá dòng-cột */}
          <div className="bg-white border-b border-slate-200/70 px-3 py-1.5 flex flex-wrap items-center gap-1.5 shrink-0 print:hidden z-10">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 pr-1 hidden lg:inline">Định dạng</span>
            <button
              onClick={() => toggleStyleOnSelectedCell('bold')}
              disabled={!selectedCell}
              title="In đậm"
              aria-label="In đậm"
              className={`w-7 h-7 flex items-center justify-center rounded-md text-[13px] font-bold border transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer ${
                selectedCell?.style?.bold ? 'bg-teal-600 text-white border-teal-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-100'
              }`}
            >
              B
            </button>
            <button
              onClick={() => toggleStyleOnSelectedCell('italic')}
              disabled={!selectedCell}
              title="In nghiêng"
              aria-label="In nghiêng"
              className={`w-7 h-7 flex items-center justify-center rounded-md text-[13px] italic font-semibold border transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer ${
                selectedCell?.style?.italic ? 'bg-teal-600 text-white border-teal-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-100'
              }`}
            >
              I
            </button>
            <button
              onClick={() => toggleStyleOnSelectedCell('underline')}
              disabled={!selectedCell}
              title="Gạch chân"
              aria-label="Gạch chân"
              className={`w-7 h-7 flex items-center justify-center rounded-md text-[13px] underline font-semibold border transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer ${
                selectedCell?.style?.underline ? 'bg-teal-600 text-white border-teal-600' : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-100'
              }`}
            >
              U
            </button>

            <div className="w-px h-5 bg-slate-200 mx-1" />

            {(['left', 'center', 'right'] as const).map((align) => (
              <button
                key={align}
                onClick={() => applyStyleToSelectedCell({ align })}
                disabled={!selectedCell}
                title={align === 'left' ? 'Căn trái' : align === 'center' ? 'Căn giữa' : 'Căn phải'}
                aria-label={align === 'left' ? 'Căn trái' : align === 'center' ? 'Căn giữa' : 'Căn phải'}
                className={`w-7 h-7 flex items-center justify-center rounded-md border transition-colors disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer ${
                  (selectedCell?.style?.align ?? 'left') === align
                    ? 'bg-teal-600 text-white border-teal-600'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-100'
                }`}
              >
                <Icon.AlignIcon align={align} className="w-4 h-4" />
              </button>
            ))}

            <div className="w-px h-5 bg-slate-200 mx-1" />

            <label title="Màu chữ" className="w-7 h-7 flex items-center justify-center rounded-md border border-slate-200 bg-white hover:bg-slate-100 cursor-pointer relative">
              <span className="text-[13px] font-bold" style={{ color: selectedCell?.style?.color || '#334155' }}>A</span>
              <input
                type="color"
                disabled={!selectedCell}
                value={selectedCell?.style?.color || '#334155'}
                onChange={(e) => applyStyleToSelectedCell({ color: e.target.value })}
                className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed"
                aria-label="Màu chữ"
              />
            </label>
            <label title="Màu nền ô" className="w-7 h-7 flex items-center justify-center rounded-md border border-slate-200 bg-white hover:bg-slate-100 cursor-pointer relative">
              <Icon.Fill className="w-4 h-4" style={{ color: selectedCell?.style?.bg || '#94a3b8' }} />
              <input
                type="color"
                disabled={!selectedCell}
                value={selectedCell?.style?.bg || '#ffffff'}
                onChange={(e) => applyStyleToSelectedCell({ bg: e.target.value })}
                className="absolute inset-0 opacity-0 cursor-pointer disabled:cursor-not-allowed"
                aria-label="Màu nền ô"
              />
            </label>

            <div className="w-px h-5 bg-slate-200 mx-1" />
            <div className="relative">
              <button type="button" disabled={!selectedCell} onClick={() => setIsExcelBorderMenuOpen(v => !v)} title="Viền ô" className="inline-flex items-center gap-1 px-2.5 h-7 text-[11.5px] font-semibold text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-md disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer">▦ Viền</button>
              {isExcelBorderMenuOpen && <div className="absolute left-0 top-8 z-[80] w-[190px] rounded-lg border border-slate-200 bg-white p-1.5 shadow-xl">{excelBorderOptions.map(([k,l]) => <button key={k} type="button" onClick={() => { applyExcelBorderToSelection(k); setIsExcelBorderMenuOpen(false); }} className="w-full text-left px-2.5 py-1.5 text-[11px] rounded hover:bg-slate-50">{l}</button>)}</div>}
            </div>
            <div className="w-px h-5 bg-slate-200 mx-1" />
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 pr-1 hidden lg:inline">Gộp ô</span>
            <div className="relative">
              <button type="button" disabled={!selectedCell} onClick={() => setIsExcelMergeMenuOpen(v => !v)} className="inline-flex items-center gap-1 px-2.5 h-7 text-[11.5px] font-semibold text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-md disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer">Gộp ▾</button>
              {isExcelMergeMenuOpen && <div className="absolute left-0 top-8 z-[80] w-[175px] rounded-lg border border-slate-200 bg-white p-1.5 shadow-xl">{excelMergeOptions.map(([k,l]) => <button key={k} type="button" onClick={() => { mergeExcelCells(k); setIsExcelMergeMenuOpen(false); }} className="w-full text-left px-2.5 py-1.5 text-[11px] rounded hover:bg-slate-50">{l}</button>)}</div>}
            </div>



            <div className="w-px h-5 bg-slate-200 mx-1" />
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 pr-1 hidden lg:inline">Dòng / Cột</span>

            <button
              onClick={() => selectedCell && handleInsertRow(selectedCell.r)}
              disabled={!selectedCell}
              className="inline-flex items-center gap-1 px-2.5 h-7 text-[11.5px] font-semibold text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-md disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
            >
              + Dòng
            </button>
            <button
              onClick={() => selectedCell && handleDeleteRow(selectedCell.r)}
              disabled={!selectedCell}
              className="inline-flex items-center gap-1 px-2.5 h-7 text-[11.5px] font-semibold text-rose-600 bg-white hover:bg-rose-50 border border-slate-200 rounded-md disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
            >
              − Dòng
            </button>
            <button
              onClick={() => selectedCell && handleInsertColumn(selectedCell.c)}
              disabled={!selectedCell}
              className="inline-flex items-center gap-1 px-2.5 h-7 text-[11.5px] font-semibold text-slate-600 bg-white hover:bg-slate-100 border border-slate-200 rounded-md disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
            >
              + Cột
            </button>
            <button
              onClick={() => selectedCell && handleDeleteColumn(selectedCell.c)}
              disabled={!selectedCell}
              className="inline-flex items-center gap-1 px-2.5 h-7 text-[11.5px] font-semibold text-rose-600 bg-white hover:bg-rose-50 border border-slate-200 rounded-md disabled:opacity-30 disabled:cursor-not-allowed cursor-pointer"
            >
              − Cột
            </button>
          </div>

          {/* Thanh công thức (Formula bar) kiểu Excel — gõ trực tiếp và Enter để áp dụng */}
          <div className="bg-white border-b border-slate-200/70 px-3 py-1.5 flex items-center gap-2 shrink-0 print:hidden z-10">
            <div className="px-2.5 py-1 border border-slate-300 rounded bg-slate-50 font-mono font-bold text-[12px] text-slate-700 min-w-[60px] text-center shrink-0">
              {selectedCell ? `${colLetter(selectedCell.c)}${selectedCell.r + 1}` : ''}
            </div>
            <span className="italic text-slate-400 text-[13px] font-serif px-0.5 shrink-0">fx</span>
            <input
              value={excelFormulaValue}
              disabled={!selectedCell}
              onChange={(e) => setExcelFormulaValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && selectedCell) {
                  handleExcelCellEdit(activeSheet, selectedCell.r, selectedCell.c, excelFormulaValue);
                  (e.target as HTMLInputElement).blur();
                }
              }}
              onBlur={() => {
                if (selectedCell && excelFormulaValue !== selectedCell.text) {
                  handleExcelCellEdit(activeSheet, selectedCell.r, selectedCell.c, excelFormulaValue);
                }
              }}
              placeholder={selectedCell ? '' : 'Chọn một ô để xem/sửa nội dung'}
              className="flex-1 min-w-0 px-2.5 py-1 border border-slate-200 rounded bg-white text-slate-800 text-[12.5px] outline-none focus:border-teal-400 focus:ring-2 focus:ring-teal-500/15 disabled:bg-slate-50"
            />
          </div>

          {/* Vùng lưới bảng tính */}
          <div data-excel-editor className="flex-1 min-w-0 overflow-auto bg-white">
            {isExcelLoading ? (
              <div className="p-4 space-y-1.5 animate-riseIn">
                {Array.from({ length: 10 }).map((_, i) => (
                  <div key={i} className="flex gap-1.5">
                    {Array.from({ length: 8 }).map((__, j) => (
                      <div
                        key={j}
                        className="h-6 rounded bg-slate-100 animate-pulse"
                        style={{ width: 92, animationDelay: `${(i * 8 + j) * 20}ms` }}
                      />
                    ))}
                  </div>
                ))}
              </div>
            ) : hasError ? (
              <div className="p-6">
                <div className="flex items-start gap-3 p-4 text-amber-700 bg-amber-50 border border-amber-200 rounded-xl text-[13px] font-medium">
                  <svg viewBox="0 0 24 24" className="w-5 h-5 shrink-0 mt-0.5" fill="none" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M12 9v4m0 4h.01M10.3 3.9 2.7 17.5A1.6 1.6 0 0 0 4.1 20h15.8a1.6 1.6 0 0 0 1.4-2.5L13.7 3.9a1.6 1.6 0 0 0-2.8 0Z" />
                  </svg>
                  <div>
                    <p className="font-bold mb-0.5">Không thể tải bảng tính</p>
                    <p>{activeSheetData?.error}</p>
                  </div>
                </div>
              </div>
            ) : !activeSheetData || activeSheetData.rows.length === 0 ? (
              <div className="p-10 text-center text-slate-400 text-[13px]">Bảng tính trống.</div>
            ) : (
              <div className="flex flex-col items-center w-full">
                    <table
                      className="border-collapse select-none"
                      style={{ fontFamily: 'Calibri, Arial, sans-serif' }}
                    >
                      <thead>
                        <tr>
                          <th onClick={selectAllExcelPage} title="Chọn toàn bộ trang tính" aria-label="Chọn toàn bộ trang tính" className="sticky top-0 left-0 z-30 bg-slate-100 border border-slate-300 w-11 h-6 text-[11px] cursor-pointer hover:bg-slate-200">◢</th>
                          {Array.from({ length: activeSheetData.endCol - activeSheetData.startCol + 1 }, (_, i) => activeSheetData.startCol + i).map((c) => (
                            <th
                              key={c}
                              style={{ width: getExcelColWidth(activeSheet, c) }}
                              className="relative sticky top-0 z-20 bg-slate-100 border border-slate-300 text-[11px] font-semibold text-slate-600 px-2 h-6"
                            >
                              {colLetter(c)}
                              <div
                                onMouseDown={(e) => startExcelColResize(e, c)}
                                onTouchStart={(e) => startExcelColResize(e, c)}
                                title="Kéo để đổi độ rộng cột"
                                className="absolute top-0 right-0 h-full w-2 -mr-1 cursor-col-resize z-30 touch-none group/handle"
                              >
                                <div className="h-full w-px mx-auto bg-transparent group-hover/handle:bg-teal-400" />
                              </div>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {activeSheetData.rows.map((rowCells) => {
                          const absRow = rowCells[0]?.r ?? 0;
                          const rowHeight = getExcelRowHeight(activeSheet, absRow);
                          return (
                            <tr key={absRow} style={{ height: rowHeight }}>
                              <td
                                style={{ height: rowHeight }}
                                className="relative sticky left-0 z-10 bg-slate-100 border border-slate-300 text-[11px] font-semibold text-slate-600 text-center w-11"
                              >
                                {absRow + 1}
                                <div
                                  onMouseDown={(e) => startExcelRowResize(e, absRow)}
                                  onTouchStart={(e) => startExcelRowResize(e, absRow)}
                                  title="Kéo để đổi chiều cao dòng"
                                  className="absolute bottom-0 left-0 w-full h-2 -mb-1 cursor-row-resize z-30 touch-none group/handle"
                                >
                                  <div className="w-full h-px my-auto bg-transparent group-hover/handle:bg-teal-400" />
                                </div>
                              </td>
                              {rowCells.map((cell) => {
                                const key = `${cell.r}-${cell.c}`;
                                if (activeSheetData.skip.has(key)) return null;
                                const merge = activeSheetData.merges[key];
                                const isSelected = isExcelCellSelected(cell.r, cell.c);
                                const isActiveCell = selectedCell?.r === cell.r && selectedCell?.c === cell.c;
                                return (
                                  <td
                                    key={key}
                                    rowSpan={merge?.rowSpan}
                                    colSpan={merge?.colSpan}
                                    contentEditable
                                    suppressContentEditableWarning
                                    spellCheck={false}
                                    onFocus={() => selectExcelCell(cell, false)}
                                    onClick={(e) => selectExcelCell(cell, e.shiftKey)}
                                    onPaste={(e) => handleExcelPaste(e, cell)}
                                    onBlur={(e) => {
                                      const newText = (e.currentTarget.textContent || '').trim();
                                      if (newText !== cell.text) {
                                        handleExcelCellEdit(activeSheet, cell.r, cell.c, newText);
                                      }
                                    }}
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter') {
                                        e.preventDefault();
                                        (e.currentTarget as HTMLTableCellElement).blur();
                                      }
                                      if (e.key === 'Escape') {
                                        e.currentTarget.textContent = cell.text;
                                        (e.currentTarget as HTMLTableCellElement).blur();
                                      }
                                    }}
                                    style={{
                                      fontWeight: cell.style?.bold ? 700 : 400,
                                      fontStyle: cell.style?.italic ? 'italic' : 'normal',
                                      textDecoration: cell.style?.underline ? 'underline' : 'none',
                                      textAlign: cell.style?.align || 'left',
                                      color: cell.style?.color || undefined,
                                      backgroundColor: isSelected ? undefined : cell.style?.bg || undefined,
                                      fontSize: cell.style?.fontSize ? `${cell.style.fontSize}px` : undefined,
                                      borderTop: cell.style?.border?.top,
                                      borderRight: cell.style?.border?.right,
                                      borderBottom: cell.style?.border?.bottom,
                                      borderLeft: cell.style?.border?.left,
                                      width: getExcelColWidth(activeSheet, cell.c),
                                      height: rowHeight,
                                    }}
                                    className={`border px-2 py-1 text-[12.5px] align-top whitespace-normal break-words overflow-hidden cursor-text outline-none focus:bg-amber-50/60 ${
                                      isSelected
                                        ? (isActiveCell ? 'border-teal-600 ring-2 ring-inset ring-teal-600 bg-teal-50/70' : 'border-teal-300 bg-teal-50/40')
                                        : 'border-slate-200 hover:bg-slate-50'
                                    }`}
                                  >
                                    {cell.text}
                                  </td>
                                );
                              })}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
              </div>
            )}
          </div>

          {/* Tab sheet kiểu Excel + thước zoom, đặt ở dưới cùng */}
          <div className="bg-[#E8ECEA] border-t border-slate-300 px-2 py-1 flex items-center justify-between gap-3 shrink-0 print:hidden z-10">
            <button type="button" title="Sheet trước" disabled={activeSheet<=0} onClick={() => { const i=Math.max(0,activeSheet-1); setActiveSheet(i); setSelectedCell(null); setExcelSelection(null); }} className="w-7 h-7 rounded border border-slate-300 bg-white text-slate-600 hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed shrink-0">◀</button>
            <div className="flex items-end gap-0.5 overflow-x-auto no-scrollbar min-w-0 flex-1">
              {excelSheets.map((sheet, idx) => (
                <button
                  key={sheet.name + idx}
                  onClick={() => {
                    setActiveSheet(idx);
                    setSelectedCell(null);
                    setExcelSelection(null);
                  }}
                  className={`px-3.5 py-1.5 text-[11.5px] font-semibold rounded-t-md border-t border-l border-r transition-all duration-150 cursor-pointer shrink-0 -mb-px ${
                    activeSheet === idx
                      ? 'bg-white text-teal-800 border-slate-300 border-b-2 border-b-teal-700 relative z-10'
                      : 'bg-[#DCE1DE] text-slate-500 border-transparent hover:bg-white/70'
                  }`}
                >
                  {sheet.name}
                </button>
              ))}
            </div>
            <div className="flex items-center gap-1 shrink-0">
              <button type="button" title="Sheet sau" disabled={activeSheet>=excelSheets.length-1} onClick={() => { const i=Math.min(excelSheets.length-1,activeSheet+1); setActiveSheet(i); setSelectedCell(null); setExcelSelection(null); }} className="w-7 h-7 rounded border border-slate-300 bg-white text-slate-600 hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed">▶</button>
              <span className="text-[11px] text-slate-400 font-medium shrink-0 pr-1">Nhấp vào ô để chỉnh sửa &middot; 100%</span>
            </div>
          </div>
        </div>
      );
    }

    if (fileType === 'word') {
      return (
        <div key={contentKey} className="flex flex-col h-full w-full min-w-0 bg-[#EEF4F3] overflow-hidden animate-riseIn">
          <div className="bg-white/85 backdrop-blur-md border-b border-slate-200/70 px-6 py-3 flex items-center justify-between shadow-sm print:hidden shrink-0 z-10 animate-slideDown">
            <div className="flex items-center gap-3">
              <span className="inline-flex items-center gap-1.5 text-[11px] font-semibold px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-full border border-emerald-200/80">
                <Icon.Pen className="w-3.5 h-3.5" />
                Cho phép điền trực tiếp
              </span>
              <span className="inline-flex items-center gap-1.5 text-[11px] font-medium text-slate-500" style={{ fontFamily: '"IBM Plex Mono", ui-monospace, monospace' }}>
                {isSaved ? (
                  <>
                    <span className="relative flex h-1.5 w-1.5">
                      <span className="absolute inline-flex h-full w-full rounded-full bg-teal-400 opacity-60" />
                      <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-teal-500" />
                    </span>
                    Đã đồng bộ Cloud
                  </>
                ) : (
                  <>
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse" />
                    Đang lưu dữ liệu...
                  </>
                )}
              </span>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleReset}
                className="group inline-flex items-center gap-1.5 px-3.5 py-1.5 text-[12px] font-semibold text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Refresh className="w-3.5 h-3.5 transition-transform duration-500 group-hover:rotate-180" />
                Đặt lại mẫu gốc
              </button>

              <div className="w-px h-6 bg-slate-200 mx-0.5" />

              <button
                onClick={handleDownloadWord}
                title="Tải file Word hiện tại"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-teal-700 bg-teal-50 hover:bg-teal-100 border border-teal-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Download className="w-3.5 h-3.5" />
                Word
              </button>
              <button
                onClick={handleDownloadPdf}
                title="Tải / Xuất file PDF"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-rose-600 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Pdf className="w-3.5 h-3.5" />
                PDF
              </button>
              <button
                onClick={handlePrint}
                title="In tài liệu hiện tại"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-[12px] font-semibold text-slate-600 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer"
              >
                <Icon.Print className="w-3.5 h-3.5" />
                In File
              </button>
            </div>
          </div>

          {/* THANH CÔNG CỤ ĐỊNH DẠNG (RIBBON) — kiểu Word */}
          {!isLoading && (
            <div className="bg-[#F8FAF9] border-b border-slate-200/70 px-3 pt-1.5 pb-2 flex flex-wrap items-start gap-x-2.5 gap-y-1.5 overflow-visible print:hidden shrink-0 animate-slideDown">
              {/* Nhóm: Hoàn tác */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  <RibbonBtn title="Hoàn tác" onClick={() => execFormat('undo')}>
                    <Icon.Undo className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Làm lại" onClick={() => execFormat('redo')}>
                    <Icon.Redo className="w-4 h-4" />
                  </RibbonBtn>
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Hoàn tác</span>
              </div>

              <div className="w-px h-11 bg-slate-200 mt-1 shrink-0" />

              {/* Nhóm: Chức năng cơ bản Word */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  <RibbonBtn title="Sao chép" onClick={copyWordSelection}>
                    <span className="text-[11px] font-bold">Copy</span>
                  </RibbonBtn>
                  <RibbonBtn title="Cắt" onClick={cutWordSelection}>
                    <span className="text-[11px] font-bold">Cut</span>
                  </RibbonBtn>
                  <RibbonBtn title="Dán" onClick={() => { void pasteWordText(); }}>
                    <span className="text-[11px] font-bold">Paste</span>
                  </RibbonBtn>
                  <RibbonBtn title="Chọn tất cả" onClick={selectAllWord}>
                    <span className="text-[10px] font-bold">All</span>
                  </RibbonBtn>
                  <RibbonBtn title="Xóa định dạng" onClick={clearWordFormatting}>
                    <span className="text-[11px] font-bold">Tx</span>
                  </RibbonBtn>
                  <RibbonBtn title="Tìm và thay thế" onClick={findAndReplaceWordText}>
                    <span className="text-[10px] font-bold">Tìm</span>
                  </RibbonBtn>
                  <RibbonBtn title="Chèn liên kết" onClick={insertWordLink}>
                    <span className="text-[12px] font-bold">🔗</span>
                  </RibbonBtn>
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Cơ bản</span>
              </div>

              <div className="w-px h-11 bg-slate-200 mt-1 shrink-0" />

              {/* Nhóm: Phông chữ */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  <select
                    title="Phông chữ"
                    aria-label="Phông chữ"
                    defaultValue="Times New Roman"
                    onMouseDown={saveSelection}
                    onChange={(e) => execFormat('fontName', e.target.value)}
                    className="h-8 px-2 text-[12.5px] bg-white border border-slate-200 rounded-lg text-slate-700 shrink-0 cursor-pointer focus:outline-none focus:ring-2 focus:ring-teal-500/20 max-w-[110px]"
                  >
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Arial">Arial</option>
                    <option value="Calibri">Calibri</option>
                    <option value="Verdana">Verdana</option>
                    <option value="Georgia">Georgia</option>
                  </select>

                  <select
                    title="Cỡ chữ"
                    aria-label="Cỡ chữ"
                    defaultValue="13"
                    onMouseDown={saveSelection}
                    onChange={(e) => applyFontSize(e.target.value)}
                    className="h-8 px-2 text-[12.5px] bg-white border border-slate-200 rounded-lg text-slate-700 shrink-0 cursor-pointer focus:outline-none focus:ring-2 focus:ring-teal-500/20 w-[60px]"
                  >
                    {['8', '9', '10', '11', '12', '13', '14', '16', '18', '20', '24', '28', '32', '36', '40'].map((s) => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Phông chữ</span>
              </div>

              <div className="w-px h-11 bg-slate-200 mt-1 shrink-0" />

              {/* Nhóm: Định dạng chữ + màu */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  <RibbonBtn title="In đậm" active={activeFormats.bold} onClick={() => execFormat('bold')}>
                    <Icon.Bold className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="In nghiêng" active={activeFormats.italic} onClick={() => execFormat('italic')}>
                    <Icon.Italic className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Gạch chân" active={activeFormats.underline} onClick={() => execFormat('underline')}>
                    <Icon.Underline className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Gạch ngang" active={activeFormats.strikeThrough} onClick={() => execFormat('strikeThrough')}>
                    <Icon.Strike className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Màu chữ" onClick={openWordTextColorPicker}>
                    <Icon.TextColor className="w-4 h-4" />
                  </RibbonBtn>
                  <input
                    id="word-text-color-picker"
                    type="color"
                    className="hidden"
                    onChange={(e) => applyColor('foreColor', e.target.value)}
                    aria-label="Chọn màu chữ"
                  />
                  <RibbonBtn title="Tô sáng" onClick={openWordHighlightPicker}>
                    <Icon.Highlight className="w-4 h-4" />
                  </RibbonBtn>
                  <input
                    id="word-highlight-picker"
                    type="color"
                    defaultValue="#fef08a"
                    className="hidden"
                    onChange={(e) => applyColor('hiliteColor', e.target.value)}
                    aria-label="Chọn màu tô sáng"
                  />
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Định dạng</span>
              </div>

              <div className="w-px h-11 bg-slate-200 mt-1 shrink-0" />

              {/* Nhóm: Căn lề + danh sách */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  <RibbonBtn title="Căn trái" active={activeFormats.justifyLeft} onClick={() => execFormat('justifyLeft')}>
                    <Icon.AlignLeft className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Căn giữa" active={activeFormats.justifyCenter} onClick={() => execFormat('justifyCenter')}>
                    <Icon.AlignCenter className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Căn phải" active={activeFormats.justifyRight} onClick={() => execFormat('justifyRight')}>
                    <Icon.AlignRight className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Căn đều" active={activeFormats.justifyFull} onClick={() => execFormat('justifyFull')}>
                    <Icon.AlignJustify className="w-4 h-4" />
                  </RibbonBtn>
                  <div className="w-px h-6 bg-slate-200 mx-0.5 shrink-0" />
                  <RibbonBtn title="Danh sách chấm" active={activeFormats.insertUnorderedList} onClick={() => execFormat('insertUnorderedList')}>
                    <Icon.ListBullet className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Danh sách số" active={activeFormats.insertOrderedList} onClick={() => execFormat('insertOrderedList')}>
                    <Icon.ListNumber className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Giảm thụt lề" onClick={() => execFormat('outdent')}>
                    <Icon.Outdent className="w-4 h-4" />
                  </RibbonBtn>
                  <RibbonBtn title="Tăng thụt lề" onClick={() => execFormat('indent')}>
                    <Icon.Indent className="w-4 h-4" />
                  </RibbonBtn>
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Đoạn văn</span>
              </div>

              <div className="w-px h-11 bg-slate-200 mt-1 shrink-0" />

              {/* Nhóm: Style Gallery — thay cho <select> kiểu văn bản, vẫn gọi đúng applyHeading() như cũ */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  {(
                    [
                      { tag: 'p', label: 'Bình thường', preview: 'text-[10px] font-medium' },
                      { tag: 'h1', label: 'Tiêu đề 1', preview: 'text-[13px] font-bold' },
                      { tag: 'h2', label: 'Tiêu đề 2', preview: 'text-[12px] font-bold' },
                      { tag: 'h3', label: 'Tiêu đề 3', preview: 'text-[11px] font-semibold' },
                      { tag: 'blockquote', label: 'Trích dẫn', preview: 'text-[10px] italic' },
                    ] as const
                  ).map((style) => {
                    const isActive = (currentBlock || 'p') === style.tag;
                    return (
                      <button
                        key={style.tag}
                        type="button"
                        title={style.label}
                        aria-label={style.label}
                        aria-pressed={isActive}
                        onMouseDown={(e) => {
                          e.preventDefault();
                          saveSelection();
                        }}
                        onClick={() => applyHeading(style.tag)}
                        className={`flex flex-col items-center justify-center w-[52px] h-9 rounded-lg border shrink-0 transition-colors duration-150 cursor-pointer ${
                          isActive
                            ? 'bg-teal-700 border-teal-700 text-white shadow-sm shadow-teal-900/20'
                            : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                        }`}
                      >
                        <span className={`leading-none ${style.preview}`}>Aa</span>
                        <span className="text-[7.5px] font-semibold mt-0.5 truncate w-full text-center px-0.5 leading-tight">
                          {style.label}
                        </span>
                      </button>
                    );
                  })}
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Kiểu văn bản</span>
              </div>

              <div className="w-px h-11 bg-slate-200 mt-1 shrink-0" />

              {/* Nhóm: Chèn bảng + xóa ô/bảng Word */}
              <div className="flex flex-col items-center gap-1 shrink-0">
                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    title="Chèn bảng"
                    aria-label="Chèn bảng"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      saveSelection();
                    }}
                    onClick={insertWordTable}
                    className="flex items-center gap-1.5 px-3 h-9 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 shrink-0 transition-colors duration-150 cursor-pointer"
                  >
                    <Icon.Table className="w-4 h-4" />
                    <span className="text-[11.5px] font-semibold">Chèn bảng</span>
                  </button>
                  <div className="relative shrink-0">
                    <button
                      type="button"
                      title="Chèn Shapes"
                      aria-label="Chèn Shapes"
                      aria-expanded={isShapesMenuOpen}
                      onMouseDown={(e) => {
                        e.preventDefault();
                        saveSelection();
                      }}
                      onClick={() => setIsShapesMenuOpen((v) => !v)}
                      className="flex items-center gap-1.5 px-3 h-9 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 shrink-0 transition-colors duration-150 cursor-pointer"
                    >
                      <Icon.Shapes className="w-4 h-4" />
                      <span className="text-[11.5px] font-semibold">Shapes</span>
                    </button>
                    {isShapesMenuOpen && (
                      <div className="absolute left-0 top-[42px] z-[80] w-[235px] rounded-xl border border-slate-200 bg-white p-2 shadow-xl">
                        <div className="px-2 pb-1 text-[10px] font-semibold uppercase tracking-wide text-slate-400">Hình khối & đường nối</div>
                        <div className="px-2 pb-2">
                          <div className="mb-1 text-[10px] font-semibold text-slate-500">Màu viền</div>
                          <div className="flex flex-wrap gap-1" aria-label="Chọn màu viền Shape">
                            {['#000000','#ffffff','#ef4444','#f97316','#f59e0b','#eab308','#84cc16','#22c55e','#10b981','#06b6d4','#0ea5e9','#2563eb','#4f46e5','#7c3aed','#a855f7','#ec4899','#f43f5e','#64748b','#475569','#78350f'].map((color) => (
                              <button
                                key={color}
                                type="button"
                                title={`Màu viền ${color}`}
                                aria-label={`Màu viền ${color}`}
                                onMouseDown={(e) => e.preventDefault()}
                                onClick={() => {
                                  setShapeBorderColor(color);
                                  const shape = selectedShapeRef.current;
                                  if (shape && editorRef.current?.contains(shape)) {
                                    shape.dataset.shapeBorderColor = color;
                                    if (shape.dataset.smartShape === 'line' || shape.dataset.smartShape === 'dashedLine' || shape.getAttribute('data-smart-shape') === 'line' || shape.getAttribute('data-smart-shape') === 'dashedLine') {
                                      const inner = shape.querySelector(':scope > span') as HTMLElement | null;
                                      if (inner) inner.style.borderTopColor = color;
                                    } else if (shape.getAttribute('data-smart-shape') === 'downArrow' || shape.getAttribute('data-smart-shape') === 'rightArrow') {
                                      shape.style.color = color;
                                    } else if (shape.querySelector('svg')) {
                                      shape.style.color = color;
                                      shape.style.borderColor = color;
                                    } else {
                                      shape.style.borderColor = color;
                                    }
                                    handleInput();
                                  }
                                }}
                                className={`h-5 w-5 rounded-full border border-slate-300 shadow-sm hover:scale-110 transition-transform ${shapeBorderColor === color ? 'ring-2 ring-slate-500 ring-offset-1' : ''}`}
                                style={{ backgroundColor: color }}
                              />
                            ))}
                            <label title="Màu tùy chỉnh" className="relative h-5 w-5 overflow-hidden rounded-full border border-slate-300 cursor-pointer shadow-sm hover:scale-110 transition-transform">
                              <input
                                type="color"
                                value={shapeBorderColor}
                                onChange={(e) => {
                                  const color = e.target.value;
                                  setShapeBorderColor(color);
                                  const shape = selectedShapeRef.current;
                                  if (shape && editorRef.current?.contains(shape)) {
                                    shape.dataset.shapeBorderColor = color;
                                    if ((shape.getAttribute('data-smart-shape') === 'line' || shape.getAttribute('data-smart-shape') === 'dashedLine')) {
                                      const inner = shape.querySelector(':scope > span') as HTMLElement | null;
                                      if (inner) inner.style.borderTopColor = color;
                                    } else if (shape.getAttribute('data-smart-shape') === 'downArrow' || shape.getAttribute('data-smart-shape') === 'rightArrow') {
                                      shape.style.color = color;
                                    } else if (shape.querySelector('svg')) {
                                      shape.style.color = color;
                                      shape.style.borderColor = color;
                                    } else {
                                      shape.style.borderColor = color;
                                    }
                                    handleInput();
                                  }
                                }}
                                className="absolute inset-0 h-full w-full cursor-pointer opacity-0"
                              />
                              <span className="pointer-events-none flex h-full w-full items-center justify-center text-[11px] text-slate-500">+</span>
                            </label>
                          </div>
                        </div>
                        <button
                          type="button"
                          title="Xóa Shape đang chọn"
                          aria-label="Xóa Shape đang chọn"
                          onMouseDown={(e) => e.preventDefault()}
                          onClick={deleteSelectedWordShape}
                          className="mb-1 flex w-full items-center gap-2 rounded-lg border border-rose-100 bg-rose-50 px-2 py-2 text-left text-[11px] font-semibold text-rose-700 hover:bg-rose-100"
                        >
                          <span className="flex h-7 w-7 items-center justify-center rounded border border-rose-200 bg-white">×</span>
                          <span>Xóa Shape đang chọn</span>
                        </button>
                        <div className="mb-1 text-[10px] font-semibold text-slate-500">Hình khối</div>
                        <div className="grid grid-cols-2 gap-1">
                          {[
                            ['textBox','Hộp văn bản','▤'],['rect','Hình chữ nhật','□'],['round','Chữ nhật bo góc','▢'],['ellipse','Hình elip','○'],
                            ['triangle','Tam giác','△'],['diamond','Hình thoi','◇'],['parallelogram','Hình bình hành','▱'],['trapezoid','Hình thang','⏢'],
                            ['pentagon','Ngũ giác','⬠'],['hexagon','Lục giác','⬡'],['octagon','Bát giác','🛑'],['rightChevron','Mũi tên góc','›'],
                            ['rightArrow','Mũi tên phải','→'],['leftArrow','Mũi tên trái','←'],['downArrow','Mũi tên xuống','↓'],['upArrow','Mũi tên lên','↑'],
                            ['star','Ngôi sao','☆'],['braceLeft','Ngoặc trái','{'],['braceRight','Ngoặc phải','}'],['doubleBrace','Ngoặc đôi','}{']
                          ].map(([kind,label,icon]) => (
                            <button key={kind} type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => insertWordShape(kind as WordShapeKind)} className="flex items-center gap-2 rounded-lg px-2 py-2 text-left text-[11px] text-slate-700 hover:bg-slate-50">
                              <span className="flex h-7 w-7 items-center justify-center rounded border border-slate-200 bg-slate-50 text-slate-600 text-[15px]">{icon}</span><span>{label}</span>
                            </button>
                          ))}
                        </div>
                        <div className="mt-2 mb-1 text-[10px] font-semibold text-slate-500">Đường nối</div>
                        <div className="grid grid-cols-2 gap-1">
                          {[
                            ['line','Đường thẳng','—'],['arrowLine','Đường có mũi tên','→'],['elbow','Đường gấp khúc','⌜'],['elbowArrow','Gấp khúc có mũi tên','↗'],
                            ['curve','Đường cong','⌒'],['curveArrow','Cong có mũi tên','↗'],['uLine','Đường chữ U','∪'],['arc','Cung','⌒'],
                            ['dashedLine','Đường nét đứt','┅']
                          ].map(([kind,label,icon]) => (
                            <button key={kind} type="button" onMouseDown={(e) => e.preventDefault()} onClick={() => insertWordShape(kind as WordShapeKind)} className="flex items-center gap-2 rounded-lg px-2 py-2 text-left text-[11px] text-slate-700 hover:bg-slate-50">
                              <span className="flex h-7 w-7 items-center justify-center rounded border border-slate-200 bg-slate-50 text-slate-600 text-[15px]">{icon}</span><span>{label}</span>
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                  <div className="relative">
                    <button type="button" title="Viền bảng/vùng chọn" onMouseDown={(e)=>{e.preventDefault();saveSelection();}} onClick={()=>setIsWordBorderMenuOpen(v=>!v)} className="flex items-center gap-1.5 px-3 h-9 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 shrink-0 transition-colors duration-150 cursor-pointer">▦ Viền</button>
                    {isWordBorderMenuOpen && <div className="absolute left-0 top-[42px] z-[90] w-[190px] rounded-xl border border-slate-200 bg-white p-1.5 shadow-xl">{wordBorderOptions.map(([k,l])=><button key={k} type="button" onMouseDown={(e)=>e.preventDefault()} onClick={()=>{applyWordBorder(k);setIsWordBorderMenuOpen(false)}} className="w-full text-left px-2.5 py-1.5 text-[11px] rounded-lg hover:bg-slate-50">{l}</button>)}</div>}
                  </div>
                  <div className="relative">
                    <button type="button" title="Gộp ô bảng Word" onMouseDown={(e)=>{e.preventDefault();saveSelection();}} onClick={()=>setIsWordMergeMenuOpen(v=>!v)} className="flex items-center gap-1.5 px-3 h-9 rounded-lg border border-slate-200 bg-white text-slate-600 hover:bg-slate-50 shrink-0 transition-colors duration-150 cursor-pointer">Gộp ▾</button>
                    {isWordMergeMenuOpen && <div className="absolute left-0 top-[42px] z-[90] w-[175px] rounded-xl border border-slate-200 bg-white p-1.5 shadow-xl">{wordMergeOptions.map(([k,l])=><button key={k} type="button" onMouseDown={(e)=>e.preventDefault()} onClick={()=>{mergeWordCells(k);setIsWordMergeMenuOpen(false)}} className="w-full text-left px-2.5 py-1.5 text-[11px] rounded-lg hover:bg-slate-50">{l}</button>)}</div>}
                  </div>
                  <button
                    type="button"
                    title="Xóa ô bảng Word"
                    aria-label="Xóa ô bảng Word"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={deleteWordTableCell}
                    className="flex items-center gap-1.5 px-3 h-9 rounded-lg border border-slate-200 bg-white text-rose-600 hover:bg-rose-50 shrink-0 transition-colors duration-150 cursor-pointer"
                  >
                    <Icon.DeleteCell className="w-4 h-4" />
                    <span className="text-[11.5px] font-semibold">Xóa ô</span>
                  </button>
                  <button
                    type="button"
                    title="Xóa bảng Word"
                    aria-label="Xóa bảng Word"
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={deleteWordTable}
                    className="flex items-center gap-1.5 px-3 h-9 rounded-lg border border-slate-200 bg-white text-rose-700 hover:bg-rose-50 shrink-0 transition-colors duration-150 cursor-pointer"
                  >
                    <Icon.DeleteTable className="w-4 h-4" />
                    <span className="text-[11.5px] font-semibold">Xóa bảng</span>
                  </button>
                </div>
                <span className="text-[8.5px] font-semibold uppercase tracking-wide text-slate-400">Chèn</span>
              </div>
            </div>
          )}

          {/* Thanh trạng thái tài liệu Word: số từ + zoom thật (không phải trang trí) */}
          {!isLoading && (
            <div className="bg-white border-b border-slate-200/70 px-4 py-1 flex items-center justify-between text-[11px] text-slate-400 font-medium shrink-0 print:hidden">
              <span>{wordCount.toLocaleString('vi-VN')} từ</span>
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setWordZoom((z) => Math.max(WORD_ZOOM_MIN, z - WORD_ZOOM_STEP))}
                  disabled={wordZoom <= WORD_ZOOM_MIN}
                  aria-label="Thu nhỏ"
                  title="Thu nhỏ"
                  className="w-6 h-6 flex items-center justify-center rounded-md text-slate-400 hover:text-teal-700 hover:bg-teal-50 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
                >
                  <Icon.ZoomOut className="w-3.5 h-3.5" />
                </button>
                <input
                  type="range"
                  min={WORD_ZOOM_MIN}
                  max={WORD_ZOOM_MAX}
                  step={WORD_ZOOM_STEP}
                  value={wordZoom}
                  onChange={(e) => setWordZoom(Number(e.target.value))}
                  aria-label="Mức thu phóng tài liệu"
                  className="w-24 accent-teal-600 cursor-pointer"
                />
                <button
                  onClick={() => setWordZoom((z) => Math.min(WORD_ZOOM_MAX, z + WORD_ZOOM_STEP))}
                  disabled={wordZoom >= WORD_ZOOM_MAX}
                  aria-label="Phóng to"
                  title="Phóng to"
                  className="w-6 h-6 flex items-center justify-center rounded-md text-slate-400 hover:text-teal-700 hover:bg-teal-50 disabled:opacity-30 disabled:cursor-not-allowed transition-colors cursor-pointer"
                >
                  <Icon.ZoomIn className="w-3.5 h-3.5" />
                </button>
                <button
                  onClick={() => setWordZoom(100)}
                  title="Đặt lại mức thu phóng 100%"
                  className="w-12 text-center font-mono text-[10.5px] text-slate-500 hover:text-teal-700 cursor-pointer"
                >
                  {wordZoom}%
                </button>
              </div>
            </div>
          )}

          <div className="flex-1 min-w-0 overflow-auto p-10 flex justify-center">
            {isLoading ? (
              <div className="bg-white shadow-[0_1px_1px_rgba(15,50,55,0.05),0_20px_40px_-16px_rgba(15,50,55,0.18)] border border-slate-200 p-16 w-[210mm] max-w-full shrink-0 rounded-sm self-start animate-riseIn">
                <div className="space-y-3">
                  <div className="h-5 w-2/3 rounded bg-slate-100 animate-pulse" />
                  <div className="h-3.5 w-full rounded bg-slate-100 animate-pulse" />
                  <div className="h-3.5 w-full rounded bg-slate-100 animate-pulse" />
                  <div className="h-3.5 w-5/6 rounded bg-slate-100 animate-pulse" />
                  <div className="h-3.5 w-full rounded bg-slate-100 animate-pulse mt-6" />
                  <div className="h-3.5 w-11/12 rounded bg-slate-100 animate-pulse" />
                  <div className="h-3.5 w-4/5 rounded bg-slate-100 animate-pulse" />
                </div>
                <div className="flex items-center gap-2.5 text-teal-700 font-semibold text-[13px] mt-8">
                  <Icon.Spinner className="w-4 h-4 animate-spin" />
                  Đang đồng bộ dữ liệu từ Cloud...
                </div>
              </div>
            ) : (
              <div
                className="relative flex flex-col w-[210mm] max-w-full shrink-0 self-start mb-12"
                style={{ zoom: `${wordZoom}%` } as React.CSSProperties}
              >
                <div className="flex flex-col w-[210mm] max-w-full shrink-0">
                    <div
                  ref={editorRef}
                  onMouseUp={refreshActiveFormats}
                  onKeyUp={refreshActiveFormats}
                  onMouseDown={handleEditorMouseDown}
                  onMouseMove={handleEditorMouseMoveHover}
                  contentEditable
                  suppressContentEditableWarning
                  spellCheck={false}
                  onInput={handleInput}
                  dangerouslySetInnerHTML={{ __html: htmlContent }}
                  className="bg-white shadow-[0_1px_1px_rgba(15,50,55,0.05),0_20px_40px_-16px_rgba(15,50,55,0.18)] border border-slate-200 p-16 min-h-[297mm] h-auto outline-none text-black prose prose-slate focus:ring-4 focus:ring-teal-500/20 focus:border-teal-300 rounded-sm transition-shadow duration-300 animate-popIn [&_table]:w-full [&_table]:table-fixed [&_table]:border-collapse [&_table]:my-3 [&_td]:border [&_td]:border-black [&_td]:p-1.5 [&_td]:overflow-hidden [&_td]:text-xs [&_th]:border [&_th]:border-black [&_th]:p-1.5 print:shadow-none print:border-none print:w-full print:p-0 print:m-0"
                  style={{
                    boxSizing: 'border-box',
                    position: 'relative',
                    wordBreak: 'break-word',
                    fontFamily: '"Times New Roman", Times, serif',
                    fontSize: '13pt',
                    lineHeight: '1.4',
                  }}
                    />
                </div>

              </div>
            )}
          </div>
        </div>
      );
    }

    return null;
  };

  return (
    <div className="relative h-screen w-screen overflow-hidden bg-[#F4F9F8] font-sans">
      <style jsx global>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500&display=swap');

        .font-display { font-family: 'Fraunces', 'Times New Roman', serif; }
        .font-ui { font-family: 'Inter', ui-sans-serif, system-ui, sans-serif; }

        .no-scrollbar::-webkit-scrollbar { display: none; }
        .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }

        /* Khoảng trống giữa các trang chỉ phục vụ giao diện soạn thảo, không phải nội dung Word. */
        [data-page-gap] { break-inside: avoid; }
        @media print {
          [data-page-gap], [data-page-badge] { display: none !important; }
        }

        @keyframes riseIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-riseIn { animation: riseIn 0.45s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes popIn {
          from { opacity: 0; transform: translateY(8px) scale(0.985); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-popIn { animation: popIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes slideDown {
          from { opacity: 0; transform: translateY(-10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-slideDown { animation: slideDown 0.4s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes veilFade {
          from { opacity: 0; }
          to { opacity: 1; }
        }
        .animate-veilFade { animation: veilFade 0.4s ease both; }

        @keyframes cardIn {
          from { opacity: 0; transform: translateY(18px) scale(0.97); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        .animate-cardIn { animation: cardIn 0.55s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes fieldIn {
          from { opacity: 0; transform: translateY(10px); }
          to { opacity: 1; transform: translateY(0); }
        }
        .animate-fieldIn { animation: fieldIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both; }

        @keyframes ringSpin {
          to { transform: rotate(360deg); }
        }
        .animate-ringSpin { animation: ringSpin 10s linear infinite; }

        @keyframes ringSpinReverse {
          to { transform: rotate(-360deg); }
        }
        .animate-ringSpinReverse { animation: ringSpinReverse 14s linear infinite; }

        @keyframes shakeX {
          10%, 90% { transform: translateX(-1px); }
          20%, 80% { transform: translateX(2px); }
          30%, 50%, 70% { transform: translateX(-4px); }
          40%, 60% { transform: translateX(4px); }
        }
        .animate-shakeX { animation: shakeX 0.5s cubic-bezier(0.36,0.07,0.19,0.97) both; }

        @keyframes sidebarWidth {
          from { width: 280px; }
        }

        @media (prefers-reduced-motion: reduce) {
          .animate-riseIn, .animate-popIn, .animate-slideDown, .animate-veilFade,
          .animate-cardIn, .animate-fieldIn, .animate-ringSpin, .animate-ringSpinReverse, .animate-shakeX {
            animation: none !important;
          }
        }
      `}</style>

      {/* ============== MÀN HÌNH ĐĂNG NHẬP ============== */}
      {!isAuthenticated && (
        <div
          className={`absolute inset-0 z-50 flex transition-opacity duration-500 ${isLoggingOut ? 'opacity-0' : 'opacity-100 animate-veilFade'}`}
        >
          {/* Cột trái — thương hiệu / bối cảnh (ẩn trên màn nhỏ) */}
          <div
            className="relative hidden lg:flex lg:w-[46%] xl:w-[42%] flex-col justify-between p-12 overflow-hidden"
            style={{
              background: 'radial-gradient(120% 120% at 50% 0%, #0E3A41 0%, #0A2A30 55%, #081F24 100%)',
            }}
          >
            <div
              className="pointer-events-none absolute inset-0 opacity-[0.08]"
              style={{
                backgroundImage:
                  'radial-gradient(circle at 15% 15%, #fff 0, transparent 42%), radial-gradient(circle at 85% 75%, #14B8AA 0, transparent 45%)',
              }}
            />
            <div
              className="pointer-events-none absolute inset-0 opacity-[0.05]"
              style={{
                backgroundImage:
                  'linear-gradient(#fff 1px, transparent 1px), linear-gradient(90deg, #fff 1px, transparent 1px)',
                backgroundSize: '42px 42px',
              }}
            />

            <div className="relative flex items-center gap-3 font-ui">
              <div className="w-10 h-10 rounded-xl bg-white/10 border border-white/15 flex items-center justify-center backdrop-blur-sm">
                <Icon.Shield className="w-5 h-5 text-teal-200" />
              </div>
              <div className="leading-tight">
                <p className="text-white font-bold text-[14px] tracking-wide">HỆ THỐNG HSCL</p>
                <p className="text-teal-200/70 text-[11px] font-medium">Khoa Vi sinh &ndash; Miễn dịch</p>
              </div>
            </div>

            <div className="relative font-ui">
              <h1 className="font-display text-white text-[32px] xl:text-[36px] font-semibold leading-[1.15] mb-4 max-w-md">
                Quản lý hồ sơ chất lượng, gọn gàng và an toàn.
              </h1>
              <p className="text-teal-100/70 text-[13.5px] leading-relaxed max-w-sm">
                Truy cập SOP, quy trình, biểu mẫu và tài liệu hướng dẫn của khoa &mdash; chỉnh sửa,
                lưu tự động và trích xuất PDF ngay trên trình duyệt.
              </p>
              <div className="flex items-center gap-5 mt-8 text-teal-100/60 text-[11.5px] font-semibold">
                <span className="inline-flex items-center gap-1.5"><Icon.Doc className="w-3.5 h-3.5" /> Word</span>
                <span className="inline-flex items-center gap-1.5"><Icon.Table className="w-3.5 h-3.5" /> Excel</span>
                <span className="inline-flex items-center gap-1.5"><Icon.File className="w-3.5 h-3.5" /> PDF</span>
              </div>
            </div>

            <p className="relative text-teal-200/40 text-[11px] font-medium font-ui">
              © 2026 Bệnh viện Phong &ndash; Da liễu TW Quy Hòa
            </p>
          </div>

          {/* Cột phải — thẻ đăng nhập */}
          <div className="flex-1 flex items-center justify-center p-4 sm:p-8 bg-[#F4F9F8]">
            <div className="relative w-full max-w-md">
              <div className="bg-white rounded-[28px] shadow-[0_30px_80px_-20px_rgba(15,58,65,0.25)] border border-slate-200/70 p-8 animate-cardIn">
                <div className="text-center mb-7">
                  <div className="relative inline-flex items-center justify-center w-16 h-16 mb-4 lg:hidden">
                    <span className="absolute inset-0 rounded-full border border-teal-200 animate-ringSpin" style={{ borderStyle: 'dashed' }} />
                    <span className="absolute inset-[3px] rounded-full border border-teal-300/60 animate-ringSpinReverse" style={{ borderStyle: 'dotted' }} />
                    <span className="absolute inset-1.5 rounded-full bg-teal-50" />
                    <Icon.Lock className="relative w-6 h-6 text-teal-700" />
                  </div>
                  <span className="hidden lg:inline-flex items-center justify-center w-11 h-11 rounded-2xl bg-teal-50 border border-teal-200/70 mb-4">
                    <Icon.Lock className="w-5 h-5 text-teal-700" />
                  </span>
                  <h2 className="font-display text-[19px] font-semibold text-[#0E3A41] tracking-tight leading-snug">
                    HỒ SƠ QUẢN LÝ CHẤT LƯỢNG QĐ-2429/BYT
                  </h2>
                  <h2 className="font-display text-[19px] font-semibold text-[#0E3A41] tracking-tight leading-snug">
                    KHOA VI SINH - MIỄN DỊCH
                  </h2>
                  <p className="text-[11.5px] text-slate-400 font-medium mt-1.5 font-ui">
                    Đăng nhập bằng Pass Key được cấp để tiếp tục
                  </p>
                </div>

                <form onSubmit={handleLogin} className="space-y-4 font-ui">
                  <div className="animate-fieldIn" style={{ animationDelay: '60ms' }}>
                    <label htmlFor="login-username" className="block text-[11px] font-bold text-slate-500 mb-1.5 uppercase tracking-wider">
                      Tên đăng nhập
                    </label>
                    <div className="relative">
                      <Icon.User className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input
                        id="login-username"
                        type="text"
                        value={username}
                        onChange={(e) => setUsername(e.target.value)}
                        placeholder="Nhập tên đăng nhập"
                        autoComplete="username"
                        className="w-full pl-10 pr-4 py-2.75 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:ring-4 focus:ring-teal-500/15 focus:border-teal-400 font-medium text-[13px] text-slate-800 transition-all duration-200 placeholder:text-slate-400"
                        required
                      />
                    </div>
                  </div>

                  <div className="animate-fieldIn" style={{ animationDelay: '140ms' }}>
                    <label htmlFor="login-password" className="block text-[11px] font-bold text-slate-500 mb-1.5 uppercase tracking-wider">
                      Mật khẩu (Pass Key)
                    </label>
                    <div className="relative">
                      <Icon.Key className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input
                        id="login-password"
                        type={showPassword ? 'text' : 'password'}
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        placeholder="Nhập mật khẩu"
                        autoComplete="current-password"
                        className="w-full pl-10 pr-11 py-2.75 bg-slate-50 border border-slate-200 rounded-2xl focus:outline-none focus:ring-4 focus:ring-teal-500/15 focus:border-teal-400 font-medium text-[13px] text-slate-800 transition-all duration-200 placeholder:text-slate-400"
                        required
                      />
                      <button
                        type="button"
                        onClick={() => setShowPassword((v) => !v)}
                        aria-label={showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
                        className="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer transition-colors"
                      >
                        {showPassword ? <Icon.EyeOff className="w-4 h-4" /> : <Icon.Eye className="w-4 h-4" />}
                      </button>
                    </div>
                  </div>

                  {loginError && (
                    <div className="p-3 bg-rose-50 border border-rose-200 text-rose-600 rounded-2xl text-[11.5px] font-semibold text-center animate-shakeX">
                      {loginError}
                    </div>
                  )}

                  <button
                    type="submit"
                    disabled={isLoggingIn}
                    className="w-full py-3 bg-[#0E3A41] hover:bg-[#0A2C31] disabled:opacity-70 disabled:cursor-not-allowed text-white rounded-2xl font-bold shadow-lg shadow-teal-950/20 transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-[0.98] cursor-pointer text-[12.5px] uppercase tracking-wider mt-1 animate-fieldIn inline-flex items-center justify-center gap-2"
                    style={{ animationDelay: '220ms' }}
                  >
                    {isLoggingIn ? (
                      <>
                        <Icon.Spinner className="w-4 h-4 animate-spin" />
                        Đang xác thực...
                      </>
                    ) : (
                      'Đăng nhập'
                    )}
                  </button>
                </form>

                <div className="mt-6 pt-4 border-t border-slate-100 text-center text-[10.5px] text-slate-400 font-medium font-ui animate-fieldIn" style={{ animationDelay: '280ms' }}>
                  <span className="font-bold text-teal-700">© 2026 Khoa Vi sinh - Miễn dịch, Bệnh viện Phong - Da liễu TW Quy Hòa</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ============== GIAO DIỆN CHÍNH ============== */}
      <div
        className={`h-full w-full flex flex-col overflow-hidden transition-all duration-700 ease-out ${
          isAuthenticated ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4 pointer-events-none'
        }`}
      >
        {/* ---- HEADER ---- */}
        <header className="h-14 shrink-0 bg-white border-b border-slate-200/80 shadow-[0_1px_2px_rgba(15,50,55,0.04)] flex items-center px-4 gap-4 z-30 font-ui">
          <button
            onClick={() => setIsMobileSidebarOpen(true)}
            aria-label="Mở danh mục tài liệu"
            title="Danh mục tài liệu"
            className="lg:hidden w-8 h-8 flex items-center justify-center rounded-lg text-slate-500 hover:text-teal-700 hover:bg-teal-50 transition-colors cursor-pointer shrink-0 -ml-1"
          >
            <Icon.Menu className="w-4.5 h-4.5" />
          </button>

          <div className="flex items-center gap-2.5 shrink-0">
            <div className="w-8 h-8 rounded-lg bg-[#0E3A41] flex items-center justify-center shrink-0">
              <Icon.Shield className="w-4 h-4 text-teal-200" />
            </div>
            <div className="leading-tight hidden sm:block">
              <p className="text-[13px] font-bold text-slate-800 tracking-tight">Hệ thống HSCL</p>
              <p className="text-[10.5px] text-slate-400 font-medium -mt-0.5">Khoa Vi sinh &ndash; Miễn dịch</p>
            </div>
          </div>

          <div className="w-px h-6 bg-slate-200 hidden md:block" />

          {/* Tên bệnh viện — hiển thị nổi bật ở giữa header (thay cho vị trí ô tìm kiếm trước đây, ô tìm kiếm đã dời xuống thanh bên dưới) */}
          <div className="flex-1 min-w-0 flex flex-col items-center justify-center text-center leading-tight px-2">
            <p className="text-[13.5px] font-bold text-[#0E3A41] tracking-tight truncate max-w-full">
              Bệnh viện Phong - Da liễu TW Quy Hòa
            </p>
            <p className="text-[11px] text-slate-400 font-medium truncate max-w-full">
              Khoa Vi sinh - Miễn dịch
            </p>
          </div>

          {selectedFile && (
            <div className="hidden sm:flex items-center gap-1.5 text-[12px] font-medium text-slate-400 max-w-[240px] min-w-0">
              {currentSyncOk ? (
                <>
                  <span className="relative flex h-1.5 w-1.5 shrink-0">
                    <span className="absolute inline-flex h-full w-full rounded-full bg-teal-400 opacity-60" />
                    <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-teal-500" />
                  </span>
                  <span className="truncate">Đã lưu &middot; {selectedFile.title}</span>
                </>
              ) : (
                <>
                  <Icon.Spinner className="w-3 h-3 animate-spin text-amber-500 shrink-0" />
                  <span className="truncate">Đang lưu &middot; {selectedFile.title}</span>
                </>
              )}
            </div>
          )}

          <div className="w-px h-6 bg-slate-200" />

          <div className="flex items-center gap-2 shrink-0">
            <div className="w-8 h-8 rounded-full bg-teal-50 border border-teal-200/70 flex items-center justify-center">
              <Icon.User className="w-4 h-4 text-teal-700" />
            </div>
            <span className="text-[12.5px] font-semibold text-slate-700 hidden sm:inline">Quản trị viên</span>
          </div>

          <button
            onClick={handleLogout}
            aria-label="Đăng xuất"
            className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200/80 rounded-xl font-bold transition-all duration-200 hover:-translate-y-px active:translate-y-0 active:scale-95 cursor-pointer text-[12px] shrink-0"
          >
            <Icon.Logout className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Đăng xuất</span>
          </button>
        </header>

        {/* ---- THANH TÌM KIẾM (đã dời xuống từ giữa Header) ---- */}
        <div className="h-12 shrink-0 bg-white border-b border-slate-200/80 flex items-center px-4 z-20 font-ui">
          <div className="flex-1 min-w-0 max-w-2xl mx-auto relative">
            <Icon.Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
            <input
              ref={headerSearchRef}
              type="text"
              value={docSearchQuery}
              onChange={(e) => setDocSearchQuery(e.target.value)}
              placeholder="Tìm kiếm tài liệu, biểu mẫu, hồ sơ..."
              className="w-full pl-10 pr-16 py-1.75 text-[13px] bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-4 focus:ring-teal-500/12 focus:border-teal-400 focus:bg-white transition-all duration-200 placeholder:text-slate-400"
            />
            {docSearchQuery ? (
              <button
                onClick={() => setDocSearchQuery('')}
                aria-label="Xoá tìm kiếm"
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
              >
                <Icon.Close className="w-3.5 h-3.5" />
              </button>
            ) : (
              <kbd className="absolute right-2.5 top-1/2 -translate-y-1/2 hidden md:inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-md border border-slate-200 bg-white text-[10px] font-semibold text-slate-400 font-mono">
                Ctrl K
              </kbd>
            )}
          </div>
        </div>

        {/* ---- THÂN CHÍNH: SIDEBAR + WORKSPACE ---- */}
        <div className="flex-1 min-h-0 flex overflow-hidden relative">
          {/* Lớp phủ nền cho drawer di động — chỉ hiển thị trên màn hình nhỏ khi sidebar mở */}
          {isMobileSidebarOpen && (
            <div
              onClick={() => setIsMobileSidebarOpen(false)}
              aria-hidden="true"
              className="lg:hidden fixed inset-0 top-[104px] bg-slate-900/40 backdrop-blur-[1px] z-30 animate-veilFade"
            />
          )}

          {/* ---- SIDEBAR ----
              Di động: drawer trượt từ trái, phủ (fixed), đóng bằng nút X hoặc chạm nền.
              Desktop (lg+): nằm trong luồng bố cục, có thể thu gọn còn dải icon,
              hoặc kéo giãn/thu nhỏ tự do bằng thanh chia ở viền phải. */}
          <aside
            ref={sidebarRef}
            style={{ '--sidebar-width': `${sidebarWidth}px` } as React.CSSProperties}
            className={`relative bg-white border-r border-slate-200/80 flex flex-col overflow-hidden font-ui z-40
              fixed left-0 top-[104px] bottom-0 w-[var(--sidebar-width)] transition-transform duration-200 ease-out
              ${isMobileSidebarOpen ? 'translate-x-0' : '-translate-x-full'}
              lg:static lg:top-auto lg:bottom-auto lg:translate-x-0 lg:h-full lg:shrink-0
              ${isResizingSidebar ? '' : 'transition-[width]'}
              ${isSidebarCollapsed ? 'lg:w-[60px]' : ''}`}
          >
            <div className={`flex items-center justify-between px-3.5 py-3 border-b border-slate-100 shrink-0 ${isSidebarCollapsed ? 'lg:px-0 lg:justify-center' : ''}`}>
              <span className={`text-[11px] font-bold uppercase tracking-wider text-slate-400 ${isSidebarCollapsed ? 'lg:hidden' : ''}`}>
                Kho tài liệu
              </span>

              {/* Nút đóng — chỉ hiện trên di động */}
              <button
                onClick={() => setIsMobileSidebarOpen(false)}
                title="Đóng danh mục"
                aria-label="Đóng danh mục"
                className="lg:hidden w-7 h-7 flex items-center justify-center rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors cursor-pointer shrink-0"
              >
                <Icon.Close className="w-4 h-4" />
              </button>

              {/* Nút thu gọn/mở rộng — chỉ hiện trên desktop */}
              <button
                onClick={() => setIsSidebarCollapsed((v) => !v)}
                title={isSidebarCollapsed ? 'Mở rộng danh mục' : 'Thu gọn danh mục'}
                aria-label={isSidebarCollapsed ? 'Mở rộng danh mục' : 'Thu gọn danh mục'}
                className="hidden lg:flex w-7 h-7 items-center justify-center rounded-lg text-slate-400 hover:text-teal-700 hover:bg-teal-50 transition-colors cursor-pointer shrink-0"
              >
                {isSidebarCollapsed ? <Icon.ChevronsRight className="w-4 h-4" /> : <Icon.ChevronsLeft className="w-4 h-4" />}
              </button>
            </div>

            {/* Vùng cây thư mục — giữ nguyên FolderTree/DocumentNode, chỉ bọc khung UI mới.
                Không unmount khi thu gọn để không mất trạng thái nội bộ của FolderTree. */}
            <div className={`flex-1 min-h-0 overflow-y-auto overflow-x-hidden no-scrollbar min-w-[220px] lg:min-w-0 ${isSidebarCollapsed ? 'lg:opacity-0 lg:pointer-events-none lg:w-0' : 'opacity-100'}`}>
              <FolderTree
                onSelectFile={(file) => {
                  setSelectedFile(file);
                  setIsMobileSidebarOpen(false);
                }}
                selectedFile={selectedFile}
                searchQuery={docSearchQuery}
                onSearchQueryChange={setDocSearchQuery}
              />
            </div>

            {/* Thanh kéo giãn/thu nhỏ — hoạt động cả bằng chuột (desktop) lẫn chạm (di động/tablet),
                ẩn khi sidebar đang thu gọn. Nhấp đúp (hoặc chạm giữ) để đặt lại độ rộng mặc định. */}
            {!isSidebarCollapsed && (
              <div
                onMouseDown={handleSidebarResizeStart}
                onTouchStart={handleSidebarResizeStart}
                onDoubleClick={() => setSidebarWidth(SIDEBAR_DEFAULT_WIDTH)}
                title="Kéo để thay đổi độ rộng danh mục"
                role="separator"
                aria-orientation="vertical"
                className="flex absolute top-0 right-0 h-full w-4 -mr-2 lg:w-2.5 lg:-mr-1 cursor-col-resize items-stretch justify-center group z-10 touch-none"
              >
                <div
                  className={`h-full w-px transition-colors duration-150 ${
                    isResizingSidebar ? 'bg-teal-500' : 'bg-transparent group-hover:bg-teal-400'
                  }`}
                />
              </div>
            )}
          </aside>

          {/* ---- WORKSPACE ---- */}
          <main className="flex-1 min-w-0 h-full flex flex-col overflow-hidden print:p-0">
            <div className="flex-1 min-h-0 bg-white m-3 rounded-2xl shadow-[0_1px_2px_rgba(15,50,55,0.04),0_16px_36px_-16px_rgba(15,50,55,0.12)] border border-slate-200/70 overflow-hidden flex flex-col print:border-none print:m-0 print:rounded-none print:shadow-none">
              {renderContent()}
            </div>

            {/* ---- STATUS BAR ---- */}
            <div className="h-8 shrink-0 border-t border-slate-200/80 bg-white/70 backdrop-blur-sm px-4 flex items-center gap-4 text-[11px] font-medium text-slate-400 font-ui print:hidden">
              <span className="truncate max-w-[240px]">
                {selectedFile ? selectedFile.title : 'Chưa có tài liệu nào được mở'}
              </span>
              {currentFileType && (
                <span
                  className={`inline-flex items-center px-1.5 py-0.5 rounded border text-[10px] font-bold uppercase tracking-wide shrink-0 ${
                    (fileTypeMeta[currentFileType] || fileTypeMeta.text).className
                  }`}
                >
                  {(fileTypeMeta[currentFileType] || fileTypeMeta.text).label}
                </span>
              )}
              <div className="flex-1" />
              {selectedFile && (
                <span className="inline-flex items-center gap-1.5 shrink-0">
                  {currentSyncOk ? (
                    <>
                      <span className="w-1.5 h-1.5 rounded-full bg-teal-500 shrink-0" />
                      Đã lưu
                    </>
                  ) : (
                    <>
                      <Icon.Spinner className="w-3 h-3 animate-spin text-amber-500 shrink-0" />
                      Đang lưu...
                    </>
                  )}
                </span>
              )}
              <span className="inline-flex items-center gap-1.5 shrink-0">
                <span className="w-1.5 h-1.5 rounded-full bg-teal-500 shrink-0" />
                Đã kết nối
              </span>
              <span className="hidden md:inline shrink-0">Quản trị viên</span>
              <div className="w-px h-3.5 bg-slate-200 shrink-0 hidden md:block" />
              <span className="hidden md:inline shrink-0 text-slate-400 truncate">
                © 2026 Khoa Vi sinh - Miễn dịch, Bệnh viện Phong - Da liễu TW Quy Hòa
              </span>
            </div>
          </main>
        </div>
      </div>
    </div>
  );
}
