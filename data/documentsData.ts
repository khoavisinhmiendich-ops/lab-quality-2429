export interface DocumentItem {
  id: string;
  title: string;
  fileName?: string;
  type?: 'docx' | 'pdf' | 'folder' | string;
  path?: string;
  children?: DocumentItem[];
}

export type DocumentNode = DocumentItem;

export const documentsData: DocumentItem[] = [
  {
    id: 'c1',
    title: 'Chương I: Tổ chức và Quản lý',
    children: [
      {
        id: 'c1-1',
        title: 'Sơ đồ tổ chức',
        fileName: 'XN-BM 5.1.1.01 So do to chuc.docx',
        type: 'docx',
        path: '2429.2026/1. Chuong I To chuc quan ly/XN-BM 5.1.1.01 So do to chuc.docx',
      },
      {
        id: 'c1-2',
        title: 'Thẩm quyền ký',
        fileName: 'XN-BM 5.1.1.02 Tham quyen ky.docx',
        type: 'docx',
        path: '2429.2026/1. Chuong I To chuc quan ly/XN-BM 5.1.1.02 Tham quyen ky.docx',
      },
    ],
  },
  {
    id: 'c2',
    title: 'Chương II: Tài liệu và Hồ sơ',
    children: [
      {
        id: 'c2-1',
        title: 'Quy trình kiểm soát hồ sơ',
        fileName: 'Quy trinh kiem soat ho so.docx',
        type: 'docx',
        path: '2429.2026/2. Chuong II Tai lieu ho so/Quy trinh kiem soat ho so.docx',
      },
      {
        id: 'c2-2',
        title: 'Biểu mẫu theo dõi hồ sơ lưu trữ',
        fileName: 'Bieu mau theo doi ho so luu tru.docx',
        type: 'docx',
        path: '2429.2026/2. Chuong II Tai lieu ho so/Bieu mau theo doi ho so luu tru.docx',
      },
    ],
  },
  {
    id: 'c3',
    title: 'Chương III: Quản lý nhân sự',
    children: [],
  },
  {
    id: 'c4',
    title: 'Chương IV: Dịch vụ và Quản lý khách hàng',
    children: [],
  },
  {
    id: 'c5',
    title: 'Chương V: Quản lý trang thiết bị',
    children: [],
  },
  {
    id: 'c6',
    title: 'Chương VI: Đánh giá nội bộ',
    children: [],
  },
  {
    id: 'c7',
    title: 'Chương VII: Quản lý mua sắm vật tư, TTB',
    children: [],
  },
  {
    id: 'c8',
    title: 'Chương VIII: Quản lý quá trình thực hiện xét nghiệm',
    children: [],
  },
  {
    id: 'c9',
    title: 'Chương IX: Báo cáo sai sót, sự cố và Quản lý thông tin',
    children: [],
  },
  {
    id: 'c10',
    title: 'Chương X: Sự KPH, HĐKP, HĐPN',
    children: [],
  },
  {
    id: 'c11',
    title: 'Chương XI: Cải tiến liên tục',
    children: [],
  },
  {
    id: 'c12',
    title: 'Chương XII: Trang thiết bị công nghệ thông tin',
    children: [
      {
        id: 'c12-1',
        title: 'Quy định an toàn sinh học (PDF)',
        fileName: 'Quy dinh an toan sinh hoc.pdf',
        type: 'pdf',
        path: '2429.2026/12. Chuong XII CSVC va An toan/Quy dinh an toan sinh hoc.pdf',
      },
      {
        id: 'c12-2',
        title: 'Quy trình quản lý hệ thống LIS',
        fileName: 'Quy trinh quan ly he thong LIS.docx',
        type: 'docx',
        path: '2429.2026/12. Chuong XII CSVC va An toan/Quy trinh quan ly he thong LIS.docx',
      },
    ],
  },
];