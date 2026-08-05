export interface DocDetail {
  id: string;
  code: string;
  title: string;
  type: 'SOP' | 'Form' | 'Policy';
  chapter: string;
  content?: {
    purpose?: string;
    scope?: string;
    responsibilities?: string;
    procedureSteps?: string[];
  };
}

export const ALL_LAB_DOCUMENTS: Record<string, DocDetail> = {
  'VS-QT-5.6.1.01': {
    id: 'VS-QT-5.6.1.01',
    code: 'VS-QT 5.6.1.01',
    title: 'SOP Lấy mẫu và nhận bệnh phẩm Vi sinh',
    type: 'SOP',
    chapter: 'Chương VI',
    content: {
      purpose: 'Đảm bảo bệnh phẩm vi sinh được lấy và bảo quản đúng kỹ thuật.',
      scope: 'Bộ phận tiếp nhận bệnh phẩm Khoa Vi sinh.',
      responsibilities: 'KTV tiếp nhận kiểm tra tiêu chuẩn chấp nhận/từ chối mẫu.',
      procedureSteps: [
        '1. Đối chiếu thông tin trên mẫu phẩm và phiếu chỉ định.',
        '2. Kiểm tra thể tích, dụng cụ chứa mẫu vô trùng.',
        '3. Nhập mã Barcode vào hệ thống LIS và chuyển mẫu vào khu vực thao tác.'
      ]
    }
  }
};