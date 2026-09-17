// Nhan, 2026-09-16: "cho tôi thêm 1 chỗ xuất chứng từ mẫu ... biên bản giao
// nhận hay biên bản nghiệm thu hay bảng kê hóa đơn" -- sinh file Word (.docx)
// cho tung NCC trong trang Ho So > Hoa Don NCC, dua tren mau "BIEN BAN GIAO
// NHAN" Nhan da gui.
//
// Quyet dinh theo cau tra loi cua Nhan (AskUserQuestion 2026-09-16):
//  - "Ten hang hoa" lay nguyen tu truong noiDung da co san cua hoa don (KHONG
//    co du lieu DVT/So luong rieng tung mat hang -- de trong 2 cot do).
//  - Thong tin Ben A (dia chi/dien thoai/dai dien/chuc vu): neu ho so hop
//    dong NCC (phap_danh_hop_dong_ncc) co san cac truong nay thi dung, khong
//    co thi DE TRONG (gach chan de vien tay).
//  - Thong tin Ben B (K&H): lay tu utils/companies.js (COMPANIES[key].daiDien
//    /.chucVu/.diaChiNhanHang) -- hien dang de trong cho den khi Nhan cung
//    cap, luc do chi can sua file do la moi bien ban tu dong dien theo.
//  - Chon NHIEU hoa don cung luc -> Giao nhan/Nghiem thu ra MOI HOA DON 1
//    FILE rieng (vi moi bien ban chi co DUNG 1 ngay = ngay hoa don), goi
//    lam nhieu Document roi nen ZIP o route goi ham nay.
//  - Bang ke hoa don: LUON 1 file duy nhat du chon bao nhieu hoa don.
const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, VerticalAlign, HeadingLevel,
} = require("docx");

const NO_BORDER = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};
const THIN_BORDER = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
};

function p(text, opts) {
  opts = opts || {};
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { after: opts.after != null ? opts.after : 100 },
    children: [new TextRun({
      text: text || "",
      bold: !!opts.bold,
      italics: !!opts.italic,
      size: opts.size || 22, // 22 half-point = 11pt
      underline: opts.underline ? {} : undefined,
    })],
  });
}

function blank(label, value, width) {
  // "Nhãn: giá trị" -- neu khong co value thi ke gach chan de vien tay
  const v = (value || "").trim();
  return v ? (label + ": " + v) : (label + ": " + "…".repeat(width || 30));
}

function formatDateVn(isoDate) {
  // isoDate dang "YYYY-MM-DD" (co the rong)
  if (!isoDate) {
    const d = new Date();
    return { d: d.getDate(), m: d.getMonth() + 1, y: d.getFullYear() };
  }
  const [y, m, d] = String(isoDate).split("-").map(Number);
  if (!y || !m || !d) {
    const dd = new Date();
    return { d: dd.getDate(), m: dd.getMonth() + 1, y: dd.getFullYear() };
  }
  return { d, m, y };
}

// Bang khong vien: 2 cot -- trai = ten NCC/tieu de, phai = quoc hieu + ngay
function buildTopBlock(leftLines, rightLines) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDER,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 48, type: WidthType.PERCENTAGE },
            borders: NO_BORDER,
            verticalAlign: VerticalAlign.TOP,
            children: leftLines,
          }),
          new TableCell({
            width: { size: 52, type: WidthType.PERCENTAGE },
            borders: NO_BORDER,
            verticalAlign: VerticalAlign.TOP,
            children: rightLines,
          }),
        ],
      }),
    ],
  });
}

function buildSignatureBlock(labelA, labelB) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDER,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: NO_BORDER,
            children: [
              p("ĐẠI DIỆN BÊN A", { bold: true, align: AlignmentType.CENTER, after: 20 }),
              p(labelA || "(Bên giao hàng)", { italic: true, align: AlignmentType.CENTER, after: 1200 }),
            ],
          }),
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: NO_BORDER,
            children: [
              p("ĐẠI DIỆN BÊN B", { bold: true, align: AlignmentType.CENTER, after: 20 }),
              p(labelB || "(Bên nhận hàng)", { italic: true, align: AlignmentType.CENTER, after: 1200 }),
            ],
          }),
        ],
      }),
    ],
  });
}

// Nhan, 2026-09-17: "hóa đơn 115 có nhiều hàng hóa lắm mà 1 cái là 1 dòng cho
// tôi nhá đvt số lượng và cả thành tiền và cả tổng cho tôi" -- them cot DON
// GIA / THANH TIEN + dong TONG CONG. items: [{ten, dvt, soLuong, donGia,
// thanhTien}] (donGia/thanhTien co the rong neu khong trich xuat duoc tu PDF
// hoa don -- van hien dong, chi de trong 2 cot do). tongCong: chuoi da format
// san (vd "4.960.000"), rong thi khong hien dong tong.
function buildGoodsTable(items, tongCong) {
  const cols = ["STT", "TÊN HÀNG HÓA / DỊCH VỤ", "ĐVT", "SỐ LƯỢNG", "ĐƠN GIÁ", "THÀNH TIỀN"];
  const widths = [6, 34, 10, 12, 18, 20];
  const headerRow = new TableRow({
    tableHeader: true,
    children: cols.map((h, i) =>
      new TableCell({
        borders: THIN_BORDER,
        width: { size: widths[i], type: WidthType.PERCENTAGE },
        shading: { fill: "DCE6F1" },
        children: [p(h, { bold: true, align: AlignmentType.CENTER, after: 0 })],
      })
    ),
  });
  const rows = (items && items.length ? items : [{ ten: "", dvt: "", soLuong: "" }]).map((it, i) =>
    new TableRow({
      children: [
        new TableCell({ borders: THIN_BORDER, children: [p(String(i + 1), { align: AlignmentType.CENTER, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(it.ten || "", { after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(it.dvt || "", { align: AlignmentType.CENTER, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(it.soLuong != null ? String(it.soLuong) : "", { align: AlignmentType.CENTER, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(it.donGia != null ? String(it.donGia) : "", { align: AlignmentType.RIGHT, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(it.thanhTien != null ? String(it.thanhTien) : "", { align: AlignmentType.RIGHT, after: 0 })] }),
      ],
    })
  );
  const totalRow = tongCong
    ? new TableRow({
        children: [
          new TableCell({
            borders: THIN_BORDER,
            columnSpan: 5,
            children: [p("TỔNG CỘNG", { bold: true, align: AlignmentType.RIGHT, after: 0 })],
          }),
          new TableCell({
            borders: THIN_BORDER,
            children: [p(String(tongCong), { bold: true, align: AlignmentType.RIGHT, after: 0 })],
          }),
        ],
      })
    : null;
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: THIN_BORDER,
    rows: totalRow ? [headerRow, ...rows, totalRow] : [headerRow, ...rows],
  });
}

// bangKeRows: [{stt, soHD, ngay, noiDung, tongTien}]
function buildBangKeTable(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: ["STT", "SỐ HĐ", "NGÀY", "NỘI DUNG", "TỔNG TIỀN"].map((h, i) =>
      new TableCell({
        borders: THIN_BORDER,
        width: { size: i === 0 ? 6 : i === 1 ? 14 : i === 2 ? 12 : i === 3 ? 48 : 20, type: WidthType.PERCENTAGE },
        shading: { fill: "DCE6F1" },
        children: [p(h, { bold: true, align: AlignmentType.CENTER, after: 0 })],
      })
    ),
  });
  const body = rows.map((r) =>
    new TableRow({
      children: [
        new TableCell({ borders: THIN_BORDER, children: [p(String(r.stt), { align: AlignmentType.CENTER, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(r.soHD || "", { align: AlignmentType.CENTER, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(r.ngay || "", { align: AlignmentType.CENTER, after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(r.noiDung || "", { after: 0 })] }),
        new TableCell({ borders: THIN_BORDER, children: [p(r.tongTien || "", { align: AlignmentType.RIGHT, after: 0 })] }),
      ],
    })
  );
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: THIN_BORDER, rows: [headerRow, ...body] });
}

function nccPartyLines(ncc) {
  return [
    p("BÊN A (Bên giao hàng): " + (ncc.tenDayDuNCC || ncc.nccShort || ""), { bold: true }),
    p(blank("Địa chỉ", ncc.diaChi, 40)),
    p(blank("Điện thoại", ncc.dienThoai, 20)),
    p(blank("Đại diện", ncc.daiDien, 20) + "          " + blank("Chức vụ", ncc.chucVu, 15)),
  ];
}

function companyPartyLines(company) {
  return [
    p("BÊN B (Bên nhận hàng): " + (company.fullName || ""), { bold: true }),
    p(blank("Đại diện", company.daiDien, 20) + "          " + blank("Chức vụ", company.chucVu, 15)),
    p(blank("Địa chỉ nhận hàng", company.diaChiNhanHang, 40)),
  ];
}

// ── Biên bản giao nhận (1 hóa đơn = 1 file) ──────────────────────────────
function buildBienBanGiaoNhan({ ncc, invoice, company }) {
  const { d, m, y } = formatDateVn(invoice.ngay);
  const rightLines = [
    p("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { bold: true, align: AlignmentType.CENTER, after: 0 }),
    p("Độc lập – Tự do – Hạnh phúc", { bold: true, align: AlignmentType.CENTER, underline: true, after: 200 }),
    p("TP.HCM, ngày " + d + " tháng " + m + " năm " + y, { italic: true, align: AlignmentType.CENTER }),
  ];
  const leftLines = [p(ncc.tenDayDuNCC || ncc.nccShort || "", { bold: true, after: 0 })];

  const children = [
    buildTopBlock(leftLines, rightLines),
    p(""),
    p("BIÊN BẢN GIAO NHẬN", { bold: true, align: AlignmentType.CENTER, size: 30, after: 300 }),
    p("Số: " + (invoice.soHoaDon || "") + "/BBGN", { align: AlignmentType.CENTER, italic: true, after: 300 }),
    ...nccPartyLines(ncc),
    p(""),
    ...companyPartyLines(company),
    p(""),
    p("Hôm nay, ngày " + d + " tháng " + m + " năm " + y + ", hai bên chúng tôi tiến hành giao nhận hàng hóa theo Hóa đơn số " + (invoice.soHoaDon || "") + ", chi tiết như sau:", { after: 150 }),
    buildGoodsTable(
      invoice.pdfItems && invoice.pdfItems.length ? invoice.pdfItems : [{ ten: invoice.noiDung || "", dvt: "", soLuong: "" }],
      invoice.pdfTongCong || ""
    ),
    p(""),
    p("Bên B xác nhận Bên A đã giao cho Bên B đúng chủng loại và đủ số lượng hàng hóa như trên. Hai bên đồng ý, thống nhất ký tên. Biên bản được lập thành 02 bản, mỗi bên giữ 01 bản có giá trị pháp lý như nhau.", { after: 400 }),
    buildSignatureBlock(),
  ];
  return new Document({ sections: [{ children }] });
}

// ── Biên bản nghiệm thu (1 hóa đơn = 1 file, dùng chung khung với giao
// nhận nhưng doi loi van + tham chieu hop dong neu tim thay) ──────────────
function buildBienBanNghiemThu({ ncc, invoice, company, hopDong }) {
  const { d, m, y } = formatDateVn(invoice.ngay);
  const rightLines = [
    p("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { bold: true, align: AlignmentType.CENTER, after: 0 }),
    p("Độc lập – Tự do – Hạnh phúc", { bold: true, align: AlignmentType.CENTER, underline: true, after: 200 }),
    p("TP.HCM, ngày " + d + " tháng " + m + " năm " + y, { italic: true, align: AlignmentType.CENTER }),
  ];
  const leftLines = [p(ncc.tenDayDuNCC || ncc.nccShort || "", { bold: true, after: 0 })];

  const hopDongLine = hopDong && hopDong.soHopDong
    ? "theo Hợp đồng số " + hopDong.soHopDong + (hopDong.ngayKy ? " ngày " + hopDong.ngayKy : "") + " và Hóa đơn số " + (invoice.soHoaDon || "")
    : "theo Hóa đơn số " + (invoice.soHoaDon || "");

  const children = [
    buildTopBlock(leftLines, rightLines),
    p(""),
    p("BIÊN BẢN NGHIỆM THU", { bold: true, align: AlignmentType.CENTER, size: 30, after: 300 }),
    p("Số: " + (invoice.soHoaDon || "") + "/BBNT", { align: AlignmentType.CENTER, italic: true, after: 300 }),
    ...nccPartyLines(ncc),
    p(""),
    ...companyPartyLines(company),
    p(""),
    p("Hôm nay, ngày " + d + " tháng " + m + " năm " + y + ", hai bên tiến hành nghiệm thu hàng hóa/dịch vụ " + hopDongLine + ", chi tiết như sau:", { after: 150 }),
    buildGoodsTable(
      invoice.pdfItems && invoice.pdfItems.length ? invoice.pdfItems : [{ ten: invoice.noiDung || "", dvt: "", soLuong: "" }],
      invoice.pdfTongCong || ""
    ),
    p(""),
    p("Bên B xác nhận đã nghiệm thu hàng hóa/dịch vụ do Bên A cung cấp, đảm bảo đúng chất lượng, chủng loại, số lượng theo hợp đồng/hóa đơn đã nêu trên. Hai bên thống nhất ký tên xác nhận. Biên bản được lập thành 02 bản, mỗi bên giữ 01 bản có giá trị pháp lý như nhau.", { after: 400 }),
    buildSignatureBlock(),
  ];
  return new Document({ sections: [{ children }] });
}

// ── Bảng kê hóa đơn (luôn 1 file, gộp nhiều hóa đơn) ──────────────────────
function buildBangKeHoaDon({ ncc, invoices, company }) {
  const now = new Date();
  const rightLines = [
    p("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", { bold: true, align: AlignmentType.CENTER, after: 0 }),
    p("Độc lập – Tự do – Hạnh phúc", { bold: true, align: AlignmentType.CENTER, underline: true, after: 200 }),
    p("TP.HCM, ngày " + now.getDate() + " tháng " + (now.getMonth() + 1) + " năm " + now.getFullYear(), { italic: true, align: AlignmentType.CENTER }),
  ];
  const leftLines = [p(ncc.tenDayDuNCC || ncc.nccShort || "", { bold: true, after: 0 })];

  const rows = invoices.map((inv, i) => ({
    stt: i + 1,
    soHD: inv.soHoaDon || "",
    ngay: inv.ngay ? inv.ngay.split("-").reverse().join("/") : "",
    noiDung: inv.noiDung || "",
    tongTien: inv.tongTien || "",
  }));

  const children = [
    buildTopBlock(leftLines, rightLines),
    p(""),
    p("BẢNG KÊ HÓA ĐƠN", { bold: true, align: AlignmentType.CENTER, size: 30, after: 100 }),
    p("Nhà cung cấp: " + (ncc.tenDayDuNCC || ncc.nccShort || ""), { align: AlignmentType.CENTER, italic: true, after: 300 }),
    buildBangKeTable(rows),
    p(""),
    p("Tổng số " + invoices.length + " hóa đơn.", { bold: true, after: 400 }),
    buildSignatureBlock("(Bên lập bảng kê)", "(Bên xác nhận)"),
  ];
  return new Document({ sections: [{ children }] });
}

module.exports = { buildBienBanGiaoNhan, buildBienBanNghiemThu, buildBangKeHoaDon };
