const XLSX = require("xlsx");

// Chi Nhan, 2026-07-28: "từ cái bảng này thêm cho tôi chỗ úp file này luôn
// nhá ... lấy ra cái nào chung hóa đơn gộp lại nhưng diễn giải có hàng hóa
// chia ra nhá" -- parser cho file xuat tu he thong hoa don dien tu, ten sheet
// "Bang_ke_hoa_don_hh_mua_vao_CT" ("BẢNG KÊ HÓA ĐƠN HÀNG HOÁ, DỊCH VỤ MUA
// VÀO"). Cau truc thuc te (kiem tra truc tiep tren file Chi Nhan gui,
// 2026-07-28): 6 dong tieu de/khoang trang, dong tieu de cot o INDEX 6 (dong
// thu 7), du lieu tu dong 7 tro di, 1 dong CUOI CUNG la dong "Tổng cộng" (co
// gia tri o cot "Mẫu số" thay vi STT so).
//
// MOT hoa don (1 So hoa don + Ky hieu) co the co NHIEU dong chi tiet (1 dong
// / 1 hang hoa-dich vu) -- Chi Nhan xac nhan ro: GIU NGUYEN moi dong la 1
// ban ghi rieng (khong gop lai thanh 1 dong / 1 hoa don), vi moi dong co
// Dien giai + tien truoc thue/thue/tong tien RIENG cho tung hang hoa.
const HEADER_KEYWORDS = ["stt", "số hóa đơn", "tên người bán"];

function normHeaderCell(v) {
  return String(v || "").trim().toLowerCase();
}

function findHeaderRowIdx(grid) {
  for (let r = 0; r < Math.min(grid.length, 20); r++) {
    const row = (grid[r] || []).map(normHeaderCell);
    if (HEADER_KEYWORDS.every((kw) => row.some((c) => c.includes(kw)))) {
      return r;
    }
  }
  return -1;
}

function findCol(headerRow, keywords) {
  for (let c = 0; c < headerRow.length; c++) {
    const h = normHeaderCell(headerRow[c]);
    if (keywords.some((kw) => h.includes(kw))) return c;
  }
  return -1;
}

function parseDateCell(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    const utcMillis = (Math.round(v) - 25569) * 86400 * 1000;
    return new Date(utcMillis).toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m2) {
    const [, y, mo, d] = m2;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return "";
}

function parseNumberCell(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/,/g, "").trim());
  return isNaN(n) ? 0 : n;
}

// Tra ve { rows, skippedNoInvoiceNo } -- rows la mang cac ban ghi da chuan
// hoa, GIU NGUYEN 1 dong nguon = 1 dong ket qua (khong gop hoa don).
function parseHoaDonDauVaoWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const headerIdx = findHeaderRowIdx(grid);
  if (headerIdx === -1) {
    throw new Error(
      'Không nhận diện được file: cần có dòng tiêu đề với các cột "STT", "Số hóa đơn", "Tên người bán" (đúng định dạng "Bảng kê hóa đơn hàng hóa, dịch vụ mua vào").'
    );
  }
  const headerRow = grid[headerIdx];
  const col = {
    stt: findCol(headerRow, ["stt"]),
    kyHieu: findCol(headerRow, ["ký hiệu"]),
    soHoaDon: findCol(headerRow, ["số hóa đơn"]),
    ngayHD: findCol(headerRow, ["ngày hóa đơn"]),
    tenNguoiBan: findCol(headerRow, ["tên người bán"]),
    mstNguoiBan: findCol(headerRow, ["mst người bán"]),
    dienGiai: findCol(headerRow, ["diễn giải"]),
    truocThue: findCol(headerRow, ["doanh số bán chưa thuế", "chưa thuế"]),
    thueGtgt: findCol(headerRow, ["thuế gtgt", "tiền thuế"]),
    tongThanhToan: findCol(headerRow, ["tổng tiền thanh toán"]),
  };
  const missing = Object.entries(col)
    .filter(([k, v]) => v === -1 && k !== "stt" && k !== "kyHieu")
    .map(([k]) => k);
  if (missing.length > 0) {
    throw new Error(`Không tìm thấy cột: ${missing.join(", ")} trong file.`);
  }

  const rows = [];
  let skippedNoInvoiceNo = 0;
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const soHoaDon = row[col.soHoaDon];
    // Dong "Tong cong" o cuoi file khong co So hoa don -- bo qua lang le.
    if (soHoaDon === null || soHoaDon === undefined || String(soHoaDon).trim() === "") {
      skippedNoInvoiceNo++;
      continue;
    }
    rows.push({
      soHoaDon: String(soHoaDon).trim(),
      kyHieuHD: col.kyHieu !== -1 ? String(row[col.kyHieu] || "").trim() : "",
      ngayHD: parseDateCell(row[col.ngayHD]),
      tenNCC: String(row[col.tenNguoiBan] || "").trim(),
      mstNCC: String(row[col.mstNguoiBan] || "").trim(),
      dienGiai: String(row[col.dienGiai] || "").trim(),
      soTienTruocThue: parseNumberCell(row[col.truocThue]),
      tienThue: parseNumberCell(row[col.thueGtgt]),
      soTien: parseNumberCell(row[col.tongThanhToan]),
    });
  }

  return { sheetName, rows, skippedNoInvoiceNo };
}

module.exports = { parseHoaDonDauVaoWorkbook };
