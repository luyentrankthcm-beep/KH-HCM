const XLSX = require("xlsx");

// Luyen, 2026-07-21: "có cái nút tự cập nhật cho tôi nhá" (trang Hợp Đồng NCC)
// -- parser doc file Google Sheet "THEO DÕI HĐ HN-HCM" (sheet "Tổng hợp
// HCM-HN") ma Luyen xuat lai moi lan can cap nhat, cung 1 file da dung de
// import 48 hop dong NCC + 69 gian ban dau (xem theo_doi_hd.xlsx trong outputs
// luc phan tich). Doc theo VI TRI COT (khong theo ten header) vi hang header
// co vai cot bi trung ten do merge cell trong file goc -- vi tri cot da doi
// chieu 1-1 voi hcm_rows.json (787 dong, da parse va dung lam nguon du lieu
// cho 48 NCC + 69 gian truoc do) nen dam bao dung.
const SHEET_NAME = "Tổng hợp HCM-HN";

const COLS = [
  "kv", // 0
  "congTyRaw", // 1
  "diaDiem", // 2
  "tenKH", // 3
  "mstKH", // 4
  "diaChiKH", // 5
  "noiDungHD", // 6
  "phanLoaiHD", // 7
  "thuocBP", // 8
  "soHopDong", // 9
  "viTri", // 10
  "soGheMayDienTich", // 11
  "viTriLuuBanCung", // 12
  "hinhThucHD", // 13
  "khach", // 14
  "congTy2", // 15
  "soTKNHkhach", // 16
  "khachMoTaiNH", // 17
  "ngayKyHD", // 18
  "ngayBatDauHD", // 19
  "ngayHetHan", // 20
  "baoHanHD", // 21
  "tienCocDamBao", // 22
  "tienCocThiCong", // 23
  "soTienHD", // 24
  "tongTienThueThang", // 25
  "thongTinVAT", // 26
  "phiDichVuKhac", // 27
  "dinhKyThanhToan", // 28
  "ghiChu", // 29
  "tinhTrangCoc", // 30
  "billNhanCoc", // 31
  "tinh", // 32
  "linkHDduDau", // 33
  "linkHDchuaDuDau", // 34
];

function pad2(n) {
  return String(n).padStart(2, "0");
}

// Cong thuc excelSerialToIso da kiem chung o utils/bankStatementParser.js va
// utils/chiPhiSheetParser.js -- dung UTC getters, tranh bug lech 1 ngay cua
// cellDates:true.
function excelSerialToIso(serial) {
  const days = Math.round(serial);
  const utcMillis = (days - 25569) * 86400 * 1000;
  const d = new Date(utcMillis);
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

function cellDate(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") return excelSerialToIso(v);
  if (v instanceof Date) return ""; // khong dung cellDates, khong nen gap truong hop nay
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{2,4})$/);
  if (m) {
    let dd = Number(m[1]);
    let mm = Number(m[2]);
    let yy = Number(m[3]);
    if (yy < 100) yy += 2000;
    return `${yy}-${pad2(mm)}-${pad2(dd)}`;
  }
  return "";
}

function cellText(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

function cellNumber(v) {
  if (typeof v === "number") return v;
  if (!v) return null;
  const n = Number(String(v).replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) && n !== 0 ? n : null;
}

// "K&H mới (705)" -> kh_moi, "K&H cũ (989)" -> kh_cu (giong quy uoc da dung
// khi import 48 hop dong NCC ban dau).
function normCongTy(v) {
  const s = cellText(v).toLowerCase();
  if (s.includes("mới") || s.includes("moi") || s.includes("705")) return "kh_moi";
  if (s.includes("cũ") || s.includes("cu") || s.includes("989")) return "kh_cu";
  return "";
}

// Luyen, 2026-07-21 (lan 3): BUG FIX -- khi doc CSV export tu Google Sheet
// (nut "Cập nhật hợp đồng" doc thang tu link, khong qua upload file .xlsx
// nua), XLSX.read(buffer, {type:"buffer"}) doan sai encoding cho CSV thuan
// text (khong co header ZIP nhu file .xlsx that), lam mat dau tieng Viet
// kieu mojibake (vd "Khác" bi doc thanh "KhÃ¡c" -- 2 byte UTF-8 cua "á" bi
// hieu nham thanh 2 ky tu rieng). File .xlsx that luon bat dau bang chu ky
// ZIP "PK" nen van doc dung binh thuong qua nhanh "buffer". Voi CSV (khong
// co "PK" o dau), tu giai ma bang Buffer.toString("utf8") TRUOC roi moi dua
// cho XLSX duoi dang "string" -- tranh hoan toan buoc doan sai encoding cua
// XLSX cho truong hop nay.
function parseHopDongHcmWorkbook(buffer) {
  const isZip = Buffer.isBuffer(buffer) && buffer.length > 1 && buffer[0] === 0x50 && buffer[1] === 0x4b;
  const wb = isZip ? XLSX.read(buffer, { type: "buffer" }) : XLSX.read(buffer.toString("utf8"), { type: "string" });
  const sheetName = wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const rows = [];
  for (let r = 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const kv = cellText(row[0]);
    if (kv.toUpperCase() !== "HCM") continue; // web nay chi theo doi khu vuc HCM
    const soHopDong = cellText(row[9]);
    const noiDungHD = cellText(row[6]);
    const thuocBP = cellText(row[8]);
    const diaDiem = cellText(row[2]);
    if (!soHopDong && !noiDungHD && !thuocBP && !diaDiem) continue; // dong rong
    rows.push({
      kv,
      congTy: normCongTy(row[1]),
      congTyRaw: cellText(row[1]),
      diaDiem,
      tenKH: cellText(row[3]),
      mstKH: cellText(row[4]),
      diaChiKH: cellText(row[5]),
      noiDungHD,
      phanLoaiHD: cellText(row[7]),
      thuocBP,
      soHopDong,
      viTri: cellText(row[10]),
      soTKNHkhach: cellText(row[16]),
      khachMoTaiNH: cellText(row[17]),
      ngayKyHD: cellDate(row[18]),
      ngayBatDauHD: cellDate(row[19]),
      ngayHetHan: cellDate(row[20]),
      baoHanHD: cellText(row[21]),
      tienCocDamBao: cellNumber(row[22]),
      tienCocThiCong: cellNumber(row[23]),
      soTienHD: row[24] === null || row[24] === undefined ? null : row[24],
      tongTienThueThang: row[25] === null || row[25] === undefined ? null : row[25],
      ghiChu: cellText(row[29]),
      linkHDduDau: cellText(row[33]),
      linkHDchuaDuDau: cellText(row[34]),
      _row: r + 1,
    });
  }
  return { sheetName, rows };
}

const NCC_THUOC_BP = new Set(["khác", "khac", "mua bán vocher", "mua ban vocher"]);
function isNccRow(row) {
  const n = (row.thuocBP || "").trim().toLowerCase();
  return NCC_THUOC_BP.has(n);
}

module.exports = { parseHopDongHcmWorkbook, isNccRow, normCongTy };
