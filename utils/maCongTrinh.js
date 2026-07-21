// ---------- Danh sach Ma cong trinh chuan (rieng theo tung cong ty) ----------
// Luyen, 2026-07-20: "thêm chỗ tải mã công trình kh cũ và kh mới đều khác
// nhau như 2 pháp nhân ... nếu hiển thị giống thì lấy còn không thì cái gần
// giống nhất đổi qua mã công trình đó cho tôi map đúng mã công trình chuẩn" --
// tai file "DANH SACH CONG TRINH" (vd tu phan mem quan ly cong trinh) rieng
// cho KH Cu / KH Moi (2 phap nhan khac nhau nen 2 danh sach ma khac nhau),
// dung lam nguon "chuan" de doi chieu/chuan hoa cac ten gian/cong trinh xuat
// hien o noi khac trong he thong (chi phi, doi soat...): neu ten khop chinh
// xac (khong phan biet dau/hoa thuong) thi lay luon ma tuong ung, neu khong
// khop tuyet doi thi tim ten gan giong nhat trong danh sach chuan va doi ve
// dung ma cong trinh do.
//
// File mau da nhan (Danh_sach_cong_trinh.xlsx, KH Moi): 1 sheet duy nhat ten
// "DANH SACH CONG TRINH", tieu de o dong 1, dong trong o dong 2, HANG TIEU
// DE THUC SU o dong 3: STT | Ma cong trinh | Ten cong trinh | Loai cong
// trinh | Tinh trang | Ngay bat dau | Ngay ket thuc | Du toan | Chu dau tu |
// Chi nhanh | Trang thai. Parser ben duoi do dong dau tien co chua "ma cong
// trinh" (khong phan biet dau/hoa thuong) lam hang tieu de, roi anh xa MOI
// cot theo ten tieu de (khong hard-code thu tu cot) de neu file KH Cu sau
// nay đổi thứ tự cột vẫn đọc đúng.

const XLSX = require("xlsx");
const { normText, normCode } = require("./momoReconcile");

// Anh xa tu tieu de cot (da normText) sang ten field luu trong store.
const HEADER_FIELD_MAP = {
  stt: "stt",
  "ma cong trinh": "maCongTrinh",
  "ten cong trinh": "tenCongTrinh",
  "loai cong trinh": "loaiCongTrinh",
  "tinh trang": "tinhTrang",
  "ngay bat dau": "ngayBatDau",
  "ngay ket thuc": "ngayKetThuc",
  "du toan": "duToan",
  "chu dau tu": "chuDauTu",
  "chi nhanh": "chiNhanh",
  "trang thai": "trangThai",
};

function parseMaCongTrinhSheet(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  // Uu tien sheet co ten chua "cong trinh", neu khong co thi lay sheet dau tien.
  const sheetName =
    wb.SheetNames.find((n) => normText(n).includes("cong trinh")) || wb.SheetNames[0];
  if (!sheetName) return { sheetName: null, rows: [] };
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let maCongTrinhCol = -1;
  const colFieldMap = {}; // col index -> field name
  for (let r = 0; r < Math.min(grid.length, 15); r++) {
    const row = grid[r] || [];
    let mc = -1;
    const map = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (s.includes("ma cong trinh")) mc = c;
      if (HEADER_FIELD_MAP[s]) map[c] = HEADER_FIELD_MAP[s];
    });
    if (mc !== -1) {
      headerRowIdx = r;
      maCongTrinhCol = mc;
      Object.assign(colFieldMap, map);
      break;
    }
  }
  if (headerRowIdx < 0) return { sheetName, rows: [] };

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const maCongTrinh = normCode(row[maCongTrinhCol]);
    if (!maCongTrinh) continue;
    const entry = { maCongTrinh };
    Object.entries(colFieldMap).forEach(([col, field]) => {
      if (field === "maCongTrinh") return;
      const v = row[Number(col)];
      entry[field] = v === null || v === undefined ? "" : String(v).trim();
    });
    rows.push(entry);
  }
  return { sheetName, rows };
}

// ---------- Doi chieu / chuan hoa 1 ten thanh Ma cong trinh chuan ----------
// 1) Khop tuyet doi (normText) voi "Ten cong trinh" hoac chinh "Ma cong
//    trinh" trong danh sach chuan -> lay ngay.
// 2) Neu khong co, tim ten GAN GIONG NHAT (chua nhau theo 1 trong 2 chieu,
//    uu tien chuoi khop dai nhat) -> doi ve Ma cong trinh cua dong do.
// 3) Neu khong tim duoc gi ca -> tra ve null (giu nguyen text goc, KHONG
//    doan bua) de con nguoi tu kiem tra thay vi doi nham.
function findBestMaCongTrinh(rawName, masterRows) {
  if (!rawName || !masterRows || masterRows.length === 0) return null;
  const needle = normText(rawName);
  if (!needle) return null;

  // 1) khop tuyet doi
  for (const r of masterRows) {
    if (normText(r.tenCongTrinh) === needle || normText(r.maCongTrinh) === needle) {
      return { maCongTrinh: r.maCongTrinh, matchType: "exact", matchedOn: r.tenCongTrinh || r.maCongTrinh };
    }
  }

  // 2) gan giong nhat: uu tien ten chua nhau (chieu nao cung duoc), lay
  // truong hop co do dai chuoi khop (overlap) lon nhat.
  let best = null;
  let bestScore = 0;
  for (const r of masterRows) {
    const cand = normText(r.tenCongTrinh);
    if (!cand) continue;
    let score = 0;
    if (cand.includes(needle) || needle.includes(cand)) {
      score = Math.min(cand.length, needle.length);
    }
    if (score > bestScore) {
      bestScore = score;
      best = r;
    }
  }
  if (best && bestScore >= 4) {
    return { maCongTrinh: best.maCongTrinh, matchType: "fuzzy", matchedOn: best.tenCongTrinh };
  }
  return null;
}

module.exports = { parseMaCongTrinhSheet, findBestMaCongTrinh };
