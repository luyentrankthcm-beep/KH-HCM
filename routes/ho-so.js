// Luyen, 2026-08-24: "thêm cho tôi 1 mục nữa là hồ sơ" -- trang Hồ Sơ Hóa
// Đơn NCC: tổng hợp hóa đơn đầu vào theo NCC, link thẳng về file PDF trên
// Google Drive. Mỗi hóa đơn có thêm: tên đầy đủ NCC, tổng tiền, nội dung
// tóm tắt, hồ sơ liên quan.
const express = require("express");
const { load, save } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");

const router = express.Router();
router.use(requireLogin);

function parseInvoiceFileName(fileName) {
  const base = fileName.replace(/\.pdf$/i, "");
  const parts = base.split("_");
  const ncc = parts[0] || "";
  const soHoaDon = parts[1] || "";
  const dateStr = parts[2] || ""; // DD-MM-YYYY
  let ngay = "";
  if (dateStr && dateStr.match(/^\d{2}-\d{2}-\d{4}$/)) {
    const [d, m, y] = dateStr.split("-");
    ngay = `${y}-${m}-${d}`;
  }
  const thang = ngay ? ngay.slice(0, 7) : "";
  return { ncc, soHoaDon, ngay, thang };
}

function ensureHoSo(store) {
  if (!store.ho_so_hoa_don) store.ho_so_hoa_don = [];
}

// Luyen, 2026-08-25: tu dong dien Ten NCC va Tong tien tu bang hoa_don_dau_vao
// (doi chieu theo so hoa don tu ten file). Neu da sua tay thi uu tien du lieu tay.
function buildHddvLookup(store) {
  const lookup = {}; // soHoaDon -> { tenNCC, tongTien }
  (store.hoa_don_dau_vao || []).forEach((r) => {
    const no = String(r.soHoaDon || "").trim();
    if (!no) return;
    if (!lookup[no]) lookup[no] = { tenNCC: r.tenNCC || "", tongTien: 0 };
    lookup[no].tongTien += r.soTien || 0;
  });
  return lookup;
}

// Fallback: dò tên đầy đủ NCC từ danh sách NCC (chi_phi_ncc_list) khi số HĐ
// không khớp trong hoa_don_dau_vao. Chuẩn hoá bằng cách bỏ ký tự đặc biệt,
// lowercase rồi kiểm tra nccShort có nằm trong tên NCC không.
function normForMatch(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}
function lookupNccNameFallback(nccShort, store) {
  const shortNorm = normForMatch(nccShort);
  if (shortNorm.length < 3) return "";
  const allLists = [
    ...(store.chi_phi_ncc_list_cu || []),
    ...(store.chi_phi_ncc_list_moi || []),
    ...(store.chi_phi_ncc_list || []),
    ...(store.danh_muc_ma_nha_cung_cap_cu || []).map((r) => ({ tenNCC: r.ten })),
    ...(store.danh_muc_ma_nha_cung_cap_moi || []).map((r) => ({ tenNCC: r.ten })),
    ...(store.danh_muc_ma_nha_cung_cap || []).map((r) => ({ tenNCC: r.ten })),
  ];
  for (const rec of allLists) {
    const name = rec.tenNCC || "";
    if (!name) continue;
    if (normForMatch(name).includes(shortNorm)) return name;
  }
  return "";
}

function enrichRow(r, hddvLookup, store) {
  const parsed = parseInvoiceFileName(r.fileName || "");
  const inv = hddvLookup && parsed.soHoaDon ? hddvLookup[parsed.soHoaDon] : null;
  const autoTenFromHD = inv ? inv.tenNCC : "";
  const autoTenFromNCC = !autoTenFromHD ? lookupNccNameFallback(parsed.ncc, store) : "";
  const autoTen = autoTenFromHD || autoTenFromNCC;
  const autoTien = inv && inv.tongTien ? inv.tongTien.toLocaleString("vi-VN") + "đ" : "";
  return {
    ...r,
    nccShort: parsed.ncc,
    soHoaDon: parsed.soHoaDon,
    ngay: parsed.ngay,
    thang: parsed.thang,
    driveLink: r.driveId ? `https://drive.google.com/file/d/${r.driveId}/view` : "",
    tenDayDuNCC: r.tenDayDuNCC || autoTen,
    tongTien: r.tongTien || autoTien,
    noiDung: r.noiDung || "",
    hoSoLienQuan: r.hoSoLienQuan || "",
    autoFilled: !r.tenDayDuNCC && !!autoTen,
  };
}

router.get("/ho-so/hoa-don-ncc", (req, res) => {
  const store = load();
  ensureHoSo(store);
  const thangFilter = req.query.thang || "";
  const hddvLookup = buildHddvLookup(store);

  let rows = store.ho_so_hoa_don.map((r) => enrichRow(r, hddvLookup, store));
  if (thangFilter) rows = rows.filter((r) => r.thang === thangFilter);
  rows.sort((a, b) => {
    const nccCmp = a.nccShort.toLowerCase().localeCompare(b.nccShort.toLowerCase());
    return nccCmp !== 0 ? nccCmp : a.ngay.localeCompare(b.ngay);
  });

  const allThang = [...new Set(store.ho_so_hoa_don.map((r) => {
    return parseInvoiceFileName(r.fileName || "").thang;
  }).filter(Boolean))].sort().reverse();

  res.render("ho-so-hoa-don", {
    userName: req.session.userName,
    rows,
    thangFilter,
    allThang,
    total: rows.length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Luu cac truong bo sung cho 1 hoa don
router.post("/ho-so/hoa-don-ncc/:driveId/sua", requireAdmin, (req, res) => {
  const store = load();
  ensureHoSo(store);
  const r = store.ho_so_hoa_don.find((x) => x.driveId === req.params.driveId);
  if (!r) return res.redirect("/ho-so/hoa-don-ncc?error=Không+tìm+thấy");
  const { tenDayDuNCC, tongTien, noiDung, hoSoLienQuan, thang } = req.body;
  r.tenDayDuNCC = (tenDayDuNCC || "").trim();
  r.tongTien = (tongTien || "").trim();
  r.noiDung = (noiDung || "").trim();
  r.hoSoLienQuan = (hoSoLienQuan || "").trim();
  save(store);
  const qs = thang ? "?thang=" + encodeURIComponent(thang) : "";
  res.redirect("/ho-so/hoa-don-ncc" + qs + "&success=Đã+lưu");
});

module.exports = router;
