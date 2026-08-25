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

function enrichRow(r) {
  const parsed = parseInvoiceFileName(r.fileName || "");
  return {
    ...r,
    nccShort: parsed.ncc,
    soHoaDon: parsed.soHoaDon,
    ngay: parsed.ngay,
    thang: parsed.thang,
    driveLink: r.driveId ? `https://drive.google.com/file/d/${r.driveId}/view` : "",
    tenDayDuNCC: r.tenDayDuNCC || "",
    tongTien: r.tongTien || "",
    noiDung: r.noiDung || "",
    hoSoLienQuan: r.hoSoLienQuan || "",
  };
}

router.get("/ho-so/hoa-don-ncc", (req, res) => {
  const store = load();
  ensureHoSo(store);
  const thangFilter = req.query.thang || "";

  let rows = store.ho_so_hoa_don.map(enrichRow);
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
