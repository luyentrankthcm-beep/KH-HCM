// Luyen, 2026-08-24: "thêm cho tôi 1 mục nữa là hồ sơ" -- trang Hồ Sơ Hóa
// Đơn NCC: tổng hợp hóa đơn đầu vào theo NCC, link thẳng về file PDF trên
// Google Drive. Mỗi tháng Luyến sync lại bằng cách paste folder Drive URL.
const express = require("express");
const { load, save, nextId } = require("../store");
const { requireLogin } = require("../middleware/auth");

const router = express.Router();
router.use(requireLogin);

function parseInvoiceFileName(fileName) {
  // Format: NCC_SoHoaDon_DD-MM-YYYY.pdf (hoac NCC_SoHoaDon_DD-MM-YYYY.pdf)
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

router.get("/ho-so/hoa-don-ncc", (req, res) => {
  const store = load();
  ensureHoSo(store);
  const thangFilter = req.query.thang || "";

  let rows = store.ho_so_hoa_don.map((r) => {
    const parsed = parseInvoiceFileName(r.fileName || "");
    return {
      ...r,
      nccParsed: parsed.ncc,
      soHoaDon: parsed.soHoaDon,
      ngay: parsed.ngay,
      thang: parsed.thang,
      driveLink: r.driveId ? `https://drive.google.com/file/d/${r.driveId}/view` : "",
    };
  });

  if (thangFilter) rows = rows.filter((r) => r.thang === thangFilter);

  // Group by NCC
  const groups = {};
  rows.forEach((r) => {
    const key = r.nccParsed || "(Chưa rõ)";
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });
  const nccGroups = Object.entries(groups)
    .map(([ncc, items]) => ({ ncc, items: items.sort((a, b) => a.ngay.localeCompare(b.ngay)) }))
    .sort((a, b) => a.ncc.localeCompare(b.ncc));

  // Distinct months for filter
  const allThang = [...new Set(store.ho_so_hoa_don.map((r) => {
    const p = parseInvoiceFileName(r.fileName || "");
    return p.thang;
  }).filter(Boolean))].sort().reverse();

  res.render("ho-so-hoa-don", {
    userName: req.session.userName,
    nccGroups,
    thangFilter,
    allThang,
    total: rows.length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

module.exports = router;
