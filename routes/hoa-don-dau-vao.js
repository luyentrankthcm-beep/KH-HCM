const express = require("express");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan, 2026-07-28: "thêm cho tôi 1 trang hóa đơn đầu vào kh cũ và kh mới
// nhá nội dung tôi sẽ nối sao" -- trang MOI, khac voi "Đối soát Chi phí"
// (doi soat so tien tren sao ke ngan hang) va "Chi Phí" (so ghi chi phi theo
// Mien Nam/Bac). Day la SO GHI HOA DON DAU VAO (hoa don mua hang/dich vu tu
// NCC ma cong ty nhan duoc), tach theo cong ty qua nut Cu/Moi o topbar (giong
// Phap danh/Chi Phi). Theo dung tien le "Phap danh" (routes/phap-danh.js,
// 2026-07-19): tao khung trang RONG truoc (bang + form them dong voi bo
// truong pho bien nhat cho hoa don dau vao), Chi Nhan se mo ta them cot/
// truong chi tiet sau khi biet ro can gi (vd MST NCC, so tien truoc/sau VAT
// tach rieng, ky hieu-mau so hoa don...).
function ensureShape(store) {
  if (!store.hoa_don_dau_vao) store.hoa_don_dau_vao = [];
}

function ensureDefaults(row) {
  return Object.assign(
    {
      congTy: "kh_cu",
      ngayHD: "",
      tenNCC: "",
      mstNCC: "",
      soHoaDon: "",
      dienGiai: "",
      soTien: 0,
      linkHoaDon: "",
      daHachToan: false,
      ghiChu: "",
    },
    row
  );
}

router.get("/hoa-don-dau-vao", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hachToanFilter = req.query.hachToan || ""; // "" = tat ca, "1" = da hach toan, "0" = chua hach toan
  const thangFilter = req.query.thang || "";

  let rows = store.hoa_don_dau_vao.map(ensureDefaults).filter((r) => r.congTy === activeCompany);
  const totalForCompany = rows.length;

  const monthSet = new Set();
  rows.forEach((r) => {
    const m = (r.ngayHD || "").slice(0, 7);
    if (m) monthSet.add(m);
  });
  const availableMonths = [...monthSet].sort().reverse();

  if (hachToanFilter === "1") rows = rows.filter((r) => r.daHachToan);
  else if (hachToanFilter === "0") rows = rows.filter((r) => !r.daHachToan);
  if (thangFilter) rows = rows.filter((r) => (r.ngayHD || "").slice(0, 7) === thangFilter);

  rows.sort((a, b) => (a.ngayHD < b.ngayHD ? 1 : -1));
  const tongTien = rows.reduce((s, r) => s + (r.soTien || 0), 0);

  res.render("hoa-don-dau-vao", {
    userName: req.session.userName,
    rows,
    totalForCompany,
    hachToanFilter,
    thangFilter,
    availableMonths,
    tongTien,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.get("/hoa-don-dau-vao/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hachToanFilter = req.query.hachToan || "";
  const thangFilter = req.query.thang || "";
  let rows = store.hoa_don_dau_vao.map(ensureDefaults).filter((r) => r.congTy === activeCompany);
  if (hachToanFilter === "1") rows = rows.filter((r) => r.daHachToan);
  else if (hachToanFilter === "0") rows = rows.filter((r) => !r.daHachToan);
  if (thangFilter) rows = rows.filter((r) => (r.ngayHD || "").slice(0, 7) === thangFilter);
  rows.sort((a, b) => (a.ngayHD < b.ngayHD ? 1 : -1));

  const exportRows = rows.map((r) => ({
    "Ngày HĐ": r.ngayHD,
    "Tên NCC": r.tenNCC,
    "MST NCC": r.mstNCC,
    "Số hóa đơn": r.soHoaDon,
    "Diễn giải": r.dienGiai,
    "Số tiền": r.soTien,
    "Link hóa đơn": r.linkHoaDon,
    "Đã hạch toán": r.daHachToan ? "Có" : "Không",
    "Ghi chú": r.ghiChu,
  }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hoa don dau vao");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=hoa-don-dau-vao-${activeCompany}.xlsx`);
  res.send(buf);
});

router.post("/hoa-don-dau-vao", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const { ngayHD, tenNCC, mstNCC, soHoaDon, dienGiai, soTien, linkHoaDon, ghiChu } = req.body;
    if (!ngayHD) throw new Error("Thiếu ngày hóa đơn.");
    if (!tenNCC || !tenNCC.trim()) throw new Error("Thiếu tên NCC.");
    const amt = soTien ? Number(String(soTien).replace(/[^\d]/g, "")) : 0;
    store.hoa_don_dau_vao.push({
      id: nextId(store, "hoa_don_dau_vao_seq") || Date.now(),
      congTy: activeCompany,
      ngayHD,
      tenNCC: tenNCC.trim(),
      mstNCC: (mstNCC || "").trim(),
      soHoaDon: (soHoaDon || "").trim(),
      dienGiai: (dienGiai || "").trim(),
      soTien: amt,
      linkHoaDon: (linkHoaDon || "").trim(),
      daHachToan: false,
      ghiChu: (ghiChu || "").trim(),
      createdAt: new Date().toISOString(),
      source: "nhap tay",
    });
    save(store);
    res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent("Đã lưu hóa đơn."));
  } catch (e) {
    res.redirect("/hoa-don-dau-vao?error=" + encodeURIComponent(e.message));
  }
});

router.post("/hoa-don-dau-vao/:id/hach-toan", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.hoa_don_dau_vao.find((x) => String(x.id) === req.params.id);
  if (r) {
    r.daHachToan = req.body.daHachToan === "1";
    save(store);
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang) qs.push("thang=" + encodeURIComponent(req.body.thang));
  qs.push("success=" + encodeURIComponent("Đã cập nhật trạng thái hạch toán."));
  res.redirect("/hoa-don-dau-vao?" + qs.join("&"));
});

router.post("/hoa-don-dau-vao/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  store.hoa_don_dau_vao = store.hoa_don_dau_vao.filter((r) => String(r.id) !== req.params.id);
  save(store);
  res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent("Đã xóa hóa đơn."));
});

module.exports = router;
