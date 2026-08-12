// Luyen, 2026-08-12: "thêm cho tôi thêm 1 trang hóa đơn đầu ra nữa á"
// Trang quan ly hoa don dau ra (hoa don ban hang, xuat cho khach).
// Luu trong store.hoa_don_dau_ra -- mang cac doi tuong {id, congTy, ...}.
const express = require("express");
const router = express.Router();
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { COMPANIES, getCompany } = require("../utils/companies");

function ensureShape(store) {
  if (!store.hoa_don_dau_ra) store.hoa_don_dau_ra = [];
}

// GET /hoa-don-dau-ra
router.get("/hoa-don-dau-ra", requireLogin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);

  // Bo loc
  const selectedMonth = req.query.thang || "";
  const selectedLoai = req.query.loai || ""; // "khach-le", "hop-dong", ...
  const searchQ = (req.query.q || "").trim().toLowerCase();

  let rows = store.hoa_don_dau_ra.filter(
    (r) => (r.congTy || "kh_cu") === activeCompany
  );

  // Loc thang
  if (selectedMonth) {
    rows = rows.filter((r) => (r.ngayHD || "").startsWith(selectedMonth));
  }
  // Loc loai
  if (selectedLoai) {
    rows = rows.filter((r) => (r.loaiHD || "") === selectedLoai);
  }
  // Tim kiem tu khoa
  if (searchQ) {
    rows = rows.filter((r) =>
      [r.soHD, r.tenKhachHang, r.dienGiai, r.maKH].some(
        (v) => v && String(v).toLowerCase().includes(searchQ)
      )
    );
  }

  // Sort: moi nhat len tren
  rows = [...rows].sort((a, b) => (b.ngayHD || "").localeCompare(a.ngayHD || ""));

  // Tong tien
  const tongTien = rows.reduce((s, r) => s + (Number(r.soTien) || 0), 0);
  const tongTienVAT = rows.reduce((s, r) => s + (Number(r.soTienVAT) || 0), 0);

  // Danh sach thang co du lieu (de dropdown bo loc)
  const allRows = store.hoa_don_dau_ra.filter(
    (r) => (r.congTy || "kh_cu") === activeCompany
  );
  const monthSet = new Set(allRows.map((r) => (r.ngayHD || "").slice(0, 7)).filter(Boolean));
  const months = [...monthSet].sort().reverse();

  res.render("hoa-don-dau-ra", {
    COMPANIES, activeCompany,
    userName: req.session.userName, isAdmin: req.session.isAdmin,
    userRole: req.session.userRole,
    rows, months, tongTien, tongTienVAT,
    selectedMonth, selectedLoai, searchQ: req.query.q || "",
    currentPath: req.path,
    success: req.query.success || "",
    error: req.query.error || "",
  });
});

// POST /hoa-don-dau-ra/them -- them moi 1 dong
router.post("/hoa-don-dau-ra/them", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);

  const row = {
    id: nextId(store),
    congTy: activeCompany,
    ngayHD: (req.body.ngayHD || "").trim(),
    soHD: (req.body.soHD || "").trim(),
    maKH: (req.body.maKH || "").trim(),
    tenKhachHang: (req.body.tenKhachHang || "").trim(),
    loaiHD: (req.body.loaiHD || "").trim(),
    dienGiai: (req.body.dienGiai || "").trim(),
    soTien: Number(String(req.body.soTien || "0").replace(/[^\d.-]/g, "")) || 0,
    soTienVAT: Number(String(req.body.soTienVAT || "0").replace(/[^\d.-]/g, "")) || 0,
    thueVAT: (req.body.thueVAT || "").trim(),
    ghiChu: (req.body.ghiChu || "").trim(),
    createdAt: new Date().toISOString(),
  };

  store.hoa_don_dau_ra.push(row);
  save(store);
  res.redirect("/hoa-don-dau-ra?success=Đã thêm hóa đơn " + encodeURIComponent(row.soHD || String(row.id)));
});

// POST /hoa-don-dau-ra/sua/:id -- cap nhat 1 dong
router.post("/hoa-don-dau-ra/sua/:id", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const id = Number(req.params.id);
  const row = store.hoa_don_dau_ra.find((r) => r.id === id);
  if (!row) return res.redirect("/hoa-don-dau-ra?error=Không tìm thấy hóa đơn id=" + id);

  row.ngayHD      = (req.body.ngayHD || "").trim();
  row.soHD        = (req.body.soHD || "").trim();
  row.maKH        = (req.body.maKH || "").trim();
  row.tenKhachHang = (req.body.tenKhachHang || "").trim();
  row.loaiHD      = (req.body.loaiHD || "").trim();
  row.dienGiai    = (req.body.dienGiai || "").trim();
  row.soTien      = Number(String(req.body.soTien || "0").replace(/[^\d.-]/g, "")) || 0;
  row.soTienVAT   = Number(String(req.body.soTienVAT || "0").replace(/[^\d.-]/g, "")) || 0;
  row.thueVAT     = (req.body.thueVAT || "").trim();
  row.ghiChu      = (req.body.ghiChu || "").trim();
  row.updatedAt   = new Date().toISOString();

  save(store);
  res.redirect("/hoa-don-dau-ra?success=Đã cập nhật hóa đơn " + encodeURIComponent(row.soHD || String(row.id)));
});

// POST /hoa-don-dau-ra/xoa/:id
router.post("/hoa-don-dau-ra/xoa/:id", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const id = Number(req.params.id);
  const idx = store.hoa_don_dau_ra.findIndex((r) => r.id === id);
  if (idx === -1) return res.redirect("/hoa-don-dau-ra?error=Không tìm thấy hóa đơn id=" + id);
  const removed = store.hoa_don_dau_ra.splice(idx, 1)[0];
  save(store);
  res.redirect("/hoa-don-dau-ra?success=Đã xóa hóa đơn " + encodeURIComponent(removed.soHD || String(removed.id)));
});

// POST /hoa-don-dau-ra/import-json -- nhap hang loat tu JSON array (admin only)
// Body: { replace: true, records: [{congTy,ngayHD,soHD,maKH,...}] }
const multer = require("multer");
const uploadMem = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
router.post("/hoa-don-dau-ra/import-json", requireAdmin, uploadMem.none(), (req, res) => {
  try {
    const raw = req.body.records;
    if (!raw) return res.json({ ok: false, error: "Thiếu trường records" });
    const records = typeof raw === "string" ? JSON.parse(raw) : raw;
    if (!Array.isArray(records)) return res.json({ ok: false, error: "records phải là array" });

    const store = load();
    ensureShape(store);
    if (req.body.replace === "true" || req.body.replace === true) {
      store.hoa_don_dau_ra = [];
    }

    const now = new Date().toISOString();
    let maxId = store.hoa_don_dau_ra.reduce((m, x) => (x.id > m ? x.id : m), 0);
    const added = [];
    for (const r of records) {
      maxId++;
      added.push({
        id: maxId,
        congTy: r.congTy || "kh_moi",
        ngayHD: r.ngayHD || "",
        soHD: r.soHD || "",
        maKH: r.maKH || "",
        tenKhachHang: r.tenKhachHang || "",
        loaiHD: r.loaiHD || "khach-le",
        dienGiai: r.dienGiai || "",
        soTien: Number(r.soTien) || 0,
        soTienVAT: Number(r.soTienVAT) || 0,
        thueVAT: r.thueVAT || "",
        ghiChu: r.ghiChu || "",
        createdAt: now,
      });
    }
    store.hoa_don_dau_ra.push(...added);
    if (!store.seq) store.seq = {};
    store.seq.hoa_don_dau_ra = maxId;
    save(store);
    res.json({ ok: true, added: added.length, total: store.hoa_don_dau_ra.length });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

module.exports = router;
