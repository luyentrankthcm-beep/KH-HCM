// Luyen, 2026-08-13: "thêm Danh mục trong Tổng hợp: Mã Công Trình, Mã Khách
// Hàng, Mã Nhà Cung Cấp" -- 3 trang CRUD đơn giản cho danh mục mã code.
const express = require("express");
const router = express.Router();
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");

// Moi trang dung 1 store key + 1 URL prefix + 1 view name -- cau hinh chung
const PAGES = {
  "ma-cong-trinh":   { storeKey: "danh_muc_ma_cong_trinh",   title: "Mã Công Trình",    colLabel: "Mã công trình" },
  "ma-khach-hang":   { storeKey: "danh_muc_ma_khach_hang",   title: "Mã Khách Hàng",    colLabel: "Mã khách hàng" },
  "ma-nha-cung-cap": { storeKey: "danh_muc_ma_nha_cung_cap", title: "Mã Nhà Cung Cấp",  colLabel: "Mã nhà cung cấp" },
};

function ensureList(store, key) {
  if (!Array.isArray(store[key])) store[key] = [];
}

// GET /danh-muc/:slug
router.get("/danh-muc/:slug", requireLogin, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const store = load();
  ensureList(store, page.storeKey);
  const searchQ = (req.query.q || "").trim().toLowerCase();
  let rows = store[page.storeKey];
  if (searchQ) {
    rows = rows.filter((r) =>
      [r.ma, r.ten, r.ghiChu].some((v) => v && String(v).toLowerCase().includes(searchQ))
    );
  }
  res.render("danh-muc", {
    slug: req.params.slug,
    page,
    rows,
    searchQ: req.query.q || "",
    success: req.query.success || "",
    error: req.query.error || "",
  });
});

// POST /danh-muc/:slug/them
router.post("/danh-muc/:slug/them", requireDataEntry, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const store = load();
  ensureList(store, page.storeKey);
  const ma = (req.body.ma || "").trim();
  const ten = (req.body.ten || "").trim();
  if (!ma) return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent("Mã không được để trống."));
  // Kiem tra trung ma
  const dup = store[page.storeKey].find((r) => r.ma.toLowerCase() === ma.toLowerCase());
  if (dup) return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent(`Mã "${ma}" đã tồn tại.`));
  store[page.storeKey].push({
    id: nextId(store),
    ma,
    ten: ten || "",
    ghiChu: (req.body.ghiChu || "").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect(`/danh-muc/${req.params.slug}?success=` + encodeURIComponent(`Đã thêm mã "${ma}".`));
});

// POST /danh-muc/:slug/sua/:id
router.post("/danh-muc/:slug/sua/:id", requireDataEntry, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const store = load();
  ensureList(store, page.storeKey);
  const id = Number(req.params.id);
  const row = store[page.storeKey].find((r) => r.id === id);
  if (!row) return res.redirect(`/danh-muc/${req.params.slug}?error=Không tìm thấy mục.`);
  const ma = (req.body.ma || "").trim();
  if (!ma) return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent("Mã không được để trống."));
  // Kiem tra trung ma (loai tru chinh no)
  const dup = store[page.storeKey].find((r) => r.id !== id && r.ma.toLowerCase() === ma.toLowerCase());
  if (dup) return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent(`Mã "${ma}" đã tồn tại.`));
  row.ma = ma;
  row.ten = (req.body.ten || "").trim();
  row.ghiChu = (req.body.ghiChu || "").trim();
  row.updatedAt = new Date().toISOString();
  save(store);
  res.redirect(`/danh-muc/${req.params.slug}?success=` + encodeURIComponent(`Đã cập nhật mã "${ma}".`));
});

// POST /danh-muc/:slug/xoa/:id
router.post("/danh-muc/:slug/xoa/:id", requireAdmin, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const store = load();
  ensureList(store, page.storeKey);
  const id = Number(req.params.id);
  store[page.storeKey] = store[page.storeKey].filter((r) => r.id !== id);
  save(store);
  res.redirect(`/danh-muc/${req.params.slug}?success=Đã xóa mục.`);
});

module.exports = router;
