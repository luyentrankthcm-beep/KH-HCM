// Luyen, 2026-08-13: "thêm Danh mục trong Tổng hợp: Mã Công Trình, Mã Khách
// Hàng, Mã Nhà Cung Cấp" -- 3 trang CRUD đơn giản cho danh mục mã code.
// Luyen, 2026-08-17: thêm Excel upload/sync (upsert by mã, xóa dòng bị bỏ).
const express = require("express");
const router = express.Router();
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// Seed du lieu ban dau tu seeds/danh-muc.json neu danh sach con trong.
// Luyen, 2026-08-17: khi deploy lan dau, 3 danh muc duoc tu dong nap tu file
// Excel da parse san (157 cong trinh, 164 KH, 907 NCC).
const SEED_FILE = require("path").join(__dirname, "../seeds/danh-muc.json");
const SEED_FILE_MOI = require("path").join(__dirname, "../seeds/danh-muc-moi.json");
const SEED_KEYS = {
  "ma-cong-trinh":   "ma_cong_trinh",
  "ma-khach-hang":   "ma_khach_hang",
  "ma-nha-cung-cap": "ma_nha_cung_cap",
};
const SEED_KEYS_MOI = {
  "ma-cong-trinh":   "ma_cong_trinh_moi",
  "ma-khach-hang":   "ma_khach_hang_moi",
  "ma-nha-cung-cap": "ma_nha_cung_cap_moi",
};

function seedIfEmpty(store, slug, company) {
  const page = PAGES[slug];
  if (!page) return;
  const storeKey = getStoreKey(page, company);
  ensureList(store, storeKey);
  if (store[storeKey].length > 0) return; // da co du lieu
  const isMoi = company === "kh_moi";
  let seedData;
  try {
    seedData = require(isMoi ? SEED_FILE_MOI : SEED_FILE);
  } catch (e) {
    return;
  }
  const seedKey = isMoi ? SEED_KEYS_MOI[slug] : SEED_KEYS[slug];
  const rows = seedData[seedKey] || [];
  rows.forEach((r) => {
    if (!r.ma) return;
    store[storeKey].push({
      id: nextId(store),
      ma: r.ma,
      ten: r.ten || "",
      ghiChu: "",
      createdAt: new Date().toISOString(),
      fromExcel: true,
    });
  });
  // Khi seed NCC moi thi dong bo luon vao chi_phi_ncc_list_moi
  if (isMoi && slug === "ma-nha-cung-cap") {
    if (!Array.isArray(store.chi_phi_ncc_list_moi)) store.chi_phi_ncc_list_moi = [];
    rows.forEach((r) => {
      if (!r.ma) return;
      if (!store.chi_phi_ncc_list_moi.find((x) => x.maNCC === r.ma)) {
        store.chi_phi_ncc_list_moi.push({ maNCC: r.ma, tenNCC: r.ten || "" });
      }
    });
  }
}

// Moi trang dung 1 base key, tach thanh _cu / _moi theo cong ty.
// Luyen, 2026-08-21: KH Cu va KH Moi la 2 cong ty khac nhau trong MISA nen
// ma cong trinh, ma KH, ma NCC cung khac nhau hoan toan.
const PAGES = {
  "ma-cong-trinh":   { baseKey: "danh_muc_ma_cong_trinh",   title: "Mã Công Trình",    colLabel: "Mã công trình" },
  "ma-khach-hang":   { baseKey: "danh_muc_ma_khach_hang",   title: "Mã Khách Hàng",    colLabel: "Mã khách hàng" },
  "ma-nha-cung-cap": { baseKey: "danh_muc_ma_nha_cung_cap", title: "Mã Nhà Cung Cấp",  colLabel: "Mã nhà cung cấp" },
};

// Tra ve store key theo cong ty (kh_cu / kh_moi)
function getStoreKey(page, company) {
  const suffix = company === "kh_moi" ? "_moi" : "_cu";
  return page.baseKey + suffix;
}

function ensureList(store, key) {
  if (!Array.isArray(store[key])) store[key] = [];
}

// GET /danh-muc/map-nh-gian -- danh sach rules keyword → gian tu dien gian
router.get("/danh-muc/map-nh-gian", requireLogin, (req, res) => {
  const store = load();
  const rules = store.description_gian_rules || [];
  res.render("danh-muc-map-nh-gian", {
    rules,
    success: req.query.success || "",
    error: req.query.error || "",
    userName: req.session.userName || "",
  });
});

// GET /danh-muc/:slug
router.get("/danh-muc/:slug", requireLogin, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const company = req.query.company === "kh_moi" ? "kh_moi" : "kh_cu";
  const storeKey = getStoreKey(page, company);
  const store = load();
  ensureList(store, storeKey);
  // Tu dong seed du lieu ban dau neu danh sach con trong (chi KH Cu)
  const beforeSeed = store[storeKey].length;
  seedIfEmpty(store, req.params.slug, company);
  if (store[storeKey].length > beforeSeed) save(store);
  const searchQ = (req.query.q || "").trim().toLowerCase();
  let rows = store[storeKey];
  if (searchQ) {
    rows = rows.filter((r) =>
      [r.ma, r.ten, r.ghiChu].some((v) => v && String(v).toLowerCase().includes(searchQ))
    );
  }
  const uploadMeta = ((store.danh_muc_upload_meta || {})[req.params.slug] || {})[company] || null;
  res.render("danh-muc", {
    slug: req.params.slug,
    page,
    company,
    rows,
    uploadMeta,
    searchQ: req.query.q || "",
    success: req.query.success || "",
    error: req.query.error || "",
    userName: req.session.userName || "",
  });
});

// POST /danh-muc/:slug/them
router.post("/danh-muc/:slug/them", requireDataEntry, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const company = req.body.company === "kh_moi" ? "kh_moi" : "kh_cu";
  const storeKey = getStoreKey(page, company);
  const store = load();
  ensureList(store, storeKey);
  const ma = (req.body.ma || "").trim();
  const ten = (req.body.ten || "").trim();
  if (!ma) return res.redirect(`/danh-muc/${req.params.slug}?company=${company}&error=` + encodeURIComponent("Mã không được để trống."));
  const dup = store[storeKey].find((r) => r.ma.toLowerCase() === ma.toLowerCase());
  if (dup) return res.redirect(`/danh-muc/${req.params.slug}?company=${company}&error=` + encodeURIComponent(`Mã "${ma}" đã tồn tại.`));
  store[storeKey].push({
    id: nextId(store),
    ma,
    ten: ten || "",
    ghiChu: (req.body.ghiChu || "").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect(`/danh-muc/${req.params.slug}?company=${company}&success=` + encodeURIComponent(`Đã thêm mã "${ma}".`));
});

// POST /danh-muc/:slug/sua/:id
router.post("/danh-muc/:slug/sua/:id", requireDataEntry, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const company = req.body.company === "kh_moi" ? "kh_moi" : "kh_cu";
  const storeKey = getStoreKey(page, company);
  const store = load();
  ensureList(store, storeKey);
  const id = Number(req.params.id);
  const row = store[storeKey].find((r) => r.id === id);
  if (!row) return res.redirect(`/danh-muc/${req.params.slug}?company=${company}&error=Không tìm thấy mục.`);
  const ma = (req.body.ma || "").trim();
  if (!ma) return res.redirect(`/danh-muc/${req.params.slug}?company=${company}&error=` + encodeURIComponent("Mã không được để trống."));
  const dup = store[storeKey].find((r) => r.id !== id && r.ma.toLowerCase() === ma.toLowerCase());
  if (dup) return res.redirect(`/danh-muc/${req.params.slug}?company=${company}&error=` + encodeURIComponent(`Mã "${ma}" đã tồn tại.`));
  row.ma = ma;
  row.ten = (req.body.ten || "").trim();
  row.ghiChu = (req.body.ghiChu || "").trim();
  row.updatedAt = new Date().toISOString();
  save(store);
  res.redirect(`/danh-muc/${req.params.slug}?company=${company}&success=` + encodeURIComponent(`Đã cập nhật mã "${ma}".`));
});

// POST /danh-muc/:slug/xoa/:id
router.post("/danh-muc/:slug/xoa/:id", requireAdmin, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const company = req.body.company === "kh_moi" ? "kh_moi" : "kh_cu";
  const storeKey = getStoreKey(page, company);
  const store = load();
  ensureList(store, storeKey);
  const id = Number(req.params.id);
  store[storeKey] = store[storeKey].filter((r) => r.id !== id);
  save(store);
  res.redirect(`/danh-muc/${req.params.slug}?company=${company}&success=Đã xóa mục.`);
});

// POST /danh-muc/:slug/upload-excel
// Parse Excel (header row 3, data from row 4): col[1]=Mã, col[2]=Tên.
// Sync: upsert by mã (case-insensitive), xóa dòng không còn trong file.
router.post("/danh-muc/:slug/upload-excel", requireDataEntry, upload.single("file"), (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  // company duoc gui qua hidden input trong form
  // (da duoc them vao view danh-muc.ejs)
  if (!req.file) {
    return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent("Chưa chọn file."));
  }
  let excelRows;
  try {
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    // aoa = array of arrays, bao gom ca hang trong
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    // Tim dong header: dong co "Mã" o col[1]
    let dataStart = 3; // mac dinh row 4 (index 3)
    for (let i = 0; i < Math.min(aoa.length, 6); i++) {
      const cell = String(aoa[i][1] || "").trim();
      if (cell.startsWith("Mã") || cell === "STT") {
        dataStart = i + 1;
        break;
      }
    }
    excelRows = [];
    for (let i = dataStart; i < aoa.length; i++) {
      const row = aoa[i];
      const ma = String(row[1] || "").trim();
      const ten = String(row[2] || "").trim();
      if (!ma) continue; // bo qua dong trong
      excelRows.push({ ma, ten });
    }
  } catch (e) {
    return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent("Lỗi đọc file Excel: " + e.message));
  }
  if (excelRows.length === 0) {
    return res.redirect(`/danh-muc/${req.params.slug}?error=` + encodeURIComponent("File không có dữ liệu hợp lệ."));
  }

  const company = req.body.company === "kh_moi" ? "kh_moi" : "kh_cu";
  const storeKey = getStoreKey(page, company);
  const store = load();
  ensureList(store, storeKey);
  const list = store[storeKey];

  // Tao map ma (lowercase) -> row de upsert nhanh
  const excelMap = new Map();
  excelRows.forEach((r) => excelMap.set(r.ma.toLowerCase(), r));

  let added = 0, updated = 0, deleted = 0;

  // Upsert: duyet qua tat ca excelRows
  excelMap.forEach((exRow, maKey) => {
    const existing = list.find((r) => r.ma.toLowerCase() === maKey);
    if (!existing) {
      list.push({
        id: nextId(store),
        ma: exRow.ma,
        ten: exRow.ten,
        ghiChu: "",
        createdAt: new Date().toISOString(),
        fromExcel: true,
      });
      added++;
    } else if (existing.ten !== exRow.ten) {
      existing.ten = exRow.ten;
      existing.updatedAt = new Date().toISOString();
      updated++;
    }
  });

  // Xoa: cac dong co fromExcel=true nhung khong con trong file moi
  const before = store[storeKey].length;
  store[storeKey] = store[storeKey].filter((r) => {
    if (!r.fromExcel) return true; // giu nguyen dong them tay
    return excelMap.has(r.ma.toLowerCase());
  });
  deleted = before - store[storeKey].length;

  // Luu lich su upload
  if (!store.danh_muc_upload_meta) store.danh_muc_upload_meta = {};
  if (!store.danh_muc_upload_meta[req.params.slug]) store.danh_muc_upload_meta[req.params.slug] = {};
  store.danh_muc_upload_meta[req.params.slug][company] = {
    uploaded_at: new Date().toISOString(),
    file_name: req.file.originalname,
    count: store[storeKey].length,
  };
  save(store);
  const label = company === "kh_moi" ? "KH Mới" : "KH Cũ";
  const msg = `[${label}] Đồng bộ xong: +${added} mới, ~${updated} cập nhật, -${deleted} xóa (tổng ${store[storeKey].length} mục).`;
  res.redirect(`/danh-muc/${req.params.slug}?company=${company}&success=` + encodeURIComponent(msg));
});

// POST /danh-muc/:slug/xoa-het -- xoa toan bo danh sach (chi dong fromExcel) theo cong ty
router.post("/danh-muc/:slug/xoa-het", requireAdmin, (req, res) => {
  const page = PAGES[req.params.slug];
  if (!page) return res.status(404).send("Không tìm thấy trang.");
  const company = req.body.company === "kh_moi" ? "kh_moi" : "kh_cu";
  const storeKey = getStoreKey(page, company);
  const store = load();
  ensureList(store, storeKey);
  const before = store[storeKey].length;
  store[storeKey] = store[storeKey].filter((r) => !r.fromExcel);
  // Xoa meta upload
  if (store.danh_muc_upload_meta && store.danh_muc_upload_meta[req.params.slug]) {
    delete store.danh_muc_upload_meta[req.params.slug][company];
  }
  save(store);
  const label = company === "kh_moi" ? "KH Mới" : "KH Cũ";
  const deleted = before - store[storeKey].length;
  res.redirect(`/danh-muc/${req.params.slug}?company=${company}&success=` + encodeURIComponent(`[${label}] Đã xóa ${deleted} mục từ Excel.`));
});

module.exports = router;
