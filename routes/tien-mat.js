const express = require("express");
const { load, save, nextId } = require("../store");
const { requireLogin, requireDataEntry } = require("../middleware/auth");

const router = express.Router();
router.use(requireLogin);

// GET /tien-mat
router.get("/tien-mat", (req, res) => {
  const { activeCompany, COMPANIES } = req;
  const q = (req.query.q || "").trim().toLowerCase();
  const loaiFilter = req.query.loai || "";
  const thangFilter = req.query.thang || "";

  const data = load();
  let rows = (data.tienMat || []).filter(r => r.company === activeCompany);

  if (q) {
    rows = rows.filter(r =>
      (r.diGiai || "").toLowerCase().includes(q) ||
      (r.nguon || "").toLowerCase().includes(q) ||
      (r.ghiChu || "").toLowerCase().includes(q)
    );
  }
  if (loaiFilter) rows = rows.filter(r => r.loai === loaiFilter);
  if (thangFilter) rows = rows.filter(r => (r.ngay || "").startsWith(thangFilter));

  // Sort mới nhất trước
  rows = rows.slice().sort((a, b) => (b.ngay || "").localeCompare(a.ngay || ""));

  res.render("tien-mat", {
    rows, q, loaiFilter, thangFilter,
    activeCompany, COMPANIES,
    success: req.query.success || "",
    error: req.query.error || "",
    currentPath: req.path,
  });
});

// POST /tien-mat/them
router.post("/tien-mat/them", requireDataEntry, (req, res) => {
  const { activeCompany } = req;
  const data = load();
  if (!data.tienMat) data.tienMat = [];

  const { loai, ngay, soTien, diGiai, nguon, ghiChu } = req.body;
  if (!loai || !ngay || !soTien) {
    return res.redirect("/tien-mat?error=Thiếu+thông+tin+bắt+buộc");
  }

  const id = nextId(data.tienMat);
  data.tienMat.push({
    id,
    company: activeCompany,
    loai,      // "Thu" hoặc "Chi"
    ngay,
    soTien: soTien.replace(/\./g, "").replace(/,/g, ""),
    diGiai: (diGiai || "").trim(),
    nguon: (nguon || "").trim(),
    ghiChu: (ghiChu || "").trim(),
    createdAt: new Date().toISOString(),
  });
  save(data);
  res.redirect("/tien-mat?success=Đã+thêm+bản+ghi");
});

// POST /tien-mat/:id/xoa
router.post("/tien-mat/:id/xoa", requireDataEntry, (req, res) => {
  const { activeCompany } = req;
  const id = parseInt(req.params.id, 10);
  const data = load();
  if (!data.tienMat) return res.redirect("/tien-mat");

  const idx = data.tienMat.findIndex(r => r.id === id && r.company === activeCompany);
  if (idx === -1) return res.redirect("/tien-mat?error=Không+tìm+thấy");

  data.tienMat.splice(idx, 1);
  save(data);
  res.redirect("/tien-mat?success=Đã+xóa");
});

module.exports = router;
