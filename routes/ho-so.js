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

// Bo so 0 dau de doi chieu so HD khong phu thuoc format (2072 = 00002072)
function normSoHD(s) {
  const str = String(s || "").trim();
  return str.replace(/^0+/, "") || str;
}

// Luyen, 2026-08-25: tu dong dien Ten NCC va Tong tien tu bang hoa_don_dau_vao
// (doi chieu theo so hoa don tu ten file). Luu ca key goc lan key bo so 0 dau.
function buildHddvLookup(store) {
  const lookup = {}; // soHoaDon (raw + normalized) -> { tenNCC, tongTien }
  (store.hoa_don_dau_vao || []).forEach((r) => {
    const no = String(r.soHoaDon || "").trim();
    if (!no) return;
    const noNorm = normSoHD(no);
    [no, noNorm].forEach((k) => {
      if (!lookup[k]) lookup[k] = { tenNCC: r.tenNCC || "", tongTien: 0, dienGiai: r.dienGiai || r.tenHangHoaMisa || "" };
      lookup[k].tongTien += r.soTien || 0;
      if (!lookup[k].dienGiai && (r.dienGiai || r.tenHangHoaMisa))
        lookup[k].dienGiai = r.dienGiai || r.tenHangHoaMisa || "";
    });
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
  const inv = hddvLookup && parsed.soHoaDon
    ? (hddvLookup[parsed.soHoaDon] || hddvLookup[normSoHD(parsed.soHoaDon)])
    : null;
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
    noiDung: r.noiDung || (inv ? inv.dienGiai : "") || "",
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

  // Nhom theo nccShort (ten ngan trong ten file) -- moi NCC 1 nhom du co
  // chi nhanh hay ten khac nhau. Ten day du NCC lay tu hoa don dau vao / ncc list.
  const groupMap = new Map();
  rows.forEach((r) => {
    const key = r.nccShort.toLowerCase();
    if (!groupMap.has(key)) {
      groupMap.set(key, { nccShort: r.nccShort, tenDayDuNCC: "", invoices: [], hopDong: [], chungTu: "" });
    }
    const g = groupMap.get(key);
    if (!g.tenDayDuNCC && r.tenDayDuNCC) g.tenDayDuNCC = r.tenDayDuNCC;
    g.invoices.push(r);
  });

  // Auto-match hop dong NCC tu phap_danh_hop_dong_ncc
  const hdNccList = store.phap_danh_hop_dong_ncc || [];
  const nccChungTuMap = store.ho_so_ncc_chungtu || {};
  groupMap.forEach((g, key) => {
    const shortNorm = normForMatch(g.nccShort);
    const fullNorm = normForMatch(g.tenDayDuNCC);
    g.hopDong = hdNccList.filter((hd) => {
      const hdShort = normForMatch(hd.tenNCC || "");
      const hdFull = normForMatch(hd.tenDayDuNCC || "");
      return (shortNorm.length >= 3 && (hdShort.includes(shortNorm) || shortNorm.includes(hdShort))) ||
             (fullNorm.length >= 5 && (hdFull.includes(fullNorm) || fullNorm.includes(hdFull)));
    });
    g.chungTu = (nccChungTuMap[key] || {}).chungTu || "";
    g.ghiChuNCC = (nccChungTuMap[key] || {}).ghiChu || "";
  });

  const groups = Array.from(groupMap.values()).sort((a, b) =>
    a.nccShort.toLowerCase().localeCompare(b.nccShort.toLowerCase())
  );

  res.render("ho-so-hoa-don", {
    userName: req.session.userName,
    rows,
    groups,
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

// Luu chung tu / ghi chu cap NCC (bien ban, bao gia, link, v.v.)
router.post("/ho-so/hoa-don-ncc/ncc/:nccKey/luu-chungtu", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_ncc_chungtu) store.ho_so_ncc_chungtu = {};
  const key = req.params.nccKey;
  store.ho_so_ncc_chungtu[key] = {
    chungTu: (req.body.chungTu || "").trim(),
    ghiChu: (req.body.ghiChu || "").trim(),
    updatedAt: new Date().toISOString(),
  };
  save(store);
  const qs = req.body.thang ? "?thang=" + encodeURIComponent(req.body.thang) : "";
  res.json({ ok: true });
});

// ─── Phí Cảng Phú Quốc ────────────────────────────────────────────────────
router.get("/ho-so/phi-cang-phu-quoc", (req, res) => {
  const store = load();
  const rows = (store.ho_so_phi_cang || []).slice().sort((a,b) => (b.ngay||'').localeCompare(a.ngay||''));
  res.render("ho-so-phi-cang", {
    userName: req.session.userName,
    rows,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/ho-so/phi-cang-phu-quoc/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  const { ngay, loaiPhi, soTien, linkFile, ghiChu } = req.body;
  store.ho_so_phi_cang.push({
    id: nextId(store),
    ngay: (ngay||"").trim(),
    loaiPhi: (loaiPhi||"").trim(),
    soTien: (soTien||"").trim(),
    linkFile: (linkFile||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+thêm");
});

router.post("/ho-so/phi-cang-phu-quoc/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  store.ho_so_phi_cang = (store.ho_so_phi_cang||[]).filter(r=>String(r.id)!==req.params.id);
  save(store);
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+xóa");
});

router.post("/ho-so/phi-cang-phu-quoc/:id/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_phi_cang||[]).find(x=>String(x.id)===req.params.id);
  if (r) { r.ngay=(req.body.ngay||'').trim(); r.loaiPhi=(req.body.loaiPhi||'').trim(); r.soTien=(req.body.soTien||'').trim(); r.linkFile=(req.body.linkFile||'').trim(); r.ghiChu=(req.body.ghiChu||'').trim(); save(store); }
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+lưu");
});

// ─── Doanh Thu Chia Sẻ ────────────────────────────────────────────────────
router.get("/ho-so/doanh-thu-chia-se", (req, res) => {
  const store = load();
  const rows = (store.ho_so_doanhthu_chiase || []).slice().sort((a,b) => (b.ngay||'').localeCompare(a.ngay||''));
  res.render("ho-so-doanhthu-chiase", {
    userName: req.session.userName,
    rows,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/ho-so/doanh-thu-chia-se/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_doanhthu_chiase) store.ho_so_doanhthu_chiase = [];
  const { ngay, gian, loai, soTien, linkFile, ghiChu } = req.body;
  store.ho_so_doanhthu_chiase.push({
    id: nextId(store),
    ngay: (ngay||"").trim(),
    gian: (gian||"").trim(),
    loai: (loai||"").trim(),
    soTien: (soTien||"").trim(),
    linkFile: (linkFile||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/doanh-thu-chia-se?success=Đã+thêm");
});

router.post("/ho-so/doanh-thu-chia-se/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  store.ho_so_doanhthu_chiase = (store.ho_so_doanhthu_chiase||[]).filter(r=>String(r.id)!==req.params.id);
  save(store);
  res.redirect("/ho-so/doanh-thu-chia-se?success=Đã+xóa");
});

router.post("/ho-so/doanh-thu-chia-se/:id/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_doanhthu_chiase||[]).find(x=>String(x.id)===req.params.id);
  if (r) { r.ngay=(req.body.ngay||'').trim(); r.gian=(req.body.gian||'').trim(); r.loai=(req.body.loai||'').trim(); r.soTien=(req.body.soTien||'').trim(); r.linkFile=(req.body.linkFile||'').trim(); r.ghiChu=(req.body.ghiChu||'').trim(); save(store); }
  res.redirect("/ho-so/doanh-thu-chia-se?success=Đã+lưu");
});

module.exports = router;
