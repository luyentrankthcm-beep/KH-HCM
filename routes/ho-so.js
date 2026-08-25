// Luyen, 2026-08-24: "thêm cho tôi 1 mục nữa là hồ sơ" -- trang Hồ Sơ Hóa
// Đơn NCC: tổng hợp hóa đơn đầu vào theo NCC, link thẳng về file PDF trên
// Google Drive. Mỗi hóa đơn có thêm: tên đầy đủ NCC, tổng tiền, nội dung
// tóm tắt, hồ sơ liên quan.
const express = require("express");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const driveApi = require("../utils/driveApi");

const uploadMem = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

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
    warn: req.query.warn || null,
  });
});

router.post("/ho-so/phi-cang-phu-quoc/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  const { ngay, soHoaDon, loaiPhi, soTien, linkFile, ghiChu } = req.body;
  store.ho_so_phi_cang.push({
    id: nextId(store),
    ngay: (ngay||"").trim(),
    soHoaDon: (soHoaDon||"").trim(),
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
  if (r) { r.ngay=(req.body.ngay||'').trim(); r.soHoaDon=(req.body.soHoaDon||'').trim(); r.loaiPhi=(req.body.loaiPhi||'').trim(); r.soTien=(req.body.soTien||'').trim(); r.linkFile=(req.body.linkFile||'').trim(); r.ghiChu=(req.body.ghiChu||'').trim(); save(store); }
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+lưu");
});

// ─── Doanh Thu Chia Sẻ ────────────────────────────────────────────────────
const DTCS_TABS = [
  { key: 'mall-giu-tien',      label: '🏦 Mall giữ tiền' },
  { key: 'tien-thue-vuot',     label: '📈 Tiền thuê vượt' },
  { key: 'tien-thue-co-dinh',  label: '📋 Tiền thuê chia cố định' },
];
router.get("/ho-so/doanh-thu-chia-se", (req, res) => {
  const store = load();
  const activeTab = DTCS_TABS.some(t => t.key === req.query.tab) ? req.query.tab : DTCS_TABS[0].key;
  const allRows = (store.ho_so_doanhthu_chiase || []).slice().sort((a,b) => (b.ngay||'').localeCompare(a.ngay||''));
  const rows = allRows.filter(r => (r.loaiTab || 'mall-giu-tien') === activeTab);
  res.render("ho-so-doanhthu-chiase", {
    userName: req.session.userName,
    rows,
    tabs: DTCS_TABS,
    activeTab,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/ho-so/doanh-thu-chia-se/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_doanhthu_chiase) store.ho_so_doanhthu_chiase = [];
  const { ngay, gian, loai, soTien, linkFile, ghiChu, loaiTab } = req.body;
  const tab = DTCS_TABS.some(t => t.key === loaiTab) ? loaiTab : DTCS_TABS[0].key;
  store.ho_so_doanhthu_chiase.push({
    id: nextId(store),
    ngay: (ngay||"").trim(),
    gian: (gian||"").trim(),
    loai: (loai||"").trim(),
    soTien: (soTien||"").trim(),
    linkFile: (linkFile||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    loaiTab: tab,
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/doanh-thu-chia-se?tab=" + tab + "&success=Đã+thêm");
});

router.post("/ho-so/doanh-thu-chia-se/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_doanhthu_chiase||[]).find(x=>String(x.id)===req.params.id);
  const tab = r ? (r.loaiTab || DTCS_TABS[0].key) : DTCS_TABS[0].key;
  store.ho_so_doanhthu_chiase = (store.ho_so_doanhthu_chiase||[]).filter(r=>String(r.id)!==req.params.id);
  save(store);
  res.redirect("/ho-so/doanh-thu-chia-se?tab=" + tab + "&success=Đã+xóa");
});

router.post("/ho-so/doanh-thu-chia-se/:id/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_doanhthu_chiase||[]).find(x=>String(x.id)===req.params.id);
  if (r) { r.ngay=(req.body.ngay||'').trim(); r.gian=(req.body.gian||'').trim(); r.loai=(req.body.loai||'').trim(); r.soTien=(req.body.soTien||'').trim(); r.linkFile=(req.body.linkFile||'').trim(); r.ghiChu=(req.body.ghiChu||'').trim(); save(store); }
  const tab = r ? (r.loaiTab || DTCS_TABS[0].key) : DTCS_TABS[0].key;
  res.redirect("/ho-so/doanh-thu-chia-se?tab=" + tab + "&success=Đã+lưu");
});

// ─── Đọc PDF biên lai → trả về fields để auto-fill form ─────────────────────
// Luyen, 2026-08-25: "tải file PDF, đọc nội dung, điền vào form" -- nhan PDF
// qua multer (memoryStorage), goi driveApi.extractPhiCangInfo, tra JSON.
// Khong can auth (chi doc, khong luu gi).
router.post("/ho-so/phi-cang-phu-quoc/doc-file", uploadMem.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Vui lòng chọn file PDF." });
    const mt = (req.file.mimetype || "").toLowerCase();
    if (!mt.includes("pdf")) return res.status(400).json({ error: "Chỉ hỗ trợ file PDF để đọc tự động." });
    const info = await driveApi.extractPhiCangInfo(req.file.buffer);
    return res.json({ success: true, ...info });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ─── Upload ảnh Phí Cảng → PDF → Google Drive ────────────────────────────────
// Luyen, 2026-08-25: "tải lên ảnh, lưu ảnh qua PDF đặt tên số HĐ + ngày, lưu
// vào Drive, điền thông tin cơ bản hiển thị trên web" -- nhan anh tu form,
// chuyen sang PDF bang pdf-lib, upload len folder "Phi Cang HK" tren Drive,
// luu record vao store.ho_so_phi_cang voi link Drive.
router.post("/ho-so/phi-cang-phu-quoc/upload-anh", requireDataEntry, uploadMem.array("anh", 10), async (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  if (!req.files || req.files.length === 0) {
    return res.redirect("/ho-so/phi-cang-phu-quoc?error=" + encodeURIComponent("Vui lòng chọn ít nhất 1 file."));
  }
  const { ngay, soHoaDon, loaiPhi, soTien, ghiChu } = req.body;
  const soHD = (soHoaDon || "").trim();
  // Dinh dang ten file: PhiCang_SoHD_DD-MM-YYYY.pdf
  const dateStr = ngay
    ? ngay.split("-").reverse().join("-") // YYYY-MM-DD → DD-MM-YYYY
    : new Date().toLocaleDateString("vi-VN").replace(/\//g, "-");

  const driveWarnings = [];
  let added = 0;

  for (let i = 0; i < req.files.length; i++) {
    const f = req.files[i];
    const suffix = req.files.length > 1 ? `_${i + 1}` : "";
    const fileName = `PhiCang_${soHD || "HoaDon"}${suffix}_${dateStr}.pdf`;
    let linkFile = "";
    let driveId = "";

    // Thu upload Drive; neu loi (chua ket noi OAuth) thi van luu record khong co link
    try {
      const isPdf = (f.mimetype || "").toLowerCase().includes("pdf");
      const pdfBuffer = isPdf ? f.buffer : await driveApi.imageToPdf(f.buffer, f.mimetype);
      const driveFile = await driveApi.uploadFileToDrive(
        store, driveApi.DRIVE_PHI_CANG_FOLDER_ID, fileName, pdfBuffer, "application/pdf"
      );
      linkFile = "https://drive.google.com/file/d/" + driveFile.id + "/view";
      driveId = driveFile.id;
    } catch (e) {
      driveWarnings.push("⚠️ Chưa kết nối Drive — đã lưu hồ sơ trên web nhưng chưa upload file lên Google Drive. Vào Chi Phí → Kết nối Gmail để cấp quyền.");
    }

    store.ho_so_phi_cang.push({
      id: nextId(store),
      ngay: ngay || "",
      soHoaDon: soHD,
      loaiPhi: (loaiPhi || "").trim(),
      soTien: (soTien || "").trim(),
      linkFile,
      driveId,
      ghiChu: (ghiChu || "").trim(),
      tenFile: fileName,
      createdAt: new Date().toISOString(),
    });
    added++;
  }

  save(store);
  let msg = "Đã lưu " + added + " hồ sơ" + (driveWarnings.length > 0 ? " (chưa có link Drive)." : " lên Drive.");
  if (driveWarnings.length > 0) {
    return res.redirect("/ho-so/phi-cang-phu-quoc?success=" + encodeURIComponent(msg) + "&warn=" + encodeURIComponent(driveWarnings[0]));
  }
  res.redirect("/ho-so/phi-cang-phu-quoc?success=" + encodeURIComponent(msg));
});

// ─── Sync Hóa Đơn NCC từ Google Drive ────────────────────────────────────────
// Luyen, 2026-08-25: "cập nhật hóa đơn hàng ngày từ gg drive" -- list PDF
// trong folder Drive, upsert vao store.ho_so_hoa_don (dedup theo driveId).
// Scope drive.readonly da duoc them vao buildAuthUrl trong gmailApi.js -- neu
// token cu chua co scope nay, Drive API tra 403, error message huong dan re-auth.
router.post("/ho-so/sync-drive-hoa-don", requireAdmin, async (req, res) => {
  const store = load();
  ensureHoSo(store);
  try {
    const files = await driveApi.listFolderFiles(store, driveApi.DRIVE_HOADON_FOLDER_ID, "application/pdf");
    const existingIds = new Set(store.ho_so_hoa_don.map((r) => r.driveId));
    let added = 0;
    files.forEach((f) => {
      if (!existingIds.has(f.id)) {
        store.ho_so_hoa_don.push({
          driveId: f.id,
          fileName: f.name,
          addedAt: new Date().toISOString(),
        });
        added++;
      }
    });
    save(store);
    res.redirect(
      "/ho-so/hoa-don-ncc?success=" +
        encodeURIComponent(
          "Đã sync Drive: thêm " + added + " file mới (tổng " + files.length + " file trong folder)."
        )
    );
  } catch (e) {
    res.redirect("/ho-so/hoa-don-ncc?error=" + encodeURIComponent(e.message));
  }
});

// ─── Sync Phí Cảng từ Google Drive ───────────────────────────────────────────
// Luyen, 2026-08-25: sync PDF tu folder "Phi Cang HK" Drive vao ho_so_phi_cang.
// Moi file tao 1 dong moi voi linkFile = link Drive, ten file lam loaiPhi tam,
// ngay/soTien de trong cho Luyen tu dien sau khi kiem tra.
router.post("/ho-so/phi-cang-phu-quoc/sync-drive", requireAdmin, async (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  try {
    const files = await driveApi.listFolderFiles(store, driveApi.DRIVE_PHI_CANG_FOLDER_ID, "application/pdf");
    const existingLinks = new Set(
      store.ho_so_phi_cang.map((r) => r.linkFile).filter(Boolean)
    );
    let added = 0;
    files.forEach((f) => {
      const link = "https://drive.google.com/file/d/" + f.id + "/view";
      if (!existingLinks.has(link)) {
        store.ho_so_phi_cang.push({
          id: nextId(store),
          ngay: "",
          loaiPhi: f.name.replace(/\.pdf$/i, ""),
          soTien: "",
          linkFile: link,
          ghiChu: "Sync từ Drive tự động",
          createdAt: new Date().toISOString(),
        });
        added++;
      }
    });
    save(store);
    res.redirect(
      "/ho-so/phi-cang-phu-quoc?success=" +
        encodeURIComponent(
          "Đã sync Drive: thêm " + added + " file mới (tổng " + files.length + " file trong folder)."
        )
    );
  } catch (e) {
    res.redirect("/ho-so/phi-cang-phu-quoc?error=" + encodeURIComponent(e.message));
  }
});

module.exports = router;
