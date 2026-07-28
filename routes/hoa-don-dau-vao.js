const express = require("express");
const XLSX = require("xlsx");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");
const { parseHoaDonDauVaoWorkbook } = require("../utils/hoaDonDauVaoParser");

const router = express.Router();
router.use(requireLogin);
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 80 * 1024 * 1024 } });

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
      kyHieuHD: "",
      dienGiai: "",
      // Chi Nhan, 2026-07-28: them tach rieng tien truoc thue/tien thue (tu
      // "Bảng kê hóa đơn hàng hóa dịch vụ mua vào chi tiết") -- soTien van la
      // TONG TIEN THANH TOAN (giu nguyen y nghia cu, cac dong nhap tay truoc
      // do khong bi anh huong).
      soTienTruocThue: 0,
      tienThue: 0,
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
    "Ký hiệu": r.kyHieuHD,
    "Số hóa đơn": r.soHoaDon,
    "Diễn giải": r.dienGiai,
    "Tiền trước thuế": r.soTienTruocThue,
    "Tiền thuế": r.tienThue,
    "Tổng tiền thanh toán": r.soTien,
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

function parseAmt(v) {
  return v ? Number(String(v).replace(/[^\d]/g, "")) : 0;
}

router.post("/hoa-don-dau-vao", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const { ngayHD, tenNCC, mstNCC, soHoaDon, kyHieuHD, dienGiai, soTienTruocThue, tienThue, soTien, linkHoaDon, ghiChu } = req.body;
    if (!ngayHD) throw new Error("Thiếu ngày hóa đơn.");
    if (!tenNCC || !tenNCC.trim()) throw new Error("Thiếu tên NCC.");
    store.hoa_don_dau_vao.push({
      id: nextId(store, "hoa_don_dau_vao_seq") || Date.now(),
      congTy: activeCompany,
      ngayHD,
      tenNCC: tenNCC.trim(),
      mstNCC: (mstNCC || "").trim(),
      soHoaDon: (soHoaDon || "").trim(),
      kyHieuHD: (kyHieuHD || "").trim(),
      dienGiai: (dienGiai || "").trim(),
      soTienTruocThue: parseAmt(soTienTruocThue),
      tienThue: parseAmt(tienThue),
      soTien: parseAmt(soTien),
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

// Chi Nhan, 2026-07-28: "thêm cho tôi chỗ úp file này luôn nhá" -- tai len
// thang file "Bảng kê hóa đơn hàng hóa, dịch vụ mua vào chi tiết" (xuat tu he
// thong hoa don dien tu).
//
// BUG suyt gay MAT DU LIEU (phat hien va sua truoc khi bao Chi Nhan): ban dau
// dung khoa upsert congTy+soHoaDon+kyHieuHD+dienGiai (KHONG co so tien) --
// nhung file thuc te CO 1 hoa don cung Ky hieu/So hoa don voi 2 dong CUNG
// dien giai (vd "NƯỚC UỐNG ĐÓNG CHAI LA VIE...") nhung so tien KHAC NHAU (2
// lo hang khac nhau cung 1 ten hang). Dung khoa khong co so tien se GOP 2
// dong nay lam 1 (dong sau ghi de dong truoc), MAT 1 dong tien that. Sua:
// them soTien vao khoa upsert (congTy+soHoaDon+kyHieuHD+dienGiai+soTien) --
// 2 dong noi dung giong het (ke ca so tien) moi duoc coi la "da co" (tai lai
// dung file cu se khong tao trung); dong nao so tien khac (kha nang la dong
// that khac, khong phai sua loi) se duoc THEM MOI thay vi ghi de, an toan
// hon (Chi Nhan tu xoa tay neu that su la 1 ban ghi trung do tai file 2 lan).
router.post("/hoa-don-dau-vao/upload", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { rows, skippedNoInvoiceNo } = parseHoaDonDauVaoWorkbook(req.file.buffer);
    if (rows.length === 0) {
      throw new Error("File không đọc được dòng hóa đơn nào.");
    }
    const existingByKey = new Map();
    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany) return;
      const key = [r.congTy, r.soHoaDon, r.kyHieuHD || "", r.dienGiai || "", r.soTien].join("|");
      existingByKey.set(key, r);
    });
    let added = 0;
    let updated = 0;
    rows.forEach((r) => {
      const key = [activeCompany, r.soHoaDon, r.kyHieuHD || "", r.dienGiai || "", r.soTien].join("|");
      const existing = existingByKey.get(key);
      if (existing) {
        // Da co dong y het (cung hoa don + dien giai + so tien) -- khong tao
        // trung, chi dam bao ngay/NCC dong bo (thuong khong doi).
        existing.ngayHD = r.ngayHD || existing.ngayHD;
        existing.tenNCC = r.tenNCC || existing.tenNCC;
        existing.mstNCC = r.mstNCC || existing.mstNCC;
        updated++;
        return;
      }
      const newRow = {
        id: nextId(store, "hoa_don_dau_vao_seq") || Date.now(),
        congTy: activeCompany,
        ngayHD: r.ngayHD,
        tenNCC: r.tenNCC,
        mstNCC: r.mstNCC,
        soHoaDon: r.soHoaDon,
        kyHieuHD: r.kyHieuHD,
        dienGiai: r.dienGiai,
        soTienTruocThue: r.soTienTruocThue,
        tienThue: r.tienThue,
        soTien: r.soTien,
        linkHoaDon: "",
        daHachToan: false,
        ghiChu: "",
        createdAt: new Date().toISOString(),
        source: `upload "${req.file.originalname}" ${new Date().toISOString().slice(0, 10)}`,
      };
      store.hoa_don_dau_vao.push(newRow);
      existingByKey.set(key, newRow);
      added++;
    });
    save(store);
    let msg = `Đã đọc "${req.file.originalname}": thêm ${added} dòng mới, cập nhật ${updated} dòng đã có.`;
    if (skippedNoInvoiceNo > 0) {
      msg += ` (Bỏ qua ${skippedNoInvoiceNo} dòng không có số hóa đơn -- dòng tổng cộng cuối file.)`;
    }
    res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent(msg));
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
