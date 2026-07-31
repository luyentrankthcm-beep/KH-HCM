const express = require("express");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireDataEntry } = require("../middleware/auth");
const { getCompany, COMPANIES } = require("../utils/companies");
const { parseAmount } = require("../utils/parse");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan, 2026-07-31: "thêm 1 mục mới kế pháp danh là Tổng hợp cho tôi sổ
// xuống sẽ có các mục nhỏ mỗi mục là 1 trang cho tôi 1 là Lương 2 là theo dõi
// đầu vào" -- muc nav moi "Tong hop", 2 trang con: Luong (khung trong, nhap
// tay -- CHUA co du lieu luong nao trong he thong, giong cach lam voi Phap
// danh/Hoa Don Dau Vao/Chi Phi luc moi tao: bang trong + form them dong,
// Chi Nhan tu nhap roi minh dieu chinh cot sau) va Theo Doi Dau Vao (BAO
// CAO, khong phai nhap tay -- gop lai tu du lieu Hoa Don Dau Vao DA CO SAN,
// loc rieng hang hoa TK No 156/242 nhu Chi Nhan yeu cau).
function ensureShape(store) {
  if (!store.tong_hop_luong) store.tong_hop_luong = [];
}

function ensureLuongDefaults(row) {
  return Object.assign(
    {
      congTy: "kh_cu",
      thang: "",
      hoTen: "",
      chucVu: "",
      gian: "",
      luongCoBan: 0,
      phuCap: 0,
      thuong: 0,
      khauTru: 0,
      ghiChu: "",
    },
    row
  );
}

// ---------------- Luong (nhap tay, khung trong) ----------------

router.get("/tong-hop/luong", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  let rows = store.tong_hop_luong.map(ensureLuongDefaults).filter((r) => r.congTy === activeCompany);

  const monthSet = new Set();
  rows.forEach((r) => { if (r.thang) monthSet.add(r.thang); });
  const availableMonths = [...monthSet].sort().reverse();
  const thangFilter = req.query.thang !== undefined ? req.query.thang : (availableMonths[0] || "");
  if (thangFilter) rows = rows.filter((r) => r.thang === thangFilter);

  rows.sort((a, b) => (a.hoTen || "").localeCompare(b.hoTen || ""));
  rows.forEach((r) => {
    r.thucNhan = (r.luongCoBan || 0) + (r.phuCap || 0) + (r.thuong || 0) - (r.khauTru || 0);
  });
  const tongThucNhan = rows.reduce((s, r) => s + r.thucNhan, 0);

  res.render("tonghop-luong", {
    userName: req.session.userName,
    activeCompany,
    COMPANIES,
    rows,
    availableMonths,
    thangFilter,
    tongThucNhan,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/tong-hop/luong", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const { thang, hoTen, chucVu, gian, luongCoBan, phuCap, thuong, khauTru, ghiChu } = req.body;
    if (!thang) throw new Error("Thiếu tháng.");
    if (!hoTen || !hoTen.trim()) throw new Error("Thiếu họ tên.");
    store.tong_hop_luong.push({
      id: nextId(store, "tong_hop_luong_seq") || Date.now(),
      congTy: activeCompany,
      thang,
      hoTen: hoTen.trim(),
      chucVu: (chucVu || "").trim(),
      gian: (gian || "").trim(),
      luongCoBan: Math.abs(parseAmount(luongCoBan) || 0),
      phuCap: Math.abs(parseAmount(phuCap) || 0),
      thuong: Math.abs(parseAmount(thuong) || 0),
      khauTru: Math.abs(parseAmount(khauTru) || 0),
      ghiChu: (ghiChu || "").trim(),
      createdAt: new Date().toISOString(),
    });
    save(store);
    res.redirect("/tong-hop/luong?thang=" + encodeURIComponent(thang) + "&success=" + encodeURIComponent("Đã lưu lương."));
  } catch (e) {
    res.redirect("/tong-hop/luong?error=" + encodeURIComponent(e.message));
  }
});

router.post("/tong-hop/luong/:id/delete", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.tong_hop_luong.find((x) => String(x.id) === req.params.id);
  store.tong_hop_luong = store.tong_hop_luong.filter((x) => String(x.id) !== req.params.id);
  save(store);
  const qs = r && r.thang ? "?thang=" + encodeURIComponent(r.thang) + "&" : "?";
  res.redirect("/tong-hop/luong" + qs + "success=" + encodeURIComponent("Đã xóa dòng lương."));
});

router.get("/tong-hop/luong/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const thangFilter = req.query.thang || "";
  let rows = store.tong_hop_luong.map(ensureLuongDefaults).filter((r) => r.congTy === activeCompany);
  if (thangFilter) rows = rows.filter((r) => r.thang === thangFilter);
  rows.sort((a, b) => (a.hoTen || "").localeCompare(b.hoTen || ""));

  const exportRows = rows.map((r) => ({
    "Tháng": r.thang,
    "Họ tên": r.hoTen,
    "Chức vụ": r.chucVu,
    "Gian/Bộ phận": r.gian,
    "Lương cơ bản": r.luongCoBan,
    "Phụ cấp": r.phuCap,
    "Thưởng": r.thuong,
    "Khấu trừ": r.khauTru,
    "Thực nhận": (r.luongCoBan || 0) + (r.phuCap || 0) + (r.thuong || 0) - (r.khauTru || 0),
    "Ghi chú": r.ghiChu,
  }));
  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Luong");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=luong-${activeCompany}-${thangFilter || "TatCa"}.xlsx`);
  res.send(buf);
});

// ---------------- Theo doi dau vao: hang hoa TK No 156/242 ----------------
// Chi Nhan, 2026-07-31: "theo dỗi đầu vào sẽ có hàng hóa từ 156 với 242 bên
// hóa đơn đầu vào nhá" -- BAO CAO tu du lieu Hoa Don Dau Vao da co (truong
// taiKhoanNo, xem routes/hoa-don-dau-vao.js), CHI loc 2 TK Chi Nhan yeu cau
// (156 = hang hoa, 242 = chi phi tra truoc/cong cu phan bo), gom theo Ten
// hang hoa (Misa) trong thang dang xem.
router.get("/tong-hop/theo-doi-dau-vao", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const allRows = (store.hoa_don_dau_vao || []).filter(
    (r) => r.congTy === activeCompany && ["156", "242"].includes((r.taiKhoanNo || "").trim())
  );

  const monthSet = new Set();
  allRows.forEach((r) => { const m = (r.ngayHD || "").slice(0, 7); if (m) monthSet.add(m); });
  const availableMonths = [...monthSet].sort().reverse();
  const thangFilter = req.query.thang !== undefined ? req.query.thang : (availableMonths[0] || "");
  const rows = thangFilter ? allRows.filter((r) => (r.ngayHD || "").slice(0, 7) === thangFilter) : allRows;

  const tk156Total = rows.filter((r) => r.taiKhoanNo === "156").reduce((s, r) => s + (r.soTienTruocThue || 0), 0);
  const tk242Total = rows.filter((r) => r.taiKhoanNo === "242").reduce((s, r) => s + (r.soTienTruocThue || 0), 0);

  const byHangHoa = {};
  rows.forEach((r) => {
    const key = (r.tenHangHoaMisa || "").trim() || "(chưa có tên hàng hóa)";
    if (!byHangHoa[key]) byHangHoa[key] = { tenHangHoa: key, taiKhoanNo: r.taiKhoanNo, soTien: 0, soHoaDon: 0 };
    byHangHoa[key].soTien += r.soTienTruocThue || 0;
    byHangHoa[key].soHoaDon += 1;
  });
  const hangHoaRows = Object.values(byHangHoa).sort((a, b) => b.soTien - a.soTien);

  const detailRows = [...rows].sort((a, b) => (a.ngayHD < b.ngayHD ? 1 : -1));

  res.render("tonghop-theodoidauvao", {
    userName: req.session.userName,
    activeCompany,
    COMPANIES,
    availableMonths,
    thangFilter,
    tk156Total,
    tk242Total,
    hangHoaRows,
    detailRows,
  });
});

module.exports = router;
