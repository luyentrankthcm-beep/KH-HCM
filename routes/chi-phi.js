const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireDataEntry } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");
const {
  parseChiPhiSheetWorkbook,
  parseKvcMienBacWorkbook,
  parseKvcMienBacAutoWorkbook,
} = require("../utils/chiPhiSheetParser");
const {
  extractGianRentText,
  matchGianRecord,
  isDoanhThuChiaSeRecord,
  buildGianAliasIndex,
  findContractForGianText,
} = require("../utils/rentPaymentMatcher");
const gmailApi = require("../utils/gmailApi");
const gmailInvoiceMatcher = require("../utils/gmailInvoiceMatcher");
const { parseAmount } = require("../utils/parse");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

// Luyen, 2026-07-20: "thêm cho tôi 1 trang bên kh cũ và kh mới tên là Chi phí
// nhá" -- trang MOI HOAN TOAN, khac voi "Doi soat Chi phi" (doisoat-chiphi.js
// -- doi soat chi phi voi so tien tren sao ke ngan hang). Trang nay la SO GHI
// CHI PHI, tach rieng theo cong ty (giong Phap danh: 1 nut Cu/Moi o topbar).
//
// Luyen, 2026-07-20 (lan 2): "vao gg sheet ... chon sheet tháng 7.2026 với tt
// tiền mặt ... tách ra cho tôi số hóa đơn, ncc, gian ... link dẫn tới hóa đơn
// gốc" -- mo rong schema tu khung don gian ban dau (loaiChiPhi/ghiChu) sang
// cau truc chi tiet hon de chua du lieu import tu Google Sheet "ĐI ỦY NHIỆM
// CHI KVC + MTĐ MN" (tab THÁNG 7.2026 + TT TIỀN MẶT): gian, ncc, soHoaDon,
// dienGiai, linkHoaDon, trangThaiHoaDon (ghi chu rieng cho biet da doc duoc
// hoa don tu Drive hay can Luyen tu mo link kiem tra), nguon (biet dong nao
// tu tab nao). Van giu loaiChiPhi/ghiChu cho cac dong nhap tay truoc do.
//
// Luyen, 2026-07-20 (lan 3): "bỏ từ xóa này đi thay vào đó là tích chọn đã
// hạch toán ... với lại cập nhật số chứng từ liên quan ... với lại unc" -- bo
// nut Xoa moi dong (du lieu import khong nen tu y xoa tay), thay bang
// checkbox "daHachToan" (mac dinh false cho toan bo du lieu cu/moi import).
// Them 2 truong: "soUNC" (so uy nhiem chi, lay tu cot "Ngay can di tien" cua
// sheet -- cot nay bi dat ten sai, gia tri thuc te la ma UNC vd "UNC 01/7",
// KHONG phai ngay) va "soChungTuLienQuan" (voi cac dong tien thue gian ma
// chua co so hoa don ro rang, tra cuu Phap Danh Hop Dong Thue Gian Hang theo
// ten gian de lay so Hop Dong lam chung tu tham chieu).
//
// Luyen, 2026-07-20 (lan 4): "thêm cái lọc theo tháng ... tt tiền mặt nhiều
// quá bạn mở link bên cạnh xem hóa đơn nào ngày nào" -- sua lai truong `ngay`
// cho 153 dong THANG 7.2026 (truoc do bi gan nham gia tri "UNC 01/7"/"PAYMENT
// 30/6" thay vi ngay that -- da fix bang script rieng, lay tu ngayCanDiTien).
// Voi 105 dong TT TIEN MAT: 26 dong da doc duoc tu Drive OCR thi lay ngay hoa
// don tu ocr_results; 77 dong con lai "can mo link" thi mo tung link (Chrome)
// de lay so hoa don + ngay hoa don that, cap nhat truc tiep vao chi_phi (viec
// nay lam bang script rieng, khong qua route). Them bo loc "thang" (theo
// prefix YYYY-MM cua truong ngay) o day, tuong tu bo loc hachToan.
function ensureShape(store) {
  if (!store.chi_phi) store.chi_phi = [];
}

function ensureChiPhiDefaults(row) {
  return Object.assign(
    {
      congTy: "kh_cu",
      // Luyen, 2026-07-23: "chia cho tôi thành 2 trang 1 trang Miền Nam và 1
      // trang miền Bắc" -- them chieu moi "mien" ("nam"/"bac"), doc lap voi
      // congTy (Cu/Moi). TAT CA du lieu hien co (781 dong, tu sheet "ĐI ỦY
      // NHIỆM CHI KVC + MTĐ MN") deu la Mien Nam -- Luyen xac nhan qua
      // AskUserQuestion ("Đúng vậy") -- nen mac dinh "nam" cho moi dong
      // KHONG co san truong nay (du lieu cu), giong cach congTy mac dinh
      // "kh_cu" o tren.
      mien: "nam",
      ngay: "",
      gian: "",
      ncc: "",
      soHoaDon: "",
      soUNC: "",
      soChungTuLienQuan: "",
      dienGiai: "",
      loaiChiPhi: "",
      soTien: 0,
      soTienHoaDonGoc: null,
      linkHoaDon: "",
      trangThaiHoaDon: "",
      daHachToan: false,
      nguon: "",
      ghiChu: "",
    },
    row
  );
}

// "mien-nam"/"mien-bac" (doan URL) <-> "nam"/"bac" (gia tri luu trong du
// lieu). Dung 1 regex rang buoc ngay tren route (":mien(mien-nam|mien-bac)")
// de Express tu chan 404 neu ai go sai doan URL, khong can validate tay o
// tung route ben duoi.
function mienFromSeg(seg) {
  return seg === "mien-bac" ? "bac" : "nam";
}
function mienSeg(mien) {
  return mien === "bac" ? "mien-bac" : "mien-nam";
}

// Luyen, 2026-07-23: "thêm cho tôi 1 cột tài khoản ... doanh thu chia sẻ gian
// thì đưa vô 1388 còn lại thì để 131 ... dựa vào hợp đồng thuê gian á của
// miền nam trước nhá chi phí của miền nam cả 2 kh" -- moi dong Chi Phi duoc
// gan TK 1388 khi ten gian (r.gian) khop CHAC CHAN (qua matchGianRecord, cung
// nguong tin cay >=0.6 da dung o quet-ngan-hang) voi 1 hop dong trong "Hop
// Dong Thue Gian Hang" (store.phap_danh_hop_dong_thue) duoc danh dau la
// "doanh thu chia se" (xem isDoanhThuChiaSeRecord). Con lai (khong khop duoc
// gian nao, hoac khop nhung khong phai doanh thu chia se) -> mac dinh TK 131.
// TINH TAI THOI DIEM XEM/XUAT (khong luu vao dong Chi Phi) de tu dong cap
// nhat khi hop dong thay doi/them moi, khong can chay lai import gi ca.
function computeTaiKhoanChiPhi(r, gianList, aliasIndex) {
  if (!r.gian || !r.gian.trim()) return "131";
  const rec = findContractForGianText(r.gian, gianList, aliasIndex);
  return rec && isDoanhThuChiaSeRecord(rec) ? "1388" : "131";
}

// Luyen, 2026-07-23: link cu "/chi-phi" (khong co doan mien) -- giu lai de
// khong hong cac link/bookmark cu (vd "/chi-phi?thang=2026-07&hoaDon=1" da
// dung truoc do), chuyen thang ve "/chi-phi/mien-nam" (moi du lieu cu deu la
// Mien Nam) va giu nguyen toan bo querystring dang co.
router.get("/chi-phi", (req, res) => {
  const qIdx = req.originalUrl.indexOf("?");
  const qs = qIdx >= 0 ? req.originalUrl.slice(qIdx) : "";
  res.redirect("/chi-phi/mien-nam" + qs);
});

router.get("/chi-phi/:mien(mien-nam|mien-bac)", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const mien = mienFromSeg(req.params.mien);
  const hachToanFilter = req.query.hachToan || ""; // "" = tat ca, "1" = da hach toan, "0" = chua hach toan
  // Luyen, 2026-07-21: "thêm chỗ lọc chưa có số hóa đơn" -- loc theo con thieu
  // soHoaDon hay khong, giup tim nhanh cac dong con can dien so hoa don (thu
  // cong hoac tra cuu them) thay vi phai doc het danh sach.
  const hoaDonFilter = req.query.hoaDon || ""; // "" = tat ca, "1" = da co so HD, "0" = chua co so HD
  let rows = store.chi_phi.map(ensureChiPhiDefaults).filter((r) => r.congTy === activeCompany && r.mien === mien);
  const totalForCompany = rows.length;
  const chuaHachToanCount = rows.filter((r) => !r.daHachToan).length;
  const chuaSoHoaDonCount = rows.filter((r) => !(r.soHoaDon || "").trim()).length;

  // Danh sach cac thang co du lieu (YYYY-MM), moi nhat truoc, de do vao <select>.
  const monthSet = new Set();
  rows.forEach((r) => { const m = (r.ngay || "").slice(0, 7); if (m) monthSet.add(m); });
  const availableMonths = [...monthSet].sort().reverse();
  // Mac dinh loc thang 7/2026 truoc ("loc trước tháng 7 cho tôi trước nhá"),
  // neu thang do khong co du lieu cho cong ty dang chon thi hien tat ca.
  const defaultMonth = availableMonths.includes("2026-07") ? "2026-07" : "";
  const thangFilter = req.query.thang !== undefined ? req.query.thang : defaultMonth;

  if (hachToanFilter === "1") rows = rows.filter((r) => r.daHachToan);
  else if (hachToanFilter === "0") rows = rows.filter((r) => !r.daHachToan);
  if (thangFilter) rows = rows.filter((r) => (r.ngay || "").slice(0, 7) === thangFilter);
  if (hoaDonFilter === "1") rows = rows.filter((r) => (r.soHoaDon || "").trim());
  else if (hoaDonFilter === "0") rows = rows.filter((r) => !(r.soHoaDon || "").trim());

  rows.sort((a, b) => (a.ngay < b.ngay ? 1 : -1));
  const tongTien = rows.reduce((s, r) => s + (r.soTien || 0), 0);
  const gianListForTaiKhoan = store.phap_danh_hop_dong_thue || [];
  const aliasIndexForTaiKhoan = buildGianAliasIndex(gianListForTaiKhoan);
  rows.forEach((r) => { r.taiKhoan = computeTaiKhoanChiPhi(r, gianListForTaiKhoan, aliasIndexForTaiKhoan); });
  res.render("chi-phi", {
    userName: req.session.userName,
    mien,
    mienLabel: mien === "bac" ? "Miền Bắc" : "Miền Nam",
    rows,
    totalForCompany,
    chuaHachToanCount,
    chuaSoHoaDonCount,
    hachToanFilter,
    thangFilter,
    hoaDonFilter,
    availableMonths,
    tongTien,
    error: req.query.error || null,
    success: req.query.success || null,
    gmailConfigured: gmailApi.isConfigured(),
    gmailConnected: !!(store.gmail_oauth && store.gmail_oauth.refresh_token),
  });
});

// Chi Nhan, 2026-07-22: "thêm chỗ xuất ra excel nhá" -- xuat danh sach DANG
// XEM (theo cong ty + bo loc hach toan/thang/so hoa don dang chon tren man
// hinh, giong cach lam voi trang Hop Dong Thue Gian Hang) ra file Excel.
router.get("/chi-phi/:mien(mien-nam|mien-bac)/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const mien = mienFromSeg(req.params.mien);
  const hachToanFilter = req.query.hachToan || "";
  const hoaDonFilter = req.query.hoaDon || "";
  const thangFilter = req.query.thang !== undefined ? req.query.thang : "";
  let rows = store.chi_phi.map(ensureChiPhiDefaults).filter((r) => r.congTy === activeCompany && r.mien === mien);
  if (hachToanFilter === "1") rows = rows.filter((r) => r.daHachToan);
  else if (hachToanFilter === "0") rows = rows.filter((r) => !r.daHachToan);
  if (thangFilter) rows = rows.filter((r) => (r.ngay || "").slice(0, 7) === thangFilter);
  if (hoaDonFilter === "1") rows = rows.filter((r) => (r.soHoaDon || "").trim());
  else if (hoaDonFilter === "0") rows = rows.filter((r) => !(r.soHoaDon || "").trim());
  rows.sort((a, b) => (a.ngay < b.ngay ? 1 : -1));

  const gianListForTaiKhoanExport = store.phap_danh_hop_dong_thue || [];
  const aliasIndexForTaiKhoanExport = buildGianAliasIndex(gianListForTaiKhoanExport);
  const exportRows = rows.map((r) => ({
    "Ngày chi": r.ngay,
    "Gian/Cơ sở": r.gian,
    "Tài khoản": computeTaiKhoanChiPhi(r, gianListForTaiKhoanExport, aliasIndexForTaiKhoanExport),
    NCC: r.ncc,
    "Số hóa đơn": r.soHoaDon,
    "Số UNC": r.soUNC,
    "Số chứng từ liên quan": r.soChungTuLienQuan,
    "Loại chi phí": r.loaiChiPhi,
    "Diễn giải": r.dienGiai,
    "Số tiền": r.soTien,
    "Link hóa đơn": r.linkHoaDon,
    "Trạng thái hóa đơn": r.trangThaiHoaDon,
    "Đã hạch toán": r.daHachToan ? "Có" : "Không",
    "Ghi chú": r.ghiChu,
    Nguồn: r.nguon,
  }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Chi phi");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=chi-phi-${activeCompany}-${req.params.mien}.xlsx`);
  res.send(buf);
});

// Luyen, 2026-07-22: "cập nhật thì giữ nguyên trang, đừng đưa lên đầu trang" --
// khi request la AJAX (goi tu JS trong chi-phi.ejs, khong phai form submit
// thuong) thi tra ve JSON thay vi redirect, de trang KHONG reload/scroll len
// dau va bo loc dang chon KHONG bi doi -- dong duoc sua van nam nguyen tai
// cho chi den khi chi tu tay doi bo loc.
function isAjaxChiPhiRequest(req) {
  return req.get("X-Requested-With") === "XMLHttpRequest";
}

router.post("/chi-phi/:id/hach-toan", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.chi_phi.find((x) => String(x.id) === req.params.id);
  if (r) {
    r.daHachToan = req.body.daHachToan === "1";
    save(store);
  }
  if (isAjaxChiPhiRequest(req)) {
    if (!r) return res.status(404).json({ error: "Không tìm thấy khoản chi này." });
    return res.json({ success: true, daHachToan: r.daHachToan });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật trạng thái hạch toán."));
  res.redirect("/chi-phi/" + mienSeg(r ? r.mien : req.body.mien) + "?" + qs.join("&"));
});

// Luyen, 2026-07-21: "thêm số hóa đơn thủ công được đi" -- nhieu dong (nhat la
// TT TIEN MAT bi chan captcha/khong mo duoc link) van chua co so hoa don tu
// dong tra ra duoc, can Luyen tu dien tay sau khi tra cuu/xem hoa don giay.
// Cho sua truc tiep tu bang danh sach, khong can vao form rieng.
router.post("/chi-phi/:id/so-hoa-don", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.chi_phi.find((x) => String(x.id) === req.params.id);
  if (!r) {
    if (isAjaxChiPhiRequest(req)) return res.status(404).json({ error: "Không tìm thấy khoản chi này." });
    const qs = [];
    if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
    if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
    if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
    qs.push("error=" + encodeURIComponent("Không tìm thấy khoản chi này."));
    return res.redirect("/chi-phi/" + mienSeg(req.body.mien) + "?" + qs.join("&"));
  }
  r.soHoaDon = (req.body.soHoaDon || "").trim();
  // Ghi nhan la da dien tay, xoa cac ghi chu canh bao cu (vd "can mo link
  // kiem tra") vi Luyen da tu xu ly xong dong nay.
  //
  // Chi Nhan (2026-07-22): "hóa đơn xăng dầu mà sao lại ra cái này" -- dong
  // id 239 bi gan nham so hoa don/link cua cong ty khac (canh bao mismatch tu
  // luc import). Neu chi XOA TRANG so hoa don (thay vi dien so moi) VA dong
  // dang co canh bao mismatch cu, coi day la "chi dang don lai du lieu sai",
  // xoa luon canh bao + so tien hoa don goc di kem (khong con y nghia gi nua
  // khi da xoa so hoa don), thay bang ghi chu ngan gon nhac chi dien lai.
  if (!r.soHoaDon && (r.trangThaiHoaDon || "").includes("CẢNH BÁO")) {
    r.trangThaiHoaDon = "Đã xóa số hóa đơn cũ (bị gán nhầm) -- cần điền lại số hóa đơn đúng.";
    r.soTienHoaDonGoc = null;
  } else {
    r.trangThaiHoaDon = r.soHoaDon ? "Đã điền số hóa đơn thủ công" : r.trangThaiHoaDon;
  }
  save(store);
  if (isAjaxChiPhiRequest(req)) {
    return res.json({ success: true, soHoaDon: r.soHoaDon, trangThaiHoaDon: r.trangThaiHoaDon || "" });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật số hóa đơn."));
  res.redirect("/chi-phi/" + mienSeg(r.mien) + "?" + qs.join("&"));
});

// Chi Nhan (2026-07-23): "có công ty ncc mà đưa vô cho tôi nhá" -- cac dong
// Chi Phi tao tu dong qua "quet ngan hang" (/chi-phi/quet-ngan-hang) CHI co
// duoc ngay/dien giai/so tien tu giao dich ngan hang -- KHONG co ten NCC
// (thong tin nay chi co trong sheet "DI UY NHIEM CHI" goc, cot "Ten don vi
// thu huong", khong nam trong sao ke ngan hang) nen luon bi de trong. Truoc
// gio khong co cach dien NCC qua web -- them route sua truc tiep, cung 1
// kieu voi so-hoa-don/link-hoa-don/so-tien o tren.
router.post("/chi-phi/:id/ncc", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.chi_phi.find((x) => String(x.id) === req.params.id);
  if (!r) {
    if (isAjaxChiPhiRequest(req)) return res.status(404).json({ error: "Không tìm thấy khoản chi này." });
    const qs = [];
    if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
    if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
    if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
    qs.push("error=" + encodeURIComponent("Không tìm thấy khoản chi này."));
    return res.redirect("/chi-phi/" + mienSeg(req.body.mien) + "?" + qs.join("&"));
  }
  r.ncc = (req.body.ncc || "").trim();
  save(store);
  if (isAjaxChiPhiRequest(req)) {
    return res.json({ success: true, ncc: r.ncc });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật NCC."));
  res.redirect("/chi-phi/" + mienSeg(r.mien) + "?" + qs.join("&"));
});

// Chi Nhan (2026-07-23): "số tiền có 8tr mấy mà bạn lấy lên mất tỷ dữ vậy" --
// dong Chi Phi tao tu dong qua "quet ngan hang" lay dung so tien cua giao
// dich ngan hang goc (co the bi nhap sai tu luc nhap giao dich); truoc gio
// khong co cach sua so tien 1 dong Chi Phi da co san. Them route sua truc
// tiep tu bang danh sach, cung 1 kieu voi so-hoa-don/link-hoa-don o tren.
router.post("/chi-phi/:id/so-tien", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.chi_phi.find((x) => String(x.id) === req.params.id);
  if (!r) {
    if (isAjaxChiPhiRequest(req)) return res.status(404).json({ error: "Không tìm thấy khoản chi này." });
    const qs = [];
    if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
    if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
    if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
    qs.push("error=" + encodeURIComponent("Không tìm thấy khoản chi này."));
    return res.redirect("/chi-phi/" + mienSeg(req.body.mien) + "?" + qs.join("&"));
  }
  const parsedAmount = Math.abs(parseAmount(req.body.soTien));
  if (isNaN(parsedAmount)) {
    const msg = "Số tiền không hợp lệ.";
    if (isAjaxChiPhiRequest(req)) return res.status(400).json({ error: msg });
    return res.redirect("/chi-phi/" + mienSeg(r.mien) + "?error=" + encodeURIComponent(msg));
  }
  r.soTien = parsedAmount;
  // Neu dong dang co canh bao mismatch/bat thuong cu (vd tu quet-ngan-hang khi
  // so tien qua lon so voi cac thang truoc), coi viec chi tu sua so tien la da
  // xu ly xong, xoa canh bao cu di tranh hien thi lac long.
  if ((r.trangThaiHoaDon || "").includes("CẢNH BÁO") && (r.trangThaiHoaDon || "").includes("bất thường")) {
    r.trangThaiHoaDon = "Đã sửa số tiền (trước đó bị cảnh báo bất thường).";
  }
  save(store);
  if (isAjaxChiPhiRequest(req)) {
    return res.json({ success: true, soTien: r.soTien, trangThaiHoaDon: r.trangThaiHoaDon || "" });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật số tiền."));
  res.redirect("/chi-phi/" + mienSeg(r.mien) + "?" + qs.join("&"));
});

// Luyen, 2026-07-22: "tìm hóa đơn qua Gmail lưu Drive rồi gán lên đây" -- can
// route rieng de cap nhat link hoa don (link Google Drive) cho 1 dong Chi Phi
// da co san, khong dung chung voi /so-hoa-don (chi sua so hoa don thu cong).
router.post("/chi-phi/:id/link-hoa-don", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.chi_phi.find((x) => String(x.id) === req.params.id);
  if (!r) {
    if (isAjaxChiPhiRequest(req)) return res.status(404).json({ error: "Không tìm thấy khoản chi này." });
    const qs = [];
    if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
    if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
    if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
    qs.push("error=" + encodeURIComponent("Không tìm thấy khoản chi này."));
    return res.redirect("/chi-phi/" + mienSeg(req.body.mien) + "?" + qs.join("&"));
  }
  r.linkHoaDon = (req.body.linkHoaDon || "").trim();
  // Chi Nhan (2026-07-22): tuong tu /so-hoa-don -- neu chi XOA TRANG link (dang
  // don dep 1 mismatch cu) thi xoa luon canh bao + so tien hoa don goc di kem.
  if (!r.linkHoaDon && (r.trangThaiHoaDon || "").includes("CẢNH BÁO")) {
    r.trangThaiHoaDon = "Đã xóa link hóa đơn cũ (bị gán nhầm) -- cần dán lại link đúng.";
    r.soTienHoaDonGoc = null;
  }
  save(store);
  if (isAjaxChiPhiRequest(req)) {
    return res.json({ success: true, linkHoaDon: r.linkHoaDon, trangThaiHoaDon: r.trangThaiHoaDon || "" });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật link hóa đơn."));
  res.redirect("/chi-phi/" + mienSeg(r.mien) + "?" + qs.join("&"));
});

router.post("/chi-phi/:mien(mien-nam|mien-bac)", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const mien = mienFromSeg(req.params.mien);
  try {
    const { ngay, gian, ncc, soHoaDon, soUNC, soChungTuLienQuan, dienGiai, loaiChiPhi, soTien, linkHoaDon, ghiChu } = req.body;
    if (!ngay) throw new Error("Thiếu ngày chi.");
    if (!loaiChiPhi || !loaiChiPhi.trim()) throw new Error("Thiếu loại chi phí.");
    const amt = soTien ? Number(String(soTien).replace(/[^\d]/g, "")) : 0;
    store.chi_phi.push({
      id: nextId(store, "chi_phi_seq") || Date.now(),
      congTy: activeCompany,
      mien,
      ngay,
      gian: (gian || "").trim(),
      ncc: (ncc || "").trim(),
      soHoaDon: (soHoaDon || "").trim(),
      soUNC: (soUNC || "").trim(),
      soChungTuLienQuan: (soChungTuLienQuan || "").trim(),
      dienGiai: (dienGiai || "").trim(),
      loaiChiPhi: loaiChiPhi.trim(),
      soTien: amt,
      soTienHoaDonGoc: null,
      linkHoaDon: (linkHoaDon || "").trim(),
      trangThaiHoaDon: "",
      daHachToan: false,
      nguon: "nhap tay",
      ghiChu: (ghiChu || "").trim(),
      createdAt: new Date().toISOString(),
    });
    save(store);
    res.redirect("/chi-phi/" + req.params.mien + "?success=" + encodeURIComponent("Đã lưu khoản chi."));
  } catch (e) {
    res.redirect("/chi-phi/" + req.params.mien + "?error=" + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-21: "thêm cho tôi ở trên có nút cập nhật ... hình icon rồi có
// chữ phía dưới như nút bấm" -- nut "Cap nhat chi phi" tren trang, up lai
// chinh file Google Sheet "ĐI ỦY NHIỆM CHI KVC + MTĐ MN" (xuat lai tu Sheet)
// de nap them du lieu ky/thang moi, KHONG can lam script rieng nhu truoc gio.
// Tu do cac sheet "Tháng N.YYYY" bang regex (utils/chiPhiSheetParser.js).
//
// BUG suyt gay nhan doi hang tram dong khi test truoc khi upload that (da
// phat hien va rollback): ban dau khoa upsert co dung "gian" -- nhung ban ghi
// CU (thang 1-6, import truoc do bang script rieng) luu THANG THUAN cot A tho
// (vd "POSH"/"JP", nhan nhom chung chung), trong khi parser MOI (theo yeu cau
// 2026-07-21 cua Luyen) tach ten gian that tu dien giai (vd "Lotte Nam Sai
// Gòn") cho CUNG dong do -- 2 gia tri gian khac nhau lam khoa khong khop,
// tuong la dong moi nhung thuc ra la dong CU, suyt nhan doi ~400 dong. Sua:
// khoa upsert KHONG dung "gian" nua (dung dien giai + ncc + so tien + ngay +
// cong ty -- deu la du lieu goc, khong doi theo logic tach gian), on dinh qua
// moi lan nang cap logic tach gian sau nay. Rieng gian: dong da co ma dang
// luu nhan nhom chung chung (posh/JP/...) thi NANG CAP len ten that neu lan
// nay tach duoc (cai thien du lieu cu), dong da co gian cu the roi thi giu
// nguyen (khong ghi de). Cac truong Luyen tu dien tay (soHoaDon/daHachToan/
// ghiChu/soChungTuLienQuan/linkHoaDon) LUON giu nguyen, khong bao gio ghi de.
//
// Rieng tab "TT TIỀN MẶT" co cau truc cot khac han (khong co ngay/NCC/cong ty
// ro rang, truoc gio phai mo tung link tra cuu/OCR bang tay) -- KHONG tu dong
// parse o day de tranh nhap sai/trung 105 dong da xu ly thu cong, chi bao cho
// Luyen biet da bo qua tab nay.
// Luyen, 2026-07-23: them khoa "mien" (nam/bac) de khong lan du lieu Mien
// Nam/Mien Bac neu ngau nhien trung congTy+ngay+ncc+soTien+dienGiai (2 nguon
// KHAC HAN nhau nen kha nang trung that su rat thap, nhung van nen tach cho
// chac). QUAN TRONG: dung "(r.mien || 'nam')" (KHONG phai r.mien tho) --
// 781 dong Mien Nam co san TRUOC KHI co truong "mien" nay hoan toan KHONG co
// field nay trong du lieu da luu (chi duoc gan mac dinh "nam" luc DOC qua
// ensureChiPhiDefaults, khong phai luc luu) -- neu dung r.mien tho o day, khoa
// cua 781 dong cu se la "undefined|..." trong khi dong MOI parse tu sheet
// Mien Nam co r.mien = "nam" ro rang, 2 khoa se KHONG khop nhau va gay nhan
// doi toan bo 781 dong o lan "Cập nhật chi phí" tiep theo.
// mienOverride: cac dong VUA PARSE tu sheet (mr.rows, chua duoc luu) KHONG TU
// mang san field ".mien" (chi duoc gan luc tao newRow ben duoi) -- neu chi
// dung "r.mien || 'nam'" o day cho ca 2 phia (dong luu roi VA dong vua parse),
// dong Mien Bac vua parse se luon bi tinh nham thanh khoa "nam" (vi r.mien
// luon undefined tren dong parse), khong bao gio khop voi khoa "bac" cua
// chinh no sau khi da luu -- gay nhap TRUNG LAP moi lan bam "Cập nhật chi
// phí" tiep theo (da phat hien qua dry-run truoc khi dung that, xem
// applyChiPhiMonthRowsToStore ben duoi). Truyen rieng mienOverride khi tinh
// khoa cho dong MOI VUA PARSE (dong DA LUU thi khong can, tu no da co r.mien
// dung roi hoac mac dinh dung "nam" cho du lieu cu).
function chiPhiRowKey(r, mienOverride) {
  return [(r.mien || mienOverride || "nam"), r.congTy, r.ngay, (r.ncc || "").trim(), r.soTien, (r.dienGiai || "").trim()].join("|");
}

// Cac gian la nhan nhom chung chung tu cot A (khong phai ten gian that) --
// dung de nhan biet dong CU nao nen duoc "nang cap" gian khi khop lai.
const GENERIC_GIAN_LABELS = new Set(["posh", "jp", "posh+jp", "jp+posh", ""]);
function isUpgradableGian(gianText) {
  const n = (gianText || "").trim().toLowerCase();
  return GENERIC_GIAN_LABELS.has(n);
}

// Luyen, 2026-07-21 (lan 3): tach logic ap dung monthRows vao store thanh ham
// dung chung cho CA route upload file VA route moi doc thang tu Google Sheet
// link (xem ben duoi) -- "cập nhật chi phí nữa từ gg sheet". Hanh vi upsert
// GIU NGUYEN nhu truoc (chi them dong MOI, khong dong tay dong da co -- Luyen
// xac nhan rieng "không cập nhật cái mới thôi cái cũ vẫn giữ nguyên") --
// doi o day CHI la nguon lay du lieu (file vs link), khong doi cach xu ly.
function applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel, mien = "nam") {
  const existingByKey = new Map();
  store.chi_phi.forEach((r) => existingByKey.set(chiPhiRowKey(r), r));
  let added = 0;
  let addedNoGian = 0;
  let upgradedGian = 0;
  const monthSummary = [];
  for (const mr of monthRows) {
    let addedThisSheet = 0;
    for (const r of mr.rows) {
      const key = chiPhiRowKey(r, mien);
      const existingRow = existingByKey.get(key);
      if (existingRow) {
        // Dong da co roi -- KHONG dong tay, chi nang cap Gian tu nhan nhom
        // chung chung len ten that neu lan nay tach duoc cu the hon.
        if (r.gian && isUpgradableGian(existingRow.gian) && !isUpgradableGian(r.gian)) {
          existingRow.gian = r.gian;
          upgradedGian++;
        }
        continue;
      }
      const newRow = {
        id: nextId(store, "chi_phi_seq") || Date.now(),
        congTy: r.congTy,
        // Luyen, 2026-07-23: gio ham nay dung chung cho CA 2 nguon (sheet
        // "ĐI ỦY NHIỆM CHI KVC + MTĐ MN" = Mien Nam, VA sheet "Tạo lệnh UNC
        // KVC MB" = Mien Bac) -- nguoi goi (router.post upload/cap-nhat-tu-sheet
        // ben duoi) truyen dung tham so mien vao, KHONG con hardcode "nam" o day.
        mien,
        ngay: r.ngay,
        gian: r.gian,
        ncc: r.ncc,
        soHoaDon: r.soHoaDon || "",
        soUNC: r.soUNC || "",
        soChungTuLienQuan: "",
        dienGiai: r.dienGiai,
        loaiChiPhi: "",
        soTien: r.soTien,
        soTienHoaDonGoc: null,
        linkHoaDon: "",
        trangThaiHoaDon: r.gian ? "" : "Chưa xác định Gian (tự động) -- cần điền tay",
        daHachToan: false,
        nguon: mr.sheetName,
        ghiChu: "",
        createdAt: new Date().toISOString(),
        source: sourceLabel,
      };
      store.chi_phi.push(newRow);
      existingByKey.set(key, newRow);
      added++;
      addedThisSheet++;
      if (!r.gian) addedNoGian++;
    }
    if (addedThisSheet > 0) monthSummary.push(`${mr.sheetName}: +${addedThisSheet}`);
  }
  return { added, addedNoGian, upgradedGian, monthSummary };
}

// Luyen, 2026-08-01: "gg sheet có chi phí mới ... cập nhật hết của tháng 7
// cho tôi đi 2 sheet TT tiền mặt với lại Tháng 7 2026 á" -- tab "TT TIỀN MẶT"
// khong co cot Ngay/NCC (xem ghi chu o utils/chiPhiSheetParser.js) nen KHONG
// dung chiPhiRowKey (can r.ngay) duoc -- upsert rieng, uu tien khoa theo
// linkHoaDon (gan nhu luon duy nhat 1-1 voi 1 hoa don that, khop dung voi ca
// 105 dong da import bang tay/OCR truoc do dang co san linkHoaDon), dong
// khong co link (vd ghi "chưa có hóa đơn") moi fallback ve nguoiMua+coSo+soTien.
// Luyen xac nhan (AskUserQuestion, 2026-08-01): "nhập phần biết được, còn
// không biết thì như cũ đưa lên trước đã" -- nap San Nguoi mua/Co so/So
// tien/Link hoa don, DE TRONG Ngay+NCC+So hoa don cho Luyen tu dien sau khi
// mo link (khong doan bua).
function ttTienMatRowKey(r) {
  return r.linkHoaDon ? `link|${r.linkHoaDon}` : `nolink|${r.nguoiMua}|${r.coSo}|${r.soTien}`;
}
function applyTtTienMatRowsToStore(store, ttRows, sourceLabel) {
  const existingByKey = new Map();
  store.chi_phi.forEach((r) => {
    if (r.nguon === "TT TIỀN MẶT") existingByKey.set(ttTienMatRowKey({ linkHoaDon: r.linkHoaDon, nguoiMua: (r.dienGiai || "").trim(), coSo: r.gian, soTien: r.soTien }), r);
  });
  let added = 0;
  for (const r of ttRows) {
    const key = ttTienMatRowKey(r);
    if (existingByKey.has(key)) continue;
    const dienGiai = `Chi tiền mặt - người mua: ${r.nguoiMua} - cơ sở: ${r.coSo}`;
    const newRow = {
      id: nextId(store, "chi_phi_seq") || Date.now(),
      congTy: "kh_cu",
      mien: "nam",
      ngay: "",
      gian: r.coSo,
      ncc: "",
      soHoaDon: "",
      soUNC: "",
      soChungTuLienQuan: "",
      dienGiai,
      loaiChiPhi: "Chi tiền mặt",
      soTien: r.soTien,
      soTienHoaDonGoc: null,
      linkHoaDon: r.linkHoaDon,
      trangThaiHoaDon: r.linkHoaDon
        ? "Cần mở link kiểm tra (chưa tự đọc được) -- điền Ngày + NCC"
        : r.linkRawText || "Chưa có hóa đơn -- điền Ngày + NCC",
      daHachToan: false,
      nguon: "TT TIỀN MẶT",
      ghiChu: "",
      createdAt: new Date().toISOString(),
      source: sourceLabel,
    };
    store.chi_phi.push(newRow);
    existingByKey.set(key, newRow);
    added++;
  }
  return { added };
}

function buildChiPhiResultMessage(prefix, result, skippedSheets, ttResult) {
  let msg = `${prefix} thêm ${result.added} khoản chi mới${
    result.monthSummary.length ? " (" + result.monthSummary.join(", ") + ")" : ""
  }.`;
  if (result.upgradedGian > 0) {
    msg += ` Đã nâng cấp Gian (từ nhãn nhóm chung chung như "posh"/"JP" sang tên gian cụ thể) cho ${result.upgradedGian} dòng đã có sẵn.`;
  }
  if (result.addedNoGian > 0) {
    msg += ` CẢNH BÁO: ${result.addedNoGian} dòng không tự xác định được Gian (diễn giải không có cụm "ghế ...") -- cần chị tự điền tay, lọc theo cột Gian trống.`;
  }
  // Luyen, 2026-08-01: tab "TT TIỀN MẶT" gio DA duoc tu dong nap (Nguoi mua/Co
  // so/So tien/Link hoa don), khong con nam trong skippedSheets nua -- them
  // dong rieng bao so dong moi + nhac lai can tu dien Ngay+NCC qua link.
  if (ttResult && ttResult.added > 0) {
    msg += ` TT TIỀN MẶT: thêm ${ttResult.added} khoản chi mới (đã có Cơ sở/Số tiền/Link hóa đơn, CẦN chị tự mở link điền Ngày + NCC + Số hóa đơn -- lọc theo Nguồn "TT TIỀN MẶT" và Ngày trống).`;
  }
  if (skippedSheets.length > 0) {
    msg += ` Đã bỏ qua sheet không khớp cấu trúc nào: ${skippedSheets.join(", ")}.`;
  }
  return msg;
}

router.post("/chi-phi/upload", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { monthRows, ttTienMatRows, skippedSheets } = parseChiPhiSheetWorkbook(req.file.buffer);
    if (monthRows.length === 0 && (!ttTienMatRows || ttTienMatRows.length === 0)) {
      throw new Error(
        `File không có sheet nào khớp tên dạng "Tháng N.YYYY" hay tab "TT TIỀN MẶT" (đã thấy: ${skippedSheets.join(", ") || "(không có sheet)"}).`
      );
    }
    const sourceLabel = `upload "${req.file.originalname}" ${new Date().toISOString().slice(0, 10)}`;
    const result = applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel, "nam");
    const ttResult = applyTtTienMatRowsToStore(store, ttTienMatRows || [], sourceLabel);
    save(store);
    const msg = buildChiPhiResultMessage(`Đã nạp "${req.file.originalname}":`, result, skippedSheets, ttResult);
    // Sheet nay luon la du lieu Mien Nam (xem applyChiPhiMonthRowsToStore) --
    // ve thang Mien Nam de thay ngay du lieu vua nap, bat ke dang bam nut tu
    // trang Mien Nam hay Mien Bac.
    res.redirect("/chi-phi/mien-nam?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/chi-phi/mien-nam?error=" + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-21 (lan 3): "cập nhật chi phí nữa từ gg sheet" -- doc THANG
// tu Google Sheet (khong can tai file ve roi tai len nua), giong nut Cap nhat
// hop dong NCC. Dung export?format=xlsx (khong phai format=csv nhu NCC) vi
// parser nay can CA workbook nhieu sheet (tu do sheet "Tháng N.YYYY" bang ten)
// VA thong tin merged-cell (cot Gian forward-fill) -- ca 2 thu nay CSV export
// (chi 1 tab, khong merge) khong co duoc, chi export?format=xlsx (toan bo file,
// giu nguyen dinh dang .xlsx that) moi dam bao dung.
const CHI_PHI_SHEET_XLSX_URL =
  process.env.CHI_PHI_SHEET_XLSX_URL ||
  "https://docs.google.com/spreadsheets/d/17PdHu2ji8y1sO1KK6H-BOJTXhsbcSziEG9hHRd9-iL0/export?format=xlsx";

// Luyen, 2026-07-23: 2 nguon rieng cho Mien Bac -- (1) file "Tạo lệnh UNC KVC
// MB" (Luyen: "link này là mảng kvc miền bắc có cột kh cũ kh mới", nhap tay,
// 1 tab lien tuc), (2) file "TẠO LỆNH UNC MTĐ MB 2025" (Luyen: "đây là của
// máy tự động ... chia ra cho tôi 2 miền kh mới cũ gian tên nhà cung cấp hay
// các tiền thuê hay số hóa đơn giống như cách lấy của miền nam nhá" -- nguon
// xuat tu dong, tuong duong vai tro voi file MTĐ MN cua Mien Nam nhung cau
// truc cot khac han). Ca 2 deu doc va gop chung 1 lan bam "Cập nhật chi phí"
// (giong Mien Nam gop nhieu sheet "Tháng N.YYYY" trong 1 file).
const KVC_MB_SHEET_XLSX_URL =
  process.env.KVC_MB_SHEET_XLSX_URL ||
  "https://docs.google.com/spreadsheets/d/12QOccXdRPnhb3JmRXUfdbszHioMaumsz/export?format=xlsx";
const MTD_MB_AUTO_SHEET_XLSX_URL =
  process.env.MTD_MB_AUTO_SHEET_XLSX_URL ||
  "https://docs.google.com/spreadsheets/d/132gDgtxG_3X-WlkS4LLksLsuYd_7YRYNeUQoflsy2Ic/export?format=xlsx";

router.post("/chi-phi/:mien(mien-nam|mien-bac)/cap-nhat-tu-sheet", requireDataEntry, async (req, res) => {
  const store = load();
  ensureShape(store);
  const mien = mienFromSeg(req.params.mien);
  try {
    if (mien === "nam") {
      const resp = await fetch(CHI_PHI_SHEET_XLSX_URL);
      if (!resp.ok) {
        throw new Error(
          `Không đọc được Google Sheet (mã lỗi ${resp.status}). Kiểm tra lại sheet đã chia sẻ "Bất kỳ ai có link đều xem được" chưa, hoặc link có bị đổi không.`
        );
      }
      const buf = Buffer.from(await resp.arrayBuffer());
      const { monthRows, ttTienMatRows, skippedSheets } = parseChiPhiSheetWorkbook(buf);
      if (monthRows.length === 0 && (!ttTienMatRows || ttTienMatRows.length === 0)) {
        throw new Error(
          `Google Sheet không có sheet nào khớp tên dạng "Tháng N.YYYY" hay tab "TT TIỀN MẶT" (đã thấy: ${skippedSheets.join(", ") || "(không có sheet)"}).`
        );
      }
      const sourceLabel = "GG Sheet " + new Date().toISOString().slice(0, 10);
      const result = applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel, "nam");
      const ttResult = applyTtTienMatRowsToStore(store, ttTienMatRows || [], sourceLabel);
      save(store);
      const msg = buildChiPhiResultMessage("Đã đọc thẳng từ Google Sheet:", result, skippedSheets, ttResult);
      res.redirect("/chi-phi/mien-nam?success=" + encodeURIComponent(msg));
    } else {
      const monthRows = [];
      let totalUnclassified = 0;
      const partLabels = [];

      const [respManual, respAuto] = await Promise.all([
        fetch(KVC_MB_SHEET_XLSX_URL),
        fetch(MTD_MB_AUTO_SHEET_XLSX_URL),
      ]);
      if (!respManual.ok && !respAuto.ok) {
        throw new Error(
          `Không đọc được cả 2 Google Sheet Miền Bắc (mã lỗi ${respManual.status}/${respAuto.status}). Kiểm tra lại đã chia sẻ "Bất kỳ ai có link đều xem được" chưa, hoặc link có bị đổi không.`
        );
      }
      if (respManual.ok) {
        const buf = Buffer.from(await respManual.arrayBuffer());
        const { sheetName, rows, unclassifiedCount } = parseKvcMienBacWorkbook(buf);
        if (sheetName) {
          monthRows.push({ sheetName: `UNC KVC MB (${sheetName})`, rows });
          totalUnclassified += unclassifiedCount;
          partLabels.push(`UNC KVC MB: +${rows.length}`);
        }
      }
      if (respAuto.ok) {
        const buf = Buffer.from(await respAuto.arrayBuffer());
        const { sheetName, rows, unclassifiedCount } = parseKvcMienBacAutoWorkbook(buf);
        if (sheetName) {
          monthRows.push({ sheetName: `MTĐ MB tự động (${sheetName})`, rows });
          totalUnclassified += unclassifiedCount;
          partLabels.push(`MTĐ MB tự động: +${rows.length}`);
        }
      }
      if (monthRows.length === 0) {
        throw new Error(
          `Không tìm thấy tab nào khớp đúng cấu trúc mong đợi trong cả 2 Google Sheet Miền Bắc (kiểm tra lại tên cột chưa bị đổi).`
        );
      }

      const sourceLabel = "GG Sheet Mien Bac " + new Date().toISOString().slice(0, 10);
      const result = applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel, "bac");
      save(store);
      let msg = `Đã đọc thẳng từ Google Sheet Miền Bắc (${partLabels.join(", ")} dòng khớp cấu trúc): thêm ${result.added} khoản chi mới.`;
      if (result.addedNoGian > 0) {
        msg += ` ${result.addedNoGian} dòng không tự xác định được Gian -- cần chị tự điền tay.`;
      }
      if (totalUnclassified > 0) {
        msg += ` CẢNH BÁO: bỏ qua ${totalUnclassified} dòng không xác định được KH Cũ/Mới -- không tự đoán để tránh gán sai công ty.`;
      }
      if (!respManual.ok || !respAuto.ok) {
        msg += ` (Lưu ý: chỉ đọc được 1/2 sheet lần này -- sheet ${!respManual.ok ? "UNC KVC MB" : "MTĐ MB tự động"} lỗi mã ${!respManual.ok ? respManual.status : respAuto.status}.)`;
      }
      res.redirect("/chi-phi/mien-bac?success=" + encodeURIComponent(msg));
    }
  } catch (e) {
    res.redirect("/chi-phi/" + mienSeg(mien) + "?error=" + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-21: "từ cái nội dung với 2 chi 2 ngân hàng này á bạn sẽ liên
// kết qua cái bên file chi phí á có thanh toán kh cũ và kh mới của gian nào á
// note lại chi ngày mấy ngân hàng nào bên chi phí cho tôi nhá cập nhật hàng
// ngày cho tôi luôn nhá" -- quet giao dich "chi" (tien thue gian) tren TAT CA
// tai khoan ngan hang dang theo doi (Luyen xac nhan qua AskUserQuestion), tu
// dong tao dong Chi Phi tuong ung (Luyen xac nhan "tu tao dong moi"). Pham vi
// thang: CHI THANG HIEN TAI (Luyen xac nhan "tháng 7 thôi" khi duoc hoi, vi
// Chi Phi hien chi dang theo doi tu thang 1.2026 -- khong lui ve ca nam 2025
// de tranh nhoi hang loat dong ngoai pham vi dang dung). Dedup bang truong
// rieng "bankTxId" (id giao dich nguon) luu tren moi dong Chi Phi da tao --
// chay lai (hang ngay) se tu bo qua giao dich da xu ly, khong tao trung.
router.post("/chi-phi/:mien(mien-nam|mien-bac)/quet-ngan-hang", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const mien = mienFromSeg(req.params.mien);
  try {
    const now = new Date();
    const monthPrefix = now.toISOString().slice(0, 7); // "YYYY-MM"
    const banksById = {};
    store.banks.forEach((b) => (banksById[b.id] = b));

    const existingBankTxIds = new Set(
      store.chi_phi.filter((r) => r.bankTxId).map((r) => r.bankTxId)
    );

    // Luyen, 2026-07-23: "chia cho tôi thành 2 trang 1 trang Miền Nam và 1
    // trang miền Bắc ... liên kết link với ngân hàng tôi sẽ liên kết sau" --
    // chi quet cac tai khoan ngan hang DA duoc gan dung mien nay (b.mien, xem
    // routes/banks.js -- mac dinh "nam" cho tai khoan cu). Mien Bac hien CHUA
    // co tai khoan nao gan "bac" nen quet o trang Mien Bac se tra ve 0 giao
    // dich, cho toi khi Luyen tu them/gan ngan hang Mien Bac ben trang Ngan
    // hang.
    const chiTx = store.transactions.filter((t) => {
      if (t.type !== "chi" || !t.date || t.date.slice(0, 7) !== monthPrefix) return false;
      if (existingBankTxIds.has(t.id)) return false;
      const bank = banksById[t.bank_id];
      if (!bank) return false;
      return (bank.mien || "nam") === mien;
    });

    let added = 0;
    let flaggedSuspicious = 0;
    const unmatched = [];
    const AMOUNT_SUSPICIOUS_THRESHOLD = 200000000; // 200 trieu -- nguong canh bao, khong chan tao dong

    chiTx.forEach((t) => {
      const bank = banksById[t.bank_id];
      if (!bank) return;
      const gianText = extractGianRentText(t.description);
      if (!gianText) return; // khong nhan dien duoc mau "tien thue ..." -- bo qua lang le (qua nhieu loai GD khac: luong, phi NH, mua hang NCC...)
      const rec = matchGianRecord(gianText, store.phap_danh_hop_dong_thue);
      if (!rec) {
        unmatched.push({ bank: bank.name, date: t.date, amount: t.amount, gianText, description: t.description });
        return;
      }
      const congTy = bank.company || rec.congTy;
      const suspicious = t.amount >= AMOUNT_SUSPICIOUS_THRESHOLD;
      if (suspicious) flaggedSuspicious++;
      store.chi_phi.push({
        id: nextId(store, "chi_phi_seq") || Date.now(),
        congTy,
        mien,
        ngay: t.date,
        gian: rec.gian,
        ncc: "",
        soHoaDon: "",
        soUNC: "",
        soChungTuLienQuan: "",
        dienGiai: t.description,
        loaiChiPhi: "Tiền thuê gian hàng",
        soTien: t.amount,
        soTienHoaDonGoc: null,
        linkHoaDon: "",
        trangThaiHoaDon: suspicious
          ? `CẢNH BÁO: số tiền ${t.amount.toLocaleString("vi-VN")}đ bất thường so với các tháng trước -- chị kiểm tra lại sao kê gốc.`
          : "",
        daHachToan: false,
        nguon: `Ngân hàng ${bank.name}`,
        ghiChu: `Tự động liên kết từ giao dịch ngân hàng "${bank.name}" ngày ${t.date} (mã GD #${t.id}), khớp gian qua diễn giải "${gianText}".`,
        createdAt: new Date().toISOString(),
        source: `auto bank-link ${bank.name} ${new Date().toISOString().slice(0, 10)}`,
        bankTxId: t.id,
      });
      added++;
    });

    save(store);

    const mienLabel = mien === "bac" ? "Miền Bắc" : "Miền Nam";
    let msg = `Đã quét ${chiTx.length} giao dịch Chi tháng ${monthPrefix} (chưa xử lý) trên các tài khoản ${mienLabel}. Thêm ${added} khoản chi phí thuê gian mới.`;
    if (chiTx.length === 0 && mien === "bac") {
      msg += ` (Chưa có tài khoản ngân hàng nào gán "Miền Bắc" -- vào trang Ngân hàng để gán khi chị liên kết ngân hàng Miền Bắc.)`;
    }
    if (flaggedSuspicious > 0) {
      msg += ` CẢNH BÁO: ${flaggedSuspicious} khoản có số tiền bất thường (>= 200 triệu) -- chị xem cột Trạng thái/ghi chú.`;
    }
    if (unmatched.length > 0) {
      msg += ` ${unmatched.length} giao dịch nhận diện là "tiền thuê" nhưng KHÔNG khớp được gian nào rõ ràng -- chị tự thêm tay: ${unmatched
        .map((u) => `${u.bank} ${u.date} ${u.amount.toLocaleString("vi-VN")}đ ("${u.gianText}")`)
        .join("; ")}.`;
    }
    res.redirect("/chi-phi/" + mienSeg(mien) + "?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/chi-phi/" + mienSeg(mien) + "?error=" + encodeURIComponent(e.message));
  }
});

// Nhan, 2026-07-22: "thêm cho tôi 1 nút cập nhật tìm hóa đơn ... tìm trên
// gmail giống như vậy ... tìm cái thời gian mới nhất với lại các hóa đơn chưa
// có thôi không cần tìm cái cũ đâu nhá" -- nut that trong web, tu dong tim qua
// Gmail API (utils/gmailApi.js + utils/gmailInvoiceMatcher.js) thay vi Claude
// tra bang tay. Pham vi CHINH XAC theo yeu cau: (1) chi thang MOI NHAT dang co
// du lieu cho cong ty dang chon (khong lui ve cac thang cu da xu ly roi), (2)
// chi cac dong dang THIEU link hoa don (bo qua dong da co linkHoaDon, du co
// hay khong co soHoaDon rieng -- tranh tim lai nhung dong da xong).
//
// Gioi han so dong xu ly 1 lan bam (MAX_ROWS_PER_RUN) vi moi dong can vai lan
// goi Gmail API (search + doc tung email ung vien) -- xu ly qua nhieu dong 1
// luc de bi timeout giua chung (dung nhu da gap khi Claude chay script tay
// truoc do). Bam nhieu lan se tiep tuc xu ly cac dong con lai (da xong roi thi
// tu dong bi bo qua o lan bam sau, vi luc do da co linkHoaDon).
const GMAIL_AUTO_MAX_ROWS_PER_RUN = 25;

router.post("/chi-phi/:mien(mien-nam|mien-bac)/tim-hoa-don-gmail", requireDataEntry, async (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const mien = mienFromSeg(req.params.mien);
  try {
    if (!gmailApi.isConfigured()) {
      throw new Error(
        "Chưa cấu hình Google OAuth trên Railway (thiếu GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET/GOOGLE_REDIRECT_URI)."
      );
    }
    if (!store.gmail_oauth || !store.gmail_oauth.refresh_token) {
      throw new Error("Chưa kết nối Gmail -- bấm nút \"Kết nối Gmail\" trước.");
    }

    let rows = store.chi_phi.map(ensureChiPhiDefaults).filter((r) => r.congTy === activeCompany && r.mien === mien);
    const monthSet = new Set();
    rows.forEach((r) => {
      const m = (r.ngay || "").slice(0, 7);
      if (m) monthSet.add(m);
    });
    const newestMonth = [...monthSet].sort().reverse()[0] || "";

    // Chi Nhan (2026-07-22): "các dữ liệu tôi cập nhật rồi lưu ... để khi tôi
    // nhấn cập nhật cái lại mất phải làm lại" -- truoc day chi bo qua dong da
    // co linkHoaDon, nen dong nao chi da tu dien tay soHoaDon (nhung chua dan
    // linkHoaDon) van bi nut "Cap nhat tim hoa don" quet lai va GHI DE so hoa
    // don tim duoc tu Gmail len tren, mat du lieu chi da luu. Sua: bo qua ca
    // dong da co soHoaDon (khong chi linkHoaDon) -- dong nao chi da dien/luu 1
    // trong 2 truong deu duoc coi la xong, khong tu dong dong vao nua.
    const candidates = store.chi_phi.filter(
      (r) =>
        r.congTy === activeCompany &&
        (r.mien || "nam") === mien &&
        (r.ngay || "").slice(0, 7) === newestMonth &&
        !(r.linkHoaDon || "").trim() &&
        !(r.soHoaDon || "").trim()
    );
    const toProcess = candidates.slice(0, GMAIL_AUTO_MAX_ROWS_PER_RUN);

    let found = 0;
    let errors = 0;
    for (const r of toProcess) {
      try {
        const match = await gmailInvoiceMatcher.findInvoiceForRow(store, r);
        if (match) {
          r.linkHoaDon = match.linkHoaDon || r.linkHoaDon;
          if (match.soHoaDon) {
            r.soHoaDon = match.soHoaDon;
            r.trangThaiHoaDon = "Đã tìm thấy qua Gmail (tự động)";
          } else {
            r.trangThaiHoaDon = "Đã tìm thấy email qua Gmail (tự động) -- chưa đọc được số hóa đơn, kiểm tra lại link.";
          }
          found++;
        }
      } catch (e) {
        errors++;
      }
    }
    save(store);

    let msg = `Đã quét ${toProcess.length}/${candidates.length} dòng thiếu hóa đơn của tháng ${newestMonth || "(không rõ)"}. Tìm thấy ${found} dòng.`;
    if (candidates.length > toProcess.length) {
      msg += ` Còn ${candidates.length - toProcess.length} dòng chưa quét (bấm "Cập nhật tìm hóa đơn" thêm lần nữa để tiếp tục).`;
    }
    if (errors > 0) {
      msg += ` (${errors} dòng gặp lỗi khi tra Gmail, có thể do hết hạn mức API -- thử lại sau.)`;
    }
    const qs = [];
    if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
    if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
    if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
    qs.push("success=" + encodeURIComponent(msg));
    res.redirect("/chi-phi/" + mienSeg(mien) + "?" + qs.join("&"));
  } catch (e) {
    res.redirect("/chi-phi/" + mienSeg(mien) + "?error=" + encodeURIComponent(e.message));
  }
});

module.exports = router;
// Chi Nhan, 2026-07-29: cho routes/hoa-don-dau-vao.js dung LAI (khong doan
// lai) logic doc+gop 3 Google Sheet UNC vao store.chi_phi -- "Cập nhật Gian
// Hàng" ben Hoa Don Dau Vao se tu lam moi store.chi_phi (giong het nut "Cập
// nhật chi phí" o day) TRUOC KHI do Gian tu do, khong can chi tu qua trang
// Chi Phi bam truoc nua.
module.exports.ensureShape = ensureShape;
module.exports.applyChiPhiMonthRowsToStore = applyChiPhiMonthRowsToStore;
module.exports.CHI_PHI_SHEET_XLSX_URL = CHI_PHI_SHEET_XLSX_URL;
module.exports.KVC_MB_SHEET_XLSX_URL = KVC_MB_SHEET_XLSX_URL;
module.exports.MTD_MB_AUTO_SHEET_XLSX_URL = MTD_MB_AUTO_SHEET_XLSX_URL;
