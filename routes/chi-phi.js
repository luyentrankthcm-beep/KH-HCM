const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");
const { parseChiPhiSheetWorkbook } = require("../utils/chiPhiSheetParser");
const { extractGianRentText, matchGianRecord } = require("../utils/rentPaymentMatcher");
const gmailApi = require("../utils/gmailApi");
const gmailInvoiceMatcher = require("../utils/gmailInvoiceMatcher");

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

router.get("/chi-phi", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hachToanFilter = req.query.hachToan || ""; // "" = tat ca, "1" = da hach toan, "0" = chua hach toan
  // Luyen, 2026-07-21: "thêm chỗ lọc chưa có số hóa đơn" -- loc theo con thieu
  // soHoaDon hay khong, giup tim nhanh cac dong con can dien so hoa don (thu
  // cong hoac tra cuu them) thay vi phai doc het danh sach.
  const hoaDonFilter = req.query.hoaDon || ""; // "" = tat ca, "1" = da co so HD, "0" = chua co so HD
  let rows = store.chi_phi.map(ensureChiPhiDefaults).filter((r) => r.congTy === activeCompany);
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
  res.render("chi-phi", {
    userName: req.session.userName,
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
router.get("/chi-phi/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hachToanFilter = req.query.hachToan || "";
  const hoaDonFilter = req.query.hoaDon || "";
  const thangFilter = req.query.thang !== undefined ? req.query.thang : "";
  let rows = store.chi_phi.map(ensureChiPhiDefaults).filter((r) => r.congTy === activeCompany);
  if (hachToanFilter === "1") rows = rows.filter((r) => r.daHachToan);
  else if (hachToanFilter === "0") rows = rows.filter((r) => !r.daHachToan);
  if (thangFilter) rows = rows.filter((r) => (r.ngay || "").slice(0, 7) === thangFilter);
  if (hoaDonFilter === "1") rows = rows.filter((r) => (r.soHoaDon || "").trim());
  else if (hoaDonFilter === "0") rows = rows.filter((r) => !(r.soHoaDon || "").trim());
  rows.sort((a, b) => (a.ngay < b.ngay ? 1 : -1));

  const exportRows = rows.map((r) => ({
    "Ngày chi": r.ngay,
    "Gian/Cơ sở": r.gian,
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
  res.setHeader("Content-Disposition", `attachment; filename=chi-phi-${activeCompany}.xlsx`);
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

router.post("/chi-phi/:id/hach-toan", requireAdmin, (req, res) => {
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
  res.redirect("/chi-phi?" + qs.join("&"));
});

// Luyen, 2026-07-21: "thêm số hóa đơn thủ công được đi" -- nhieu dong (nhat la
// TT TIEN MAT bi chan captcha/khong mo duoc link) van chua co so hoa don tu
// dong tra ra duoc, can Luyen tu dien tay sau khi tra cuu/xem hoa don giay.
// Cho sua truc tiep tu bang danh sach, khong can vao form rieng.
router.post("/chi-phi/:id/so-hoa-don", requireAdmin, (req, res) => {
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
    return res.redirect("/chi-phi?" + qs.join("&"));
  }
  r.soHoaDon = (req.body.soHoaDon || "").trim();
  // Ghi nhan la da dien tay, xoa cac ghi chu canh bao cu (vd "can mo link
  // kiem tra") vi Luyen da tu xu ly xong dong nay.
  r.trangThaiHoaDon = r.soHoaDon ? "Đã điền số hóa đơn thủ công" : r.trangThaiHoaDon;
  save(store);
  if (isAjaxChiPhiRequest(req)) {
    return res.json({ success: true, soHoaDon: r.soHoaDon, trangThaiHoaDon: r.trangThaiHoaDon || "" });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật số hóa đơn."));
  res.redirect("/chi-phi?" + qs.join("&"));
});

// Luyen, 2026-07-22: "tìm hóa đơn qua Gmail lưu Drive rồi gán lên đây" -- can
// route rieng de cap nhat link hoa don (link Google Drive) cho 1 dong Chi Phi
// da co san, khong dung chung voi /so-hoa-don (chi sua so hoa don thu cong).
router.post("/chi-phi/:id/link-hoa-don", requireAdmin, (req, res) => {
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
    return res.redirect("/chi-phi?" + qs.join("&"));
  }
  r.linkHoaDon = (req.body.linkHoaDon || "").trim();
  save(store);
  if (isAjaxChiPhiRequest(req)) {
    return res.json({ success: true, linkHoaDon: r.linkHoaDon });
  }
  const qs = [];
  if (req.body.hachToan) qs.push("hachToan=" + encodeURIComponent(req.body.hachToan));
  if (req.body.thang !== undefined) qs.push("thang=" + encodeURIComponent(req.body.thang));
  if (req.body.hoaDon) qs.push("hoaDon=" + encodeURIComponent(req.body.hoaDon));
  qs.push("success=" + encodeURIComponent("Đã cập nhật link hóa đơn."));
  res.redirect("/chi-phi?" + qs.join("&"));
});

router.post("/chi-phi", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const { ngay, gian, ncc, soHoaDon, soUNC, soChungTuLienQuan, dienGiai, loaiChiPhi, soTien, linkHoaDon, ghiChu } = req.body;
    if (!ngay) throw new Error("Thiếu ngày chi.");
    if (!loaiChiPhi || !loaiChiPhi.trim()) throw new Error("Thiếu loại chi phí.");
    const amt = soTien ? Number(String(soTien).replace(/[^\d]/g, "")) : 0;
    store.chi_phi.push({
      id: nextId(store, "chi_phi_seq") || Date.now(),
      congTy: activeCompany,
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
    res.redirect("/chi-phi?success=" + encodeURIComponent("Đã lưu khoản chi."));
  } catch (e) {
    res.redirect("/chi-phi?error=" + encodeURIComponent(e.message));
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
function chiPhiRowKey(r) {
  return [r.congTy, r.ngay, (r.ncc || "").trim(), r.soTien, (r.dienGiai || "").trim()].join("|");
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
function applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel) {
  const existingByKey = new Map();
  store.chi_phi.forEach((r) => existingByKey.set(chiPhiRowKey(r), r));
  let added = 0;
  let addedNoGian = 0;
  let upgradedGian = 0;
  const monthSummary = [];
  for (const mr of monthRows) {
    let addedThisSheet = 0;
    for (const r of mr.rows) {
      const key = chiPhiRowKey(r);
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

function buildChiPhiResultMessage(prefix, result, skippedSheets) {
  let msg = `${prefix} thêm ${result.added} khoản chi mới${
    result.monthSummary.length ? " (" + result.monthSummary.join(", ") + ")" : ""
  }.`;
  if (result.upgradedGian > 0) {
    msg += ` Đã nâng cấp Gian (từ nhãn nhóm chung chung như "posh"/"JP" sang tên gian cụ thể) cho ${result.upgradedGian} dòng đã có sẵn.`;
  }
  if (result.addedNoGian > 0) {
    msg += ` CẢNH BÁO: ${result.addedNoGian} dòng không tự xác định được Gian (diễn giải không có cụm "ghế ...") -- cần chị tự điền tay, lọc theo cột Gian trống.`;
  }
  if (skippedSheets.length > 0) {
    msg += ` Đã bỏ qua sheet không phải "Tháng N.YYYY": ${skippedSheets.join(", ")}${
      skippedSheets.includes("TT TIỀN MẶT") ? " (tab này cần mở từng link tra cứu thủ công như trước giờ, không tự động parse)" : ""
    }.`;
  }
  return msg;
}

router.post("/chi-phi/upload", requireAdmin, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { monthRows, skippedSheets } = parseChiPhiSheetWorkbook(req.file.buffer);
    if (monthRows.length === 0) {
      throw new Error(
        `File không có sheet nào khớp tên dạng "Tháng N.YYYY" (đã thấy: ${skippedSheets.join(", ") || "(không có sheet)"}).`
      );
    }
    const sourceLabel = `upload "${req.file.originalname}" ${new Date().toISOString().slice(0, 10)}`;
    const result = applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel);
    save(store);
    const msg = buildChiPhiResultMessage(`Đã nạp "${req.file.originalname}":`, result, skippedSheets);
    res.redirect("/chi-phi?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/chi-phi?error=" + encodeURIComponent(e.message));
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

router.post("/chi-phi/cap-nhat-tu-sheet", requireAdmin, async (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const resp = await fetch(CHI_PHI_SHEET_XLSX_URL);
    if (!resp.ok) {
      throw new Error(
        `Không đọc được Google Sheet (mã lỗi ${resp.status}). Kiểm tra lại sheet đã chia sẻ "Bất kỳ ai có link đều xem được" chưa, hoặc link có bị đổi không.`
      );
    }
    const buf = Buffer.from(await resp.arrayBuffer());
    const { monthRows, skippedSheets } = parseChiPhiSheetWorkbook(buf);
    if (monthRows.length === 0) {
      throw new Error(
        `Google Sheet không có sheet nào khớp tên dạng "Tháng N.YYYY" (đã thấy: ${skippedSheets.join(", ") || "(không có sheet)"}).`
      );
    }
    const sourceLabel = "GG Sheet " + new Date().toISOString().slice(0, 10);
    const result = applyChiPhiMonthRowsToStore(store, monthRows, sourceLabel);
    save(store);
    const msg = buildChiPhiResultMessage("Đã đọc thẳng từ Google Sheet:", result, skippedSheets);
    res.redirect("/chi-phi?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/chi-phi?error=" + encodeURIComponent(e.message));
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
router.post("/chi-phi/quet-ngan-hang", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const now = new Date();
    const monthPrefix = now.toISOString().slice(0, 7); // "YYYY-MM"
    const banksById = {};
    store.banks.forEach((b) => (banksById[b.id] = b));

    const existingBankTxIds = new Set(
      store.chi_phi.filter((r) => r.bankTxId).map((r) => r.bankTxId)
    );

    const chiTx = store.transactions.filter(
      (t) => t.type === "chi" && t.date && t.date.slice(0, 7) === monthPrefix && !existingBankTxIds.has(t.id)
    );

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

    let msg = `Đã quét ${chiTx.length} giao dịch Chi tháng ${monthPrefix} (chưa xử lý) trên tất cả tài khoản. Thêm ${added} khoản chi phí thuê gian mới.`;
    if (flaggedSuspicious > 0) {
      msg += ` CẢNH BÁO: ${flaggedSuspicious} khoản có số tiền bất thường (>= 200 triệu) -- chị xem cột Trạng thái/ghi chú.`;
    }
    if (unmatched.length > 0) {
      msg += ` ${unmatched.length} giao dịch nhận diện là "tiền thuê" nhưng KHÔNG khớp được gian nào rõ ràng -- chị tự thêm tay: ${unmatched
        .map((u) => `${u.bank} ${u.date} ${u.amount.toLocaleString("vi-VN")}đ ("${u.gianText}")`)
        .join("; ")}.`;
    }
    res.redirect("/chi-phi?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/chi-phi?error=" + encodeURIComponent(e.message));
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

router.post("/chi-phi/tim-hoa-don-gmail", requireAdmin, async (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    if (!gmailApi.isConfigured()) {
      throw new Error(
        "Chưa cấu hình Google OAuth trên Railway (thiếu GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET/GOOGLE_REDIRECT_URI)."
      );
    }
    if (!store.gmail_oauth || !store.gmail_oauth.refresh_token) {
      throw new Error("Chưa kết nối Gmail -- bấm nút \"Kết nối Gmail\" trước.");
    }

    let rows = store.chi_phi.map(ensureChiPhiDefaults).filter((r) => r.congTy === activeCompany);
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
    res.redirect("/chi-phi?" + qs.join("&"));
  } catch (e) {
    res.redirect("/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

module.exports = router;
