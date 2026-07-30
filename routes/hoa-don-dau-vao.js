const express = require("express");
const XLSX = require("xlsx");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");
const { parseHoaDonDauVaoWorkbook } = require("../utils/hoaDonDauVaoParser");
const {
  parseNccListWorkbook,
  parseHangHoaWorkbook,
  parseGianSheetWorkbook,
  matchByMstOrName,
  matchGianViaUncContent,
  matchGianViaChiPhiLedger,
  classifyPhanLoai,
  matchHangHoaRecord,
  classifyFromHangHoaMatch,
  normVN,
} = require("../utils/hoaDonDauVaoEnrich");
const chiPhi = require("./doisoat-chiphi");
// Chi Nhan, 2026-07-29: "đi tìm từ 3 unc gg sheet á lấy ra gian cho tôi" --
// store.chi_phi (routes/chi-phi.js) da dong bo san tu 3 Google Sheet "UNC",
// co san cot gian/ncc/soHoaDon/soTien -- dung lam nguon do them Gian Hang
// (xem matchGianViaChiPhiLedger). Va rentPaymentMatcher de xac dinh Tai
// khoan Co (1388/331) tu ten gian tim duoc, dung LAI logic da co san o
// chi-phi.js (computeTaiKhoanChiPhi) thay vi doan lai tu dau.
const { findContractForGianText, buildGianAliasIndex, isDoanhThuChiaSeRecord } = require("../utils/rentPaymentMatcher");
// Chi Nhan, 2026-07-29: "3 cái đi unc lấy ra chỗ gian ... dựa vào hóa đơn hay
// ncc hay số tiền" -- dung LAI (khong doan lai) 3 URL Google Sheet UNC + ham
// gop du lieu da co san o routes/chi-phi.js (trang "Chi Phí"), de nut "Cập
// nhật Gian Hàng" ben nay TU LAM MOI store.chi_phi truoc khi do Gian, khong
// bat chi phai qua trang Chi Phi bam "Cập nhật chi phí" truoc nua.
const chiPhiSheetRoutes = require("./chi-phi");
const {
  parseChiPhiSheetWorkbook,
  parseKvcMienBacWorkbook,
  parseKvcMienBacAutoWorkbook,
} = require("../utils/chiPhiSheetParser");
// Chi Nhan, 2026-07-28: "check trên UNC ra tên gian" -- khi khong khop duoc
// Gian Hang qua Google Sheet (theo Ten NCC/MST), thu tim tiep qua bang lenh
// chi UNC (chi_phi_unc_list, cung du lieu voi trang Doi Soat Chi Phi): khop
// UNC theo Ten NCC + So tien hoa don, roi do noi dung UNC xem co chua ma
// diem noi bo/ma diem thue nao trong danh sach gian (cua dung cong ty) khong
// -- CHI dien khi khop DUY NHAT 1 gian, tranh doan nham.
const { buildUncIndex, matchUncForPayment } = require("../utils/chiphiReconcile");

const router = express.Router();
router.use(requireLogin);
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 80 * 1024 * 1024 } });

// Chi Nhan, 2026-07-28: link Google Sheet "Danh sách các gian" (3 tab: HÀ NỘI
// MTD, HÀ NỘI KVC, HỒ CHÍ MINH) -- dung de tu dien Gian Hàng + Tài khoản Có
// (1388 neu Hinh Thuc Hop Tac = CSE, con lai 331), khop theo "Tên khách hàng"
// (ben cho thue/NCC) + "Pháp nhân" (KH cũ/KH mới) trong chinh sheet do.
const GIAN_SHEET_XLSX_URL =
  process.env.HOA_DON_DAU_VAO_GIAN_SHEET_URL ||
  "https://docs.google.com/spreadsheets/d/1Fd5v128o6eVuzHqtYzV5Eef62YtW6x5X/export?format=xlsx";

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
  if (!store.hoa_don_dau_vao_ncc_list) store.hoa_don_dau_vao_ncc_list = [];
  if (!store.hoa_don_dau_vao_ncc_meta) store.hoa_don_dau_vao_ncc_meta = null;
  if (!store.hoa_don_dau_vao_hang_hoa_list) store.hoa_don_dau_vao_hang_hoa_list = [];
  if (!store.hoa_don_dau_vao_hang_hoa_meta) store.hoa_don_dau_vao_hang_hoa_meta = null;
  if (!store.hoa_don_dau_vao_gian_list) store.hoa_don_dau_vao_gian_list = [];
  if (!store.hoa_don_dau_vao_gian_meta) store.hoa_don_dau_vao_gian_meta = null;
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
      // Chi Nhan, 2026-07-29: "hóa đơn lấy thêm thông tin đơn vị tính ... số
      // lượng và đơn giá cho tôi nhá xuất ra file excel mới có thôi chứ
      // không cần hiển thị thêm trên đây" -- CHI dung cho export.xlsx.
      donViTinh: "",
      soLuong: 0,
      donGia: 0,
      // Chi Nhan, 2026-07-28: 8 cot "lam giau" moi -- tat ca deu la GOI Y tu
      // dong (tu doi chieu voi Danh sach NCC/hang hoa/cac gian, hoac tu Dien
      // giai), Chi Nhan tu sua lai tung dong qua nut sua tren bang neu sai.
      gianHang: "",
      hinhThucHopTac: "",
      taiKhoanCo: "",
      daChiTien: false,
      maDoiTuongNCC: "",
      tenHangHoaMisa: "",
      phanLoai: "",
      taiKhoanNo: "",
    },
    row
  );
}

// Ap dung phan loai/tai khoan No + gian hang/tai khoan Co/ma doi tuong NCC/
// ten hang hoa (neu da co san du lieu doi chieu trong store) ngay khi 1 dong
// moi duoc tao (nhap tay hoac upload bang ke) -- tranh phai bam "Cap nhat" tay
// ngay sau khi vua tai/nhap xong.
function enrichNewRow(store, row) {
  // Chi Nhan, 2026-07-29: "tên hàng hóa á thì thêm cho tôi tài khoản là 156
  // đi" -- khop duoc voi danh muc Hang hoa/dich vu (file MISA Chi Nhan da tai
  // len) la nguon DANG TIN CAY NHAT (do CHINH XAC tung mat hang, khong phai
  // doan tu keyword) nen uu tien truoc: dung TK Kho cua chinh mat hang do neu
  // file co ghi, khong co thi mac dinh 156 (hang hoa ban ra). CHI khi KHONG
  // khop duoc voi danh muc nao moi roi ve doan qua tu khoa trong Dien giai
  // (classifyPhanLoai, xem ghi chu tai do).
  const hangHoaList = store.hoa_don_dau_vao_hang_hoa_list || [];
  const matchedHH = hangHoaList.length > 0 ? matchHangHoaRecord(row.dienGiai, hangHoaList) : null;
  if (matchedHH) {
    const r = classifyFromHangHoaMatch(matchedHH, row.dienGiai, row.soTienTruocThue);
    row.tenHangHoaMisa = r.tenHangHoaMisa;
    row.phanLoai = r.phanLoai;
    row.taiKhoanNo = r.taiKhoanNo;
  } else {
    const classified = classifyPhanLoai(row.dienGiai, row.soTienTruocThue);
    row.phanLoai = classified.phanLoai;
    row.taiKhoanNo = classified.taiKhoanNo;
  }
  const nccList = (store.hoa_don_dau_vao_ncc_list || []).filter((n) => n.congTy === row.congTy);
  if (nccList.length > 0) {
    const m = matchByMstOrName(row.mstNCC, row.tenNCC, nccList, "mst", "tenNCC");
    if (m) row.maDoiTuongNCC = m.maNCC;
  }
  const gianList = (store.hoa_don_dau_vao_gian_list || []).filter((g) => g.congTy === row.congTy);
  if (gianList.length > 0) {
    const m = matchByMstOrName(row.mstNCC, row.tenNCC, gianList, "mstKhachHang", "tenKhachHang");
    if (m) {
      row.gianHang = m.gianHang;
      row.hinhThucHopTac = m.hinhThucHopTac;
      row.taiKhoanCo = /cse/i.test(m.hinhThucHopTac) ? "1388" : "331";
    }
  }
  return row;
}

// Chi Nhan, 2026-07-28: "mấy cái hóa đơn nhiều dòng á bạn gom lại chỗ diễn
// giải để nhiều dòng thôi á xuất ra thì mỗi cái 1 dòng cho tôi chớ trên đây
// coi gom lại cho tôi đi" -- tren TRANG WEB (khong dung cho export.xlsx, xem
// route rieng ben duoi), gop cac dong CUNG 1 hoa don (cung So hoa don + Ky
// hieu) lai thanh 1 dong hien thi: Dien giai noi cac dong con lai bang xuong
// dong, cong don Tien truoc thue/Tien thue/Tong tien. Du lieu GOC trong
// store van giu nguyen moi dong rieng (khong sua/gop that trong store.json)
// -- chi gop luc RENDER, nen export/upload/dedup phia tren khong bi anh
// huong gi ca.
//
// Dong nhap tay KHONG co So hoa don (rong) thi KHONG gop chung voi nhau (moi
// dong 1 nhom rieng, dung id lam khoa) -- tranh gop nham hang loat dong
// khong lien quan chi vi cung co soHoaDon="".
function groupRowsByInvoice(rows) {
  const map = new Map();
  const order = [];
  rows.forEach((r) => {
    const key = r.soHoaDon ? `${r.soHoaDon}|${r.kyHieuHD || ""}` : `__single__${r.id}`;
    if (!map.has(key)) {
      map.set(key, {
        ids: [],
        ngayHD: r.ngayHD,
        tenNCC: r.tenNCC,
        mstNCC: r.mstNCC,
        kyHieuHD: r.kyHieuHD,
        soHoaDon: r.soHoaDon,
        dienGiaiList: [],
        soTienTruocThue: 0,
        tienThue: 0,
        soTien: 0,
        linkHoaDon: "",
        ghiChuList: [],
        daHachToanCount: 0,
        gianHang: "",
        taiKhoanCo: "",
        daChiTienCount: 0,
        maDoiTuongNCC: "",
        // Chi Nhan, 2026-07-28: Ten hang hoa/Phan loai/Tai khoan No la khai
        // niem THEO TUNG DONG hang hoa (1 hoa don gop nhieu mat hang co the
        // moi mat hang 1 loai khac nhau) -- noi song song CUNG THU TU voi
        // Dien giai (xem dienGiaiList) de doc doi chieu tung dong cho dung,
        // khac voi Gian Hang/Tai khoan Co/Ma doi tuong NCC la khai niem CHUNG
        // CA HOA DON (cung 1 NCC/gian) nen chi lay 1 gia tri dai dien.
        tenHangHoaList: [],
        phanLoaiList: [],
        taiKhoanNoList: [],
      });
      order.push(key);
    }
    const g = map.get(key);
    g.ids.push(r.id);
    if (r.dienGiai) g.dienGiaiList.push(r.dienGiai);
    g.soTienTruocThue += r.soTienTruocThue || 0;
    g.tienThue += r.tienThue || 0;
    g.soTien += r.soTien || 0;
    if (!g.linkHoaDon && r.linkHoaDon) g.linkHoaDon = r.linkHoaDon;
    if (r.ghiChu) g.ghiChuList.push(r.ghiChu);
    if (r.daHachToan) g.daHachToanCount++;
    if (!g.gianHang && r.gianHang) g.gianHang = r.gianHang;
    if (!g.taiKhoanCo && r.taiKhoanCo) g.taiKhoanCo = r.taiKhoanCo;
    if (r.daChiTien) g.daChiTienCount++;
    if (!g.maDoiTuongNCC && r.maDoiTuongNCC) g.maDoiTuongNCC = r.maDoiTuongNCC;
    g.tenHangHoaList.push(r.tenHangHoaMisa || "");
    g.phanLoaiList.push(r.phanLoai || "");
    g.taiKhoanNoList.push(r.taiKhoanNo || "");
  });
  return order.map((key) => {
    const g = map.get(key);
    return {
      idsCsv: g.ids.join(","),
      soDong: g.ids.length,
      ngayHD: g.ngayHD,
      tenNCC: g.tenNCC,
      mstNCC: g.mstNCC,
      kyHieuHD: g.kyHieuHD,
      soHoaDon: g.soHoaDon,
      dienGiai: g.dienGiaiList.join("\n"),
      soTienTruocThue: g.soTienTruocThue,
      tienThue: g.tienThue,
      soTien: g.soTien,
      linkHoaDon: g.linkHoaDon,
      ghiChu: g.ghiChuList.join("; "),
      daHachToan: g.daHachToanCount === g.ids.length,
      gianHang: g.gianHang,
      taiKhoanCo: g.taiKhoanCo,
      daChiTien: g.daChiTienCount === g.ids.length,
      maDoiTuongNCC: g.maDoiTuongNCC,
      tenHangHoa: g.tenHangHoaList.join("\n"),
      phanLoai: g.phanLoaiList.join("\n"),
      taiKhoanNo: g.taiKhoanNoList.join("\n"),
    };
  });
}

// Chi Nhan, 2026-07-29: "thêm bộ lọc ... lọc theo tk 154 156 242 hay các lọc
// theo gian có/nhiều gian ... hiển thị 30 hóa đơn thôi trang 1 trang 2" --
// them 3 bo loc moi (Tai khoan No, Gian: da co/chua co) + phan trang 30 hoa
// don/trang. Loc Tai khoan No ap dung O CAP TUNG DONG hang hoa (truoc khi
// gop hien thi theo hoa don), vi 1 hoa don gop nhieu mat hang co the khac
// Tai khoan No nhau -- giu dung dong khop, cac dong khac cua cung hoa don bi
// loai (giong nguyen tac loc thang o tren). Loc Gian + phan trang ap dung SAU
// khi da gop theo hoa don (gianHang la khai niem chung ca hoa don).
router.get("/hoa-don-dau-vao", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hachToanFilter = req.query.hachToan || ""; // "" = tat ca, "1" = da hach toan, "0" = chua hach toan
  const thangFilter = req.query.thang || "";
  const tkNoFilter = req.query.tkNo || ""; // "" = tat ca, hoac 1 ma TK No cu the (vd "154")
  const gianFilter = req.query.gian || ""; // "" = tat ca, "1" = da co Gian Hang, "0" = chua co
  // Chi Nhan, 2026-07-29: "lên cho tôi cái hóa đơn nào chi chưa" -- them bo
  // loc Da chi (Tat ca / Da chi / Chua chi), giup tim nhanh hoa don con chua
  // thanh toan thay vi phai doc het danh sach.
  const daChiFilter = req.query.daChi || ""; // "" = tat ca, "1" = da chi, "0" = chua chi
  const page = Math.max(1, parseInt(req.query.page, 10) || 1);
  const PAGE_SIZE = 30;

  const allRows = store.hoa_don_dau_vao.map(ensureDefaults).filter((r) => r.congTy === activeCompany);
  const totalForCompany = groupRowsByInvoice(allRows).length;

  const monthSet = new Set();
  const tkNoSet = new Set();
  allRows.forEach((r) => {
    const m = (r.ngayHD || "").slice(0, 7);
    if (m) monthSet.add(m);
    if ((r.taiKhoanNo || "").trim()) tkNoSet.add(r.taiKhoanNo.trim());
  });
  const availableMonths = [...monthSet].sort().reverse();
  const availableTkNo = [...tkNoSet].sort();

  let rows = allRows;
  if (thangFilter) rows = rows.filter((r) => (r.ngayHD || "").slice(0, 7) === thangFilter);
  if (tkNoFilter) rows = rows.filter((r) => (r.taiKhoanNo || "").trim() === tkNoFilter);

  let groupedRows = groupRowsByInvoice(rows);
  if (hachToanFilter === "1") groupedRows = groupedRows.filter((r) => r.daHachToan);
  else if (hachToanFilter === "0") groupedRows = groupedRows.filter((r) => !r.daHachToan);
  if (gianFilter === "1") groupedRows = groupedRows.filter((r) => (r.gianHang || "").trim());
  else if (gianFilter === "0") groupedRows = groupedRows.filter((r) => !(r.gianHang || "").trim());
  if (daChiFilter === "1") groupedRows = groupedRows.filter((r) => r.daChiTien);
  else if (daChiFilter === "0") groupedRows = groupedRows.filter((r) => !r.daChiTien);
  groupedRows.sort((a, b) => (a.ngayHD < b.ngayHD ? 1 : -1));
  const tongTien = groupedRows.reduce((s, r) => s + (r.soTien || 0), 0);
  const daChiCount = groupedRows.filter((r) => r.daChiTien).length;

  const totalMatching = groupedRows.length;
  const totalPages = Math.max(1, Math.ceil(totalMatching / PAGE_SIZE));
  const currentPage = Math.min(page, totalPages);
  const pageRows = groupedRows.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);

  // Giu nguyen tat ca bo loc dang chon khi chuyen trang (chi doi "page").
  const qs = [];
  if (thangFilter) qs.push("thang=" + encodeURIComponent(thangFilter));
  if (hachToanFilter) qs.push("hachToan=" + encodeURIComponent(hachToanFilter));
  if (tkNoFilter) qs.push("tkNo=" + encodeURIComponent(tkNoFilter));
  if (gianFilter) qs.push("gian=" + encodeURIComponent(gianFilter));
  if (daChiFilter) qs.push("daChi=" + encodeURIComponent(daChiFilter));
  const baseQs = qs.join("&");

  const nccMeta = store.hoa_don_dau_vao_ncc_meta;
  const hangHoaMeta = store.hoa_don_dau_vao_hang_hoa_meta;
  const gianMeta = store.hoa_don_dau_vao_gian_meta;
  const nccCountForCompany = (store.hoa_don_dau_vao_ncc_list || []).filter((n) => n.congTy === activeCompany).length;
  const gianCountForCompany = (store.hoa_don_dau_vao_gian_list || []).filter((g) => g.congTy === activeCompany).length;

  res.render("hoa-don-dau-vao", {
    userName: req.session.userName,
    rows: pageRows,
    totalForCompany,
    totalMatching,
    hachToanFilter,
    thangFilter,
    tkNoFilter,
    gianFilter,
    daChiFilter,
    availableMonths,
    availableTkNo,
    tongTien,
    daChiCount,
    currentPage,
    totalPages,
    baseQs,
    nccMeta,
    hangHoaMeta,
    gianMeta,
    nccCountForCompany,
    gianCountForCompany,
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
    "Ma đối tượng NCC": r.maDoiTuongNCC,
    "Ký hiệu": r.kyHieuHD,
    "Số hóa đơn": r.soHoaDon,
    "Diễn giải": r.dienGiai,
    "Tên hàng hóa (Misa)": r.tenHangHoaMisa,
    "Đơn vị tính": r.donViTinh,
    "Số lượng": r.soLuong,
    "Đơn giá": r.donGia,
    "Phân loại": r.phanLoai,
    "Tài khoản Nợ": r.taiKhoanNo,
    "Tài khoản Có": r.taiKhoanCo,
    "Gian Hàng": r.gianHang,
    "Tiền trước thuế": r.soTienTruocThue,
    "Tiền thuế": r.tienThue,
    "Tổng tiền thanh toán": r.soTien,
    "Đã chi tiền": r.daChiTien ? "Có" : "Không",
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
    const newRow = {
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
      gianHang: "",
      hinhThucHopTac: "",
      taiKhoanCo: "",
      daChiTien: false,
      maDoiTuongNCC: "",
      tenHangHoaMisa: "",
      phanLoai: "",
      taiKhoanNo: "",
    };
    enrichNewRow(store, newRow);
    store.hoa_don_dau_vao.push(newRow);
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
        donViTinh: r.donViTinh || "",
        soLuong: r.soLuong || 0,
        donGia: r.donGia || 0,
        linkHoaDon: "",
        daHachToan: false,
        ghiChu: "",
        createdAt: new Date().toISOString(),
        source: `upload "${req.file.originalname}" ${new Date().toISOString().slice(0, 10)}`,
        gianHang: "",
        hinhThucHopTac: "",
        taiKhoanCo: "",
        daChiTien: false,
        maDoiTuongNCC: "",
        tenHangHoaMisa: "",
        phanLoai: "",
        taiKhoanNo: "",
      };
      enrichNewRow(store, newRow);
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

// Chi Nhan, 2026-07-28: trang gio gop nhieu dong (cung 1 hoa don) lai thanh 1
// dong hien thi (xem groupRowsByInvoice o tren) -- nut "hạch toán"/"Xóa" tren
// 1 dong hien thi phai ap dung cho TOAN BO cac dong GOC trong nhom do (id cua
// tung dong duoc gop lai truyen qua truong "ids", cach nhau boi dau phay).
// Chi Nhan, 2026-07-29: gop chung cach dung lai cac bo loc (thang/hachToan/
// tkNo/gian/page) khi redirect ve sau moi hanh dong tren 1 dong -- tranh nhay
// ve trang 1/mat bo loc dang xem sau khi bam Da chi/Hach toan/Luu/Xoa.
function buildRedirectQs(body) {
  const qs = [];
  if (body.hachToan) qs.push("hachToan=" + encodeURIComponent(body.hachToan));
  if (body.thang) qs.push("thang=" + encodeURIComponent(body.thang));
  if (body.tkNo) qs.push("tkNo=" + encodeURIComponent(body.tkNo));
  if (body.gian) qs.push("gian=" + encodeURIComponent(body.gian));
  if (body.daChi) qs.push("daChi=" + encodeURIComponent(body.daChi));
  if (body.page) qs.push("page=" + encodeURIComponent(body.page));
  return qs;
}

router.post("/hoa-don-dau-vao/hach-toan", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const ids = (req.body.ids || "").split(",").filter(Boolean);
  const daHachToan = req.body.daHachToan === "1";
  store.hoa_don_dau_vao.forEach((r) => {
    if (ids.includes(String(r.id))) r.daHachToan = daHachToan;
  });
  save(store);
  const qs = buildRedirectQs(req.body);
  qs.push("success=" + encodeURIComponent("Đã cập nhật trạng thái hạch toán."));
  res.redirect("/hoa-don-dau-vao?" + qs.join("&"));
});

// Chi Nhan, 2026-07-28: toggle tay cho "Đã chi tiền chưa" -- cung 1 kieu voi
// hach-toan o tren (ap dung cho ca nhom id cua 1 hoa don).
router.post("/hoa-don-dau-vao/da-chi-tien", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const ids = (req.body.ids || "").split(",").filter(Boolean);
  const daChiTien = req.body.daChiTien === "1";
  store.hoa_don_dau_vao.forEach((r) => {
    if (ids.includes(String(r.id))) r.daChiTien = daChiTien;
  });
  save(store);
  const qs = buildRedirectQs(req.body);
  qs.push("success=" + encodeURIComponent("Đã cập nhật trạng thái chi tiền."));
  res.redirect("/hoa-don-dau-vao?" + qs.join("&"));
});

router.post("/hoa-don-dau-vao/xoa", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const ids = (req.body.ids || "").split(",").filter(Boolean);
  store.hoa_don_dau_vao = store.hoa_don_dau_vao.filter((r) => !ids.includes(String(r.id)));
  save(store);
  res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent("Đã xóa hóa đơn."));
});

// Chi Nhan, 2026-07-28: nut sua truc tiep tren bang cho 6 cot "goi y tu dong"
// (Gian Hàng/Tài khoản Có/Mã đối tượng NCC/Tên hàng hóa/Phân loại/Tài khoản
// Nợ) -- ap dung CUNG 1 gia tri cho TOAN BO cac dong GOC trong 1 nhom hoa don
// (giong hach-toan/xoa o tren). "Tài khoản Có" sua rieng se KHONG con tu dong
// bi ghi de boi "Cập nhật Gian Hàng" lan sau NEU dang khac rong (xem cac route
// cap-nhat-* ben duoi, chi dien vao o dang TRONG).
const EDITABLE_FIELDS = new Set(["gianHang", "taiKhoanCo", "maDoiTuongNCC", "tenHangHoaMisa", "phanLoai", "taiKhoanNo"]);
// Chi Nhan, 2026-07-28: "còn phần nào sửa thì thêm vào mục các hóa đơn có
// chữ sửa chọn nào tích đó" -- gop 6 form/6 nut Luu rieng le (moi field 1
// form) thanh 1 form/1 nut Luu duy nhat. Form gui len 1 cap checkbox+gia tri
// cho MOI field (apply_<field> = "1" neu duoc tich, value_<field> = gia tri
// nhap). CHI field nao co apply_<field>="1" moi duoc ghi de; field khong
// tich thi GIU NGUYEN gia tri cu (khong dong nao bi mat du lieu ngoai y muon).
// Van giu tuong thich nguoc voi cach cu (field/value don) neu co noi nao khac
// con goi theo kieu cu.
router.post("/hoa-don-dau-vao/sua", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const ids = (req.body.ids || "").split(",").filter(Boolean);

  const fieldsToApply = {};
  if (req.body.field && EDITABLE_FIELDS.has(req.body.field)) {
    // Kieu cu: 1 field/1 value duy nhat.
    fieldsToApply[req.body.field] = (req.body.value || "").trim();
  }
  let anyChecked = false;
  EDITABLE_FIELDS.forEach((f) => {
    if (req.body["apply_" + f] === "1") {
      anyChecked = true;
      fieldsToApply[f] = (req.body["value_" + f] || "").trim();
    }
  });
  if (Object.keys(fieldsToApply).length === 0) {
    const msg = anyChecked || req.body.field ? "Cột không hợp lệ." : "Vui lòng tích chọn ít nhất 1 ô muốn sửa.";
    return res.redirect("/hoa-don-dau-vao?error=" + encodeURIComponent(msg));
  }
  store.hoa_don_dau_vao.forEach((r) => {
    if (!ids.includes(String(r.id))) return;
    Object.keys(fieldsToApply).forEach((f) => {
      r[f] = fieldsToApply[f];
    });
  });
  save(store);
  const qs = buildRedirectQs(req.body);
  qs.push("success=" + encodeURIComponent("Đã lưu."));
  res.redirect("/hoa-don-dau-vao?" + qs.join("&"));
});

// Chi Nhan, 2026-07-28: "từ cái tên ncc đây á tôi sẽ tải mã đối tượng NCC ...
// hay cái nào mới bạn để trống ô đó tôi sẽ lọc và điền vào nhá" -- tai len
// "Danh sách nhà cung cấp" (file MISA export), khop theo MST truoc (chac
// chan nhat), khong co/khong khop thi thu theo Ten NCC. CHI dien vao cac dong
// dang TRONG maDoiTuongNCC (dong da co gia tri -- du la tu lan cap nhat truoc
// hay Chi Nhan tu dien tay -- deu duoc GIU NGUYEN, khong ghi de).
router.post("/hoa-don-dau-vao/upload-ncc", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { rows } = parseNccListWorkbook(req.file.buffer);
    const taggedRows = rows.map((r) => ({ ...r, congTy: activeCompany }));
    store.hoa_don_dau_vao_ncc_list = (store.hoa_don_dau_vao_ncc_list || []).filter((r) => r.congTy !== activeCompany).concat(taggedRows);
    store.hoa_don_dau_vao_ncc_meta = { uploaded_at: new Date().toISOString(), file_name: req.file.originalname, count: rows.length, congTy: activeCompany };

    let filled = 0;
    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany || r.maDoiTuongNCC) return;
      const m = matchByMstOrName(r.mstNCC, r.tenNCC, taggedRows, "mst", "tenNCC");
      if (m) {
        r.maDoiTuongNCC = m.maNCC;
        filled++;
      }
    });
    save(store);
    res.redirect(
      "/hoa-don-dau-vao?success=" +
        encodeURIComponent(`Đã đọc "${req.file.originalname}" (${rows.length} NCC): điền Mã đối tượng NCC cho ${filled} dòng đang trống.`)
    );
  } catch (e) {
    res.redirect("/hoa-don-dau-vao?error=" + encodeURIComponent(e.message));
  }
});

// Chi Nhan, 2026-07-28: "thêm cho tôi 1 cột tên hàng hóa từ các hàng hóa map
// lại đúng mẫu misa này cho tôi nhá" -- tai len "Danh sách hàng hóa, dịch
// vụ" (file MISA export), khop chuoi con trong Dien giai (uu tien ten dai/cu
// the truoc). Danh sach nay KHONG tach theo cong ty (hang hoa/dich vu dung
// chung ca 2 phap nhan). Chi dien vao dong dang TRONG tenHangHoaMisa.
router.post("/hoa-don-dau-vao/upload-hang-hoa", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { rows } = parseHangHoaWorkbook(req.file.buffer);
    store.hoa_don_dau_vao_hang_hoa_list = rows;
    store.hoa_don_dau_vao_hang_hoa_meta = { uploaded_at: new Date().toISOString(), file_name: req.file.originalname, count: rows.length };

    // Chi Nhan, 2026-07-29: "tài khoản là 156 đi ... check lại tên nha có thể
    // tên khác" -- dung matchHangHoaRecord (khop them ca theo Ma, khong chi
    // Ten) va, khi khop duoc, DIEN LUON Tai khoan No/Phan loai theo TK Kho cua
    // mat hang do (mac dinh 156 neu file khong ghi TK Kho) -- CHI dien vao
    // dong dang TRONG (khong ghi de Chi Nhan da tu sua qua nut Sua).
    let filled = 0;
    let filledTk = 0;
    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany) return;
      const matched = matchHangHoaRecord(r.dienGiai, rows);
      if (!matched) return;
      const classified = classifyFromHangHoaMatch(matched, r.dienGiai, r.soTienTruocThue);
      if (!r.tenHangHoaMisa) {
        r.tenHangHoaMisa = classified.tenHangHoaMisa;
        filled++;
      }
      if (!r.taiKhoanNo) {
        r.phanLoai = classified.phanLoai;
        r.taiKhoanNo = classified.taiKhoanNo;
        filledTk++;
      }
    });
    save(store);
    res.redirect(
      "/hoa-don-dau-vao?success=" +
        encodeURIComponent(
          `Đã đọc "${req.file.originalname}" (${rows.length} hàng hóa/dịch vụ): điền Tên hàng hóa cho ${filled} dòng, Tài khoản Nợ cho ${filledTk} dòng đang trống.`
        )
    );
  } catch (e) {
    res.redirect("/hoa-don-dau-vao?error=" + encodeURIComponent(e.message));
  }
});

// Chi Nhan, 2026-07-28: "rồi thêm cho tôi dựa trên cái link ... có tên khách
// hàng á là ncc của mình á check theo kiểu ấy từ các sheet trên link này
// thêm nút cập nhật từng cột cho tôi nhá" -- doc thang tu Google Sheet "Danh
// sách các gian" (3 tab, dung chung ca KH Cu/Moi qua cot "Pháp nhân" ngay
// trong sheet), khop "Tên khách hàng" (ben cho thue = NCC cua minh) + dung
// cong ty. CHI dien vao dong dang TRONG gianHang (giong upload-ncc/hang-hoa).
// Chi Nhan, 2026-07-29: "mình thanh toán cho nó gần hết rồi với gian bạn tìm
// trên gg sheet cho tôi chưa á 3 cái đi unc lấy ra chỗ gian dựa vào hóa đơn
// hay ncc hay số tiền" -- goi 1 lan bam nut la LAM MOI ca 4 nguon (khong con
// bat chi phai qua trang Chi Phi bam truoc): (1) sheet "Danh sách các gian"
// (khop Ten NCC/MST), (2+3+4) 3 sheet UNC that (Mien Nam + 2 sheet Mien Bac,
// dung LAI ham gop cua routes/chi-phi.js) roi do Gian qua so Chi Phi vua lam
// moi. MOI nguon fetch rieng, 1 nguon loi KHONG chan cac nguon con lai (vd
// sheet "Danh sách các gian" bi loi van khong ngan viec lam moi + do Gian qua
// so Chi Phi, va nguoc lai) -- gom tat ca ghi chu ket qua/loi vao 1 thong bao
// duy nhat cuoi cung.
router.post("/hoa-don-dau-vao/cap-nhat-gian", requireDataEntry, async (req, res) => {
  const store = load();
  ensureShape(store);
  chiPhiSheetRoutes.ensureShape(store);
  const activeCompany = getCompany(req);
  const notes = [];
  let filled = 0;
  let gianForCompany = (store.hoa_don_dau_vao_gian_list || []).filter((g) => g.congTy === activeCompany);

  // (1) Sheet "Danh sách các gian" -- khop Ten NCC/MST (nguon co san tu truoc).
  try {
    const resp = await fetch(GIAN_SHEET_XLSX_URL);
    if (!resp.ok) {
      throw new Error(`mã lỗi ${resp.status}`);
    }
    const buf = Buffer.from(await resp.arrayBuffer());
    const { rows, sheetsRead } = parseGianSheetWorkbook(buf);
    store.hoa_don_dau_vao_gian_list = rows;
    store.hoa_don_dau_vao_gian_meta = { fetched_at: new Date().toISOString(), count: rows.length, sheetsRead };
    gianForCompany = rows.filter((g) => g.congTy === activeCompany);

    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany || r.gianHang) return;
      const m = matchByMstOrName(r.mstNCC, r.tenNCC, gianForCompany, "mstKhachHang", "tenKhachHang");
      if (m) {
        r.gianHang = m.gianHang;
        r.hinhThucHopTac = m.hinhThucHopTac;
        r.taiKhoanCo = /cse/i.test(m.hinhThucHopTac) ? "1388" : "331";
        filled++;
      }
    });
    notes.push(`Đã đọc Google Sheet Danh sách các gian (${sheetsRead.join(", ")}): điền ${filled} dòng đang trống.`);
  } catch (e) {
    notes.push(`Lỗi đọc sheet Danh sách các gian (${e.message}).`);
  }

  // Chi Nhan, 2026-07-28: "check trên unc ra tên gian" -- con dong nao van
  // trong gianHang (khong khop duoc qua Ten NCC/MST o tren) thi thu do them
  // qua bang lenh chi UNC (file tai len rieng o trang Doi soat Chi phi): khop
  // UNC theo Ten NCC + So tien hoa don truoc, roi do noi dung UNC do xem co
  // chua ma diem cua gian nao khong.
  let filledViaUnc = 0;
  const uncList = store.chi_phi_unc_list || [];
  if (uncList.length > 0) {
    const uncIndex = buildUncIndex(uncList);
    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany || r.gianHang) return;
      const uncMatch = matchUncForPayment(r.tenNCC, r.soTien, uncIndex);
      if (!uncMatch || !uncMatch.noiDungUnc) return;
      const g = matchGianViaUncContent(uncMatch.noiDungUnc, gianForCompany);
      if (g) {
        r.gianHang = g.gianHang;
        r.hinhThucHopTac = g.hinhThucHopTac;
        r.taiKhoanCo = /cse/i.test(g.hinhThucHopTac) ? "1388" : "331";
        filledViaUnc++;
      }
    });
    if (filledViaUnc > 0) notes.push(`Dò thêm qua UNC (file tải lên): +${filledViaUnc} dòng.`);
  }

  // (2+3+4) Chi Nhan, 2026-07-29: "3 cái đi unc lấy ra chỗ gian dựa vào hóa
  // đơn hay ncc hay số tiền" -- lam moi CHINH store.chi_phi tu 3 Google Sheet
  // UNC that (Mien Nam + 2 sheet Mien Bac), dung LAI ham applyChiPhiMonthRowsToStore
  // cua routes/chi-phi.js (giong het nut "Cập nhật chi phí" o trang Chi Phi) --
  // chi THEM dong moi, khong dong tay dong da co (an toan, khong sua nham
  // du lieu Chi Nhan da tu dieu chinh).
  try {
    const respNam = await fetch(chiPhiSheetRoutes.CHI_PHI_SHEET_XLSX_URL);
    if (respNam.ok) {
      const buf = Buffer.from(await respNam.arrayBuffer());
      const { monthRows, skippedSheets } = parseChiPhiSheetWorkbook(buf);
      if (monthRows.length > 0) {
        const result = chiPhiSheetRoutes.applyChiPhiMonthRowsToStore(store, monthRows, "GG Sheet " + new Date().toISOString().slice(0, 10), "nam");
        if (result.added > 0) notes.push(`Sổ Chi Phí Miền Nam: +${result.added} khoản chi mới.`);
      }
    } else {
      notes.push(`Lỗi đọc sheet Chi Phí Miền Nam (mã lỗi ${respNam.status}).`);
    }
  } catch (e) {
    notes.push(`Lỗi đọc sheet Chi Phí Miền Nam (${e.message}).`);
  }
  try {
    const [respManual, respAuto] = await Promise.all([
      fetch(chiPhiSheetRoutes.KVC_MB_SHEET_XLSX_URL),
      fetch(chiPhiSheetRoutes.MTD_MB_AUTO_SHEET_XLSX_URL),
    ]);
    const monthRowsBac = [];
    if (respManual.ok) {
      const buf = Buffer.from(await respManual.arrayBuffer());
      const { sheetName, rows } = parseKvcMienBacWorkbook(buf);
      if (sheetName) monthRowsBac.push({ sheetName: `UNC KVC MB (${sheetName})`, rows });
    }
    if (respAuto.ok) {
      const buf = Buffer.from(await respAuto.arrayBuffer());
      const { sheetName, rows } = parseKvcMienBacAutoWorkbook(buf);
      if (sheetName) monthRowsBac.push({ sheetName: `MTĐ MB tự động (${sheetName})`, rows });
    }
    if (monthRowsBac.length > 0) {
      const result = chiPhiSheetRoutes.applyChiPhiMonthRowsToStore(store, monthRowsBac, "GG Sheet Mien Bac " + new Date().toISOString().slice(0, 10), "bac");
      if (result.added > 0) notes.push(`Sổ Chi Phí Miền Bắc: +${result.added} khoản chi mới.`);
    }
    if (!respManual.ok && !respAuto.ok) {
      notes.push(`Lỗi đọc cả 2 sheet Chi Phí Miền Bắc (mã lỗi ${respManual.status}/${respAuto.status}).`);
    }
  } catch (e) {
    notes.push(`Lỗi đọc sheet Chi Phí Miền Bắc (${e.message}).`);
  }

  // Chi Nhan, 2026-07-29: "đi tìm từ 3 unc gg sheet á lấy ra gian cho tôi" --
  // dong nao van chua co Gian Hang thi do tiep qua store.chi_phi (vua duoc
  // lam moi o tren) -- khop theo Số hóa đơn truoc (chac chan nhat), khong co
  // thi khop Tên NCC + Số tiền. Neu tim duoc gian, tra cuu tiep Hợp Đồng Thuê
  // Gian Hàng (Pháp Danh) de xac dinh Tài khoản Có (1388 neu la doanh thu
  // chia se, con lai 331) -- dung LAI logic co san o chi-phi.js
  // (computeTaiKhoanChiPhi) thay vi doan lai tu dau.
  let filledViaChiPhi = 0;
  const chiPhiForCompany = (store.chi_phi || []).filter((c) => c.congTy === activeCompany && (c.gian || "").trim());
  if (chiPhiForCompany.length > 0) {
    const gianListForTaiKhoan = store.phap_danh_hop_dong_thue || [];
    const aliasIndex = buildGianAliasIndex(gianListForTaiKhoan);
    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany || r.gianHang) return;
      const g = matchGianViaChiPhiLedger(r, chiPhiForCompany);
      if (!g) return;
      r.gianHang = g;
      filledViaChiPhi++;
      if (!r.taiKhoanCo) {
        const rec = findContractForGianText(g, gianListForTaiKhoan, aliasIndex);
        if (rec) {
          const chiaSe = isDoanhThuChiaSeRecord(rec);
          r.taiKhoanCo = chiaSe ? "1388" : "331";
          if (!r.hinhThucHopTac) r.hinhThucHopTac = chiaSe ? "CSE" : "Thuê";
        }
      }
    });
    if (filledViaChiPhi > 0) notes.push(`Dò Gian qua sổ Chi Phí (3 sheet UNC): +${filledViaChiPhi} dòng.`);
  } else {
    notes.push("Sổ Chi Phí chưa có dòng nào có Gian cho công ty này để dò thêm.");
  }

  save(store);
  res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent(notes.join(" ")));
});

// Chi Nhan, 2026-07-28: "thêm cột Đã chi tiền chưa ... theo tên NCC với số
// tiền á check thử cho tôi otừng hóa đơn điền vào bảng" -- doi chieu voi
// trang Đối soát Chi Phí (3 kenh ngan hang cua dung cong ty dang xem, tai su
// dung buildChannelChiPhi da co san) -- 1 hoa don duoc coi la "Đã chi" neu co
// it nhat 1 dong Chi Phi CUNG cong ty voi so tien Debit ~ bang Tong tien
// thanh toan cua hoa don (sai so <= 1.000d, tranh le lam tron) VA khop MST
// hoac Ten NCC. Chi BAT (true) cho hoa don MOI khop duoc, KHONG tu dong TAT
// lai hoa don da duoc chi tu danh dau "Đã chi" tay truoc do (tranh xoa nham
// xac nhan thu cong cua Chi Nhan).
router.post("/hoa-don-dau-vao/cap-nhat-da-chi", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    chiPhi.ensureShape(store);
    const revMap = chiPhi.buildRevenueStatusMap(store);
    const activeKeys = chiPhi.CHANNEL_KEYS.filter((ch) => chiPhi.CHANNELS[ch].company === activeCompany);
    let chiLines = [];
    const errors = [];
    activeKeys.forEach((ch) => {
      const built = chiPhi.buildChannelChiPhi(store, ch, revMap);
      if (built.error) errors.push(`${chiPhi.CHANNELS[ch].label}: ${built.error}`);
      else chiLines = chiLines.concat(built.lines);
    });

    // Chi Nhan, 2026-07-29: "thêm cho tôi trên cái ngân hàng có hiển thị cái
    // tên đối ứng ... để qua đối soát theo tên ncc cho tôi nhá và check xem
    // bên hóa đơn đã chi hay chưa á cập nhật vào đã chi hay chưa á" -- ngoai
    // nguon Chi Phi 3 kenh (chiLines o tren), do THEM qua TOAN BO giao dich
    // "chi" tren MOI tai khoan ngan hang thuong (store.transactions, trang
    // Ngan hang/Giao dich-Sao ke) cua dung cong ty -- khop theo Ten doi ung
    // (vua them cot nay) + So tien, KHONG can rieng 3 kenh Chi Phi hardcode
    // nua vi gio da co du lieu Ten doi ung tren MOI tai khoan.
    const bankIdsForCompany = new Set(
      (store.banks || []).filter((b) => (b.company || "kh_cu") === activeCompany).map((b) => b.id)
    );
    const chiTx = (store.transactions || []).filter(
      (t) => t.type === "chi" && bankIdsForCompany.has(t.bank_id) && (t.tenDoiUng || "").trim()
    );

    const rowsForCompany = store.hoa_don_dau_vao.filter((r) => r.congTy === activeCompany);
    const grouped = groupRowsByInvoice(rowsForCompany.map(ensureDefaults));
    let matchedGroups = 0;
    let matchedViaTx = 0;
    grouped.forEach((g) => {
      if (g.daChiTien) return; // da tu tay/lan truoc danh dau roi -- bo qua
      const foundInChiPhi = chiLines.some((l) => {
        if (Math.abs(l.debit - g.soTien) > 1000) return false;
        const mstMatch = g.mstNCC && l.mstNCC && l.mstNCC === g.mstNCC;
        const tenMatch =
          g.tenNCC && (l.tenNCC || l.vendor) && normVN(l.tenNCC || l.vendor).includes(normVN(g.tenNCC).slice(0, 12));
        return mstMatch || tenMatch;
      });
      const foundInTx =
        !foundInChiPhi &&
        g.tenNCC &&
        chiTx.some((t) => {
          if (Math.abs(t.amount - g.soTien) > 1000) return false;
          return normVN(t.tenDoiUng).includes(normVN(g.tenNCC).slice(0, 12));
        });
      if (foundInChiPhi || foundInTx) {
        const ids = g.idsCsv.split(",");
        store.hoa_don_dau_vao.forEach((r) => {
          if (ids.includes(String(r.id))) r.daChiTien = true;
        });
        if (foundInChiPhi) matchedGroups++;
        else matchedViaTx++;
      }
    });
    save(store);
    let msg = `Đã đối chiếu với Chi Phí (${activeKeys.map((ch) => chiPhi.CHANNELS[ch].label).join(", ") || "chưa có kênh nào"}): đánh dấu "Đã chi" cho ${matchedGroups} hóa đơn.`;
    if (matchedViaTx > 0) msg += ` Dò thêm qua Tên đối ứng trên sao kê ngân hàng: +${matchedViaTx} hóa đơn.`;
    if (errors.length > 0) msg += ` (Lỗi: ${errors.join("; ")})`;
    res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/hoa-don-dau-vao?error=" + encodeURIComponent(e.message));
  }
});
// Chi Nhan, 2026-07-28: backfill Phan loai/Tai khoan No (tu Dien giai) cho
// CAC DONG DA CO TU TRUOC (496 dong tai len truoc khi tinh nang nay ra doi) --
// dong nao MOI tao (nhap tay/upload) da tu duoc phan loai qua enrichNewRow
// roi, nut nay chi can bam 1 lan de phu cho du lieu cu. Chi dien vao dong
// dang TRONG phanLoai (khong ghi de dong Chi Nhan da tu sua).
router.post("/hoa-don-dau-vao/cap-nhat-phan-loai", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hangHoaList = store.hoa_don_dau_vao_hang_hoa_list || [];
  let filled = 0;
  store.hoa_don_dau_vao.forEach((r) => {
    if (r.congTy !== activeCompany || r.phanLoai) return;
    // Chi Nhan, 2026-07-29: uu tien khop danh muc Hang hoa/dich vu (chinh
    // xac tung mat hang, xem enrichNewRow) truoc khi roi ve doan tu khoa.
    const matchedHH = hangHoaList.length > 0 ? matchHangHoaRecord(r.dienGiai, hangHoaList) : null;
    if (matchedHH) {
      const classified = classifyFromHangHoaMatch(matchedHH, r.dienGiai, r.soTienTruocThue);
      r.tenHangHoaMisa = r.tenHangHoaMisa || classified.tenHangHoaMisa;
      r.phanLoai = classified.phanLoai;
      r.taiKhoanNo = classified.taiKhoanNo;
      filled++;
      return;
    }
    const { phanLoai, taiKhoanNo } = classifyPhanLoai(r.dienGiai, r.soTienTruocThue);
    if (phanLoai) {
      r.phanLoai = phanLoai;
      r.taiKhoanNo = taiKhoanNo;
      filled++;
    }
  });
  save(store);
  res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent(`Đã phân loại thêm ${filled} dòng đang trống.`));
});

module.exports = router;
// Chi Nhan, 2026-07-30: export de trang "Cong No NCC" (routes/congno-ncc.js)
// tai su dung DUNG logic gom dong theo hoa don + gia tri mac dinh cot, khong
// doan lai / khong bi lech so voi trang "Hoa Don Dau Vao" goc.
module.exports.ensureShape = ensureShape;
module.exports.ensureDefaults = ensureDefaults;
module.exports.groupRowsByInvoice = groupRowsByInvoice;
