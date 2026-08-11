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
// Luyen, 2026-08-01: "từ cái hợp đồng thuê gian á nó sẽ có tên đối tác ký hợp
// đồng với mình từ cái tên đó bạn map với lại hóa đơn đầu vào" -- dung lai
// danh sach hop dong thue gian (benChoThue/mstBenChoThue/tienThueThang/...)
// cua trang Phap Danh cho trang "Đối chiếu gian XHD, Tiền thuê" ben duoi,
// khong doan lai/copy schema.
const phapDanh = require("./phap-danh");

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
  // Luyen, 2026-08-01: "thêm cho tôi trong Hóa đơn đầu vào này 2 trang á 1 là
  // hóa đơn đầu vào như hiện tại 2 là Đối chiếu gian XHD, Tiền thuê nhá thêm
  // trang trc đi tôi sẽ diễn tả sao" -- tao truoc khung RONG (giong tien le
  // Phap danh/Hoa Don Dau Vao ban dau) de co cho luu du lieu ngay khi Luyen mo
  // ta chi tiet noi dung, chua doan truoc schema/cot cu the.
  if (!store.hoa_don_dau_vao_doi_chieu_gian_xhd) store.hoa_don_dau_vao_doi_chieu_gian_xhd = [];
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
  // Luyen, 2026-08-10: "thêm cái lọc theo hóa đơn tìm nhanh" -- o text search
  // theo So Hoa Don, khop partial (contains), khong phan biet hoa thuong.
  const soHdFilter = (req.query.soHd || "").trim();
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
  if (soHdFilter) groupedRows = groupedRows.filter((r) => String(r.soHoaDon || "").toLowerCase().includes(soHdFilter.toLowerCase()));
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
  if (soHdFilter) qs.push("soHd=" + encodeURIComponent(soHdFilter));
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
    soHdFilter,
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

// Luyen, 2026-08-01: "thêm cho tôi trong Hóa đơn đầu vào này 2 trang á 1 là
// hóa đơn đầu vào như hiện tại 2 là Đối chiếu gian XHD, Tiền thuê nhá" -- trang
// MOI thu 2 trong dropdown "Hóa Đơn Đầu Vào" (xem views/partials/nav.ejs).
// Luyen mo ta chi tiet (lan 2), 2026-08-01: "từ cái hợp đồng thuê gian á nó sẽ
// có tên đối tác ký hợp đồng với mình từ cái tên đó bạn map với lại hóa đơn
// đầu vào á nó sẽ có hóa đơn nội dung số tiền theo từng tháng map theo các ncc
// trong hợp đồng thuê gian thôi nhá đọc cái nội dung diễn giải của hóa đơn xem
// nó xuất cho mình là tiền thuê chi phí điện nước hay vượt doanh thu phân tích
// ra cho tôi nhá và cả đọc luôn cái hợp đồng note ra số tiền nếu tiền nó xuất
// có trong điều khoản hợp đồng cho tôi nhá" -- logic:
//   1. Chi xet cac hop dong thue gian (phap_danh_hop_dong_thue) DA CO ten Ben
//      cho thue (benChoThue) -- day la "cac NCC trong hop dong thue gian".
//   2. Voi moi hop dong, tim TAT CA hoa don o Hoa Don Dau Vao co Ten NCC khop
//      voi Ben cho thue do (uu tien khop qua MST neu ca 2 ben deu co, an toan
//      hon ten vi nhieu chi nhanh cung 1 chuoi -- vd AEON MALL cac diem khac
//      nhau -- co MST rieng; chi fallback ve so sanh ten khi thieu MST).
//   3. Voi moi hoa don khop duoc, doc Dien giai (da gop het cac dong cung 1 so
//      hoa don) phan loai theo tu khoa: "vượt doanh thu"/"phụ thu doanh thu" ->
//      Vuot doanh thu, "điện"/"nước" -> Chi phi dien nuoc, "thuê" -> Tien thue,
//      con lai -> Khac (CANH BAO: "thuê"/"thuế" deu ve "thue" sau khi bo dau,
//      co the lan -- Luyen xem lai neu thay sai).
//   4. Voi hoa don loai "Tien thue", so sanh so tien voi cac con so doc duoc tu
//      hop dong (tienThueThang + cac so trich tu tongTienThueThangHCM/
//      dieuKhoanThanhToan qua regex) -- khop thi bao "✓ Khớp", khong thi bao
//      "Không khớp" kem cac so hop dong co de Luyen tu doi chieu (dieu khoan co
//      the phuc tap, vd "10% doanh thu neu cao hon", khong doan chac chan duoc).
function extractAmountsFromText(text) {
  const amounts = new Set();
  const re = /\d{1,3}(?:[.,]\d{3})+|\d{6,}/g;
  const matches = String(text || "").match(re) || [];
  matches.forEach((m) => {
    const n = Number(m.replace(/[.,]/g, ""));
    if (Number.isFinite(n) && n >= 100000) amounts.add(n);
  });
  return Array.from(amounts);
}

function classifyGianXhdInvoiceType(dienGiaiText) {
  const t = normVN(dienGiaiText);
  if (!t) return "Khác";
  if (t.includes("vuot doanh thu") || t.includes("doanh thu vuot") || t.includes("phu thu doanh thu")) {
    return "Vượt doanh thu";
  }
  if (t.includes("dien nuoc") || (t.includes("tien dien") && t.includes("tien nuoc")) || t.includes("dien, nuoc")) {
    return "Chi phí điện nước";
  }
  if (t.includes("dien") || t.includes("nuoc")) return "Chi phí điện nước";
  if (t.includes("thue")) return "Tiền thuê";
  return "Khác";
}

// Luyen, 2026-08-01 (lan 4): "bạn đọc hợp đồng cho tôi xem nó của gian nào đi
// rồi đưa vô cái nào khớp thì điền khớp cho tôi đi chèn soa để khoong khớp
// hết vậy" -- gian "Go Nha Trang" hien "Không khớp" GAN NHU TAT CA dong Tien
// thue, vi 2 nguyen nhan CUNG luc:
//   1) tongTienThueThangHCM co NHIEU MOC GIA THEO THOI GIAN (vd "Từ
//      27/11/2025-28/02/2026: 80.000.100đ/tháng" roi "Từ 01/03/2026-...:
//      99.99.900đ/tháng") nhung code CU so hoa don voi TAT CA cac muc gia
//      CUNG luc (khong phan biet hoa don thang nao thi ap dung muc gia nao) --
//      hoa don thang 07/2026 (thuoc muc gia thu 2) khong bao gio khop duoc voi
//      so 80.000.100 cua muc gia thu 1.
//   2) Nhieu dong "Tiền thuê" thuc ra la dong "Điều chỉnh tăng/giảm ... từ
//      ngày X đến ngày Y" (dieu chinh 1 phan thang do doi muc gia giua thang,
//      KHONG PHAI hoa don tron thang) -- ban chat KHONG THE khop voi gia thue/
//      thang tron (vd -7.351.667đ hay +3.675.834đ), so voi gia tron thang se
//      LUON ra "Khong khop" oan uong. Doi sang tra ve null (khong ap dung, hien
//      dau "-") cho cac dong nay thay vi "Khong khop" sai.
// Ghi chu rieng cho Go Nha Trang: muc gia thu 2 "99.99.900đ/tháng" ghi trong
// sheet co ve THIEU 1 CHU SO (khong dung dinh dang 3-so-1-nhom binh thuong,
// vd đúng ra phải "999.999.900" hoặc "99.999.900") -- KHONG tu doan sua so nay,
// Luyen kiem tra lai voi hop dong goc (link Drive o hop dong) va sua truc tiep
// truong "Tiền thuê/tháng (sheet HCM)" o trang Hợp Đồng Thuê Gian Hàng neu sai.
function parseDdMmYyyyToIso(s) {
  const m = String(s || "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return "";
  let [, d, mo, y] = m;
  if (y.length === 2) y = "20" + y;
  return `${y.padStart(4, "0")}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

// Doc cac dong dang "Từ DD/MM/YYYY - DD/MM/YYYY: <số tiền>đ/tháng" trong
// tongTienThueThangHCM -- tra ve danh sach {tuNgay, denNgay, soTien} (ISO) de
// so hoa don theo DUNG khoang thoi gian ap dung, thay vi so voi TAT CA cac
// muc gia cung luc. Hop dong khong viet theo dang nay (vd chi 1 gia co dinh,
// hoac "22% tổng doanh thu") thi tra ve mang rong -- code goi ham nay se tu
// dong fallback ve cach so sanh cu (so voi tat ca so doc duoc trong van ban).
function parseRentTiers(text) {
  const tiers = [];
  const re = /Từ\s+(\d{1,2}\/\d{1,2}\/\d{2,4})\s*[-–]\s*(\d{1,2}\/\d{1,2}\/\d{2,4})\s*:\s*([^\n]+)/gi;
  let m;
  while ((m = re.exec(String(text || "")))) {
    const tuNgay = parseDdMmYyyyToIso(m[1]);
    const denNgay = parseDdMmYyyyToIso(m[2]);
    const amounts = extractAmountsFromText(m[3]);
    if (tuNgay && denNgay && amounts.length > 0) {
      tiers.push({ tuNgay, denNgay, soTien: amounts[0] });
    }
  }
  return tiers;
}

// Dong "Điều chỉnh tăng/giảm ... từ ngày X đến ngày Y" la dieu chinh 1 PHAN
// thang (thuong do doi gia giua thang) -- khong phai hoa don tron thang, so
// voi gia/thang se sai lech co y nghia, khong nen bao "Khong khop".
function isProratedRentAdjustment(dienGiaiText) {
  const t = normVN(dienGiaiText);
  return t.includes("dieu chinh tang") || t.includes("dieu chinh giam");
}

// Uu tien khop qua MST (an toan hon vi nhieu chi nhanh chung 1 ten cong ty me
// nhung MST rieng, vd cac diem AEON MALL khac nhau) -- CHI fallback ve so sanh
// ten khi 1 trong 2 ben thieu MST, va chi khop 1 chieu khi phan text dai >= 8
// ky tu (tranh khop nham qua tien to chung "cong ty tnhh...").
function nccMatchesLandlord(mstNCC, tenNCC, mstBenChoThue, tenBenChoThue) {
  const mstA = String(mstNCC || "").trim();
  const mstB = String(mstBenChoThue || "").trim();
  if (mstA && mstB) return mstA === mstB;
  const nameA = normVN(tenNCC);
  const nameB = normVN(tenBenChoThue);
  if (!nameA || !nameB) return false;
  if (nameA === nameB) return true;
  if (nameB.length >= 8 && nameA.includes(nameB)) return true;
  if (nameA.length >= 8 && nameB.includes(nameA)) return true;
  return false;
}

// Luyen, 2026-08-01 (lan 3): "theo ncc đi tên ncc á xong rồi sổ xuống check
// theo hóa đơn cho tôi nhá cả nam bắc vô chung ln nhá tách theo kh thôi nhá"
// -- bang "Hóa đơn khớp TRÙNG nhiều gian" truoc gio liet ke PHANG tung hoa don
// (cung 1 NCC lap lai nhieu dong lien tiep, kho quet mat), doi sang GOM NHOM
// theo Ten NCC (giong pattern collapsible da dung o Cong No NCC/chinh trang
// nay -- xem summarizeMissingInvoiceByNcc trong routes/congno-ncc.js), so
// xuong moi thay tung hoa don. Nam/Bac da tu gop chung san (contracts o duoi
// chi loc theo congTy, khong loc theo mien) -- CHI tach theo cong ty (KH Cu/
// KH Moi, qua nut "Cũ/Mới" tren dau trang, activeCompany) nhu Luyen yeu cau.
function summarizeAmbiguousByNcc(ambiguousInvoices) {
  const map = new Map();
  ambiguousInvoices.forEach((a) => {
    const key = normVN(a.tenNCC);
    if (!map.has(key)) {
      map.set(key, { tenNCC: a.tenNCC, soLan: 0, tongSoTien: 0, cacGianTrungSet: new Set(), items: [] });
    }
    const g = map.get(key);
    g.soLan++;
    g.tongSoTien += a.soTien;
    g.cacGianTrungSet.add(a.cacGianTrung);
    g.items.push(a);
  });
  return Array.from(map.values())
    .map((g) => ({
      tenNCC: g.tenNCC,
      soLan: g.soLan,
      tongSoTien: g.tongSoTien,
      // Da so truong hop 1 NCC luon trung CUNG 1 to hop gian -- chi khi nao
      // that su co nhieu to hop khac nhau moi can hien tung dong rieng ("(nhiều
      // tổ hợp gian khác nhau)"), tranh Luyen tuong nham la 1 to hop duy nhat.
      cacGianTrung: g.cacGianTrungSet.size === 1 ? Array.from(g.cacGianTrungSet)[0] : "(nhiều tổ hợp gian khác nhau, xem từng hóa đơn bên dưới)",
      items: g.items.slice().sort((a, b) => (b.ngayHD || "").localeCompare(a.ngayHD || "")),
    }))
    .sort((a, b) => b.tongSoTien - a.tongSoTien);
}

router.get("/hoa-don-dau-vao/doi-chieu-gian-xhd-tien-thue", (req, res) => {
  const store = load();
  ensureShape(store);
  phapDanh.ensureShape(store);
  const activeCompany = getCompany(req);

  const contracts = (store.phap_danh_hop_dong_thue || [])
    .map(phapDanh.ensureThueDefaults)
    .filter((r) => r.congTy === activeCompany && (r.benChoThue || "").trim());

  const invoiceRowsRaw = (store.hoa_don_dau_vao || []).map(ensureDefaults).filter((r) => r.congTy === activeCompany);
  const invoicesGrouped = groupRowsByInvoice(invoiceRowsRaw);
  const chiPhiForCompany = (store.chi_phi || []).filter((c) => c.congTy === activeCompany);

  // Luyen, 2026-08-01 (lan 2): "note ra gian nào dựa vào hợp đồng hay bên Chi
  // phí hay bên hóa đơn mua vào á do đây là hóa đơn mua vào" -- BUG phat hien
  // qua screenshot thuc te: 2 hop dong CUNG chung 1 ten rut gon (vd "AE Bình
  // Dương JP" ghi Ben cho thue rong la "BÌNH DƯƠNG", trung voi "AE Bình Dương
  // kvc" co Ben cho thue day du "...TẠI BÌNH DƯƠNG") bi khop NHAM CA 2 cho
  // CUNG 1 nhom hoa don khi chi so sanh ten/MST. Sua: xac dinh CHINH XAC 1 hop
  // dong cho MOI hoa don (khong phai nguoc lai) theo thu tu uu tien Luyen yeu
  // cau -- (a) chinh Gian Hang da co san tren hoa don ("bên hóa đơn mua vào",
  // xem enrichNewRow o tren), (b) bang Chi Phí ("bên Chi phí", qua
  // matchGianViaChiPhiLedger -- khop theo So hoa don hoac Ten NCC+So tien voi
  // store.chi_phi, tra ve ten gian NEU duy nhat), (c) cuoi cung moi roi ve so
  // sanh Ten NCC/MST voi Ben cho thue tren tung Hop Dong (nccMatchesLandlord)
  // -- CHI gan khi khop DUY NHAT 1 hop dong; khop >=2 hop dong (con nhap nhang)
  // thi KHONG gan vao gian nao ca (tranh nhan doi), liet ke rieng o
  // "ambiguousInvoices" de Luyen tu xem va bo sung aliasGian/sua Ben cho thue
  // cho ro hon.
  const aliasIndex = buildGianAliasIndex(contracts);
  const byContractId = new Map(contracts.map((c) => [c.id, []]));
  const ambiguousInvoices = [];

  invoicesGrouped.forEach((inv) => {
    let contract = null;
    if (inv.gianHang) contract = findContractForGianText(inv.gianHang, contracts, aliasIndex);
    if (!contract) {
      const gianNameFromChiPhi = matchGianViaChiPhiLedger(inv, chiPhiForCompany);
      if (gianNameFromChiPhi) {
        contract =
          contracts.find((c) => c.gian === gianNameFromChiPhi) ||
          findContractForGianText(gianNameFromChiPhi, contracts, aliasIndex);
      }
    }
    if (!contract) {
      const nameMatches = contracts.filter((c) => nccMatchesLandlord(inv.mstNCC, inv.tenNCC, c.mstBenChoThue, c.benChoThue));
      if (nameMatches.length === 1) {
        contract = nameMatches[0];
      } else if (nameMatches.length > 1) {
        ambiguousInvoices.push({
          soHoaDon: inv.soHoaDon,
          ngayHD: inv.ngayHD,
          tenNCC: inv.tenNCC,
          soTien: inv.soTien,
          dienGiai: inv.dienGiai,
          cacGianTrung: nameMatches.map((c) => c.gian).join(", "),
        });
        return;
      } else {
        return; // khong khop hop dong thue gian nao -- khong phai NCC can doi chieu o trang nay
      }
    }
    if (byContractId.has(contract.id)) byContractId.get(contract.id).push(inv);
  });

  const gianResults = contracts.map((c) => {
    const expectedAmounts = new Set(extractAmountsFromText(c.tongTienThueThangHCM));
    extractAmountsFromText(c.dieuKhoanThanhToan).forEach((a) => expectedAmounts.add(a));
    if (c.tienThueThang) expectedAmounts.add(Math.round(c.tienThueThang));
    const rentTiers = parseRentTiers(c.tongTienThueThangHCM);

    const matchedInvoices = (byContractId.get(c.id) || [])
      .map((inv) => {
        const loai = classifyGianXhdInvoiceType(inv.dienGiai);
        let khopHopDong = null; // null = khong ap dung so sanh (khong phai Tien thue, hoac la dong dieu chinh 1 phan thang)
        if (loai === "Tiền thuê" && !isProratedRentAdjustment(inv.dienGiai)) {
          // Uu tien so theo DUNG muc gia ap dung cho ngay cua hoa don (hop
          // dong nhieu muc gia theo thoi gian) -- chi fallback ve so voi TAT
          // CA cac so doc duoc trong van ban khi khong doc duoc muc gia theo
          // ngay (hop dong 1 gia co dinh, hoac dang % doanh thu...).
          const tier = rentTiers.find((t) => inv.ngayHD && inv.ngayHD >= t.tuNgay && inv.ngayHD <= t.denNgay);
          khopHopDong = tier
            ? Math.abs(tier.soTien - inv.soTien) <= 1000
            : Array.from(expectedAmounts).some((a) => Math.abs(a - inv.soTien) <= 1000);
        }
        return {
          soHoaDon: inv.soHoaDon,
          kyHieuHD: inv.kyHieuHD,
          ngayHD: inv.ngayHD,
          thang: (inv.ngayHD || "").slice(0, 7),
          dienGiai: inv.dienGiai,
          soTien: inv.soTien,
          loai,
          khopHopDong,
          linkHoaDon: inv.linkHoaDon,
        };
      })
      .sort((a, b) => (b.ngayHD || "").localeCompare(a.ngayHD || ""));

    const tongTheoLoai = { "Tiền thuê": 0, "Chi phí điện nước": 0, "Vượt doanh thu": 0, Khác: 0 };
    matchedInvoices.forEach((inv) => {
      tongTheoLoai[inv.loai] = (tongTheoLoai[inv.loai] || 0) + inv.soTien;
    });

    return {
      gian: c.gian,
      benChoThue: c.benChoThue,
      mstBenChoThue: c.mstBenChoThue,
      tienThueThang: c.tienThueThang,
      tongTienThueThangHCM: c.tongTienThueThangHCM,
      dieuKhoanThanhToan: c.dieuKhoanThanhToan,
      expectedAmounts: Array.from(expectedAmounts),
      rentTiers,
      invoices: matchedInvoices,
      tongTheoLoai,
    };
  });

  gianResults.sort((a, b) => b.invoices.length - a.invoices.length || a.gian.localeCompare(b.gian));

  res.render("hoa-don-dau-vao-doi-chieu-gian", {
    userName: req.session.userName,
    activeCompany,
    gianResults,
    ambiguousInvoices,
    ambiguousByNcc: summarizeAmbiguousByNcc(ambiguousInvoices),
    soGianCoHopDong: contracts.length,
    soGianCoHoaDon: gianResults.filter((g) => g.invoices.length > 0).length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
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
    // Luyen, 2026-08-01 (lan 5): phat hien qua thuc te (gian Go Nha Trang) --
    // khoa upsert dung "kyHieuHD" GAY RA 1191/5786 dong TRUNG THAT tren toan
    // bo du lieu: cung 1 hoa don that duoc xuat lai o file sau (vd "...(4).
    // xlsx" ngay 31/07) nhung cot Ký hiệu HĐ LAI RONG (khac file truoc "...(1).
    // xlsx" ngay 28/07 co dien "C26TNT") -- 2 khoa khac nhau nen bi coi la 2
    // hoa don khac nhau, tao THEM 1 dong thay vi cap nhat dong cu (xem seed
    // dedupeHoaDonDauVaoRows trong store.js -- da tu dong don dep cac dong
    // trung co san). Doi khoa: bo kyHieuHD (khong on dinh giua cac lan xuat
    // file), dung ngayHD thay the (luon co, on dinh) -- van giu dienGiai+
    // soTien lam phan biet chinh (xem bug fix 2026-07-28 o tren, van dung).
    const existingByKey = new Map();
    store.hoa_don_dau_vao.forEach((r) => {
      if (r.congTy !== activeCompany) return;
      const key = [r.congTy, r.soHoaDon, r.ngayHD || "", r.dienGiai || "", r.soTien].join("|");
      existingByKey.set(key, r);
    });
    let added = 0;
    let updated = 0;
    rows.forEach((r) => {
      const key = [activeCompany, r.soHoaDon, r.ngayHD || "", r.dienGiai || "", r.soTien].join("|");
      const existing = existingByKey.get(key);
      if (existing) {
        // Da co dong y het (cung hoa don + dien giai + so tien) -- khong tao
        // trung, chi dam bao ngay/NCC dong bo (thuong khong doi). Neu dong cu
        // dang thieu Ky hieu HD ma file moi co, dien bo sung (khong ghi de
        // neu da co, tranh mat du lieu neu file sau lai thieu).
        if (!existing.kyHieuHD && r.kyHieuHD) existing.kyHieuHD = r.kyHieuHD;
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
//
// Luyen, 2026-08-01: "đối chiếu trực tiếp với ngân hàng cho tôi luôn á đối
// chiếu tên ncc số tiền đồ cập nhật trực tiếp qua đây luôn á" -- doi UU TIEN:
// truoc day doi chieu qua Chi Phi (3 kenh) TRUOC, chi khi KHONG khop moi thu
// qua Ten doi ung tren sao ke ngan hang (foundInTx bi khoa boi "!foundInChiPhi").
// Doi lai: doi chieu THANG voi sao ke ngan hang (MOI tai khoan cua cong ty,
// khop Ten doi ung + So tien) la buoc CHINH, chay doc lap khong con phu
// thuoc Chi Phi nua; Chi Phi (3 kenh) chi con la nguon DU PHONG cho cac
// truong hop hiem sao ke chua co Ten doi ung nhung Chi Phi da ghi nhan.
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
      // Buoc CHINH: doi chieu THANG voi sao ke ngan hang (Ten doi ung + So
      // tien), doc lap khong con cho Chi Phi truoc nua.
      const foundInTx =
        g.tenNCC &&
        chiTx.some((t) => {
          if (Math.abs(t.amount - g.soTien) > 1000) return false;
          return normVN(t.tenDoiUng).includes(normVN(g.tenNCC).slice(0, 12));
        });
      // Du phong: chi kiem tra Chi Phi (3 kenh) khi ngan hang chua khop duoc.
      const foundInChiPhi =
        !foundInTx &&
        chiLines.some((l) => {
          if (Math.abs(l.debit - g.soTien) > 1000) return false;
          const mstMatch = g.mstNCC && l.mstNCC && l.mstNCC === g.mstNCC;
          const tenMatch =
            g.tenNCC && (l.tenNCC || l.vendor) && normVN(l.tenNCC || l.vendor).includes(normVN(g.tenNCC).slice(0, 12));
          return mstMatch || tenMatch;
        });
      if (foundInTx || foundInChiPhi) {
        const ids = g.idsCsv.split(",");
        store.hoa_don_dau_vao.forEach((r) => {
          if (ids.includes(String(r.id))) r.daChiTien = true;
        });
        if (foundInTx) matchedViaTx++;
        else matchedGroups++;
      }
    });
    save(store);
    let msg = `Đã đối chiếu trực tiếp với ngân hàng (Tên đối ứng + Số tiền, mọi tài khoản của ${activeCompany === "kh_moi" ? "KH Mới" : "KH Cũ"}): đánh dấu "Đã chi" cho ${matchedViaTx} hóa đơn.`;
    if (matchedGroups > 0) msg += ` Dò thêm qua Chi Phí (${activeKeys.map((ch) => chiPhi.CHANNELS[ch].label).join(", ") || "chưa có kênh nào"}): +${matchedGroups} hóa đơn.`;
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

// Luyen, 2026-08-11: Chay lai phan loai TU DONG cho TAT CA dong (ke ca dong
// da co phan loai cu) theo logic moi nhat -- dung khi vua cap nhat keyword
// (vd NVL: duong/nuoc da/kem beo, dich vu: cuoc chuyen phat...).
// KHONG ghi de dong da duoc CHINH TAY (phanLoaiManual === true).
router.post("/hoa-don-dau-vao/re-phan-loai-tat-ca", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const hangHoaList = store.hoa_don_dau_vao_hang_hoa_list || [];
  let changed = 0;
  store.hoa_don_dau_vao.forEach((r) => {
    if (r.congTy !== activeCompany) return;
    // Bo qua dong da chinh tay (user tu sua bang nut Sua)
    if (r.phanLoaiManual) return;
    const matchedHH = hangHoaList.length > 0 ? matchHangHoaRecord(r.dienGiai, hangHoaList) : null;
    let newPL, newTK, newTen;
    if (matchedHH) {
      const classified = classifyFromHangHoaMatch(matchedHH, r.dienGiai, r.soTienTruocThue);
      newPL = classified.phanLoai;
      newTK = classified.taiKhoanNo;
      newTen = classified.tenHangHoaMisa;
    } else {
      const cl = classifyPhanLoai(r.dienGiai, r.soTienTruocThue);
      newPL = cl.phanLoai;
      newTK = cl.taiKhoanNo;
    }
    if (newPL && (r.phanLoai !== newPL || r.taiKhoanNo !== newTK)) {
      r.phanLoai = newPL;
      r.taiKhoanNo = newTK;
      if (newTen) r.tenHangHoaMisa = r.tenHangHoaMisa || newTen;
      changed++;
    }
  });
  save(store);
  res.redirect("/hoa-don-dau-vao?success=" + encodeURIComponent(`Đã cập nhật lại ${changed} dòng theo phân loại mới nhất.`));
});

module.exports = router;
// Chi Nhan, 2026-07-30: export de trang "Cong No NCC" (routes/congno-ncc.js)
// tai su dung DUNG logic gom dong theo hoa don + gia tri mac dinh cot, khong
// doan lai / khong bi lech so voi trang "Hoa Don Dau Vao" goc.
module.exports.ensureShape = ensureShape;
module.exports.ensureDefaults = ensureDefaults;
module.exports.groupRowsByInvoice = groupRowsByInvoice;
