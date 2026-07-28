const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const {
  extractVqrCode,
  extractVietQrSettlements,
  extractVietQrThuTransactions,
  parseVietQrRawWorkbook,
  parseCuaHangSheet,
  parseStoreExportSheet,
  parseTenDiemMaCongTrinhSheet,
  parseInvoiceWorkbookByTag,
  buildGianCandidatesFromInvoices,
  resolveGianGross,
  resolveGianGrossByBankRef,
  reconcileVietQr,
  parseVietQrMnRawWorkbook,
  parseMaCuaHangAppSheet,
  resolveGianGrossPrefix,
  extractStorePrefix,
  isoToDmy,
  displayCode,
  FF_SUFFIX,
} = require("../utils/vietqrReconcile");
const {
  mergeGianListWithMaster,
  applyGianRedirectToInvoices,
  buildOnlineProductMatcher,
  normText,
} = require("../utils/zvpReconcile");
const { getCompany } = require("../utils/companies");
const { findBestMaCongTrinh } = require("../utils/maCongTrinh");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

// Each Viet QR bank posts individual QR payments (no batch settlement), so
// the reconciliation is per calendar day per bank. Store.transactions'
// bank.name is the join key against store.banks; the tag pattern is how a
// shared invoice-list upload (same "MTT" file used by Momo/ZVP) tells this
// bank's invoices apart from the others' -- see utils/vietqrReconcile.js.
const CHANNELS = {
  bidv7704: { bankName: "BIDV7704", label: "BIDV 7704", tagPattern: /POSH\+JP\s*MB\s*\(7704\)/i, company: "kh_cu" },
  bidv77020: { bankName: "BIDV77020", label: "BIDV 77020", tagPattern: /POSH\+JP\s*MB\s*\(7020\)/i, company: "kh_cu" },
  mb11521268: { bankName: "MB11521268", label: "MB 11521268", tagPattern: /POSH\+JP\s*MB\s*\(268\)/i, company: "kh_cu" },
  // Cong ty "KH Moi" (TNHH GIAI TRI K&H) -- tai khoan Viet QR BIDV7702. Tag
  // hoa don la "MTD MN" (xac nhan tu Luyen 2026-07-16, nam tren sheet "ke ds
  // xuat HD MTT - 705" cua file MTT dung chung -- xem parseInvoiceWorkbookByTag
  // trong utils/vietqrReconcile.js).
  //
  // Luyen, 2026-07-17: chuyen sang dung truc tiep file export chuan "Transactions"
  // (giao dich QR, co cot "Noi dung TT"/"Ma tham chieu") + file "store-export"
  // rieng (danh muc Ma cua hang -> Ten diem ban, ten that vd "Posh Lotte Nam
  // Sai Gon" -- KHONG con la chu viet tat cua Ma cong trinh nhu file "VIETQR MN
  // 7702.xlsx" cu nua) -- giong het 3 kenh kia, nen bo parseMode "mn": dung
  // parser chuan (parseVietQrRawWorkbook/parseCuaHangSheet) va khop gian bang
  // fuzzy text (resolveGianGross) thay vi khop tien to (resolveGianGrossPrefix),
  // vi ten cua hang gio la ten mo ta that, khong con la tien to Ma cong trinh
  // nua nen khop tien to se khong ra ket qua nao.
  bidv7702: {
    bankName: "BIDV7702",
    label: "BIDV 7702",
    tagPattern: /MTD\s*MN/i,
    company: "kh_moi",
    // Luyen, 2026-07-18: "gian nao trong diem ban dua vo tan phu nha BIDV TAN
    // PHU" -- giao dich KHONG co ma cua hang nao het (vd chuyen khoan thuong
    // tu Liobank, khong qua QR) tu dong cong vao "AM TP PHCM" thay vi nam mai
    // trong unmapped, ap dung luon ca cho du lieu tai len sau nay.
    defaultBlankCode: "AM TP PHCM",
    // Chi Nhan, 2026-07-24 (Luyen yeu cau): "từ 22/07 trở về trước giữ nguyên
    // dữ liệu ... từ 22 trở đi" -- tu ngay nay tro di, khop gian bang "So tham
    // chieu" ngan hang (khop tuyet doi) thay vi fuzzy text nhu truoc, xem
    // resolveGianGrossByBankRef trong utils/vietqrReconcile.js. Cac ngay TRUOC
    // ngay nay van dung y nguyen cach cu (resolveGianGross ben tren).
    refMatchFrom: "2026-07-22",
  },
  // Chi Nhan, 2026-07-28: "giống như KH mới 7702 á làm cho tôi 77021 á" --
  // tai khoan BIDV moi (8690077021, cung phap nhan CONG TY TNHH GIAI TRI
  // K&H/kh_moi nhu 7702), Luyen xac nhan hoa don kenh nay duoc dan tag rieng
  // "VietQR POSH MB" (KHAC voi "MTD MN" cua 7702) tren cot "Dich vu thu ho"
  // cua file MTT dung chung. La kenh hoan toan MOI (khong co giai doan du
  // lieu cu truoc do) nen dung ngay tinh nang khop theo "So tham chieu" ngan
  // hang (refMatchFrom) ngay tu dau, khong can giai doan fuzzy-match nhu 7702.
  bidv77021: {
    bankName: "BIDV77021",
    label: "BIDV 77021",
    tagPattern: /VietQR\s*POSH\s*MB/i,
    company: "kh_moi",
    refMatchFrom: "2026-01-01",
  },
};
const CHANNEL_KEYS = Object.keys(CHANNELS);

const TKCO_VALUES = ["131", "1388", "SKIP"];
const UPDATED_NOTE = " Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.";

// Full gian merge/rename (Luyen-confirmed 2026-07-16): "JP SC VIVO" la CUNG
// 1 gian voi "SC VIVO KVCM", chia se doanh thu (CSE) tu dau -- hoa don ghi
// "JP SC VIVO" cho cac ngay dau, roi doi sang ghi "SC VIVO KVCM"/"FUNFEST
// SCVIVO" tu khoang 1-2 ngay sau (do do tre xuat hoa don). Merge permanent,
// ap dung ca hoa don cu (con ghi "JP SC VIVO") lan hoa don moi ve sau.
const GIAN_MERGE_DEFAULTS = {
  "JP SC VIVO": { maCongTrinh: "SC VIVO KVCM", isCse: true },
  // Luyen, 2026-07-24: "2 cái Phan Thiết này là 1 á cộng lại cho tôi nha" --
  // "FARM LOTTE PHAN THIET" chi la 4 ma cua hang (gan qua nut "Gan ma cong
  // trinh" cho ma cua hang MOI) thuc ra la CUNG 1 diem voi "LOTTE PHAN
  // THIET", khong phai gian rieng. Gop vinh vien vao "LOTTE PHAN THIET".
  "FARM LOTTE PHAN THIET": { maCongTrinh: "LOTTE PHAN THIET", isCse: false },
};

// Luyen, 2026-07-20: mot so ten "Mã công trình" hien tren bang doi soat Viet
// QR (lay tu Ma diem tren hoa don/QR) khac qua xa so voi ten trong danh sach
// "Mã công trình chuẩn" (vd co them chu thich "(EB Tân Phú)", "ghế"...) nen
// findBestMaCongTrinh (khop tuyet doi/gan giong bang chua chuoi) khong tu tim
// ra duoc, phai de nguyen ten cu. Luyen xac nhan tay tung truong hop qua chat
// (anh chup danh sach cong trinh) -- luu lai o day de ap dung vinh vien, cung
// co che voi GIAN_MERGE_DEFAULTS o tren nhung CHI doi TEN HIEN THI (khong
// dong den code/join key dung de doi soat/gop nhom doanh thu).
// "Go Âu Cơ (EB Tân Phú)" - ghế: Luyen xac nhan 2026-07-20 day la diem cua
// Posh, dung ma chuan "GO AU CO PHCM" (xem danh sach cong trinh KH Moi).
// "AE Tân Phú ghế" -> "AM TP PHCM", "Lotte mart Nam Sài Gòn" -> "LOTTE Q7
// (NSG) PHM": Luyen xac nhan them 2026-07-20.
// "LM NHA TRANG KVC" -> "FARM LOTTE NHA TRANG", "JP vicom 3.2 BIDV" -> "VC
// 3/2 JP-Posh", "VINCOM GAND PARK" -> "JP-POSH GRAND PARK": Luyen xac nhan
// 2026-07-24 qua anh chup danh sach cong trinh chuan (Danh_sach_cong_trinh.xlsx,
// KH Moi) -- 2 ma dau la khop tuyet doi trong danh sach nhung ten thuc te qua
// khac nen findBestMaCongTrinh khong tu tim ra; "VINCOM GAND PARK" la nham
// (chinh la diem "JP-POSH GRAND PARK" da co ma, khong phai diem rieng).
const MA_CONG_TRINH_DISPLAY_ALIAS_DEFAULTS = {
  "Go Âu Cơ (EB Tân Phú)": "GO AU CO PHCM",
  "AE Tân Phú ghế": "AM TP PHCM",
  "Lotte mart Nam Sài Gòn": "LOTTE Q7 (NSG) PHM",
  "LM NHA TRANG KVC": "FARM LOTTE NHA TRANG",
  "JP vicom 3.2 BIDV": "VC 3/2 JP-Posh",
  "VINCOM GAND PARK": "JP-POSH GRAND PARK",
};

// Chi Nhan, 2026-07-24: giong het KNOWN_INVOICE_DIEM_ALIASES ben routes/
// doisoat.js (Momo) -- Luyen xac nhan qua anh chup + doi chieu so tien/ngay
// khop tuyet doi (BIDV7702 ngay 22/07: SENSE CT PVĐ 500k, POSH LOTTE PHÚ THỌ
// 230k, POSH MN GALAXY KINH DƯƠNG VƯƠNG 80k, POSH MN GALAXY QUANG TRUNG 60k,
// KNG BÀ RỊA 70k -- deu khop dung gross cua dung gian nay). SEED SAN qua
// ensureChannelShape (chay MOI LAN load()) thay vi chi ghi 1 lan vao
// store.json truc tiep -- vi Luyen dang chay server local rieng (khong chung
// tien trinh voi may cua toi), moi lan Luyen tai file/luu gi do server cua
// Luyen se ghi de store.json bang ban store CU trong bo nho no dang giu
// (chua thay 5 dong nay), lam mat trang 3 lan lien tiep du toi da them thu
// cong qua script. Seed qua code (chay lai moi request, giong GIAN_MERGE_DEFAULTS/
// MA_CONG_TRINH_DISPLAY_ALIAS_DEFAULTS o tren) thi KHONG THE mat duoc nua, chi
// can server cua Luyen restart 1 lan de nap code moi la tu dong co lai vinh vien.
// Luyen, 2026-07-24 (lan 2): "hải phòng có 2 cái ngày 23 mà" -- Luyen xac
// nhan qua anh chup hoa don: "AM HP KVCN" (Funzone Hai Phong, hoa don 2297)
// la CSE that, NHUNG "SNOWFUN AEON HẢI PHÒNG" (hoa don 2298, san pham "SNOW
// FUN" -- KHAC voi "Funzone" -- KHONG danh dau CSE tren sheet cua chi) la
// mot san pham RIENG, KHONG CSE, cung ma cong trinh AM HP KVCN nhung khac TK
// (131). Da thu doi alias nay sang "AM HP KVCN__FF" (lan sua truoc, SAI) --
// tra lai dung "AM HP KVCN" (KHONG __FF) de khop voi dong PLAIN 360.000d
// (dung bang gross cua rieng san pham "...SALE 20% - SNOW FUN" -- xac nhan
// qua doi chieu parseFeeReportWorkbook, khop chinh xac 360.000d).
const INVOICE_DIEM_ALIAS_DEFAULTS = {
  "SENSE CT PVĐ PHCM": "SENSE PVD PHCM",
  "POSH LOTTE PHÚ THỌ": "LOTTE PHU THO PHCM",
  "POSH MN GALAXY KINH DƯƠNG VƯƠNG": "GALAXY KINH DUONG VUONG PHCM",
  "POSH MN GALAXY QUANG TRUNG": "GALAXY QUANG TRUNG PHCM",
  "KNG BÀ RỊA": "KNG BA RIA PHCM",
  "SNOWFUN AEON HẢI PHÒNG": "AM HP KVCN",
};

function ensureChannelShape(store) {
  if (!store.viet_qr_raw_uploads) store.viet_qr_raw_uploads = {};
  if (!store.viet_qr_store_names) store.viet_qr_store_names = {};
  if (!store.viet_qr_invoices) store.viet_qr_invoices = {};
  if (!store.viet_qr_manual_matches) store.viet_qr_manual_matches = {};
  if (!store.viet_qr_gian_merge) store.viet_qr_gian_merge = {};
  if (!store.viet_qr_nocode_assignments) store.viet_qr_nocode_assignments = {};
  if (!store.ma_cong_trinh_display_alias) store.ma_cong_trinh_display_alias = {};
  // Chi Nhan, 2026-07-22: "nhấn nhầm nạp nhầm chỗ này mà hk có nút xóa hay
  // lịch sử" -- upload "Danh sách điểm bán riêng" (store_export) truoc day
  // ghi THANG (Object.assign) vao viet_qr_store_names, khong luu lai lich su
  // nen khong the xoa/hoan tac 1 lan tai nham. Gio luu tung lan tai vao
  // viet_qr_store_uploads[channel] (giong het cach lam voi viet_qr_raw_uploads),
  // roi TINH LAI viet_qr_store_names[channel] = baseline + gop tat ca cac lan
  // tai con lai theo thu tu thoi gian (xem mergeStoreNames) -- xoa 1 lan tai
  // nao se tu dong tinh lai dung, khong con dinh lien vao lan do nua.
  // viet_qr_store_names_baseline chi duoc "chup" 1 LAN DUY NHAT (luc tinh nang
  // nay moi trien khai) de giu lai toan bo du lieu da co truoc do (tu cac lan
  // tai file "Dữ liệu Viet QR" hoac sua tay tung ma) -- khong bao gio ghi de
  // lai sau do, vi cac thay doi tu nhung nguon KHAC (khong phai store_export)
  // van tiep tuc ghi truc tiep vao viet_qr_store_names nhu cu.
  if (!store.viet_qr_store_uploads) store.viet_qr_store_uploads = {};
  if (!store.viet_qr_store_names_baseline) store.viet_qr_store_names_baseline = {};
  // Chi Nhan, 2026-07-24: bang tra cuu "Ten diem ban -> Ma cong trinh" rieng
  // cho tinh nang khop theo So tham chieu ngan hang (BIDV7702 tu 22/07) --
  // key da chuan hoa qua normText(tenDiem), xem parseTenDiemMaCongTrinhSheet/
  // resolveGianGrossByBankRef trong utils/vietqrReconcile.js. Ghi de truc
  // tiep (khong luu lich su tung lan tai, khac voi viet_qr_store_uploads) --
  // giam pham vi, Luyen chua yeu cau xem lai/hoan tac rieng bang nay.
  if (!store.viet_qr_ten_diem_master) store.viet_qr_ten_diem_master = {};
  // Luyen, 2026-07-24: "mã cửa hàng mới này ... chỗ chọn mã công trình để gán
  // vào nhá" -- gan THANG 1 Ma cua hang -> 1 Ma cong trinh ngay tu bang canh
  // bao "Ma cua hang MOI, chua co trong danh sach diem ban" (xem
  // resolveGianGrossByBankRef trong utils/vietqrReconcile.js), khong can doi
  // upload lai file "Danh sach diem ban" + file "Ten diem - Ma cong trinh".
  if (!store.viet_qr_store_code_override) store.viet_qr_store_code_override = {};
  // Luyen, 2026-07-27: "100k của dư ngân hàng nếu chưa có bên dữ liệu thì
  // đưa vô gian AMTP cho tôi" -- gan THANG 1 So tham chieu ngan hang -> 1 Ma
  // cong trinh, dung khi giao dich ngan hang khong co dong QR nao khop (se
  // bi tru khoi "Ngan hang" o buildChannelReconciliation neu khong gan),
  // nhung Luyen tu xac dinh duoc dung gian. Xem resolveGianGrossByBankRef.
  if (!store.viet_qr_ref_override) store.viet_qr_ref_override = {};
  // Chi Nhan, 2026-07-27: "thêm cho tôi 1 nút cập nhật viet qr ... vào ngân
  // hàng nhận tiền như 7702 chọn thời gian đối soát rồi tải về" -- moi kenh
  // (BIDV7704/77020/MB11521268/BIDV7702) ung voi 1 tai khoan rieng tren cong
  // doi tac doitac.vietqr.vn, can luu lai dung "bankId" cua tai khoan do (chi
  // tu lay tren trang doi tac, dan vao 1 lan) de nut "Cap nhat VietQR" mo
  // dung tai khoan + dung ngay hom nay, khong phai tu chon lai moi lan. Chi
  // luu ID tai khoan (khong phai mat khau/API key), Chi Nhan van tu dang
  // nhap + bam "Xuat Excel" + tai file len nhu cu (chua tu dong tai/day file).
  if (!store.viet_qr_partner_bank_id) store.viet_qr_partner_bank_id = {};
  // Luyen, 2026-07-21: "không cần chỉnh cái cũ khóa cho tôi" -- muon 1 tinh
  // nang "khoa so" that su (tung yeu cau 2 lan truoc: "khóa sổ cho tôi chỉ
  // nạp cái mới thôi"), khong phai sua tay tung dong lech cu. Luu 1 ngay
  // "khoa den het ngay X" (bao gom ca ngay X) THEO TUNG KENH -- moi ngay
  // settlementDate <= ngay nay se duoc buildChannelReconciliation coi nhu
  // "da khoa" (xem xu ly ben duoi ham buildChannelReconciliation): an het
  // canh bao "Chua co HD"/"Lech", khong con hien len de xu ly nua, nhung VAN
  // giu nguyen so lieu goc (gross/invoiceTotal) de xem lai neu can, chi doi
  // cach hien thi/tinh tong "Lech".
  if (!store.viet_qr_lock_date) store.viet_qr_lock_date = {};
  Object.keys(GIAN_MERGE_DEFAULTS).forEach((k) => {
    if (!store.viet_qr_gian_merge[k]) store.viet_qr_gian_merge[k] = GIAN_MERGE_DEFAULTS[k];
  });
  if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
  Object.keys(INVOICE_DIEM_ALIAS_DEFAULTS).forEach((k) => {
    const cur = store.invoice_diem_alias[k];
    // Ghi de ca truong hop tu-anh-xa-chinh-no (vd "SENSE CT PVĐ PHCM" ->
    // "SENSE CT PVĐ PHCM") -- day la dau vet cua 1 lan luu form "Anh xa"
    // truoc khi biet dung ma nao (hoac ban store cu bi ghi de lai qua race
    // condition, xem ghi chu tai INVOICE_DIEM_ALIAS_DEFAULTS), khong phai
    // Luyen co y dinh giu nguyen ten do.
    if (!cur || cur === k) {
      store.invoice_diem_alias[k] = INVOICE_DIEM_ALIAS_DEFAULTS[k];
    }
  });
  // Luyen, 2026-07-24 (lan 2): dao nguoc lai sua sai truoc do -- neu ban store
  // nao van con gia tri __FF (tu lan sua dau, da xac nhan la SAI), tra ve
  // "AM HP KVCN" (KHONG __FF) dung nhu INVOICE_DIEM_ALIAS_DEFAULTS o tren.
  if (store.invoice_diem_alias["SNOWFUN AEON HẢI PHÒNG"] === "AM HP KVCN__FF") {
    store.invoice_diem_alias["SNOWFUN AEON HẢI PHÒNG"] = "AM HP KVCN";
  }
  Object.keys(MA_CONG_TRINH_DISPLAY_ALIAS_DEFAULTS).forEach((k) => {
    if (!store.ma_cong_trinh_display_alias[k]) {
      store.ma_cong_trinh_display_alias[k] = MA_CONG_TRINH_DISPLAY_ALIAS_DEFAULTS[k];
    }
  });
  CHANNEL_KEYS.forEach((ch) => {
    if (!store.viet_qr_raw_uploads[ch]) store.viet_qr_raw_uploads[ch] = [];
    if (!store.viet_qr_store_names[ch]) store.viet_qr_store_names[ch] = {};
    if (!store.viet_qr_invoices[ch]) store.viet_qr_invoices[ch] = [];
    if (!store.viet_qr_manual_matches[ch]) store.viet_qr_manual_matches[ch] = {};
    if (!store.viet_qr_nocode_assignments[ch]) store.viet_qr_nocode_assignments[ch] = {};
    if (!store.viet_qr_store_uploads[ch]) store.viet_qr_store_uploads[ch] = [];
    if (!store.viet_qr_ten_diem_master[ch]) store.viet_qr_ten_diem_master[ch] = {};
    if (!store.viet_qr_ref_override[ch]) store.viet_qr_ref_override[ch] = {};
    if (!store.viet_qr_store_names_baseline[ch]) {
      // Chup 1 lan duy nhat: du lieu diem ban HIEN CO ngay truoc khi tinh
      // nang lich su/xoa nay ton tai, de khong mat du lieu cu.
      store.viet_qr_store_names_baseline[ch] = Object.assign({}, store.viet_qr_store_names[ch]);
    }
  });
}

// Gop danh sach diem ban tu nhieu lan tai "store_export" (moi lan la 1 object
// {maCuaHang: {...}}) theo thu tu THOI GIAN (cu truoc, moi sau) chong len
// baseline -- lan tai MOI HON de nguoi thang neu cung 1 ma cua hang xuat hien
// o nhieu lan tai. Xoa 1 lan tai (bo entry do khoi mang roi goi lai ham nay)
// se tu dong tinh lai dung ma khong con dinh lieu cua lan da xoa.
function mergeStoreNames(uploads, baseline) {
  let merged = Object.assign({}, baseline || {});
  const sorted = [...uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  for (const u of sorted) {
    merged = Object.assign(merged, u.map || {});
  }
  return merged;
}

// Merge raw-upload rows across multiple uploads the same way Momo/ZVP merge
// gross uploads: newest upload wins for any (date, ma cua hang) key it
// covers -- de-dup by (vqrCode) across uploads, newest first.
//
// IMPORTANT: some banks (verified on MB11521268 -- its "Noi dung TT" column
// is always literally "PaymentForOrder", no "VQR..." token at all) never
// produce a vqrCode, so EVERY row from that bank fell into a single flat
// "noCode" bucket with no de-dup at all. Luyen's workflow re-uploads a fresh
// CUMULATIVE "to date" export (not just the new days) each time, so any date
// covered by both the old and the new upload used to get counted TWICE --
// this was found causing exactly-2x gross vs invoice on every single-day
// line for MB11521268's "CHKQT CAM RANH" gian (confirmed against real data:
// every date in both uploads' overlapping range showed gross = 2 x invoice
// total, while dates only in the newer upload matched 1:1). Fixed by
// de-duping no-vqrCode rows PER DATE instead: the newest upload that
// contains ANY row for a given date fully replaces earlier uploads' rows
// for that same date, rather than appending both.
function mergeRawRows(uploads) {
  const sorted = [...uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  const byVqr = {};
  const noCodeByDate = {};
  for (const u of sorted) {
    const datesInThisUpload = new Set();
    for (const row of u.rows || []) {
      if (row.vqrCode) {
        // Luyen, 2026-07-19: mot so may/diem ban dung QR TINH (1 ma QR dan
        // san, quet lai nhieu lan cho nhieu khach khac nhau) thay vi QR dong
        // -- vd ma cua hang "45U0h5N0VU" (AE Tan Phu ghe) tai su dung DUNG 1
        // vqrCode cho nhieu giao dich THAT, KHAC ngay VA khac so tien (07-13:
        // 50k va 20k, 07-14: 20k, 07-15: 50k). De-dup CHI theo vqrCode (nhu
        // truoc day) khien MOI lan trung ma, ban ghi MOI GHI DE len ban ghi
        // CU -- khong chi trong cung 1 ngay ma con XOA LUON giao dich cua
        // NGAY KHAC, lam mat that 120.000d doanh thu (xac nhan khop voi file
        // "chia theo ma cua hang" cua Luyen: cot Tong QR Posh+JP HCM ghi
        // 160.000d cho ma nay, dung bang tong 4 giao dich rieng biet).
        //
        // Van chua du: 07-14 van thieu dung 20.000d sau fix tren -- tim ra 2
        // dong CUNG ma QR, CUNG ngay, CUNG so tien (20k) nhung "raw" (noi
        // dung goc) HOAN TOAN khac nhau -- 1 dong la ghi nhan QR ngan gon
        // ("137625171183 0365585316 VQR..."), dong kia la 1 giao dich CHUYEN
        // KHOAN ngan hang day du ("MBVCB....CT tu 0331000460371 VO PHA LUAN
        // toi V3BLC8640107702...") -- ro rang la 2 giao dich THAT khac nhau
        // vo tinh trung ma QR tinh do, khong phai 1 ban ghi xuat lai 2 lan
        // (neu la xuat lai that thi "raw" phai giong het nhau). Nen dung ca
        // "raw" lam 1 phan cua key de-dup: chi gop lai khi ca vqrCode, ngay,
        // so tien VA noi dung goc deu giong nhau (dung 1 ban ghi xuat lai),
        // con khac raw thi la 2 giao dich rieng, giu ca hai.
        const key = row.vqrCode + "|" + (row.date || "") + "|" + row.amount + "|" + (row.raw || "");
        byVqr[key] = row;
      } else if (row.date) {
        datesInThisUpload.add(row.date);
      }
    }
    // Newest upload wins: wipe any earlier noCode rows for a date this
    // upload also covers, then refill with this upload's own rows for it.
    datesInThisUpload.forEach((d) => {
      noCodeByDate[d] = [];
    });
    for (const row of u.rows || []) {
      if (!row.vqrCode && row.date) {
        noCodeByDate[row.date].push(row);
      }
    }
  }
  const noCode = Object.values(noCodeByDate).flat();
  return [...Object.values(byVqr), ...noCode];
}

function seedGianMappingDefaults(store, codes) {
  codes.forEach((c) => {
    if (!(c in store.gian_mapping)) {
      // Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- TK Co 1388
      // (doanh thu chia se/CSE) khong con duoc dung nua, moi gian moi deu
      // mac dinh 131.
      store.gian_mapping[c] = "131";
    }
  });
}

// Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- gian_mapping dung
// chung; lap lai o day de trang VietQR cung tu sua duoc du duoc mo truoc
// trang Momo/ZVP.
function ensureNo1388(store) {
  if (!store.gian_mapping) return false;
  let changed = false;
  for (const code of Object.keys(store.gian_mapping)) {
    if (store.gian_mapping[code] === "1388") {
      store.gian_mapping[code] = "131";
      changed = true;
    }
  }
  return changed;
}

// Rows from the shared daily master "gian " sheet (uploaded on the Zalo/
// VNPay/Payoo page, "0. Tai bang gian tong hop hang ngay") whose "Thuoc"
// column matches THIS Viet QR channel's own invoice tag pattern (e.g.
// "POSH+JP MB (7704) 11,12" for BIDV7704) -- reusing the exact same regex
// already used to tell this channel's invoices apart, so a row only has to
// be tagged correctly once on that one shared sheet. Needed because
// buildGianCandidatesFromInvoices alone only knows the RAW "Ma diem tren
// misa thue" text straight off each invoice, which for a site not yet typed
// in with its true Misa code (e.g. "Sân bay Phú Quốc" instead of the real
// "CHKQT PHU QUOC") would otherwise show up AS ITSELF in the reconciliation
// table and get exported to MISA under the wrong Ma cong trinh.
function getMasterRowsForVietQrChannel(store, tagPattern) {
  const rows = (store.zvp_gian_master && store.zvp_gian_master.rows) || [];
  return rows.filter((r) => tagPattern.test(r.thuoc || ""));
}

// Luyen, 2026-07-17: "nếu nó map bị sai ... cấn trừ cho nó khớp doanh thu
// giữa các hóa đơn ngày hôm đó ... lệch 50k thì xem nó đang lệch với hóa đơn
// 50k của gian nào" -- tren cung 1 ngay, neu 1 gian du hoa don (diff = invoiceTotal
// - gross > 0) va 1 gian khac thieu hoa don (diff < 0) VOI CUNG DUNG 1 SO TIEN,
// rat co the 1 hoa don bi gan nham/gop nham giua 2 gian do (vi du 1 hoa don
// gop chung cho 2 diem nhung he thong dang tinh het cho 1 ben). Goi y cap
// nay ra thay vi tu dong ap dung -- Luyen xac nhan tung cap qua nut "Dong y,
// doi tru" (xem route /doi-soat/vietqr/cross-match/:channel) truoc khi luu.
function findCrossMatchSuggestions(reconciled, channelKey) {
  const suggestions = [];
  for (const r of reconciled) {
    const candidates = r.lines.filter(
      (l) => l.tkCo !== "SKIP" && !l.matched && !l.manualOverride && Math.abs(l.diff) >= 1
    );
    const used = new Set();
    for (let i = 0; i < candidates.length; i++) {
      if (used.has(candidates[i].code)) continue;
      for (let j = 0; j < candidates.length; j++) {
        if (i === j || used.has(candidates[j].code)) continue;
        const a = candidates[i];
        const b = candidates[j];
        // a du hoa don (dang tra bot), b thieu hoa don (dang nhan them)
        if (a.diff > 0 && b.diff < 0 && Math.abs(a.diff + b.diff) < 1) {
          suggestions.push({
            channel: channelKey,
            settlementDate: r.settlementDate,
            fromCode: a.code,
            fromLabel: displayCode(a.code),
            toCode: b.code,
            toLabel: displayCode(b.code),
            amount: a.diff,
          });
          used.add(a.code);
          used.add(b.code);
          break;
        }
      }
    }
  }
  return suggestions;
}

// Luyen, 2026-07-20: "các mã công trình 7702 này nè đổi lại tên đúng theo mã
// công trình kh mới cho tôi" -- ten "Ma cong trinh" hien trong bang doi soat
// Viet QR (pivot + chi tiet tung ngay) lay thang tu cot mo ta cua hoa don/QR
// (vd "Ma diem ghi chu HT Misa"), co the khac chut it (viet tat, thieu dau...)
// so voi ten CHUAN trong danh sach "Danh sach cong trinh" rieng cua tung cong
// ty (da tai o trang Giao dich, xem utils/maCongTrinh.js). Voi nhung kenh
// thuoc 1 cong ty da co danh sach chuan, doi TEN HIEN THI (khong dong cham
// code/key dung de doi soat/gop nhom, tranh lam sai lech du lieu da khop) ve
// dung ten trong danh sach chuan do bang findBestMaCongTrinh (khop tuyet doi
// truoc, khong co thi khop gan giong nhat). Neu chua co danh sach chuan cho
// cong ty do, hoac khong tim duoc dong nao gan giong, giu nguyen ten cu
// (displayCode) nhu truoc gio.
function displayMaCongTrinhFor(code, company, store) {
  const raw = displayCode(code);
  const manualAlias = store.ma_cong_trinh_display_alias && store.ma_cong_trinh_display_alias[raw];
  if (manualAlias) return manualAlias;
  const master = store.ma_cong_trinh_master && store.ma_cong_trinh_master[company];
  if (!master || !master.rows || master.rows.length === 0) return raw;
  const best = findBestMaCongTrinh(raw, master.rows);
  return best ? best.maCongTrinh : raw;
}

function buildChannelReconciliation(store, channelKey) {
  ensureChannelShape(store);
  if (ensureNo1388(store)) save(store);
  const cfg = CHANNELS[channelKey];
  const bank = store.banks.find((b) => b.name === cfg.bankName);
  if (!bank) {
    return { error: `Chua co ngan hang "${cfg.bankName}" trong he thong.` };
  }
  const txs = store.transactions.filter((t) => t.bank_id === bank.id);
  const settlements = extractVietQrSettlements(txs);

  const rawRows = mergeRawRows(store.viet_qr_raw_uploads[channelKey]);
  const storeNames = store.viet_qr_store_names[channelKey];
  let invoices = store.viet_qr_invoices[channelKey];

  // Full gian merge/rename (Luyen-confirmed, applies to ALL invoices past +
  // future -- e.g. "JP SC VIVO" turned out to be the SAME site as "SC VIVO
  // KVCM" (invoiced from some point on as "FUNFEST SCVIVO"), doanh thu chia
  // se (CSE) tu dau. Unlike a single mid-stream CSE-status change (handled
  // generically in reconcileVietQr's per-day fallback below), this is a
  // permanent identity merge onto a DIFFERENT ma cong trinh entirely, so it
  // has to rewrite maDiem before candidates are even built.
  const gianMerge = store.viet_qr_gian_merge || {};
  const mergeKeys = Object.keys(gianMerge);
  if (mergeKeys.length > 0) {
    invoices = invoices.map((inv) => {
      const hadFF = inv.maDiem && inv.maDiem.endsWith(FF_SUFFIX);
      const baseCode = hadFF ? inv.maDiem.slice(0, -FF_SUFFIX.length) : inv.maDiem;
      const merge = gianMerge[baseCode];
      if (!merge) return inv;
      // Chi Nhan (2026-07-27): "ngan hang doi soat tat ca deu la 131 va khong
      // chia theo chia se hay khong chia se nua" -- khong con gan hau to
      // FF_SUFFIX nua du merge.isCse la gi; van giu nguyen phan doi ten/gop
      // nhan dang (baseCode -> merge.maCongTrinh).
      const newMaDiem = merge.maCongTrinh;
      if (newMaDiem === inv.maDiem) return inv;
      return { ...inv, maDiem: newMaDiem };
    });
  }

  let gianCandidates = buildGianCandidatesFromInvoices(invoices);
  // Master gian sheet wins on a name collision (it carries the real Misa
  // code, e.g. "CHKQT PHU QUOC"), same precedence rule as ZVP's Online
  // channel -- see getMasterRowsForVietQrChannel above.
  gianCandidates = mergeGianListWithMaster(gianCandidates, getMasterRowsForVietQrChannel(store, cfg.tagPattern));
  if (mergeKeys.length > 0) {
    // The physical POS/store name on the bank/QR side never changes just
    // because invoicing renamed the site -- keep the OLD name fuzzy-
    // matchable for resolveGianGross, now pointing at the merged code.
    gianCandidates = mergeGianListWithMaster(
      gianCandidates,
      mergeKeys.map((oldCode) => ({
        raw: oldCode,
        maCongTrinh: gianMerge[oldCode].maCongTrinh,
        isCse: gianMerge[oldCode].isCse,
      }))
    );
  }
  // Luyen, 2026-07-19: "Sân bay Cần Thơ / Farm Lotte Nha Trang chưa map được
  // thì cứ để tên đó luôn, chờ map được hóa đơn thì thôi, tôi check điền sau"
  // -- 1 ma cua hang co the da co TEN THAT (gan qua form "Gan Ma cong trinh"
  // o /doi-soat/vietqr/store-map, hoac tu file map goc cua Luyen) nhung CHUA
  // co hoa don nao dung ten do, nen truoc gio roi vao "unmapped"/bi gap vao
  // defaultBlankCode. Luyen muon no hien THANH GIAN RIENG cua no (doi hoa
  // don ve sau), khong gop chung vao Tan Phu hay an trong bang unmapped. Voi
  // moi matchText da co ten nhung chua khop duoc candidate nao (dung ca
  // matcher that de kiem tra, tranh trung candidate da co san), them 1 "self"
  // candidate (ten = chinh no) de no tu khop Pass1 voi chinh minh -- ap dung
  // truoc khi resolveGianGross chay, chi cho kenh dung fuzzy text (parseMode
  // != "mn", vi BIDV7702/VietQR MN dung tien to rieng, khong lien quan).
  if (cfg.parseMode !== "mn") {
    const matcherProbe = buildOnlineProductMatcher(gianCandidates);
    const seenSelfNames = new Set(gianCandidates.map((c) => normText(c.maCongTrinh)));
    // Chi Nhan, 2026-07-24 (fix): tenDiem cua 1 hoa don that (vd "AE Tân Phú
    // ghế") co the toan tu ngan (<=3 ky tu sau khi bo dau: "tan","phu","ghe")
    // nen matcherProbe (buildOnlineProductMatcher, yeu cau keyword qualifying
    // >=4 ky tu de tranh nham) KHONG BAO GIO tu khop duoc VOI CHINH NO, du
    // trung tuyet doi tung chu. Neu 1 "Ma cua hang" khac (thuong la cac ma
    // chua co ten that, matchText mac dinh/rac) TINH CO co matchText giong
    // HET tenDiem nay, no van vuot qua ca 2 check (seenSelfNames theo
    // maCongTrinh + matcherProbe fuzzy) va bi day vao lam "self" candidate
    // TRUNG applyGianRedirectToInvoices's map (key = normText(tenDiem)) --
    // vi selfCandidates duoc concat VAO SAU, no GHI DE len candidate that,
    // khien MOI hoa don cua gian that (vd AM TP PHCM) bi chuyen huong nham
    // sang 1 Ma cong trinh "ma" (vd "Aeon Tân Phú BIDV") khong co doanh thu
    // nao, hien "Chưa có HĐ" du hoa don + doanh thu deu dung. Fix: kiem tra
    // THEM trung tuyet doi (normText) theo TEN DIEM cua cac candidate CO SAN
    // (khong chi maCongTrinh) truoc khi cho phep them self-candidate.
    const existingTenDiemExact = new Set(gianCandidates.map((c) => normText(c.tenDiem || "")));
    const selfCandidates = [];
    Object.values(storeNames).forEach((info) => {
      const mt = ((info && info.matchText) || "").trim();
      if (!mt) return;
      const key = normText(mt);
      if (seenSelfNames.has(key)) return;
      if (existingTenDiemExact.has(key)) return; // trung tuyet doi voi 1 gian THAT da co, khong tao ban sao
      if (matcherProbe(mt)) return; // da khop duoc voi candidate co san, khong can fallback
      seenSelfNames.add(key);
      // Chi Nhan, 2026-07-24: Luyen yeu cau "gộp hết" -- nhieu Ma cua hang la
      // cac may/quay QR KHAC NHAU tai CUNG 1 diem ban (vd "LM VT 01".."LM VT
      // 06" deu la "LOTTE VŨNG TÀU", "lotte gò vấp 01".."08" deu la "LOTTE GÒ
      // VẤP"...) truoc day moi ma tao 1 "self" candidate voi maCongTrinh =
      // TOAN BO matchText (rieng cho tung may), nen hien thanh nhieu dong
      // rieng le du la CUNG 1 diem that. Dung "Tên điểm bán" (phan ten CHUNG,
      // khong doi giua cac may cung diem) lam ma cong trinh OUTPUT thay vi ca
      // matchText, de tu dong gop lai; "tenDiem" (dung de fuzzy-khop VOI
      // CHINH dong nay khi resolveGianGross chay) van giu nguyen matchText
      // DAY DU rieng cho tung may, khong anh huong do chinh xac khop tung
      // giao dich. Neu khong co Tên điểm bán rieng (chi co Tên cửa hàng), giu
      // nguyen hanh vi cu (moi ma 1 dong rieng).
      const groupCode = (info && info.tenDiemBan && info.tenDiemBan.trim()) || mt;
      selfCandidates.push({ tenDiem: mt, maCongTrinh: groupCode, isCse: false });
    });
    if (selfCandidates.length > 0) gianCandidates = gianCandidates.concat(selfCandidates);
  }

  // BIDV7702/VietQR MN: khop gian theo tien to ten cua hang (vd "AMTP 01"
  // -> "AMTP" ung voi ma cong trinh "AM TP KVCM"), khong dung fuzzy text
  // matcher nhu 3 kenh kia -- xem ghi chu tai CHANNELS.bidv7702 o tren.
  const nocodeAssignments = store.viet_qr_nocode_assignments[channelKey] || {};
  const resolved =
    cfg.parseMode === "mn"
      ? resolveGianGrossPrefix(rawRows, storeNames, gianCandidates)
      : resolveGianGross(
          rawRows,
          storeNames,
          gianCandidates,
          nocodeAssignments,
          cfg.defaultBlankCode,
          store.viet_qr_ten_diem_master[channelKey]
        );

  // Redirect each invoice's own "Ma diem tren misa thue" through the SAME
  // gianCandidates map (keyed by invoice's own "Ten diem xuat hoa don"),
  // same mechanism as ZVP's applyGianRedirectToInvoices -- otherwise an
  // invoice whose Ma diem column is just the site's own raw name (e.g.
  // "SÂN BAY PHÚ QUỐC" instead of "CHKQT PHU QUOC") never lines up against
  // the settlement line above (which IS already correctly redirected via
  // the master sheet), permanently showing as "chua khop" even after the
  // master gian sheet fixes the code everywhere else.
  invoices = applyGianRedirectToInvoices(invoices, gianCandidates);

  // Chi Nhan, 2026-07-24: BIDV7702 tu 22/07 tro di khop gian bang "So tham
  // chieu" ngan hang (khop tuyet doi, ngan hang la chuan) thay vi fuzzy text
  // -- xem ghi chu tai CHANNELS.bidv7702.refMatchFrom o tren va
  // resolveGianGrossByBankRef trong utils/vietqrReconcile.js. Cac ngay TRUOC
  // ngay cutover van giu y nguyen ket qua "resolved" (fuzzy) da tinh o tren;
  // chi THAY THE gross cho cac ngay >= cutover bang so ngan hang.
  let refUnmatchedBankTx = [];
  let refLateMatches = [];
  let refUnmappedStoreCodes = [];
  let refUnmappedTenDiem = [];
  if (cfg.refMatchFrom) {
    const bankThuTxs = extractVietQrThuTransactions(txs).filter((t) => t.date >= cfg.refMatchFrom);
    const tenDiemMaster = store.viet_qr_ten_diem_master[channelKey] || {};
    const storeCodeOverride = store.viet_qr_store_code_override[channelKey] || {};
    const refOverride = store.viet_qr_ref_override[channelKey] || {};
    const refResolved = resolveGianGrossByBankRef(bankThuTxs, rawRows, storeNames, tenDiemMaster, storeCodeOverride, refOverride);

    const filteredGrossByCode = {};
    const filteredCodes = new Set();
    Object.keys(resolved.grossByCode).forEach((key) => {
      const day = key.slice(0, key.indexOf("|"));
      if (day >= cfg.refMatchFrom) return; // se duoc thay bang so ngan hang phia duoi
      filteredGrossByCode[key] = resolved.grossByCode[key];
      filteredCodes.add(key.slice(key.indexOf("|") + 1));
    });
    Object.keys(refResolved.grossByCode).forEach((key) => {
      filteredGrossByCode[key] = refResolved.grossByCode[key];
      filteredCodes.add(key.slice(key.indexOf("|") + 1));
    });
    resolved.codes = Array.from(filteredCodes);
    resolved.grossByCode = filteredGrossByCode;

    refUnmatchedBankTx = refResolved.unmatchedBankTx;
    refLateMatches = refResolved.lateMatches;
    refUnmappedStoreCodes = refResolved.unmappedStoreCodes;
    refUnmappedTenDiem = refResolved.unmappedTenDiem;
  }

  // Chi Nhan, 2026-07-24: gop gross cua ma DA MERGE (GIAN_MERGE_DEFAULTS/
  // store.viet_qr_gian_merge, vd "FARM LOTTE PHAN THIET" -> "LOTTE PHAN
  // THIET") vao dung ma dich. Khoi mergeGianListWithMaster/applyGianRedirectToInvoices
  // o tren CHI doi duoc hoa don + gianCandidates (duong fuzzy resolveGianGross),
  // KHONG cham duoc gross tu resolveGianGrossByBankRef (khop theo So tham
  // chieu, hoac gan qua "Gan ma cong trinh" cho ma cua hang MOI -- ca 2 deu
  // co the tra ve thang mot ma da bi merge) -- neu khong gop lai o day, 1
  // gian bi merge van hien 2 dong rieng ben doanh thu (ngan hang) trong khi
  // hoa don goc chi co 1, gay "Chua co HD" gia tao mai mai.
  if (mergeKeys.length > 0) {
    const mergedGrossByCode = {};
    const mergedCodes = new Set();
    Object.keys(resolved.grossByCode).forEach((key) => {
      const sep = key.indexOf("|");
      const day = key.slice(0, sep);
      let code = key.slice(sep + 1);
      const hadFF = code.endsWith(FF_SUFFIX);
      const baseCode = hadFF ? code.slice(0, -FF_SUFFIX.length) : code;
      const merge = gianMerge[baseCode];
      if (merge) {
        // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong
        // rieng, nen khong gan hau to FF_SUFFIX nua o day.
        code = merge.maCongTrinh;
      }
      const newKey = `${day}|${code}`;
      mergedGrossByCode[newKey] = (mergedGrossByCode[newKey] || 0) + resolved.grossByCode[key];
      mergedCodes.add(code);
    });
    resolved.grossByCode = mergedGrossByCode;
    resolved.codes = Array.from(mergedCodes);
  }

  seedGianMappingDefaults(store, resolved.codes);

  const manualMatches = store.viet_qr_manual_matches[channelKey] || {};
  const reconciled = reconcileVietQr(
    settlements,
    resolved,
    { invoices },
    store.gian_mapping,
    manualMatches,
    store.invoice_diem_alias
  );

  // Luyen, 2026-07-27: "các giao dịch không phải của vietqr thì trừ ra nhá
  // cái nào có mã tham chiếu á" -- tu ngay cutover (refMatchFrom), 1 giao
  // dich "thu" tren sao ke KHONG khop duoc So tham chieu nao voi bat ky
  // gian nao (vd "Thanh toan lai thang 07/2026" tu ngan hang, chuyen khoan
  // IBFT ca nhan...) khong phai tien VietQR that su -- da duoc gom san
  // trong refUnmatchedBankTx (canh bao) nhung truoc gio CHUA tru khoi tong
  // "Ngan hang" cua doi soat, lam Lech gia (ngan hang > du lieu tai len chi
  // vi cong nham cac dong khong lien quan). Tru dung so tien nay khoi
  // bankAmount cua NGAY tuong ung roi tinh lai diffVsBank. Neu SAU KHI tru
  // ngan hang van con nhieu hon du lieu doi soat (tien VietQR that nhung
  // chua gian nao nhan), goi y gop vao "gian Tan Phu" (AM TP PHCM -- ma mac
  // dinh cho giao dich khong xac dinh duoc gian tren kenh nay, xem
  // cfg.defaultBlankCode) kem dung so tien du.
  if (cfg.refMatchFrom) {
    const excludedByDate = {};
    refUnmatchedBankTx.forEach((tx) => {
      excludedByDate[tx.date] = (excludedByDate[tx.date] || 0) + tx.amount;
    });
    reconciled.forEach((r) => {
      if (r.settlementDate < cfg.refMatchFrom) return;
      const excluded = excludedByDate[r.settlementDate] || 0;
      if (excluded > 0) {
        r.bankAmount -= excluded;
        r.bankAmountExcluded = excluded;
      }
      if (!r.pendingBank) {
        r.diffVsBank = r.totalNetComputed - r.bankAmount;
        const leftover = r.bankAmount - r.totalNetComputed;
        if (leftover > 1000) {
          r.tanPhuSuggestion = { amount: leftover, targetCode: "AM TP PHCM" };
        }
      }
    });
  }

  // Doi ten hien thi ve dung ten chuan (neu co danh sach chuan cho cong ty
  // nay) -- xem ghi chu tai displayMaCongTrinhFor o tren. Chi doi field hien
  // thi maCongTrinh, khong dong den code/effectiveCode dung de doi soat.
  reconciled.forEach((r) => {
    r.lines.forEach((l) => {
      l.maCongTrinh = displayMaCongTrinhFor(l.code, cfg.company, store);
    });
  });

  // "Khoa so" -- xem ghi chu tai ensureChannelShape. Danh dau nhung ngay <=
  // ngay khoa la "locked": giu nguyen gross/invoiceTotal de tra cuu, nhung
  // dua diff ve 0 (khong con cong don vao "Lech Du lieu-Hoa don" cua ngay do
  // nua) va bao view biet de hien badge "Da khoa so" thay vi "Chua co HD"/
  // "Lech", khong con nut sua/xoa nao cho cac dong nay.
  const lockDate = (store.viet_qr_lock_date && store.viet_qr_lock_date[channelKey]) || "";
  if (lockDate) {
    reconciled.forEach((r) => {
      if (r.settlementDate <= lockDate) {
        r.locked = true;
        r.lines.forEach((l) => {
          l.locked = true;
          l.diff = 0;
        });
      }
    });
  }

  // Chi Nhan, 2026-07-22: "cộng gộp các điểm mã công trình lại với nhau đi"
  // -- nhieu ma cua hang/code tho khac nhau deu duoc displayMaCongTrinhFor o
  // tren quy ve CUNG 1 ten Ma Cong Trinh chuan (vd "SB CAN THO PHCM", "GO MY
  // THO"...), nhung bang "Ket qua doi soat" chi tiet van hien MOI CODE THO
  // thanh 1 dong rieng nen bi trung ten nhieu lan. CHI gop cac dong DA KHOA
  // SO lai voi nhau (khong con nut sua/xoa nao, an toan tuyet doi khi gop) --
  // dong nao CHUA khoa (con dang can xu ly/co nut sua) thi GIU NGUYEN rieng
  // le nhu cu, tranh lam hong cac nut "Sua HD"/doi tru thu cong dang gan voi
  // dung 1 code cu the.
  reconciled.forEach((r) => {
    const groups = {};
    const order = [];
    r.lines.forEach((l) => {
      // Chi Nhan (2026-07-27): khong con chia theo chia se (CSE)/khong chia
      // se nua -- gop thang theo Ma Cong Trinh.
      const key = l.maCongTrinh;
      if (!groups[key]) {
        groups[key] = [];
        order.push(key);
      }
      groups[key].push(l);
    });
    const newLines = [];
    order.forEach((key) => {
      const group = groups[key];
      if (group.length === 1 || !group.every((l) => l.locked)) {
        newLines.push(...group);
        return;
      }
      const first = group[0];
      const gross = group.reduce((s, l) => s + l.gross, 0);
      const invoiceTotal = group.reduce((s, l) => s + l.invoiceTotal, 0);
      const invoiceNumbers = Array.from(new Set(group.flatMap((l) => l.invoiceNumbers)));
      newLines.push({
        code: first.code,
        maCongTrinh: first.maCongTrinh,
        tkCo: first.tkCo,
        gross,
        net: gross,
        invoiceNumbers,
        invoiceTotal,
        diff: 0,
        matched: invoiceNumbers.length > 0 && Math.abs(invoiceTotal - gross) < 1,
        manualOverride: group.some((l) => l.manualOverride),
        locked: true,
      });
    });
    r.lines = newLines;
  });

  const allCodes = new Set();
  reconciled.forEach((r) => r.lines.forEach((l) => allCodes.add(l.code)));

  const knownCodesForAlias = new Set([...allCodes, ...Object.keys(store.gian_mapping || {})]);
  const invoiceDiemAlias = store.invoice_diem_alias || {};
  const unmatchedInvoiceCodesSet = new Set();
  invoices.forEach((inv) => {
    if (!inv.maDiem) return;
    if (knownCodesForAlias.has(inv.maDiem)) return;
    if (invoiceDiemAlias[inv.maDiem]) return;
    unmatchedInvoiceCodesSet.add(inv.maDiem);
  });

  return {
    reconciled,
    lockDate,
    allCodes: Array.from(allCodes).sort(),
    unmappedStores: resolved.unmapped,
    // Luyen, 2026-07-17: "map giữ tên điểm nội bộ có trong file hệ thống với
    // Mã công trình -- mã nào chưa map được hiện ra cho tôi" -- structured
    // per-Ma-cua-hang version of unmappedStores so the page can render an
    // actionable mapping form (not just a read-only warning string).
    unmappedStoreDetails: resolved.unmappedDetails || [],
    unmappedBlankRows: resolved.blankRows || [],
    crossMatchSuggestions: findCrossMatchSuggestions(reconciled, channelKey),
    invoiceDiemAlias,
    unmatchedInvoiceCodes: Array.from(unmatchedInvoiceCodesSet).sort(),
    // Canh bao rieng cho co che khop theo So tham chieu ngan hang (xem
    // CHANNELS.bidv7702.refMatchFrom o tren) -- rong ([]) voi cac kenh khong
    // bat tinh nang nay.
    refUnmatchedBankTx,
    refLateMatches,
    refUnmappedStoreCodes,
    refUnmappedTenDiem,
  };
}

router.get("/doi-soat/vietqr", (req, res) => {
  const store = load();
  ensureChannelShape(store);

  // Chi hien cac ngan hang thuoc cong ty dang chon (nut chuyen cong ty tren
  // topbar) -- KH Cu van thay ca 3 ngan hang nhu truoc gio, KH Moi thay
  // BIDV 7702. Cac route upload/xoa/xuat file van nhan moi channelKey hop
  // le (khong gioi han theo cong ty) vi Luyen la nguoi dung duy nhat.
  const activeCompany = getCompany(req);
  const activeKeys = CHANNEL_KEYS.filter((ch) => CHANNELS[ch].company === activeCompany);

  const built = {};
  activeKeys.forEach((ch) => {
    built[ch] = buildChannelReconciliation(store, ch);
  });

  const monthSet = new Set();
  activeKeys.forEach((ch) => {
    (built[ch].reconciled || []).forEach((r) => monthSet.add(r.settlementDate.slice(0, 7)));
  });
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";
  // Chon ngan hang de xem: 3 ngan hang doi soat rieng, cuon qua ca 3 de tim
  // 1 dong rat mat cong -- mac dinh van hien ca 3 ("Tat ca"), nhung Luyen co
  // the loc con 1 ngan hang cho de doi chieu.
  const selectedChannel = req.query.channel !== undefined ? req.query.channel : "";

  const reconciledByChannel = {};
  activeKeys.forEach((ch) => {
    const all = built[ch].reconciled || [];
    reconciledByChannel[ch] = selectedMonth ? all.filter((r) => r.settlementDate.slice(0, 7) === selectedMonth) : all;
  });

  const allCodes = new Set();
  activeKeys.forEach((ch) => (built[ch].allCodes || []).forEach((c) => allCodes.add(c)));

  // "Bang tong quan theo gian": 1 dong = 1 ma cong trinh, 1 cot = 1 ngay
  // (trong pham vi thang/ngan hang dang loc), de Luyen nhin duoc lech cua
  // CA THANG cung 1 luc thay vi cuon qua tung ngay rieng le. Bam vao 1 o de
  // xem chi tiet (so HD, doanh thu, lech...) -- van dung dung 1 nguon du
  // lieu (reconciledByChannel) voi bang "Ket qua doi soat" chi tiet ben
  // duoi, nen 2 bang luon khop nhau tuyet doi.
  const pivotByChannel = {};
  activeKeys.forEach((ch) => {
    const rows = reconciledByChannel[ch] || [];
    const dates = Array.from(new Set(rows.map((r) => r.settlementDate))).sort();
    // Chi Nhan, 2026-07-22: "mấy cái này gộp theo mã công trình như hôm qua
    // á" -- bang nay truoc day gop dong theo "code" THO (chuoi goc tu
    // resolveGianGross, ke ca cac "self-fallback" rieng cho tung ten cua
    // hang chua co hoa don, vd "LM VT 02", "LM VT 04"... hay nhieu ma cua
    // hang khac nhau cung fuzzy-match ve 1 diem that qua displayMaCongTrinhFor
    // nhu "SB CAN THO PHCM"). Nhieu "code" khac nhau co the CUNG quy ve 1
    // TEN MA CONG TRINH sau khi hien thi (displayMaCongTrinhFor so khop voi
    // danh sach cong trinh chuan), nen truoc day hien thanh nhieu dong TRUNG
    // TEN nhau thay vi gop lam 1 -- gio gop theo TEN DA QUY VE (maCongTrinh)
    // ngay tu dau, cong don doanh thu/hoa don cua tat ca cac code con lai
    // vao chung 1 dong duy nhat cho tung ngay. Chi Nhan (2026-07-27): khong
    // con tach rieng theo chia se (CSE)/khong chia se nua.
    const cellMap = {}; // groupKey -> { date -> merged cell data }
    const groupInfo = {}; // groupKey -> { maCongTrinh, isCse: always false now }
    // Chi Nhan, 2026-07-22 (fix khan): displayMaCongTrinhFor goi
    // findBestMaCongTrinh, la 1 vong lap fuzzy-match qua TOAN BO danh sach
    // cong trinh chuan (co the vai tram dong) -- truoc day chi goi 1 LAN cho
    // MOI CODE DUY NHAT (sau khi da gop qua Set), gio neu goi lai cho TUNG
    // DONG settlement (co the hang chuc nghin dong) se cham hang chuc/tram
    // lan, gay treo/timeout server thuc te (502 tren Railway). Cache lai theo
    // code de van chi tinh 1 lan cho moi code duy nhat nhu cu.
    const maCongTrinhCache = {};
    function resolveMaCongTrinh(code) {
      if (!(code in maCongTrinhCache)) {
        maCongTrinhCache[code] = displayMaCongTrinhFor(code, CHANNELS[ch].company, store);
      }
      return maCongTrinhCache[code];
    }
    rows.forEach((r) => {
      r.lines.forEach((l) => {
        const maCongTrinh = resolveMaCongTrinh(l.code);
        // Chi Nhan (2026-07-27): khong con chia theo chia se (CSE)/khong
        // chia se nua -- gop thang theo ten Ma Cong Trinh da quy ve.
        const groupKey = maCongTrinh;
        groupInfo[groupKey] = { maCongTrinh, isCse: false };
        if (!cellMap[groupKey]) cellMap[groupKey] = {};
        const existing = cellMap[groupKey][r.settlementDate];
        if (!existing) {
          cellMap[groupKey][r.settlementDate] = {
            tkCo: l.tkCo,
            gross: l.gross,
            invoiceTotal: l.invoiceTotal,
            invoiceNumbers: [...l.invoiceNumbers],
            matched: l.matched,
            manualOverride: l.manualOverride,
            locked: !!l.locked,
          };
        } else {
          existing.gross += l.gross;
          existing.invoiceTotal += l.invoiceTotal;
          existing.invoiceNumbers = existing.invoiceNumbers.concat(l.invoiceNumbers);
          existing.matched = existing.matched && l.matched;
          existing.manualOverride = existing.manualOverride || l.manualOverride;
          // Chi coi ca nhom la "da khoa" cho ngay do neu TAT CA cac code con
          // gop vao deu da khoa -- con 1 code chua khoa thi van can hien de
          // xu ly, khong an di.
          existing.locked = existing.locked && !!l.locked;
          if ((!existing.tkCo || existing.tkCo === "SKIP") && l.tkCo && l.tkCo !== "SKIP") existing.tkCo = l.tkCo;
        }
      });
    });
    const groupKeys = Object.keys(groupInfo).sort((a, b) =>
      groupInfo[a].maCongTrinh.localeCompare(groupInfo[b].maCongTrinh)
    );
    pivotByChannel[ch] = {
      dates,
      rows: groupKeys.map((groupKey) => {
        const { maCongTrinh, isCse } = groupInfo[groupKey];
        let sumGross = 0;
        let sumInvoiceTotal = 0;
        let sumDiffUnlocked = 0;
        const cells = dates.map((d) => {
          const l = (cellMap[groupKey] || {})[d];
          if (!l) return null;
          sumGross += l.gross;
          sumInvoiceTotal += l.invoiceTotal;
          const diff = l.invoiceTotal - l.gross;
          // Khoa so: dong da khoa van tinh vao Tong DT (tien that), nhung
          // KHONG tinh vao "Tong lech" nua (khong con can xu ly) -- xem ghi
          // chu tai buildChannelReconciliation.
          if (!l.locked) sumDiffUnlocked += diff;
          let status;
          if (l.locked) status = "locked";
          else if (l.tkCo === "SKIP") status = "skip";
          else if (l.manualOverride) status = "ok";
          else if (l.invoiceNumbers.length === 0) status = "missing";
          else if (l.matched) status = "ok";
          else status = "diff";
          return {
            tkCo: l.tkCo,
            gross: l.gross,
            invoiceTotal: l.invoiceTotal,
            diff,
            invoiceNumbers: l.invoiceNumbers,
            matched: l.matched,
            manualOverride: l.manualOverride,
            locked: l.locked,
            status,
          };
        });
        return {
          code: maCongTrinh,
          maCongTrinh,
          isCse,
          cells,
          sumGross,
          sumInvoiceTotal,
          sumDiff: sumDiffUnlocked,
        };
      }),
    };
  });

  // Gop danh sach "Ma cua hang chua map duoc gian" cua tat ca kenh dang xem,
  // kem theo channel key de form biet luu vao dung viet_qr_store_names[channel]
  // nao (Luyen, 2026-07-17: muon thay ro ma nao chua map de biet vi sao
  // "Ngan hang" vs "Tinh tu du lieu tai len" bi lech).
  const allUnmappedStores = [];
  activeKeys.forEach((ch) => {
    (built[ch].unmappedStoreDetails || []).forEach((d) => {
      // Nhom "khong co ma cua hang" (raw "-") van hien de Luyen thay het tien
      // dang bi loai, nhung KHONG cho gan Ma cong trinh qua form nay -- 1 ma
      // cua hang that thi chi thuoc DUNG 1 gian, con nhom "-" gop chung giao
      // dich tu nhieu gian khac nhau (khong co code de phan biet), gan bua 1
      // ma se sai cho nhung giao dich thuc ra thuoc gian khac.
      allUnmappedStores.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...d });
    });
  });
  allUnmappedStores.sort((a, b) => b.total - a.total);

  // Tung giao dich "khong co ma cua hang" rieng le, kem nut gan Ma cong trinh
  // cho TUNG dong (Luyen, 2026-07-17: "con loi chua map dc cai nao thi bao
  // cho toi va co nut sua nha") -- khac voi allUnmappedStores o tren (chi
  // hien 1 dong tong hop, khong cho gan vi khong biet giao dich nao thuoc
  // gian nao); o day moi giao dich co rieng 1 vqrCode nen gan duoc tung cai.
  const allUnmappedBlankRows = [];
  activeKeys.forEach((ch) => {
    (built[ch].unmappedBlankRows || []).forEach((r) => {
      allUnmappedBlankRows.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...r });
    });
  });
  allUnmappedBlankRows.sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  const allCrossMatchSuggestions = [];
  activeKeys.forEach((ch) => {
    (built[ch].crossMatchSuggestions || []).forEach((s) => {
      allCrossMatchSuggestions.push({ ...s, channelLabel: CHANNELS[ch].label });
    });
  });

  // Canh bao rieng cho co che khop theo "So tham chieu" ngan hang (BIDV7702
  // tu 22/07, xem CHANNELS.bidv7702.refMatchFrom) -- gop qua tat ca kenh dang
  // xem giong cac mang allXxx khac, rong voi kenh khong bat tinh nang nay.
  const allRefUnmatchedBankTx = [];
  const allRefLateMatches = [];
  const allRefUnmappedStoreCodes = [];
  const allRefUnmappedTenDiem = [];
  activeKeys.forEach((ch) => {
    (built[ch].refUnmatchedBankTx || []).forEach((r) => allRefUnmatchedBankTx.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...r }));
    (built[ch].refLateMatches || []).forEach((r) => allRefLateMatches.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...r }));
    (built[ch].refUnmappedStoreCodes || []).forEach((r) => allRefUnmappedStoreCodes.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...r }));
    (built[ch].refUnmappedTenDiem || []).forEach((r) => allRefUnmappedTenDiem.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...r }));
  });
  allRefUnmatchedBankTx.sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  allRefLateMatches.sort((a, b) => (a.bankDate || "").localeCompare(b.bankDate || ""));
  allRefUnmappedStoreCodes.sort((a, b) => b.total - a.total);
  allRefUnmappedTenDiem.sort((a, b) => b.total - a.total);

  // Luyen, 2026-07-19: "co nut xoa hay chinh sua cac phan dien" -- liet ke lai
  // cac GD-khong-ma DA gan (truoc gio gan xong la bien mat, khong xem/xoa lai
  // duoc) kem nut Xoa de tra ve dien "chua gan" neu gan nham.
  const allNocodeAssignments = [];
  activeKeys.forEach((ch) => {
    Object.entries(store.viet_qr_nocode_assignments[ch] || {}).forEach(([vqrCode, a]) => {
      allNocodeAssignments.push({ channel: ch, channelLabel: CHANNELS[ch].label, vqrCode, ...a });
    });
  });
  allNocodeAssignments.sort((a, b) => (b.assignedAt || "").localeCompare(a.assignedAt || ""));

  // Luyen, 2026-07-20: "nhìn rối mắt, mỗi gian 1 dòng thôi, bấm vô hả ra chi
  // tiết" -- danh sach mã cửa hàng thô (mỗi mã QR terminal 1 dòng, có thể
  // hàng trăm dòng cho 1 gian) nhìn rất rối. Gom theo gian (matchText/tenDiemBan
  // -- cùng field resolveGianGross dùng để khớp) thành 1 dòng tổng hợp/gian,
  // các mã cửa hàng lẻ bên trong xem qua <details> mở rộng.
  const allStoreGroups = [];
  activeKeys.forEach((ch) => {
    const names = store.viet_qr_store_names[ch] || {};
    const groups = {};
    Object.keys(names).forEach((maCuaHang) => {
      const info = names[maCuaHang] || {};
      const label = (info.matchText || info.tenDiemBan || info.tenCuaHang || "(chưa có tên)").trim();
      const key = label.toLowerCase();
      if (!groups[key]) groups[key] = { label, maDiemBan: info.maDiemBan || "", codes: [] };
      groups[key].codes.push({ maCuaHang, tenCuaHang: info.tenCuaHang || "", maDiemBan: info.maDiemBan || "" });
    });
    Object.values(groups).forEach((g) => {
      g.codes.sort((a, b) => a.tenCuaHang.localeCompare(b.tenCuaHang));
      allStoreGroups.push({ channel: ch, channelLabel: CHANNELS[ch].label, label: g.label, maDiemBan: g.maDiemBan, codes: g.codes, count: g.codes.length });
    });
  });
  allStoreGroups.sort((a, b) => a.label.localeCompare(b.label));

  res.render("doisoat-vietqr", {
    userName: req.session.userName,
    channels: activeKeys.map((ch) => ({ key: ch, label: CHANNELS[ch].label, parseMode: CHANNELS[ch].parseMode || null })),
    selectedChannel,
    rawUploads: activeKeys.reduce((acc, ch) => {
      acc[ch] = store.viet_qr_raw_uploads[ch];
      return acc;
    }, {}),
    storeUploads: activeKeys.reduce((acc, ch) => {
      acc[ch] = store.viet_qr_store_uploads[ch];
      return acc;
    }, {}),
    invoiceCounts: activeKeys.reduce((acc, ch) => {
      acc[ch] = store.viet_qr_invoices[ch].length;
      return acc;
    }, {}),
    reconciled: reconciledByChannel,
    pivotByChannel,
    built,
    months,
    selectedMonth,
    gianMapping: store.gian_mapping,
    allCodes: Array.from(allCodes).sort(),
    allUnmappedStores,
    allUnmappedBlankRows,
    allCrossMatchSuggestions,
    allNocodeAssignments,
    allStoreGroups,
    allRefUnmatchedBankTx,
    allRefLateMatches,
    allRefUnmappedStoreCodes,
    allRefUnmappedTenDiem,
    maCongTrinhOptions: (store.ma_cong_trinh_master && store.ma_cong_trinh_master[activeCompany] && store.ma_cong_trinh_master[activeCompany].rows) || [],
    partnerBankId: store.viet_qr_partner_bank_id || {},
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Luyen, 2026-07-24: "mã cửa hàng mới này ... chỗ chọn mã công trình để gán
// vào nhá" -- gan THANG 1 Ma cua hang (con chua co trong danh sach diem ban)
// -> 1 Ma cong trinh, ap dung ngay cho ca du lieu cu va moi (xem
// resolveGianGrossByBankRef), khong can doi upload lai file "Danh sach diem
// ban" + file "Ten diem - Ma cong trinh" (2 buoc gian tiep truoc gio).
router.post("/doi-soat/vietqr/gan-ma-cua-hang/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { maCuaHang, maCongTrinh } = req.body;
    if (!maCuaHang || !maCongTrinh) throw new Error("Thieu ma cua hang hoac ma cong trinh de gan.");
    store.viet_qr_store_code_override[channelKey][maCuaHang] = maCongTrinh;
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da gan ma cua hang "${maCuaHang}" -> "${maCongTrinh}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-27: "100k của dư ngân hàng nếu chưa có bên dữ liệu thì đưa
// vô gian AMTP cho tôi" -- gan THANG 1 So tham chieu ngan hang (giao dich
// hoan toan khong co dong QR nao khop, se bi tru khoi "Ngan hang" boi buoc
// loc giao dich khong phai VietQR neu khong gan) -> 1 Ma cong trinh, ap
// dung ngay khong can cho file QR nao khop them. Xem resolveGianGrossByBankRef.
router.post("/doi-soat/vietqr/gan-ma-tham-chieu/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { reference, maCongTrinh } = req.body;
    if (!reference || !maCongTrinh) throw new Error("Thieu so tham chieu hoac ma cong trinh de gan.");
    store.viet_qr_ref_override[channelKey][reference.trim()] = maCongTrinh;
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da gan so tham chieu "${reference}" -> "${maCongTrinh}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// Chi Nhan, 2026-07-27: "thêm cho tôi 1 nút cập nhật viet qr ... bên phải
// của tài khoản nào thì lọc tài khoản đó" -- luu lai bankId (chi la ID tai
// khoan tren cong doi tac doitac.vietqr.vn, KHONG phai mat khau/token dang
// nhap) ung voi tung kenh, de nut "Cap nhat VietQR" o giao dien mo dung URL
// loc san tai khoan + ngay hom nay, khong con phai tu chon lai tai khoan
// tren trang doi tac moi lan.
router.post("/doi-soat/vietqr/partner-bank-id/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { bankId } = req.body;
    store.viet_qr_partner_bank_id[channelKey] = (bankId || "").trim();
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" + encodeURIComponent(`Da luu bankId cho ${CHANNELS[channelKey].label}.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Cross-match: 1 gian du hoa don + 1 gian thieu hoa don CUNG NGAY,
// cung 1 so tien -- Luyen xac nhan tung cap qua nut nay (khong tu dong ap
// dung) truoc khi ghi de thanh 2 ban ghi manual-match. ----------
router.post("/doi-soat/vietqr/cross-match/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { settlementDate, fromCode, toCode } = req.body;
    if (!settlementDate || !fromCode || !toCode) throw new Error("Thieu thong tin de doi tru.");
    const built = buildChannelReconciliation(store, channelKey);
    if (built.error) throw new Error(built.error);
    const settlement = (built.reconciled || []).find((r) => r.settlementDate === settlementDate);
    if (!settlement) throw new Error("Khong tim thay ngay can doi tru.");
    const lineFrom = settlement.lines.find((l) => l.code === fromCode);
    const lineTo = settlement.lines.find((l) => l.code === toCode);
    if (!lineFrom || !lineTo) throw new Error("Khong tim thay gian can doi tru (co the da thay doi, tai lai trang roi thu lai).");
    const note = `Doi tru tu dong xac nhan boi Luyen: hoa don cua "${displayCode(fromCode)}" du ${lineFrom.diff.toLocaleString(
      "vi-VN"
    )}đ, chuyen sang "${displayCode(toCode)}" dang thieu dung so do (${new Date().toLocaleDateString("vi-VN")}).`;
    store.viet_qr_manual_matches[channelKey][`${settlementDate}|${fromCode}`] = {
      invoiceNumbers: lineFrom.invoiceNumbers,
      amount: lineFrom.gross,
      grossAdjustment: 0,
      note,
      created_at: new Date().toISOString(),
    };
    store.viet_qr_manual_matches[channelKey][`${settlementDate}|${toCode}`] = {
      invoiceNumbers: Array.from(new Set([...lineTo.invoiceNumbers, ...lineFrom.invoiceNumbers])),
      amount: lineTo.gross,
      grossAdjustment: 0,
      note,
      created_at: new Date().toISOString(),
    };
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(
          `Da doi tru "${displayCode(fromCode)}" <-> "${displayCode(toCode)}" ngay ${settlementDate}.`
        )
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Cross-match hang loat: Luyen yeu cau 2026-07-18 "cho phep dong
// y dong loat" thay vi bam tung the "Dong y, doi tru" mot khi co nhieu goi
// y cung luc. Nhan 1 mang JSON cac {channel, settlementDate, fromCode,
// toCode} (dung item lay tu allCrossMatchSuggestions dang hien tren trang),
// ap dung LAI TUNG DUNG LOGIC nhu route don o tren (khong tu che, van doi
// chieu lai voi ket qua doi soat MOI NHAT truoc khi ghi -- neu 1 gian da
// thay doi/khong con dung nua thi bo qua item do, khong lam hong ca lo).
router.post("/doi-soat/vietqr/cross-match-all", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    const { items } = req.body;
    if (!items) throw new Error("Khong co goi y nao de doi tru.");
    let list;
    try {
      list = JSON.parse(items);
    } catch (e) {
      throw new Error("Du lieu goi y khong hop le.");
    }
    if (!Array.isArray(list) || list.length === 0) throw new Error("Khong co goi y nao de doi tru.");

    // Cache buildChannelReconciliation per channel (nhieu item cung 1 kenh
    // se dung chung, khong tinh lai nhieu lan cho ton).
    const builtByChannel = {};
    let applied = 0;
    let skipped = 0;
    const skippedDetails = [];
    for (const item of list) {
      const { channel: channelKey, settlementDate, fromCode, toCode } = item || {};
      if (!CHANNELS[channelKey] || !settlementDate || !fromCode || !toCode) {
        skipped++;
        continue;
      }
      if (!builtByChannel[channelKey]) {
        const built = buildChannelReconciliation(store, channelKey);
        if (built.error) {
          skipped++;
          continue;
        }
        builtByChannel[channelKey] = built;
      }
      const built = builtByChannel[channelKey];
      const settlement = (built.reconciled || []).find((r) => r.settlementDate === settlementDate);
      const lineFrom = settlement && settlement.lines.find((l) => l.code === fromCode);
      const lineTo = settlement && settlement.lines.find((l) => l.code === toCode);
      // Bo qua an toan neu gian da doi (vd Luyen vua sua tay 1 cai trong luc
      // cac goi y khac dang cho ap dung) thay vi bao loi lam dung ca lo giua
      // chung, hoac neu diem da tu khop/co manualOverride tu truoc do.
      if (!lineFrom || !lineTo || lineFrom.matched || lineFrom.manualOverride || Math.abs(lineFrom.diff + lineTo.diff) >= 1) {
        skipped++;
        skippedDetails.push(`${displayCode(fromCode)} <-> ${displayCode(toCode)} (${settlementDate})`);
        continue;
      }
      const note = `Doi tru hang loat tu dong xac nhan boi Luyen: hoa don cua "${displayCode(
        fromCode
      )}" du ${lineFrom.diff.toLocaleString("vi-VN")}đ, chuyen sang "${displayCode(
        toCode
      )}" dang thieu dung so do (${new Date().toLocaleDateString("vi-VN")}).`;
      if (!store.viet_qr_manual_matches[channelKey]) store.viet_qr_manual_matches[channelKey] = {};
      store.viet_qr_manual_matches[channelKey][`${settlementDate}|${fromCode}`] = {
        invoiceNumbers: lineFrom.invoiceNumbers,
        amount: lineFrom.gross,
        grossAdjustment: 0,
        note,
        created_at: new Date().toISOString(),
      };
      store.viet_qr_manual_matches[channelKey][`${settlementDate}|${toCode}`] = {
        invoiceNumbers: Array.from(new Set([...lineTo.invoiceNumbers, ...lineFrom.invoiceNumbers])),
        amount: lineTo.gross,
        grossAdjustment: 0,
        note,
        created_at: new Date().toISOString(),
      };
      applied++;
    }
    save(store);
    let msg = `Da doi tru hang loat ${applied} cap.`;
    if (skipped > 0) msg += ` Bo qua ${skipped} cap (du lieu da thay doi hoac khong con hop le).`;
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: raw QR export file (contains BOTH the transaction log sheet AND the "Cua hang" store-catalog sheet) ----------
router.post("/doi-soat/vietqr/upload-raw/:channel", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");

    // BIDV7702/VietQR MN dung dinh dang file rieng ("VIETQR MN 7702.xlsx":
    // sheet "VIET QR" + "Ma Cua Hang APP") -- khong co cot "Noi dung TT"/
    // token VQR nhu 3 kenh kia nen phai dung parser rieng.
    const isMn = CHANNELS[channelKey].parseMode === "mn";
    const parsed = isMn ? parseVietQrMnRawWorkbook(req.file.buffer) : parseVietQrRawWorkbook(req.file.buffer);
    // QUAN TRONG (Luyen, 2026-07-16): file giao dich re-export hang ngay
    // ("transactions_....xlsx") thuong CHI co sheet giao dich, KHONG kem sheet
    // "Cua hang"/"Ma Cua Hang APP" (sheet do da nap rieng qua nut "Danh sach
    // diem ban rieng" ben duoi, hoac tu 1 lan tai truoc). Truoc day
    // parseMaCuaHangAppSheet/parseCuaHangSheet throw loi khi thieu sheet nay
    // se lam HONG CA request -- 44xxx dong giao dich hop le cung bi mat theo,
    // dung y het loi "Chua co du lieu de doi soat" du da bam Tai len nhieu
    // lan. Bat loi rieng: thieu sheet cua hang thi chi coi la 0 cua hang MOI
    // (giu nguyen danh sach cua hang da co), KHONG chan viec luu cac dong
    // giao dich.
    let storeMap = {};
    try {
      storeMap = isMn ? parseMaCuaHangAppSheet(req.file.buffer) : parseCuaHangSheet(req.file.buffer);
    } catch (eStore) {
      storeMap = {};
    }

    // Chi Nhan, 2026-07-24: file "du lieu tai len" moi cho BIDV7702 (tu
    // 22/07) co kem theo 1 sheet "Ten diem- Ma cong trinh" trong CUNG file --
    // nap luon qua nut nay, best-effort giong storeMap o tren (khong bat buoc
    // phai co, cac kenh/lan tai khac khong co sheet nay van hoat dong binh
    // thuong).
    let tenDiemMasterMap = {};
    try {
      const parsedMaster = parseTenDiemMaCongTrinhSheet(req.file.buffer);
      tenDiemMasterMap = parsedMaster.map;
    } catch (eMaster) {
      tenDiemMasterMap = {};
    }

    store.viet_qr_raw_uploads[channelKey].push({
      id: nextId(store, "viet_qr_raw_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      rows: parsed.rows,
    });
    if (Object.keys(storeMap).length > 0) {
      store.viet_qr_store_names[channelKey] = Object.assign({}, store.viet_qr_store_names[channelKey], storeMap);
    }
    if (Object.keys(tenDiemMasterMap).length > 0) {
      store.viet_qr_ten_diem_master[channelKey] = Object.assign(
        {},
        store.viet_qr_ten_diem_master[channelKey],
        tenDiemMasterMap
      );
    }
    save(store);

    let msg = `Da nap "${parsed.sheetName}": ${parsed.rows.length} giao dich QR, ${Object.keys(storeMap).length} cua hang.${UPDATED_NOTE}`;
    if (Object.keys(tenDiemMasterMap).length > 0) {
      msg += ` Da nap them ${Object.keys(tenDiemMasterMap).length} dong "Ten diem- Ma cong trinh".`;
    }
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: danh sach diem ban rieng ("store_export"), doc lap voi
// file giao dich QR -- cho phep Luyen nap truoc bang ten diem ban moi nhat
// TRUOC KHI co giao dich, thay vi phai cho co du lieu QR moi biet ten. Neu 1
// ma cua hang la MOI (chua tung thay) hoac DOI ten diem ban so voi lan
// truoc, canh bao ngay trong thong bao de Luyen kiem tra xem co phai gian
// moi can gan vao ngan hang dang up hay khong. ----------
router.post("/doi-soat/vietqr/upload-store/:channel", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");

    const { sheetName, map } = parseStoreExportSheet(req.file.buffer);
    const existing = store.viet_qr_store_names[channelKey] || {};
    const newEntries = [];
    const changedEntries = [];
    Object.keys(map).forEach((maCuaHang) => {
      const prev = existing[maCuaHang];
      const next = map[maCuaHang];
      if (!prev) {
        newEntries.push(`${maCuaHang} -> ${next.tenDiemBan || next.matchText}`);
      } else if (next.tenDiemBan && prev.tenDiemBan !== next.tenDiemBan) {
        changedEntries.push(`${maCuaHang}: "${prev.tenDiemBan || prev.matchText}" -> "${next.tenDiemBan}"`);
      }
    });
    // Chi Nhan, 2026-07-22: luu lai LICH SU lan tai nay (thay vi ghi thang de
    // luc lo tai nham co the bam Xoa hoan tac -- xem ensureChannelShape/
    // mergeStoreNames o tren).
    store.viet_qr_store_uploads[channelKey].push({
      id: nextId(store, "viet_qr_store_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName,
      map,
    });
    store.viet_qr_store_names[channelKey] = mergeStoreNames(
      store.viet_qr_store_uploads[channelKey],
      store.viet_qr_store_names_baseline[channelKey]
    );
    save(store);

    let msg = `Da nap "${sheetName}": ${Object.keys(map).length} diem ban (${CHANNELS[channelKey].label}).${UPDATED_NOTE}`;
    if (newEntries.length > 0) {
      msg += ` CANH BAO: ${newEntries.length} ma cua hang MOI, kiem tra xem co phai gian moi khong -- ${newEntries
        .slice(0, 8)
        .join("; ")}${newEntries.length > 8 ? `... va ${newEntries.length - 8} ma khac` : ""}.`;
    }
    if (changedEntries.length > 0) {
      msg += ` CANH BAO: ${changedEntries.length} ma cua hang DOI TEN diem ban so voi lan truoc -- ${changedEntries
        .slice(0, 8)
        .join("; ")}${changedEntries.length > 8 ? `... va ${changedEntries.length - 8} ma khac` : ""}.`;
    }
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// Chi Nhan, 2026-07-22: "nhấn nhầm nạp nhầm chỗ này mà hk có nút xóa" -- xoa
// 1 lan tai "store_export" cu the, roi tinh lai viet_qr_store_names tu
// baseline + cac lan tai CON LAI (theo dung thu tu thoi gian) -- hoan tac
// dung 1 lan tai bi loi, khong dung den cac lan tai/sua tay khac.
router.post("/doi-soat/vietqr/upload-store/:channel/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const list = store.viet_qr_store_uploads[channelKey] || [];
    const idx = list.findIndex((u) => String(u.id) === req.params.id);
    if (idx === -1) throw new Error("Khong tim thay lan tai nay (co the da bi xoa roi).");
    list.splice(idx, 1);
    store.viet_qr_store_names[channelKey] = mergeStoreNames(list, store.viet_qr_store_names_baseline[channelKey]);
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da xoa lan tai danh sach diem ban do -- da tinh lai danh sach tu cac lan tai con lai.${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Khoa so: dat/xoa ngay khoa (den het ngay X) cho 1 kenh -- Luyen,
// 2026-07-21: "không cần chỉnh cái cũ khóa cho tôi" (sau khi da tim ra nguyen
// nhan Farm Lotte Nha Trang cho ngay 20/7, Luyen chon khoa CA KENH BIDV 7702
// den het 20/7 thay vi chi sua rieng ngay do/gian do). Xem xu ly tai
// buildChannelReconciliation (dua diff ve 0, danh dau locked cho moi dong tu
// ngay dau den ngay khoa). Gui lockDate rong ("") de MO khoa lai.
router.post("/doi-soat/vietqr/khoa-so/:channel", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const lockDate = (req.body.lockDate || "").trim();
    if (lockDate && !/^\d{4}-\d{2}-\d{2}$/.test(lockDate)) throw new Error("Ngay khoa khong hop le (dang YYYY-MM-DD).");
    store.viet_qr_lock_date[channelKey] = lockDate;
    save(store);
    const msg = lockDate
      ? `Da khoa so kenh "${CHANNELS[channelKey].label}" den het ngay ${lockDate}. Cac ngay tu do tro ve truoc se khong con hien canh bao lech/chua co HD nua.`
      : `Da mo khoa so kenh "${CHANNELS[channelKey].label}".`;
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Gan 1 Ma cua hang CHUA MAP duoc gian (tu unmappedStoreDetails)
// thang ve dung Ma cong trinh -- Luyen, 2026-07-17: "map giu ten diem noi bo
// co trong file he thong voi cac ma cong trinh ... ma nao chua co hay chua
// map duoc hien ra cho toi". Ghi de truc tiep vao viet_qr_store_names[channel]
// (giu nguyen tenCuaHang/tenDiemBan cu neu co de con hien thi, chi doi
// matchText -- dung dung field ma resolveGianGross dung de khop fuzzy) ap
// dung ngay, khong can tai lai file "store_export"/"Cua hang".
router.post("/doi-soat/vietqr/store-map/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { maCuaHang, targetCode } = req.body;
    if (!maCuaHang || !targetCode) throw new Error("Thieu ma cua hang hoac ma cong trinh de gan.");
    const existing = store.viet_qr_store_names[channelKey][maCuaHang] || {};
    store.viet_qr_store_names[channelKey][maCuaHang] = {
      tenCuaHang: existing.tenCuaHang || "",
      maDiemBan: existing.maDiemBan || "",
      tenDiemBan: targetCode,
      matchText: targetCode,
    };
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da gan ma cua hang "${maCuaHang}" -> "${targetCode}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Gan 1 giao dich "khong co ma cua hang" (rieng le, theo vqrCode)
// ve dung Ma cong trinh -- Luyen, 2026-07-17: "con loi chua map dc cai nao
// thi bao cho toi va co nut sua nha". Khac voi store-map o tren: moi giao
// dich blank-code chi co 1 vqrCode rieng, khong co ma cua hang chung de gan
// ca nhom, nen phai luu theo tung vqrCode (xem viet_qr_nocode_assignments,
// resolveGianGross doc lai o utils/vietqrReconcile.js). ap dung ngay, khong
// can tai lai file.
router.post("/doi-soat/vietqr/nocode-assign/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { vqrCode, targetCode } = req.body;
    if (!vqrCode || !targetCode) throw new Error("Thieu ma giao dich hoac ma cong trinh de gan.");
    store.viet_qr_nocode_assignments[channelKey][vqrCode] = {
      targetCode,
      assignedAt: new Date().toISOString(),
    };
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da gan giao dich "${vqrCode}" -> "${targetCode}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Xoa 1 gan giao dich "khong co ma cua hang" da luu -- Luyen,
// 2026-07-19: "co nut xoa hay chinh sua cac phan dien" -- truoc gio gan xong
// la het, khong xem/sua/xoa lai duoc; them nut Xoa de tra giao dich do ve lai
// dien "chua gan" (co the gan lai ma khac ngay sau do neu gan nham).
router.post("/doi-soat/vietqr/nocode-assign/:channel/:vqrCode/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    delete store.viet_qr_nocode_assignments[channelKey][req.params.vqrCode];
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" + encodeURIComponent(`Da xoa gan giao dich "${req.params.vqrCode}".${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: danh sach hoa don dung chung (file MTT), gan 3 tag Viet QR cung luc ----------
router.post("/doi-soat/vietqr/upload-hoadon", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const addedCounts = {};
    let sheetName = null;
    for (const ch of CHANNEL_KEYS) {
      const parsed = parseInvoiceWorkbookByTag(req.file.buffer, CHANNELS[ch].tagPattern);
      sheetName = parsed.sheetName;
      const existingKeys = new Set(store.viet_qr_invoices[ch].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
      let added = 0;
      for (const inv of parsed.invoices) {
        const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
        if (existingKeys.has(k)) continue;
        existingKeys.add(k);
        store.viet_qr_invoices[ch].push(inv);
        added++;
      }
      addedCounts[ch] = added;
    }
    save(store);
    const summary = CHANNEL_KEYS.map((ch) => `${CHANNELS[ch].label}: ${addedCounts[ch]}`).join(", ");
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da nap sheet "${sheetName}": them moi hoa don theo kenh -- ${summary}.${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vietqr/mapping", requireAdmin, (req, res) => {
  const store = load();
  const body = req.body || {};
  for (const [key, val] of Object.entries(body)) {
    if (key.startsWith("tkco_")) {
      const code = key.slice("tkco_".length);
      if (TKCO_VALUES.includes(val)) store.gian_mapping[code] = val;
    }
  }
  save(store);
  res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da luu bang TK Co theo gian."));
});

// ---------- Alias: dung CHUNG voi Momo/ZVP (store.invoice_diem_alias) ----------
router.post("/doi-soat/vietqr/diem-alias", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { sourceCode, targetCode } = req.body;
    if (!sourceCode || !targetCode) throw new Error("Thieu ma diem tren hoa don hoac ma cong trinh de anh xa.");
    if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
    store.invoice_diem_alias[sourceCode] = targetCode;
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" + encodeURIComponent(`Da anh xa "${sourceCode}" -> "${targetCode}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vietqr/diem-alias/delete", requireAdmin, (req, res) => {
  const store = load();
  const { sourceCode } = req.body;
  if (store.invoice_diem_alias) delete store.invoice_diem_alias[sourceCode];
  save(store);
  res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da xoa anh xa ma diem."));
});

router.post("/doi-soat/vietqr/upload-raw/:channel/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  if (!CHANNELS[channelKey]) return res.redirect("/doi-soat/vietqr");
  store.viet_qr_raw_uploads[channelKey] = store.viet_qr_raw_uploads[channelKey].filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/vietqr");
});

router.post("/doi-soat/vietqr/invoices/clear", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  CHANNEL_KEYS.forEach((ch) => (store.viet_qr_invoices[ch] = []));
  save(store);
  res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da xoa toan bo hoa don Viet QR da nap (ca 3 kenh)."));
});

// ---------- Manual match: dong "Chua co HD" da xac nhan la co HD bu ----------
router.post("/doi-soat/vietqr/manual-match", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    const { channel, settlementDate, code, invoiceNumbers, amount, note, grossAdjustment } = req.body;
    if (!CHANNELS[channel]) throw new Error("Kenh khong hop le.");
    if (!settlementDate || !code) throw new Error("Thieu thong tin dong can danh dau.");
    const key = `${settlementDate}|${code}`;
    const invoiceList = (invoiceNumbers || "")
      .split(/[,\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const amt = amount ? Number(String(amount).replace(/[^\d]/g, "")) : null;
    // grossAdjustment: cong THEM (hoac TRU BOT neu am) vao DOANH THU cua dong
    // nay (khac voi "amount" -- chi doi so tien HD hien thi) -- dung khi 1 GD
    // QR "khong co ma cua hang" thuoc ve gian nay, hoac khi Luyen biet du
    // lieu QR bi thieu/du so voi thuc te, de tong "Tinh tu du lieu tai len"
    // ca ngay khop dung voi Ngan hang, khong con hien "Chenh lech" sai.
    // Luyen, 2026-07-19: "co cho sua doanh thu" -- truoc day regex /[^\d]/g
    // xoa luon dau am, nhap "-50000" bi hieu thanh 50000 (cong nham thay vi
    // tru) -- giu lai dau "-" dau chuoi truoc khi loc so.
    let grossAdj = 0;
    if (grossAdjustment) {
      const raw = String(grossAdjustment).trim();
      const isNeg = raw.startsWith("-");
      const digits = raw.replace(/[^\d]/g, "");
      if (digits) grossAdj = (isNeg ? -1 : 1) * Number(digits);
    }
    store.viet_qr_manual_matches[channel][key] = {
      invoiceNumbers: invoiceList,
      amount: amt,
      grossAdjustment: grossAdj || 0,
      note: note || "",
      created_at: new Date().toISOString(),
    };
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" + encodeURIComponent(`Da danh dau thu cong dong "${code}" ngay ${settlementDate}.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vietqr/manual-match/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    const { channel, settlementDate, code } = req.body;
    if (store.viet_qr_manual_matches[channel]) {
      delete store.viet_qr_manual_matches[channel][`${settlementDate}|${code}`];
    }
    save(store);
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da xoa danh dau thu cong."));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Export: "MISATHue [kenh]" ----------
// Same 28-column "Mau phieu thu tien gui de nhap vao AMIS Accounting" as
// Momo/ZVP, with the SAME correct dien giai wording ("Thu tien dich vu vui
// choi giai tri theo HD ...") -- this was the whole point of building this
// as a proper reconciliation channel instead of hand-editing each bank's
// Excel file: the export now always uses that wording, no manual fixing.
function buildExportRows(reconciledList, startNo, bankAccount, bankFullName, lyDoThu) {
  let seq = startNo;
  const rows = [];
  reconciledList
    // pendingBank (Luyen, 2026-07-24): dong nay CHUA co giao dich ngan hang
    // that khop ngay -- chi hien de xem/theo doi truoc, KHONG dua vao file
    // xuat Misa cho den khi tien that ve ngan hang.
    .filter((r) => !r.pendingBank)
    .sort((a, b) => (a.settlementDate > b.settlementDate ? 1 : -1))
    .forEach((r) => {
      const exportableLines = r.lines.filter((l) => l.tkCo !== "SKIP");
      if (exportableLines.length === 0) return;

      const soCt = "NTTK" + String(seq).padStart(7, "0") + "/26";
      seq++;
      const ngayDmy = isoToDmy(r.settlementDate);
      exportableLines.forEach((l) => {
        const hdText = l.invoiceNumbers.length > 0 ? l.invoiceNumbers.join(", ") : "";
        // Hau to kenh (Luyen, 2026-07-16, sua lai 2026-07-16): "- KL" cho Viet
        // QR (thay vi "- QR") de phan biet voi Momo ("- MM") va VNPay/Zalo/
        // Payoo ("- VNP"/"- PAYOO") ngay tren file xuat Misa, dat truoc
        // "theo HD ..." neu co.
        const dienGiai = hdText
          ? `Thu tiền dịch vụ vui chơi giải trí - KL theo HĐ ${hdText}`
          : "Thu tiền dịch vụ vui chơi giải trí - KL";
        rows.push({
          "Ngày hạch toán (*)": ngayDmy,
          "Ngày chứng từ (*)": ngayDmy,
          "Số chứng từ (*)": soCt,
          "Mã đối tượng": "KL",
          "Tên đối tượng": "",
          "Địa chỉ": "",
          "Nộp vào TK": bankAccount,
          "Mở tại ngân hàng": bankFullName,
          "Lý do thu": lyDoThu,
          "Diễn giải lý do thu": dienGiai,
          "Mã nhân viên thu": "",
          "Diễn giải (hạch toán)": dienGiai,
          "TK Nợ (*)": 1121,
          "TK Có (*)": l.tkCo,
          "Số tiền": l.gross,
          "Mã đối tượng (hạch toán)": "KL",
          "Số khế ước đi vay": "",
          "Số khế ước cho vay": "",
          "Mã khoản mục chi phí": "",
          "Mã đơn vị": "",
          "Mã đối tượng THCP": "",
          "Mã công trình": l.maCongTrinh,
          "Số đơn đặt hàng": "",
          "Số đơn mua hàng": "",
          "Số hợp đồng mua": "",
          "Số hợp đồng bán": "",
          "Mã thống kê": "",
          "CP không hợp lý": "",
          "Số HĐ khớp": hdText,
          "Tổng tiền HĐ khớp": l.invoiceTotal,
          "Doanh thu (VietQR)": l.gross,
          "Chênh lệch HĐ vs doanh thu": l.diff,
          "Trạng thái": l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : l.matched ? "Khớp" : "Lệch",
        });
      });
    });
  return { rows, nextSeq: seq };
}

router.get("/doi-soat/vietqr/export.xlsx", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.query.channel;
  if (!CHANNELS[channelKey]) return res.status(400).send("Kenh khong hop le.");

  const bank = store.banks.find((b) => b.name === CHANNELS[channelKey].bankName);
  if (!bank) return res.status(400).send(`Chua co ngan hang "${CHANNELS[channelKey].bankName}" trong he thong.`);

  const built = buildChannelReconciliation(store, channelKey);
  if (built.error) return res.status(400).send(built.error);

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;
  // Luyen, 2026-07-21: "lý do thu là Thu tiền khách hàng (không theo hóa đơn)
  // đổi hết các file xuất misa nhá" -- dung 1 cau CO DINH giong het Momo, bo
  // cau rieng theo tung kenh nhu truoc (vd "...qua Viet QR (BIDV 7702)").
  const { rows } = buildExportRows(
    built.reconciled,
    startNo,
    bank.account_number,
    `Ngân hàng ${bank.bank_name}`,
    "Thu tiền khách hàng (không theo hóa đơn)"
  );

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, "MISAThue VietQR");
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=MISAThue-vietqr-${channelKey}.xlsx`);
  res.send(buf);
});

// ---------- Export: "Hóa đơn đầu ra" (mẫu VietInvoice) ----------
// Luyen, 2026-07-23: "từ chỗ Việt QR 7702 á bạn thêm cho tôi chỗ xuất ra Hóa
// đơn đầu ra nhá theo ngày lọc á file mẫu như thế này số thứ tự hóa đơn là
// cộng lên 1 số ngày là ngày lọc các dữ liệu liên quan thì cứ giữ nguyên trên
// file xuống chỉ thay đổi các số tiền theo từng dòng trên đối soát ngày hôm
// đó cột thành tiền là cột sau khi tiền trên viet qr của điểm đó trừ đi 8%
// tiền thuế là 8%" -- xuat 1 file .xlsx dung DUNG mau "VietInvoice" (file
// Luyen gui), 1 dong hoa don = 1 gian (Ma cong trinh) cua 1 NGAY doi soat cu
// the (chon qua input ngay tren trang), STT hoa don cong dan tu so bat dau
// (giong kieu "So chung tu bat dau" cua nut Xuat MISA). Cac cot khac giu
// NGUYEN gia tri mau mac dinh (Hinh thuc TT "TM/CK", Loai tien "VND", Ten
// hang hoa/dich vu "Dich vu vui choi giai tri"...), CHI doi STT/Ngay/Thanh
// tien/Tien thue GTGT/Don gia theo doanh thu VietQR that cua tung gian ngay
// do. Thanh tien = tien nhan tren VietQR (gross, DA GOM thue) / 1.08 (tach
// thue GTGT 8% ra, giong dung cong thuc cac dong vi du co san trong file mau:
// 1928000 / 1.08 = 1785185, Tien thue GTGT = 1928000 - 1785185 = 142815).
const HOA_DON_DAU_RA_HEADER_INFO = [
  ["FILE MẪU DANH SÁCH HÓA ĐƠN ĐỂ NHẬP VÀO PHẦN MỀM VIETINVOICE"],
  ["Hướng dẫn:"],
  ["- Điền dữ liệu hóa đơn cần lập trên phần mềm vào các cột tương ứng trên file này"],
  ["- Các cột có dấu (*) là những cột bắt buộc"],
  [
    "- Nếu hóa đơn chiết khấu theo tổng tiền hàng thì điền thông tin về tỷ lệ CK và tiền CK ở cột màu tím. Nếu chiết khấu theo từng mặt hàng thì điền thông tin ở cột màu vàng",
  ],
  ['- Loại tiền tệ lấy theo cột "Mã loại tiền" trong chức năng "Danh mục => Loại tiền"'],
  ['- Mã khách hàng (cột D) chỉ hợp lệ nếu đã tồn tại trong chức năng "Danh mục => Khách hàng"'],
  ['- Mã hàng chỉ hợp lệ nếu đã tồn tại trong chức năng "Danh mục => Hàng hóa, dịch vụ"'],
  ["- Các dòng dữ liệu phía dưới chỉ là ví dụ minh họa"],
  ["- Hệ thống sử dụng dấu '.' để phân tách các chữ số hàng nghìn và dấu ',' để phân tách các chữ số phần thập phân"],
  [],
];
const HOA_DON_DAU_RA_HEADER_ROW = [
  "Số thứ tự hóa đơn (*)",
  "Ngày hóa đơn",
  "Tên đơn vị mua hàng",
  "Mã khách hàng",
  "Địa chỉ",
  "Mã số thuế",
  "Người mua hàng",
  "Email",
  "CMND/CCCD",
  "Số hộ chiếu",
  "Mã DVQHNS",
  "Hình thức thanh toán",
  "Loại tiền",
  "Tỷ giá",
  "Tỷ lệ CK(%)",
  "Tiền CK",
  "% thuế GTGT",
  "Tiền thuế GTGT",
  "Tên hàng hóa/dịch vụ (*)",
  "Mã hàng",
  "ĐVT",
  "Số lượng",
  "Đơn giá",
  "Tỷ lệ CK (%)",
  "Tiền CK",
  "Thành tiền(*)",
];
const HOA_DON_DAU_RA_VAT_RATE = 0.08;

router.get("/doi-soat/vietqr/xuat-hoa-don-dau-ra", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.query.channel;
  if (!CHANNELS[channelKey]) return res.status(400).send("Kênh không hợp lệ.");
  // Luyen, 2026-07-23 (lan 3): "cho cái lọc đi từ ngày mấy tới ngày mấy á" --
  // thay vi 1 thang co dinh (lan 2) hay 1 ngay don le (lan 1), cho chon 1
  // khoang ngay tuy y (tu ngay - den ngay, ca 2 dau bao gom) -- xuat HET cac
  // ngay doi soat nam trong khoang do, lien tuc STT qua nhieu ngay.
  const fromDate = (req.query.from || "").trim();
  const toDate = (req.query.to || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate) || !/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
    return res.status(400).send("Thiếu hoặc sai định dạng ngày (YYYY-MM-DD) -- chọn đủ Từ ngày và Đến ngày trước.");
  }
  if (fromDate > toDate) {
    return res.status(400).send("Từ ngày phải nhỏ hơn hoặc bằng Đến ngày.");
  }
  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;

  const built = buildChannelReconciliation(store, channelKey);
  if (built.error) return res.status(400).send(built.error);

  const days = (built.reconciled || [])
    .filter((r) => r.settlementDate >= fromDate && r.settlementDate <= toDate)
    .sort((a, b) => (a.settlementDate > b.settlementDate ? 1 : -1));
  if (days.length === 0) {
    return res.status(400).send(`Không có dữ liệu đối soát từ ${fromDate} đến ${toDate} cho kênh này.`);
  }

  const dataRows = [];
  let seq = startNo;
  days.forEach((day) => {
    const ngayHoaDon = isoToDmy(day.settlementDate);
    day.lines
      .filter((l) => l.tkCo !== "SKIP" && Math.round(l.gross) !== 0)
      .forEach((l) => {
        const grossTotal = Math.round(l.gross);
        const thanhTien = Math.round(grossTotal / (1 + HOA_DON_DAU_RA_VAT_RATE));
        const tienThueGtgt = grossTotal - thanhTien;
        dataRows.push([
          seq,
          ngayHoaDon,
          "Bán cho người tiêu dùng ",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "TM/CK",
          "VND",
          "",
          "",
          "",
          "8",
          tienThueGtgt,
          "Dịch vụ vui chơi giải trí",
          "",
          "Kỳ ",
          "1",
          thanhTien,
          "",
          "",
          thanhTien,
        ]);
        seq++;
      });
  });

  if (dataRows.length === 0) {
    return res.status(400).send(`Từ ${fromDate} đến ${toDate} không có gian nào có doanh thu (khác 0) để xuất hóa đơn.`);
  }

  const aoa = [...HOA_DON_DAU_RA_HEADER_INFO, HOA_DON_DAU_RA_HEADER_ROW, ...dataRows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, "Hóa đơn");
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=HoaDonDauRa-${channelKey}-${fromDate}_${toDate}.xlsx`);
  res.send(buf);
});

// Chi Nhan, 2026-07-22: file "/he-thong/sao-luu/tai-xuong" (tai xuong toan bo
// du lieu) gio da qua lon, hay bi 502/timeout khi tai xuong -- route nay tra
// ve tom tat NHE (chi ten file, thoi gian tai, so dong -- KHONG kem toan bo
// du lieu giao dich) cho 1 kenh cu the, de kiem tra/xoa dung lan tai can
// thiet ma khong can tai ca file sao luu nang.
router.get("/doi-soat/vietqr/debug/:channel", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  if (!CHANNELS[channelKey]) return res.status(404).json({ error: "Kenh khong hop le." });
  const raw = (store.viet_qr_raw_uploads[channelKey] || []).map((u) => ({
    id: u.id,
    file_name: u.file_name,
    uploaded_at: u.uploaded_at,
    rowCount: (u.rows || []).length,
  }));
  const storeUps = (store.viet_qr_store_uploads[channelKey] || []).map((u) => ({
    id: u.id,
    file_name: u.file_name,
    uploaded_at: u.uploaded_at,
    entryCount: Object.keys(u.map || {}).length,
  }));
  const invoices = store.viet_qr_invoices[channelKey] || [];
  res.json({
    channel: channelKey,
    rawUploads: raw,
    storeUploads: storeUps,
    invoiceCount: invoices.length,
    storeNamesCount: Object.keys(store.viet_qr_store_names[channelKey] || {}).length,
  });
});

// Exposed so routes/dashboard.js (Tong quan / Cong no) can reuse the exact
// same per-channel reconciliation this page shows, without a second
// implementation. CHANNELS/CHANNEL_KEYS let the caller loop over all 3 Viet
// QR banks (bidv7704, bidv77020, mb11521268) without hard-coding the list
// twice.
router.buildChannelReconciliation = buildChannelReconciliation;
router.VIETQR_CHANNELS = CHANNELS;
router.VIETQR_CHANNEL_KEYS = CHANNEL_KEYS;

module.exports = router;
