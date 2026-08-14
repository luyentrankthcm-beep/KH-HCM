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
  parseVnpayPortalWorkbook,
  migrateVnpayKhMoiInvoices,
  seedMb02865168FromBidv77021,
  migrateTk168Invoices,
  fixBidv7702Day3TaggedAsDay4,
  fixBidv8613600999Day3TaggedAsDay4,
  fixBidv77021Day3TaggedAsDay4,
  parseCuaHangSheet,
  parseStoreExportSheet,
  parseTenDiemMaCongTrinhSheet,
  parseInvoiceWorkbookByTag,
  buildGianCandidatesFromInvoices,
  resolveGianGross,
  resolveGianGrossByBankRef,
  reconcileVietQr,
  applyMultiDayGroupConsolidation,
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

// Luyen, 2026-07-31: mac dinh 2 o "Tu ngay/Den ngay" cua nut Xuat file MISA
// (xem GET /doi-soat/vietqr/export.xlsx) = dau/cuoi thang dang chon o dropdown
// "Chon thang" tren trang xem, giong het ben Momo (routes/doisoat.js).
function monthBounds(m) {
  if (!m) return { first: "", last: "" };
  const [y, mo] = m.split("-").map(Number);
  const lastDay = new Date(y, mo, 0).getDate();
  return { first: `${m}-01`, last: `${m}-${String(lastDay).padStart(2, "0")}` };
}

// Each Viet QR bank posts individual QR payments (no batch settlement), so
// the reconciliation is per calendar day per bank. Store.transactions'
// bank.name is the join key against store.banks; the tag pattern is how a
// shared invoice-list upload (same "MTT" file used by Momo/ZVP) tells this
// bank's invoices apart from the others' -- see utils/vietqrReconcile.js.
const CHANNELS = {
  bidv7704: { bankName: "BIDV7704", label: "BIDV 7704", tagPattern: /POSH\+JP\s*MB\s*\(7704\)/i, company: "kh_cu" },
  bidv77020: {
    bankName: "BIDV77020",
    label: "BIDV 77020",
    tagPattern: /POSH\+JP\s*MB\s*\(7020\)/i,
    company: "kh_cu",
    // Chi Nhan, 2026-07-29: "bỏ qua các giao dịch khác ngoài viet qr nhá nhưng
    // nếu nó có tiền viet qr về mà dữ liệu hk có thì cho vô gian SB CAM RANH
    // PHN cho tôi nhá phần dư của ngân hàng mẫu diễn giải viêt qr á" -- kenh
    // nay KHONG dung refMatchFrom (khong khop tung giao dich ngan hang theo So
    // tham chieu nhu 7702/77021), nen khong the tai su dung
    // resolveGianGrossByBankRef/tanPhuSuggestion. "Ngan hang" cua kenh nay da
    // CHI tinh tu extractVietQrSettlements (loai san cac GD khong phai VietQR
    // qua NON_VQR_TX_PATTERN/excludeFromVietQrRecon) nen phan "bo qua GD khac
    // ngoai viet qr" da dung san, khong can sua them. Phan con lai (so du
    // "Ngan hang" > "Tinh tu du lieu tai len" sau khi da loai GD khong phai
    // VietQR) duoc TU DONG cong thang vao gian nay (khong chi goi y nhu
    // tanPhuSuggestion) -- xem block "bankExcessDefaultCode" trong
    // buildChannelReconciliation ben duoi.
    bankExcessDefaultCode: "SB CAM RANH PHN",
  },
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
    // Chi Nhan, 2026-07-30: "thêm 1 ngân hàng 02865168 á các điểm bán với mã
    // công trình map với 77021" -- phat hien tag "Dich vu thu ho" tren file
    // MTT dung chung THUC RA co 2 bien the: "VietQR POSH MB [ngay]" (dung cho
    // CHINH BIDV77021) va "VietQR POSH MB tk 168 [ngay]" (dung cho tai khoan
    // MOI MB02865168, xem CHANNELS.mb02865168 ben duoi) -- pattern cu
    // (/VietQR\s*POSH\s*MB/i) khop CA 2 bien the vi khong loai truong hop co
    // hau to "tk 168", khien hoa don cua MB02865168 bi gom NHAM vao pool
    // BIDV77021 tu truoc gio (2293 hoa don bidv77021, trong do co ca cap
    // "VietQR POSH MB X" + "VietQR POSH MB tk 168 X" cho MOI ngay). Them
    // negative lookahead loai bien the co "tk" ngay sau, giu nguyen bien the
    // khong "tk" cho kenh nay. migrateTk168Invoices() (utils/vietqrReconcile.js)
    // tu dong don cac hoa don "tk 168" DA LO nam trong pool nay tu truoc sang
    // dung pool mb02865168 moi lan load().
    tagPattern: /VietQR\s*POSH\s*MB(?!\s*tk)/i,
    company: "kh_moi",
    refMatchFrom: "2026-01-01",
    // Luyen, 2026-07-28: "lệch tiền á nếu nó có diễn giải giống của viet qr
    // nhưng nó không có tên điểm thì đưa vô hải phòng nhá" -- giao dich thu
    // KHONG khop duoc So tham chieu nao, hoac khop duoc nhung Ma cua hang
    // chua co ten diem, tu dong gan vao AE HP PHN thay vi nam rieng trong
    // canh bao "khong khop" (xem defaultBlankCode trong resolveGianGrossByBankRef).
    defaultBlankCode: "AE HP PHN",
  },
  // Chi Nhan, 2026-07-30: "đây là viet qr kh mới thêm cho tôi vô trong đối
  // soát viet qr nhá thêm 1 ngân hàng 02865168 á các điểm bán với mã công
  // trình map với 77021 á check cho tôi nhá" -- tai khoan MB02865168 (da tao
  // TRUOC do trong 1 phien lam viec truoc, luc do gia dinh SAI la doanh thu
  // VNPay/Payoo -- da xac nhan LAI day chinh la 1 kenh VietQR THAT (giao dich
  // tung cai 1, dinh dang "Transactions" chuan, xem sao ke da co san 7354
  // giao dich voi truong "reference" dung dinh dang FT... giong bidv7702/77021),
  // dung CHUNG may quet QR/Ma cua hang voi BIDV77021 (96/97 ma trung khop
  // 100% qua kiem tra) nen tai su dung TRUC TIEP storeNames/tenDiemMaster cua
  // bidv77021 qua seedMb02865168FromBidv77021() (utils/vietqrReconcile.js,
  // goi trong ensureChannelShape) thay vi phai cho chi Nhan tai rieng 1 file
  // "Danh sach diem ban" moi cho tai khoan nay.
  mb02865168: {
    bankName: "MB02865168",
    label: "MB 02865168",
    // Chi Nhan, 2026-07-30 (lan 2): "đây của JP VINPEARL NT đây nhá lí hiệu
    // với cách note vậy nè" -- hoa don 8021 (3.050.000d, dung cho JP VINPEARL
    // NT, ngay 1) dung tag "QR JP tk 168 1" (KHONG PHAI "VietQR POSH MB tk 168"
    // nhu da biet truoc do) -- xac nhan qua sheet goc "kê ds xuất HĐ MTT - 705"
    // cot "Dịch vụ thu hộ": biến thể "QR JP tk 168" xuat hien ~34 lan trong
    // file MTT 29.07 (cung tai khoan 168, chi khac quy uoc dan tag theo khu
    // vuc/nguoi nhap), truoc gio KHONG khop pattern cu nen hoa don loai nay
    // hoan toan khong nam trong bat ky kenh nao (khong phai chi bidv77021 nhu
    // truong hop tk168 truoc). Mo rong pattern de bat ca 2 bien the.
    tagPattern: /(?:VietQR\s*POSH\s*MB|QR\s*JP)\s*tk\s*168/i,
    company: "kh_moi",
    refMatchFrom: "2026-01-01",
    defaultBlankCode: "AE HP PHN",
  },
  // Chi Nhan, 2026-07-30 (lan 3): "thêm 1 ngân hàng 8613600999 ... thêm bên
  // đối soát viet qr cho tôi nhá kh mới" -- tai khoan BIDV 8613600999
  // (Cong Ty Tnhh Giai Tri K&H), xac nhan CUNG he thong may POS/QR voi
  // BIDV77021/MB02865168 (45/46 "Ma cua hang" tren sao ke thang 07/2026 trung
  // khop tuyet doi voi danh sach da co cua bidv77021) nen tai su dung
  // storeNames/tenDiemMaster cua bidv77021 qua seedFromBidv77021()
  // (utils/vietqrReconcile.js). "Số tham chiếu" tren sao ke + "Mã tham chiếu"
  // tren file QR cung dinh dang chuoi (vd "868rcpe-8AGY9N21t") giong het
  // bidv77021/mb02865168 nen dung duoc refMatchFrom tu dau.
  // Chi Nhan, 2026-07-30 (lan 4): "đây là các hóa đơn của 8613600999 Viet qr
  // nhá" -- gui anh chup sheet "kê ds xuất HĐ MTT - 705" xac nhan tag that la
  // "QR JP <ngay>" (vd "QR JP 30", "QR JP 1", "QR JP 4,5" cho hoa don gop 2
  // ngay) -- KHONG co hau to "tk <so>" nao ca (khac voi doan ban dau). Day
  // chinh la phan "QR JP" THUAN (khong "tk 168") ma truoc do van con ~102 dong
  // chua ro thuoc kenh nao trong file MTT 29.07 (phan biet voi "QR JP tk 168"
  // cua MB02865168 qua negative lookahead). Cac ma gian tren hoa don (JP
  // VINPEARL HOI AN, JP VC BA TRIEU, JP VC ROY, JP AE LONG BIEN, JP POSH VW
  // VU YEN, JP GO DA NANG, IPH PHN...) khop dung voi cac ma da resolve duoc
  // tu du lieu QR/store names seed tu bidv77021 (xac nhan seedFromBidv77021
  // dung kenh).
  bidv8613600999: {
    bankName: "BIDV8613600999",
    label: "BIDV 8613600999",
    tagPattern: /QR\s*JP(?!\s*tk)/i,
    company: "kh_moi",
    refMatchFrom: "2026-01-01",
    defaultBlankCode: "AE HP PHN",
  },
  // Chi Nhan, 2026-07-29: kenh "vnpayKhMoi" (MB02865168, kien truc VietQR
  // theo ngay) tung o day da bi GO BO -- chi xac nhan lai sau: "tôi nhầm rồi
  // cái vn pay này trả vè ngân hàng VTB982 á đây á với payoo cx về đây á",
  // nghia la doanh thu TUTU TRAIN/SAVICO thuc ra ve TAI KHOAN VTB982
  // (Vietinbank), tra tien THEO DOT (nhieu ngay gop 1 lan giong VNPay
  // Offline/Payoo cua KH Cu), KHONG PHAI tung giao dich 1 nhu VietQR -- nen
  // KHONG dung duoc kien truc file nay nua. Da xay lai dung kien truc ZVP
  // (extractZvpSettlements/reconcileZvpChannel) tai route rieng
  // routes/doisoat-vnpay-khmoi.js (trang /doi-soat/vnpay-khmoi), khong con
  // phu thuoc CHANNELS/pageGroup cua file nay. store.viet_qr_invoices.vnpayKhMoi
  // (hoa don da di sang qua migrateVnpayKhMoiInvoices) van con dung, chi doi
  // NOI TIEU THU sang file moi do.
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
  // Luyen, 2026-07-31: "tk 7704 đổi từ mã SB PHU QUOC PHCM qua đây cho tôi
  // nhá CHKQT PHU QUOC" -- gop vinh vien (ca gross/QR lan hoa don) ve dung ma
  // chuan da dung cho kenh mb11521268 ("CHKQT PHU QUOC", xem SÂN BAY PHÚ QUỐC
  // trong INVOICE_DIEM_ALIAS_DEFAULTS o duoi).
  "SB PHU QUOC PHCM": { maCongTrinh: "CHKQT PHU QUOC", isCse: false },
};

// Chi Nhan, 2026-07-29: "chia hóa đơn bị nhầm vô như 4 gian tôi nói á bị xuất
// chung hóa đơn á cấn trừ qua giúp tôi ... lưu luôn loại này để mốt lệch t2 á"
// -- 2 cap gian nay (kenh BIDV77021) LUON duoc nhan vien xuat hoa don GOP
// CHUNG 1 hoa don/ngay duoi TEN CHI 1 TRONG 2 gian (khong tach rieng hoa don
// cho tung diem that), nen 1 ben luon hien "Chua co HD" con ben kia luon
// "Lech" (hoa don du dung phan cua ben con lai). Truoc day phai tu tao tung
// "Doi tru thu cong" (store.viet_qr_manual_matches) rieng cho MOI ngay (~25
// lan cho ngay 1-27/7) -- Chi Nhan xac nhan 2026-07-29 day la CACH XUAT HOA
// DON CO DINH (khong phai loi 1 lan), nen chuyen thanh QUY TAC TU DONG ap
// dung cho MOI ngay (ca truoc 28/7 chua co manual match, lan tuong lai) qua
// applyInvoiceSharePairs ben duoi -- khong con phai tao tay moi ngay nua.
// Chay SAU reconcileVietQr (dung invoiceTotal DA duoc chia theo ngay cho hoa
// don gop nhieu ngay/T7+CN, xem ghi chu Math.round trong utils/vietqrReconcile.js)
// nen tu dong ket hop dung ca 2 truong hop (hoa don gop 2 gian VA hoa don gop
// 2 ngay) ma khong can code rieng cho truong hop gop ngay.
const INVOICE_SHARE_PAIRS = {
  bidv77021: [
    ["VC TUYEN QUANG PNH", "NSTV TUYEN QUANG PHN"],
    ["VC T.PHU N.TRANG PHN", "VC MAXI TN NT PHN"],
  ],
  // Chi Nhan, 2026-07-30: "2 cái này này gom lại xuất chung á bạn cấn trừ 2
  // ngày này cho tôi nhá" (ngay 18-19/7, AM TP PHCM du +177.399d, LM NHA TRANG
  // KVC thieu dung -177.399d) -- da ra soat CA THANG: da so cac ngay (01,02,
  // 03,06,07,08,10,13-17,20-24/7) nhan vien xuat 2 dong hoa don RIENG cung 1
  // "Số HĐ" (1 cho AM TP PHCM, 1 cho LM NHA TRANG KVC) nen tu khop dung san,
  // nhung MOT SO ngay (04,05,09,11,12,18,19/7) chi xuat 1 dong GOP CHUNG ca 2
  // gian duoi ten "AM TP PHCM", khong co dong rieng cho LM NHA TRANG KVC --
  // giong het co che INVOICE_SHARE_PAIRS.bidv77021 o tren (1 hoa don/ngay xuat
  // chung 2 gian). Ap dung applyInvoiceSharePairs se tu dong chia lai theo ty
  // le doanh thu MOI NGAY -- vo hai voi cac ngay da khop san (chia lai ra dung
  // y nguyen ket qua cu, da kiem chung qua toan bo thang 7).
  bidv7702: [["AM TP PHCM", "LM NHA TRANG KVC"]],
};

// Chi Nhan, 2026-07-30: "từ 29 tôi đã tách ra theo đúng tên rồi á khỏi cấn
// trừ mấy ngày trước các ngày trc vẫn giữ nguyên còn các ngày sao này từ 29
// trở đi đúng tên gian đó ln á khỏi cấn trừ tk 77021 thôi nhá" -- ke tu hoa
// don ngay doanh thu 29/7, nhan vien da tach hoa don RIENG dung ten cho tung
// cap gian (vd "VC T.PHU N.TRANG PHN" / "VC MAXI TN NT PHN" 29/7 tro di co 2
// So HD RIENG, khong con xuat chung 1 hoa don duoi 1 ten nhu truoc) -- ap
// dung applyInvoiceSharePairs/applySharePairChainNetting cho cac ngay nay se
// sai (chia lai 1 cach khong can thiet 2 hoa don da dung san, tron lan
// invoiceNumbers cua 2 dong lam 1). Gioi han CHI ap dung cho bidv77021 voi
// ngay doanh thu (r.from) TRUOC ngay nay -- CHI bidv77021 (Chi Nhan xac nhan
// "tk 77021 thoi nha"), bidv7702 khong doi.
const INVOICE_SHARE_PAIRS_STOP_FROM = {
  bidv77021: "2026-07-29",
  // Luyen, 2026-08-10: tu T8 hop dong Nha Trang het, khong con cap AM TP PHCM /
  // LM NHA TRANG KVC xuat chung hoa don nua -- tat co che chia tu dong tu ngay nay.
  bidv7702: "2026-08-01",
};

// Chi Nhan, 2026-07-30: xem ghi chu day du o cho goi -- CHI 2 kenh nay duoc
// tu dong cong phan du "Ngan hang > Du lieu" (khong co so tham chieu khop)
// vao gian mac dinh cua kenh. Cac kenh khac (mb02865168, bidv8613600999...)
// de nguyen thanh "Lech Ngan hang-Du lieu" cho Chi Nhan tu can doi tay.
const TAN_PHU_AUTO_APPLY_CHANNELS = new Set(["bidv7702", "bidv77021"]);

// Luyen, 2026-08-12: "bỏ cái bù trừ 80 ngàn này cho tôi đi" -- ngay 9/8
// kenh bidv77021 con du 80.000d (sau khi bỏ MANUAL 3 gian ngay CN), Luyen
// muon can tru thu cong, khong can he thong tu dong gop vao AE HP PHN.
// Dung Set de tat tan goc, tranh hien banner va tranh lam xau so lieu gian
// mac dinh. "Lech Ngan hang-Du lieu" van hien ra cho ngay do (dung, vi co
// tien that chua duoc can tru) -- Luyen biet va xu ly tay.
const TAN_PHU_SUPPRESS = {
  bidv77021: new Set(["2026-08-09"]),
};

function applyInvoiceSharePairs(reconciled, pairs, stopFrom) {
  if (!pairs || pairs.length === 0) return;
  reconciled.forEach((r) => {
    if (stopFrom && r.from >= stopFrom) return;
    pairs.forEach(([codeA, codeB]) => {
      const lineA = r.lines.find((l) => l.code === codeA);
      const lineB = r.lines.find((l) => l.code === codeB);
      if (!lineA || !lineB) return; // ngay nay 1 trong 2 khong co doanh thu -- khong co gi de chia
      const combinedGross = lineA.gross + lineB.gross;
      const combinedInvoiceTotal = lineA.invoiceTotal + lineB.invoiceTotal;
      const combinedInvoiceNumbers = Array.from(new Set([...lineA.invoiceNumbers, ...lineB.invoiceNumbers]));
      if (combinedGross <= 0 || combinedInvoiceNumbers.length === 0) return;
      const newTotalA = Math.round((combinedInvoiceTotal * lineA.gross) / combinedGross);
      const newTotalB = combinedInvoiceTotal - newTotalA;
      const note = `Tu dong chia theo ty le doanh thu: hoa don cua "${codeA}"/"${codeB}" do nhan vien xuat chung 1 hoa don/ngay duoi 1 ten, he thong tu chia lai theo dung doanh thu ngan hang tung gian.`;
      Object.assign(lineA, {
        invoiceNumbers: combinedInvoiceNumbers,
        invoiceTotal: newTotalA,
        diff: newTotalA - lineA.gross,
        matched: Math.abs(newTotalA - lineA.gross) < 1,
        manualOverride: true,
        manualNote: note,
      });
      Object.assign(lineB, {
        invoiceNumbers: combinedInvoiceNumbers,
        invoiceTotal: newTotalB,
        diff: newTotalB - lineB.gross,
        matched: Math.abs(newTotalB - lineB.gross) < 1,
        manualOverride: true,
        manualNote: note,
      });
    });
  });
}

// Chi Nhan, 2026-07-30: "2 cái này này gom lại xuất chung á bạn cấn trừ 2
// ngày này cho tôi nhá" -- sau khi applyInvoiceSharePairs chia lai gross theo
// TUNG NGAY, mot cap gian (vd AM TP PHCM / LM NHA TRANG KVC) van co the con
// lech CA CAP tren tung ngay rieng (vi hoa don thuc su gop ca 2 ngay lam 1,
// nhung applyInvoiceSharePairs khong biet dieu do, chi chia trong PHAM VI 1
// ngay) -- vd ngay 18 lech -177.399d, ngay 19 lech +177.399d, cong lai vua
// dung bang 0. Khac voi "chuoi ngay lien tiep" cua reconcileVietQr (chi xet
// TUNG MA rieng le, chay TRUOC applyInvoiceSharePairs nen chua thay duoc lech
// cap nay), ham nay xet TONG LECH CUA CA CAP (2 ma cong) tren tung ngay, tim
// chuoi ngay lien tiep (chi trong so cac ngay CHUA khop ca 2 ma) cong lai
// bang 0 thi coi ca chuoi da khop.
function applySharePairChainNetting(reconciled, pairs, stopFrom) {
  if (!pairs || pairs.length === 0) return;
  const sorted = reconciled.slice().sort((a, b) => (a.settlementDate < b.settlementDate ? -1 : 1));
  pairs.forEach(([codeA, codeB]) => {
    // Danh sach TOAN BO ngay co ca 2 ma (khong loc truoc theo matched) de giu
    // dung tinh KE NHAU ve ngay -- 1 ngay da khop san (ca 2 ma) dong vai tro
    // "ranh gioi" chan chuoi, giong het co che byCode trong reconcileVietQr.
    // Xem ghi chu INVOICE_SHARE_PAIRS_STOP_FROM o tren -- ngay >= stopFrom
    // (hoa don da tach rieng dung ten) bi loai KHOI danh sach hoan toan
    // (khong phai chi "khong xet") de khong lam sai tinh KE NHAU cua cac ngay
    // TRUOC do van con dung co che cu.
    const entries = [];
    sorted.forEach((day) => {
      if (stopFrom && day.from >= stopFrom) return;
      const lineA = day.lines.find((l) => l.code === codeA);
      const lineB = day.lines.find((l) => l.code === codeB);
      if (!lineA || !lineB) return;
      entries.push({
        lineA,
        lineB,
        date: day.settlementDate,
        isFree: !(lineA.matched && lineB.matched),
        combinedDiff: lineA.diff + lineB.diff,
      });
    });
    let i = 0;
    while (i < entries.length) {
      if (!entries[i].isFree) {
        i++;
        continue;
      }
      let j = i;
      while (j + 1 < entries.length && entries[j + 1].isFree) j++;
      if (j > i) {
        const run = entries.slice(i, j + 1);
        const sum = run.reduce((s, e) => s + e.combinedDiff, 0);
        if (Math.abs(sum) < 1) {
          const note = `Tu dong can tru theo cap gian "${codeA}"/"${codeB}" giua cac ngay ${run[0].date}..${run[run.length - 1].date} (hoa don xuat chung nhieu ngay, tong lech ca cap ve dung 0).`;
          run.forEach((e) => {
            e.lineA.matched = true;
            e.lineA.diff = 0;
            e.lineA.manualOverride = true;
            e.lineA.manualNote = note;
            e.lineB.matched = true;
            e.lineB.diff = 0;
            e.lineB.manualOverride = true;
            e.lineB.manualNote = note;
          });
        }
      }
      i = j + 1;
    }
  });
}

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
  // Luyen, 2026-08-10: hop dong Nha Trang (Posh/JP/Farm) het -- tu thang 8 tro
  // di cac giao dich BIDV7702 co tham chieu "LM NHA TRANG KVC" hach toan vao
  // "POSH MN CGV Vincom Xuan Khanh" thay vi "FARM LOTTE NHA TRANG".
  // Du lieu thang 6-7 (LM NHA TRANG KVC cu) da doi chieu xong, doi alias se
  // hien thi lai ten moi cho cac thang do -- chap nhan duoc vi Luyen xac nhan.
  "LM NHA TRANG KVC": "POSH MN CGV Vincom Xuân Khánh",
  "JP vicom 3.2 BIDV": "VC 3/2 JP-Posh",
  "VINCOM GAND PARK": "JP-POSH GRAND PARK",
  // Chi Nhan, 2026-07-29: "mã công trình JPSBNB đổi thành JP SB NOI BAI đúng
  // mẫu chuẩn mã công trình cho tôi nhá" -- "JPSBNB" (dung lam ma NOI BO de
  // khop gross/QR va hoa don, xem TEN_DIEM_MASTER_DEFAULTS.bidv77020 o duoi)
  // khong fuzzy-khop duoc voi ten chuan trong ma_cong_trinh_master (danh sach
  // goc co san "JP SB NOI BAI", stt 75) vi 2 chuoi qua khac nhau (viet tat vs
  // day du), nen hien thi van con "JPSBNB". Alias hien thi thang ve dung ten
  // chuan; KHONG doi ma noi bo dung de doi soat (van la "JPSBNB").
  JPSBNB: "JP SB NOI BAI",
  // Luyen, 2026-08-10: doi ten hien thi gian LOTTE GO VAP (KH Moi, BIDV7702)
  // sang ten chuong trinh chinh thuc moi.
  "LOTTE GO VAP VR-PHN": "POSH MN CGV Vincom Phan Văn Trị",
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
  // Chi Nhan, 2026-07-29: hoa don cua mb11521268 cho san bay Phu Quoc ghi
  // maDiem "SÂN BAY PHÚ QUỐC" (ten tho), trong khi ben gross/QR (sau khi map
  // theo bidv77021, xem cross-copy viet_qr_store_names trong ensureChannelShape)
  // lai tu dong quy ve ma chuan "CHKQT PHU QUOC" (qua master gian sheet dung
  // chung ca kenh) -- can alias de 2 ben khop nhau, khong thi hoa don luon
  // hien "Chua co HD" du da co hoa don that.
  "SÂN BAY PHÚ QUỐC": "CHKQT PHU QUOC",
  // Chi Nhan, 2026-07-30: "bạn đã map đúng từ mã điểm qua tên điểm qua tên mã
  // công trình giống với 77021 chưa á" -- ra soat kenh bidv77020 phat hien hoa
  // don thang 7 dung LAN LON 2 cach ghi maDiem cho CUNG 1 diem San bay Noi Bai:
  // 25 hoa don ghi tat "JPSBNB" (27.850.000d, khop dung ma noi bo dang dung --
  // xem TEN_DIEM_MASTER_DEFAULTS.bidv77020/JPSBNB o tren), nhung so con lai ghi
  // day du "JP SB NOI BAI" (28.350.000d) -- KHONG co alias nao noi 2 ten nay
  // lai voi nhau, nen toan bo 28.350.000d hoa don loai "JP SB NOI BAI" tu truoc
  // gio KHONG BAO GIO khop duoc voi gross/QR (da resolve thanh ma "JPSBNB"),
  // hien "Chua co HD" o CA 29/29 ngay thang 7 cho gian nay du tien da ve du.
  // Day la loi doc lap voi "bankExcessDefaultCode" (phan du ngan hang don vao
  // SB CAM RANH PHN, Chi Nhan tu xac nhan 2026-07-29) -- SB CAM RANH PHN thuc
  // ra da khop hoa don rat tot (85.430.000/88.760.000d, chi 6.400.000d la phan
  // du that su chua co du lieu QR), gian THUC SU chua map dung la Noi Bai.
  "JP SB NOI BAI": "JPSBNB",
  // Chi Nhan, 2026-07-30: "hóa đơn kubo bắc gian á hóa đơn để chữ như này nè
  // mốt map cho tôi nhá" -- kenh bidv77021, hoa don so 11377 (ngay 07-30, tien
  // ve ngay 29, 70.000d) ghi maDiem "FUNZONE BẮC GIANG GHẾ" thay vi dung ten
  // "Kubo Bắc Giang"/"KUBO BAC GIANG PHN" da co san alias -- day la 1 bien the
  // ten CHUA TUNG THAY (khac ca "Kubo Bắc Giang" lan "KUBO BẮC GIANG" da co
  // alias tu truoc), khien dong 70.000d nay hien "Chưa có HĐ" du tien QR that
  // ra da ve du (xem viet_qr_ten_diem_master.bidv77021 "posh funzone bac
  // giang" -> "KUBO BAC GIANG PHN", vay ben Gross/QR da tu quy dung roi, chi
  // thieu ben hoa don).
  // Nhan, 2026-08-06: sua sai lan truoc (2026-07-30) -- "FUNZONE BẮC GIANG
  // GHẾ" khong phai Kubo, Nhan xac nhan diem nay la LOTTE BAC GIANG PHN. Them
  // luon 2 ten hoa don khac cua CUNG diem Lotte Bac Giang ("POSH LOTTE BAC
  // GIANG", "PINBALL VÀ GHẾ LOTTE BAC GIANG") -- chua ten nao co alias truoc
  // do nen toan bo hoa don thang 8 cua Lotte Bac Giang deu hien "Chua co HD"
  // (Lech dung bang ca Gross du co hoa don that).
  "FUNZONE BẮC GIANG GHẾ": "LOTTE BAC GIANG PHN",
  "POSH LOTTE BAC GIANG": "LOTTE BAC GIANG PHN",
  "PINBALL VÀ GHẾ LOTTE BAC GIANG": "LOTTE BAC GIANG PHN",
  // Nhan, 2026-08-06: kenh bidv7702 hien "Lệch -2.360.000đ" cho gian "VUNG TAU
  // PHCM" (ma gross/QR) du hoa don da co day du -- hoa don ghi maDiem "POSH
  // LOTTE MART VUNG TAU" (17 hoa don) hoac "KVC LOTTE VUNG TAU" (1 hoa don,
  // cung so HD 10031 ngay 20/7 nhu dong POSH, co le do 2 nhan hieu Posh/KVC
  // cung 1 diem Lotte Mart Vung Tau xuat hoa don chung), khong ten nao khop
  // thang voi "VUNG TAU PHCM" nen tu truoc gio khong lien ket duoc.
  "POSH LOTTE MART VUNG TAU": "VUNG TAU PHCM",
  "KVC LOTTE VUNG TAU": "VUNG TAU PHCM",
  // Luyen, 2026-08-08: hoa don bidv7702 ghi maDiem "POSH MN CGV COOP BÌNH
  // DƯƠNG SQUARE" (co chu "COOP" va "SQUARE") nhung ma gross/QR resolve thanh
  // "POSH MN CGV BÌNH DƯƠNG SQUARE" (khong co "COOP") -- alias cu chi co key
  // "POSH MN CGV COOP BÌNH DƯƠNG" (thieu "SQUARE") nen hoa don co du "SQUARE"
  // van khong khop. Them alias day du de fix.
  "POSH MN CGV COOP BÌNH DƯƠNG SQUARE": "POSH MN CGV BÌNH DƯƠNG SQUARE",
  // Luyen, 2026-08-08: bidv8613600999 gian "TĐBS PHN" -- hoa don ghi maDiem
  // "TDBS PF" (khong co dau Đ, vi file MTT dung ky tu ASCII). Alias cu "TĐBS
  // PF" (co dau Đ) da co nhung KHONG bao gio match vi maDiem trong file la
  // "TDBS PF" (khong dau). Them alias ASCII de fix.
  "TDBS PF": "TĐBS PHN",
};

// Chi Nhan, 2026-07-30: "Số hóa đơn á có 20k xem nó đưa vô gian nào á bạn đưa
// vô gian đó cho tôi đi để khớp á" -- ngay 2026-07-10, kenh bidv7702, hoa don
// 9124 (140.000d) cho "GALAXY KINH DUONG VUONG PHCM" nhung du lieu QR tai len
// chi bat duoc 60.000d giao dich cho gian nay (thieu dung 20.000d, khop chinh
// xac voi "Lệch Ngân hàng-Dữ liệu: -20.000đ" ca ngay hom do). Bu 20.000d qua
// grossAdjustment. Seed qua code (giong TEN_DIEM_MASTER_DEFAULTS/
// INVOICE_DIEM_ALIAS_DEFAULTS) de KHONG BI MAT khi server cua Chi Nhan tu ghi
// de store.json bang ban cu dang giu trong bo nho (da xay ra 1 lan voi ban
// ghi thu cong truc tiep, phai chuyen sang seed code moi giu duoc).
const MANUAL_MATCH_DEFAULTS = {
  bidv7702: {
    "2026-07-10|GALAXY KINH DUONG VUONG PHCM": {
      invoiceNumbers: ["9124"],
      amount: 140000,
      grossAdjustment: 20000,
      note:
        "Bo sung 20.000d ngan hang co ve nhung du lieu QR tai len thieu dong giao dich cho gian nay (hoa don 9124 = 140.000d, du lieu QR chi co 120.000d) - Chi Nhan xac nhan 30/07.",
    },
    // Luyen, 2026-08-03: doi chieu voi file "Book5.xlsx" (phieu thu MISA thang
    // 6 chi tiet theo tung hoa don, Luyen tu lam va coi la chuan) -- 2 ngay
    // 01/06 va 05/06 co hoa don (4437 va 4966) dan tag DUY NHAT "AM TP PHCM"
    // (khong nhac gi den "LM NHA TRANG KVC") nhung mac dinh INVOICE_SHARE_PAIRS
    // (["AM TP PHCM","LM NHA TRANG KVC"]) da tu dong CHIA hoa don + doanh thu
    // 2 ngay nay giua 2 gian (vd hoa don 4437 = 5.540.000d bi chia thanh
    // 5.140.000d cho AM TP PHCM + 400.000d cho LM NHA TRANG KVC) -- ĐÚNG NHU
    // LOI LUYEN DA YEU CAU TRUOC DO (vu SB CAN THO/CON DAO AIRPORT): "nếu bị
    // nhầm 2 gian mà trên hóa đơn chỉ có 1 gian thì đưa hết vô gian đó, không
    // chia đôi". File Book5.xlsx xac nhan Luyen KHONG co dong nao cho "FARM
    // LOTTE NHA TRANG" (= LM NHA TRANG KVC) trong 2 ngay nay -- tuc la toan bo
    // doanh thu ngan hang lan hoa don ngay do deu thuoc AM TP PHCM. Dua ca
    // gross cua LM NHA TRANG KVC ve AM TP PHCM (grossAdjustment) va gan het
    // hoa don ve AM TP PHCM, LM NHA TRANG KVC con lai 0.
    // - 01/06: sau khi gop, AM TP PHCM = 5.590.000d ngan hang vs 5.540.000d hoa
    //   don (dung hoa don 4437) -> con Lệch -50.000d, GIU NGUYEN "Lệch" (khong
    //   ep khop) vi file Book5.xlsx cua Luyen cung ghi rang "Lệch" cho dong
    //   nay -- day la 1 sai lech that (hoa don thieu 50k so voi tien ngan
    //   hang), khong phai loi he thong.
    // - 05/06: sau khi gop, AM TP PHCM = 4.500.000d ngan hang vs 4.480.000d hoa
    //   don (dung hoa don 4966), con 20.000d Luyen noi "chưa tìm được hóa đơn"
    //   -- theo yeu cau "để lệch 20k đó cho tân phú đi", CHAP NHAN/HAP THU
    //   20.000d nay vao AM TP PHCM (amount = 4.500.000d, tuc = gross, coi nhu
    //   Khop) thay vi tiep tuc bao "Lệch".
    "2026-06-01|AM TP PHCM": {
      invoiceNumbers: ["4437"],
      amount: 5540000,
      grossAdjustment: 400000,
      note:
        'Gop 400.000d doanh thu ngay 01/06 tu "LM NHA TRANG KVC" ve day (hoa don 4437 chi ghi 1 gian "AM TP PHCM", khong chia doi) -- theo file Book5.xlsx Luyen xac nhan 03/08. Con Lệch -50.000d la lech that (hoa don 4437 = 5.540.000d, ngan hang 5.590.000d), khong ep khop.',
    },
    "2026-06-01|LM NHA TRANG KVC": {
      invoiceNumbers: [],
      amount: 0,
      grossAdjustment: -400000,
      note:
        'Chuyen het 400.000d doanh thu ngay 01/06 ve "AM TP PHCM" (hoa don 4437 chi ghi 1 gian, khong chia doi) -- theo file Book5.xlsx Luyen xac nhan 03/08 (khong co dong "FARM LOTTE NHA TRANG" ngay nay).',
    },
    "2026-06-05|AM TP PHCM": {
      invoiceNumbers: ["4966"],
      amount: 4500000,
      grossAdjustment: 100000,
      note:
        'Gop 100.000d doanh thu ngay 05/06 tu "LM NHA TRANG KVC" ve day (hoa don 4966 chi ghi 1 gian "AM TP PHCM") -- theo file Book5.xlsx. Con 20.000d Luyen "chưa tìm được hóa đơn" -- theo yeu cau "để lệch 20k đó cho tân phú", hap thu vao day (amount = gross, coi nhu Khop).',
    },
    "2026-06-05|LM NHA TRANG KVC": {
      invoiceNumbers: [],
      amount: 0,
      grossAdjustment: -100000,
      note:
        'Chuyen het 100.000d doanh thu ngay 05/06 ve "AM TP PHCM" (hoa don 4966 chi ghi 1 gian, khong chia doi) -- theo file Book5.xlsx (khong co dong "FARM LOTTE NHA TRANG" ngay nay).',
    },
    // Luyen, 2026-08-11: hoa don gop T7+CN ngay 1+2/08/2026 -- chia ti le tao
    // so le (vi tong hoa don != tong gross), ep khop theo dung gross tung ngay.
    "2026-08-01|AM TP PHCM": { invoiceNumbers: ["11852"], amount: 6930000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-02|AM TP PHCM": { invoiceNumbers: ["11852"], amount: 8480000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-01|LOTTE Q7 (NSG) PHM": { invoiceNumbers: ["11868"], amount: 1650000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-02|LOTTE Q7 (NSG) PHM": { invoiceNumbers: ["11868"], amount: 1520000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-01|AE BD PHCM": { invoiceNumbers: ["11855"], amount: 1410000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-02|AE BD PHCM": { invoiceNumbers: ["11855"], amount: 2600000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
  },
  // Luyen, 2026-08-11: hoa don gop T7+CN ngay 1+2/08/2026 kenh bidv77021 --
  // AE HP PHN va VC SMART PHN co ti le chia khong tron, ep khop theo gross.
  bidv77021: {
    "2026-08-01|AE HP PHN": { invoiceNumbers: ["11748"], amount: 8200000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-02|AE HP PHN": { invoiceNumbers: ["11748"], amount: 11540000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-01|VC SMART PHN": { invoiceNumbers: ["11766"], amount: 1980000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-02|VC SMART PHN": { invoiceNumbers: ["11766"], amount: 1600000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 1+2/08/2026 -- ep khop theo gross tung ngay." },
    // Luyen, 2026-08-11: hoa don gop T7+CN ngay 8+9/08/2026 kenh bidv77021 --
    // 5 gian co ti le chia so le (tong hoa don != tong gross 2 ngay), ep khop
    // theo dung gross tung ngay. Du lieu tu screenshot Luyen cung cap.
    "2026-08-08|VC TIMES PHN": { invoiceNumbers: ["12651"], amount: 4060000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 8+9/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-08|AE HP PHN": { invoiceNumbers: ["12638"], amount: 3240000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 8+9/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-08|VINKE-TCUNG": { invoiceNumbers: ["12698"], amount: 1260000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 8+9/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-08|THE GARDEN PHN": { invoiceNumbers: ["12674"], amount: 680000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 8+9/08/2026 -- ep khop theo gross tung ngay." },
    "2026-08-08|VC PHAM HUNG (SKYLAKE)": { invoiceNumbers: ["12645"], amount: 585000, grossAdjustment: 0, note: "Can tru thu cong: hoa don gop T7+CN ngay 8+9/08/2026 -- ep khop theo gross tung ngay." },
    // Luyen, 2026-08-12: ngay 9/08/2026 (CN) -- cung hoa don T7+CN voi ngay 8,
    // so HĐ giong het ngay 8, so tien = doanh thu QR tron (khong chia ty le).
    // Luyen, 2026-08-12: xoa MANUAL ngay 9/8 cho 3 gian nay -- Luyen tu can tru thu cong.
    // Luyen, 2026-08-12: xoa MANUAL ngay 10/08 VC PHAM HUNG (SKYLAKE) --
    // entry cu "Ep khop 310.000d" da het tac dung (van hien Lech -20.000d),
    // Luyen can tru thu cong.
  },
  // Luyen, 2026-08-03: "cấn trừ đưa vô 200k cho khớp cho tôi đi số 200k phú
  // quốc á" -- kenh BIDV7704, ngay 20/06 "Lệch Ngân hàng-Dữ liệu: -200.000đ"
  // (ngan hang 1.150.000d nhung du lieu QR tai len chi khop duoc 950.000d --
  // 200.000d tien that KHONG co dong QR nao khop ma tham chieu). Hoa don
  // 1445+1446 (gop chung cho "SÂN BAY PHÚ QUỐC" ca 2 ngay 20+21/06) tong dung
  // 1.100.000d = dung tong gross 2 ngay SAU KHI cong 200.000d nay vao ngay 20
  // (550.000+200.000=750.000d ngay 20, 350.000d ngay 21, cong lai 1.100.000d)
  // -- xac nhan 200.000d nay THUC SU thuoc ve "SÂN BAY PHÚ QUỐC" (dung y
  // Luyen). Vi day la hoa don gop 2 NGAY (khong phai gop 2 GIAN), khong dung
  // duoc co che applyMultiDayGroupConsolidation tu dong (ham do bo qua dong
  // da co "manualOverride"), nen ghi thang ca 2 ngay o day: gross MOI + tong
  // tien HD chia lai theo dung gross MOI cua tung ngay (tong ca 2 ngay van
  // giu dung 1.100.000d = tong 2 hoa don goc, khong doi tong that).
  bidv7704: {
    "2026-06-20|SÂN BAY PHÚ QUỐC": {
      invoiceNumbers: ["1445", "1446"],
      amount: 750000,
      grossAdjustment: 200000,
      note:
        'Cong 200.000d "Lệch Ngân hàng-Dữ liệu" ngay 20/06 (tien VietQR ve nhung khong co dong du lieu QR tai len khop ma tham chieu) vao "SÂN BAY PHÚ QUỐC" theo yeu cau Luyen 03/08. Hoa don 1445+1446 gop chung 2 ngay 20-21/06 (tong 1.100.000d) chia lai theo dung gross moi: 750.000d ngay nay + 350.000d ngay 21/06.',
    },
    "2026-06-21|SÂN BAY PHÚ QUỐC": {
      invoiceNumbers: ["1445", "1446"],
      amount: 350000,
      note:
        'Hoa don 1445+1446 gop chung 2 ngay 20-21/06 (tong 1.100.000d) -- sau khi cong 200.000d "Lệch Ngân hàng-Dữ liệu" ngay 20/06 vao gross ngay do, chia lai tong tien HD theo dung gross moi: 750.000d ngay 20/06 + 350.000d ngay nay = dung 1.100.000d.',
    },
  },
};

// Chi Nhan, 2026-07-29: "2 cái mã công trình này là 1 á gộp lại vô cái CHKQT
// CAM RANH cho tôi nhá" -- kenh mb11521268, san bay Cam Ranh dang bi tach
// thanh 2 "gian" rieng o BUOC GROSS/QR (vi 2 nhom Ma cua hang cua Luyen tu
// dat 2 "Tên điểm bán" hoi khac nhau: "1 JP SB Cam Ranh.new" va "POSH Sân bay
// Quốc Tế Cam Ranh") du hoa don THAT SU LUON dung CHUNG 1 ma diem "CHKQT CAM
// RANH" cho ca 2 nhom nay (xac nhan: 25/25 hoa don Cam Ranh cua kenh nay deu
// ghi maDiem=CHKQT CAM RANH, khong co hoa don nao dung "SB CAM RANH PHN").
// Sua tan goc bang override o viet_qr_ten_diem_master (dung TRUOC ca fuzzy
// matcher trong resolveGianGross, xem utils/vietqrReconcile.js) thay vi tao
// "Doi tru thu cong" tung ngay (~40 lan tung lam truoc do, xem doan xoa
// manual_matches ben duoi) -- Seed qua code (chay lai moi request, giong
// GIAN_MERGE_DEFAULTS/INVOICE_DIEM_ALIAS_DEFAULTS o tren) de KHONG BI MAT khi
// server cua chi Nhan tu ghi de store.json bang ban cu dang giu trong bo nho.
const TEN_DIEM_MASTER_DEFAULTS = {
  mb11521268: {
    "1 jp sb cam ranh.new": "CHKQT CAM RANH",
  },
  // Nhan, 2026-08-06: "cửa hàng POSH Funzone Bắc Giang này á là của LOTTE BAC
  // GIANG PHN tôi đưa nhầm vào KUBO BAC GIANG PHN rồi" -- sua lai override
  // truoc do (2026-07-30) tung tro "posh funzone bac giang" ve KUBO BAC GIANG
  // PHN, gio Nhan xac nhan diem nay thuc ra thuoc LOTTE BAC GIANG PHN.
  bidv77021: {
    "posh funzone bac giang": "LOTTE BAC GIANG PHN",
  },
  // Chi Nhan, 2026-07-29: "đây là 3 mã công trình chuẩn của 77020" -- chi Nhan
  // xac nhan kenh bidv77020 chi co 3 gian that (san bay Vinh/Noi Bai/Cam Ranh),
  // hoa don da dung dung 3 ma "SB VINH PHN"/"JPSBNB"/"SB CAM RANH PHN" (xem
  // maDiem trong store.viet_qr_invoices.bidv77020) nhung ben gross/QR chua co
  // override nen tu tao "gian" rieng theo dung ten Tên điểm bán tho (vd "POSH
  // Sân bay Cam ranh", "1 JP SB Cam Ranh.new", "JP Sân Bay Nội Bài" deu la
  // CUNG 1 diem that, chi khac ten do nhieu dot POS/QR khac nhau) -- gop ve
  // dung 3 ma chuan de khop voi hoa don co san.
  bidv77020: {
    "posh san bay cam ranh": "SB CAM RANH PHN",
    "posh san bay quoc te cam ranh": "SB CAM RANH PHN",
    "1 jp sb cam ranh.new": "SB CAM RANH PHN",
    "posh san bay vinh": "SB VINH PHN",
    "jp san bay noi bai": "JPSBNB",
    "1.jp san bay noi bai new": "JPSBNB",
  },
  // Chi Nhan, 2026-07-29: bang map "tên trên vn pay" -> "tên công trình" chi
  // Nhan cung cap cho kenh VNPay KH Moi (02865168) -- "Điểm thu" tren file
  // portal VNPay ("DanhSachGiaoDich...xlsx") la ten tho, chua tung khop duoc
  // hoa don nao (invoice tenDiem rong, xem VNPAY_KHMOI_INVOICE_MADIEM_MAP
  // trong utils/vietqrReconcile.js) nen PHAI dung override nay (khop truoc ca
  // fuzzy matcher) thay vi de tu hoc qua gianCandidates nhu cac kenh khac.
  // Chi Nhan, 2026-07-29: entry "vnpayKhMoi" tung o day da chuyen sang
  // TEN_DIEM_TO_MA_CONG_TRINH trong routes/doisoat-vnpay-khmoi.js (kenh nay
  // khong con dung kien truc VietQR nua -- xem ghi chu tai CHANNELS o tren).
};

function ensureChannelShape(store) {
  if (!store.viet_qr_raw_uploads) store.viet_qr_raw_uploads = {};
  if (!store.viet_qr_store_names) store.viet_qr_store_names = {};
  if (!store.viet_qr_invoices) store.viet_qr_invoices = {};
  if (!store.viet_qr_manual_matches) store.viet_qr_manual_matches = {};
  // Nhan, 2026-08-06: "đồng ý cấn trừ hiển thị lần thôi, nếu tôi đồng ý rồi
  // thì bỏ qua đi không cần hiển thị lại lần sau" -- 1 cap gian (vd "CON DAO
  // AIRPORT PHCM" du / "SB CAN THO PHCM" thieu) lap lai MOI NGAY (tai file moi
  // moi ngay lai sinh ra 1 goi y moi cho ngay do), buoc Nhan phai bam "Dong y"
  // lai tu dau moi lan. Nho lai CAP gian (khong phan biet ngay) da tung duoc
  // Nhan dong y it nhat 1 lan -- tu do tu dong ap dung cho MOI ngay sau nay co
  // cung cap nay, khong hien lai thanh goi y cho Nhan bam nua. Xem
  // applyApprovedCrossMatches ben duoi.
  if (!store.viet_qr_cross_match_approved_pairs) store.viet_qr_cross_match_approved_pairs = {};
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
  // Luyen, 2026-08-05: "còn 20k còn dư cái bạn cho hẳn vào Tân phú luôn là
  // không được nhá bạn nhớ báo tôi nhá ... không cần hiển thị lại cái nào
  // mới thì thông báo cho tôi chỗ đó thôi" -- giu nguyen co che tu dong gop
  // phan du Ngan hang-Du lieu vao gian mac dinh (AM TP PHCM/AE HP PHN, xem
  // TAN_PHU_AUTO_APPLY_CHANNELS o duoi) vi Luyen xac nhan van muon giu (hoi
  // qua AskUserQuestion 2026-08-05), nhung them 1 co che "da xem" theo tung
  // (channel, ngay, so tien) de KHONG hien lai canh bao mot khi Luyen da xac
  // nhan biet roi -- neu sau nay so tien du doi khac (vd tai them file moi
  // lam giam/tang leftover) se lai la 1 key MOI, tu dong hien lai canh bao
  // (dung y "cái nào mới thì thông báo"). Xem tanPhuAutoApplied trong
  // buildChannelReconciliation va route /doi-soat/vietqr/tanphu-ack/:channel.
  if (!store.viet_qr_tanphu_ack) store.viet_qr_tanphu_ack = {};
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
  // Chi Nhan, 2026-07-29: don dep 1 dong alias RAC "CHKQT CAM RANH" -> "Chưa
  // khớp" (rat co the do 1 dong trong file "gian master" cua chi Nhan co gia
  // tri "Chưa khớp" o cot Ma cong trinh, bi merge thang vao bang nay -- xem
  // ghi chu invoice_diem_alias trong store.js). Vi day la alias ngay TREN
  // CHINH ma dich "CHKQT CAM RANH", no khien MOI hoa don that (25/25 hoa don
  // Cam Ranh cua mb11521268) bi doi thanh "Chưa khớp" khi doi soat -- xoa moi
  // lan phat hien (khong chi 1 lan) de tru khi file goc duoc sua.
  if (store.invoice_diem_alias["CHKQT CAM RANH"] === "Chưa khớp") {
    delete store.invoice_diem_alias["CHKQT CAM RANH"];
  }
  // Luyen, 2026-08-10: hop dong Nha Trang het -- xoa gia tri cu "FARM LOTTE NHA TRANG"
  // khoi store de vong lap seed ben duoi ap dung duoc gia tri moi tu code default
  // ("POSH MN CGV Vincom Xuan Khanh"). Chi xoa gia tri CU cu the, khong anh huong
  // cac alias khac nguoi dung co the da chinh tay.
  if (store.ma_cong_trinh_display_alias["LM NHA TRANG KVC"] === "FARM LOTTE NHA TRANG") {
    delete store.ma_cong_trinh_display_alias["LM NHA TRANG KVC"];
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
    if (!store.viet_qr_cross_match_approved_pairs[ch]) store.viet_qr_cross_match_approved_pairs[ch] = {};
    if (!store.viet_qr_nocode_assignments[ch]) store.viet_qr_nocode_assignments[ch] = {};
    if (!store.viet_qr_store_uploads[ch]) store.viet_qr_store_uploads[ch] = [];
    if (!store.viet_qr_ten_diem_master[ch]) store.viet_qr_ten_diem_master[ch] = {};
    if (!store.viet_qr_tanphu_ack[ch]) store.viet_qr_tanphu_ack[ch] = {};
    // Ghi de KHONG DIEU KIEN (khac GIAN_MERGE_DEFAULTS o tren) -- day la 1 dong
    // da xac nhan SAI can sua han (dang tro toi "SB CAM RANH PHN" cu), khong
    // phai gia tri "dien vao cho trong" -- neu chi fill-if-empty se khong bao
    // gio sua duoc gia tri SAI da co san trong store.json.
    Object.keys(TEN_DIEM_MASTER_DEFAULTS[ch] || {}).forEach((k) => {
      store.viet_qr_ten_diem_master[ch][k] = TEN_DIEM_MASTER_DEFAULTS[ch][k];
    });
    // Chi Nhan, 2026-07-29: xoa cac "Doi tru thu cong" GIA TAO (~40 dong, tao
    // hang loat luc 28-29/7) tung dung de tach doi hoa don "CHKQT CAM RANH"
    // sang "SB CAM RANH PHN" moi ngay -- gio da sua tan goc bang
    // TEN_DIEM_MASTER_DEFAULTS o tren (gop thang 2 gian lam 1 tu buoc doc du
    // lieu QR), nen cac dong "doi tru" nay khong con can nua; neu con ton tai
    // (chay truoc khi co fix, hoac store.json bi ghi de lai) se tu xoa moi
    // lan load() de gian tu dong tro ve 1 dong CHKQT CAM RANH duy nhat.
    // Luyen, 2026-08-14: chi xoa manual match CU cho "SB CAM RANH PHN"
    // (ten sai cu), KHONG xoa "CHKQT CAM RANH" -- truoc day xoa ca 2 nen
    // moi lan Luyen luu "+/- Sua D'" cho CHKQT CAM RANH la bi xoa ngay
    // (khong co hieu luc gi). Gio chi don dep ten cu, giu lai ten dung.
    if (ch === "mb11521268" && store.viet_qr_manual_matches[ch]) {
      Object.keys(store.viet_qr_manual_matches[ch]).forEach((k) => {
        if (k.endsWith("|SB CAM RANH PHN")) {
          delete store.viet_qr_manual_matches[ch][k];
        }
      });
    }
    if (!store.viet_qr_ref_override[ch]) store.viet_qr_ref_override[ch] = {};
    if (!store.viet_qr_store_names_baseline[ch]) {
      // Chup 1 lan duy nhat: du lieu diem ban HIEN CO ngay truoc khi tinh
      // nang lich su/xoa nay ton tai, de khong mat du lieu cu.
      store.viet_qr_store_names_baseline[ch] = Object.assign({}, store.viet_qr_store_names[ch]);
    }
  });

  // Chi Nhan, 2026-07-29: "tôi có 2 tài khoản viet qr chia nam bắc nhưng điểm
  // bán lại chung á nên bạn lấy các map mã công trình điểm bán bên 77021
  // giống với 11521268 đc á bạn map lại cho tôi đi" -- kenh mb11521268 va
  // bidv77021 nhan tien tu CUNG 1 tap Ma cua hang vat ly (vd cac quay o San
  // bay Phu Quoc), nhung 77021 la kenh chinh nen da co san day du "Tên điểm
  // bán" cho chung; mb11521268 chi thinh thoang nhan duoc giao dich tu CUNG
  // cac quay do nen hang chuc Ma cua hang bi "chua map" moi lan co du lieu QR
  // moi (xac nhan: toan bo cac ma dang "chua map" cua mb11521268 ngay 29/7
  // deu da co san trong viet_qr_store_names.bidv77021, cung tro ve "JP SÂN
  // BAY PHÚ QUỐC QT"). Tu dong sao chep NHUNG ma nao mb11521268 CHUA CO
  // (khong ghi de ma da duoc gan rieng ben mb11521268), chay lai moi lan
  // load() de tu dong theo kip ma cua hang moi phat sinh sau nay o ca 2 kenh.
  if (store.viet_qr_store_names.bidv77021 && store.viet_qr_store_names.mb11521268) {
    const crossSource = store.viet_qr_store_names.bidv77021;
    const crossTarget = store.viet_qr_store_names.mb11521268;
    Object.keys(crossSource).forEach((maCuaHang) => {
      if (!crossTarget[maCuaHang]) {
        crossTarget[maCuaHang] = crossSource[maCuaHang];
      }
    });
  }

  // Chi Nhan, 2026-07-30: "thêm 1 ngân hàng 02865168 á các điểm bán với mã
  // công trình map với 77021 á" -- dong bo danh sach diem ban + ten diem master
  // tu bidv77021 sang mb02865168 (kenh moi, cung he thong QR/POS voi 77021, chi
  // khac tai khoan nhan tien) moi lan load(), va don cac hoa don "tk 168" tung
  // bi gom nham vao pool bidv77021 truoc khi sua tagPattern (xem CHANNELS
  // ben tren + seedMb02865168FromBidv77021/migrateTk168Invoices trong
  // utils/vietqrReconcile.js).
  seedMb02865168FromBidv77021(store);
  migrateTk168Invoices(store);

  // Chi Nhan, 2026-07-30: "7702 tôi để nhầm tên á này của ngày 3 á" -- sua 29
  // hoa don "MTD MN 4" (thuc ra la ngay 3, xem fixBidv7702Day3TaggedAsDay4
  // trong utils/vietqrReconcile.js) -- chay lai moi lan load() de khong bi mat
  // khi server cua Chi Nhan tu ghi de store.json bang ban cu dang giu trong bo
  // nho (da xay ra 1 lan voi ban sua truc tiep khong qua code).
  fixBidv7702Day3TaggedAsDay4(store);

  // Chi Nhan, 2026-07-30: "chỗ này bị nhầm của ghi lộn ngày 3 thành ngày 4"
  // -- BIDV77021 co cung loi (85 hoa don "VietQR POSH MB 4" thuc ra la ngay
  // 3, xem fixBidv77021Day3TaggedAsDay4 trong utils/vietqrReconcile.js) --
  // chay lai moi lan load() de khong bi mat khi server cua Chi Nhan tu ghi de
  // store.json bang ban cu dang giu trong bo nho.
  fixBidv77021Day3TaggedAsDay4(store);

  // Chi Nhan, 2026-07-30: BIDV8613600999 07-03/07-04 "2 ngày này cấn trừ
  // nhau á cấn trừ cho tôi đi để khớp" -- cung dang loi tag hoa don giong
  // fixBidv7702Day3TaggedAsDay4 o tren (KHONG PHAI cap tru chuoi ngay): 6 hoa
  // don so 8261-8266 dan tag rieng "QR JP 4" nhung thuc ra la doanh thu ngay
  // 3 (xem fixBidv8613600999Day3TaggedAsDay4 trong utils/vietqrReconcile.js).
  // Chay lai moi lan load() de khong bi mat khi server cua Chi Nhan tu ghi de
  // store.json bang ban cu dang giu trong bo nho.
  fixBidv8613600999Day3TaggedAsDay4(store);

  // Ghi de KHONG DIEU KIEN (giong TEN_DIEM_MASTER_DEFAULTS o tren) -- cac ban
  // ghi "Doi tru thu cong" mac dinh nay tung bi mat 1 lan vi chi luu truc tiep
  // vao store.json (khong qua code) truoc khi server cua Chi Nhan ghi de lai
  // ban cu dang giu trong bo nho -- seed lai moi lan load() de khong bao gio
  // mat nua.
  Object.keys(MANUAL_MATCH_DEFAULTS).forEach((ch) => {
    if (!store.viet_qr_manual_matches[ch]) store.viet_qr_manual_matches[ch] = {};
    Object.keys(MANUAL_MATCH_DEFAULTS[ch]).forEach((key) => {
      store.viet_qr_manual_matches[ch][key] = Object.assign(
        { created_at: new Date().toISOString() },
        store.viet_qr_manual_matches[ch][key],
        MANUAL_MATCH_DEFAULTS[ch][key]
      );
    });
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

// Nhan, 2026-08-06: "đồng ý cấn trừ hiển thị lần thôi, đồng ý rồi thì bỏ qua
// đi, có cái khác thì hiển thị 1 cái đó thôi" -- 1 cap gian (fromCode/toCode)
// hay lap lai giong nhau MOI NGAY (vd "CON DAO AIRPORT PHCM"/"SB CAN THO
// PHCM" tu 06-01 den 06-05...), buoc Nhan bam "Dong y" lai tu dau moi lan co
// ngay moi. Voi CAP da tung duoc Nhan dong y >=1 lan (luu trong
// store.viet_qr_cross_match_approved_pairs[channel]), tu dong ghi manual-match
// NGAY LUC BUILD (giong toi cac ham fixXxx/migrateXxx khac trong file nay),
// khong doi Nhan bam nua -- chi con hien nhu "goi y" cho CAP nao Nhan CHUA
// tung dong y truoc do.
function applyApprovedCrossMatches(store, channelKey, reconciled, suggestions) {
  const approved = (store.viet_qr_cross_match_approved_pairs || {})[channelKey] || {};
  const remaining = [];
  let changed = false;
  for (const s of suggestions) {
    const pairKey = `${s.fromCode}|${s.toCode}`;
    if (!approved[pairKey]) {
      remaining.push(s);
      continue;
    }
    const settlement = reconciled.find((r) => r.settlementDate === s.settlementDate);
    const lineFrom = settlement && settlement.lines.find((l) => l.code === s.fromCode);
    const lineTo = settlement && settlement.lines.find((l) => l.code === s.toCode);
    if (!lineFrom || !lineTo || lineFrom.matched || lineFrom.manualOverride || Math.abs(lineFrom.diff + lineTo.diff) >= 1) {
      continue; // du lieu da doi/khong con hop le -- bo qua, khong hien lai (Nhan da tung dong y cap nay roi)
    }
    const note = `Doi tru tu dong (da duoc Nhan dong y truoc do cho cap "${displayCode(
      s.fromCode
    )}" <-> "${displayCode(s.toCode)}"): hoa don cua "${displayCode(s.fromCode)}" du ${lineFrom.diff.toLocaleString(
      "vi-VN"
    )}đ, chuyen sang "${displayCode(s.toCode)}" dang thieu dung so do (${new Date().toLocaleDateString("vi-VN")}).`;
    if (!store.viet_qr_manual_matches[channelKey]) store.viet_qr_manual_matches[channelKey] = {};
    store.viet_qr_manual_matches[channelKey][`${s.settlementDate}|${s.fromCode}`] = {
      invoiceNumbers: lineFrom.invoiceNumbers,
      amount: lineFrom.gross,
      grossAdjustment: 0,
      note,
      created_at: new Date().toISOString(),
    };
    store.viet_qr_manual_matches[channelKey][`${s.settlementDate}|${s.toCode}`] = {
      invoiceNumbers: Array.from(new Set([...lineTo.invoiceNumbers, ...lineFrom.invoiceNumbers])),
      amount: lineTo.gross,
      grossAdjustment: 0,
      note,
      created_at: new Date().toISOString(),
    };
    changed = true;
  }
  return { remaining, changed };
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
  if (migrateVnpayKhMoiInvoices(store)) save(store);
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
    // Nhan, 2026-08-06: kenh mb02865168/bidv8613600999 dang bi giao dich
    // KHONG khop "So tham chieu" tu dong gap vao "AE HP PHN" (cfg.defaultBlankCode
    // truyen thang vao day KHONG DIEU KIEN), du Nhan da xac nhan 2026-08-03
    // (xem ghi chu TAN_PHU_AUTO_APPLY_CHANNELS o duoi) la CHI bidv7702/bidv77021
    // duoc tu dong gap vao gian mac dinh, 2 kenh nay phai de nguyen hien "Lệch
    // Ngân hàng-Dữ liệu" cho Nhan tu can doi tay (vd giao dich MBBank IBFT
    // "GLN TRANSFERO" khong mang ma VQR, dung vao AE HP PHN gia du AE HP PHN
    // khong phai gian thuc cua kenh nay). Chi truyen defaultBlankCode cho 2
    // kenh nam trong TAN_PHU_AUTO_APPLY_CHANNELS, cac kenh khac truyen null de
    // roi dung vao unmatchedBankTx (hien Lệch Ngân hàng-Dữ liệu).
    const refResolved = resolveGianGrossByBankRef(
      bankThuTxs,
      rawRows,
      storeNames,
      tenDiemMaster,
      storeCodeOverride,
      refOverride,
      TAN_PHU_AUTO_APPLY_CHANNELS.has(channelKey) ? cfg.defaultBlankCode : null
    );

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

  // Chi Nhan, 2026-07-30: "các đối soát tất cả các trang điều xếp theo ngày
  // cho tôi nhá" -- reconcileVietQr tra ve ket qua theo thu tu giao dich
  // ngan hang trong store.transactions (thu tu tai/nhap lieu), khong phai
  // thu tu ngay thang, nen trang co the hien lon xon (vd ngay 30 roi 23 roi
  // 24). Sap lai TANG DAN theo settlementDate NGAY SAU KHI TINH XONG (giong
  // cach doisoat-vnpay-khmoi.js da lam) -- an toan de sap truoc
  // applyInvoiceSharePairs/applySharePairChainNetting vi ca 2 ham do xu ly
  // TUNG DONG doc lap (khong dua vao thu tu mang de tinh), applySharePairChainNetting
  // con tu sap 1 ban sao rieng truoc khi xet chuoi ngay lien tiep.
  reconciled.sort((a, b) => (a.settlementDate < b.settlementDate ? -1 : a.settlementDate > b.settlementDate ? 1 : 0));

  // Xem ghi chu tai INVOICE_SHARE_PAIRS o tren -- chay SAU reconcileVietQr
  // (da bao gom ca chia theo ngay cho hoa don gop T7+CN) de tu dong chia lai
  // theo dung ty le doanh thu cho cac cap gian bi xuat chung 1 hoa don.
  applyInvoiceSharePairs(reconciled, INVOICE_SHARE_PAIRS[channelKey], INVOICE_SHARE_PAIRS_STOP_FROM[channelKey]);
  applySharePairChainNetting(reconciled, INVOICE_SHARE_PAIRS[channelKey], INVOICE_SHARE_PAIRS_STOP_FROM[channelKey]);

  // Chi Nhan, 2026-07-29 (xem ghi chu day du tai CHANNELS.bidv77020 o tren):
  // voi kenh co cfg.bankExcessDefaultCode, neu sau khi tinh het gross tu du
  // lieu QR tai len, "Ngan hang" (da chi gom GD VietQR that, xem
  // extractVietQrSettlements) VAN CON DU so voi tong gross da gan duoc cho
  // cac gian (r.totalNetComputed) -- tuc la co tien VietQR ve nhung KHONG co
  // dong du lieu QR nao khop -- so du do duoc cong THANG vao gian mac dinh
  // (khong chi hien "goi y" nhu tanPhuSuggestion o refMatchFrom, vi kenh nay
  // khong co co che khop tung GD ngan hang de biet CHINH XAC GD nao la du,
  // nen coi CA PHAN DU cua ca ngay la thuoc ve gian mac dinh, dung y Chi Nhan
  // "cho vô gian SB CAM RANH PHN"). Nguong 1.000d de bo qua sai so lam tron.
  if (cfg.bankExcessDefaultCode) {
    reconciled.forEach((r) => {
      if (r.pendingBank) return;
      const leftover = r.bankAmount - r.totalNetComputed;
      if (leftover <= 1000) return;
      let line = r.lines.find((l) => l.code === cfg.bankExcessDefaultCode);
      if (!line) {
        line = {
          code: cfg.bankExcessDefaultCode,
          maCongTrinh: cfg.bankExcessDefaultCode,
          tkCo: store.gian_mapping[cfg.bankExcessDefaultCode] || "131",
          gross: 0,
          net: 0,
          invoiceNumbers: [],
          invoiceTotal: 0,
          diff: 0,
          matched: false,
          manualOverride: false,
        };
        r.lines.push(line);
      }
      line.gross += leftover;
      line.net = line.gross;
      line.diff = line.invoiceTotal - line.gross;
      line.matched = line.invoiceNumbers.length > 0 && Math.abs(line.diff) < 1;
      line.bankExcessApplied = (line.bankExcessApplied || 0) + leftover;
      line.manualNote = [
        line.manualNote,
        `Da tu dong cong ${leftover.toLocaleString("vi-VN")}d phan du ngan hang (co tien VietQR ve nhung khong co du lieu QR tai len khop) vao gian nay.`,
      ]
        .filter(Boolean)
        .join(" -- ");
      r.totalNetComputed += leftover;
      r.diffVsBank = r.totalNetComputed - r.bankAmount;
      r.lines.sort((a, b) => b.gross - a.gross);
    });
  }

  // Chi Nhan, 2026-07-29: "không có lẻ với lấy ngân hàng làm chuẩn cái dư
  // bên dữ liệu thì bỏ đi không cần lấy dựa vào tham chiếu á giống các viet
  // qr khác để bt giao dịch Dữ liệu nào không có thì bỏ ra" -- chieu NGUOC
  // lai cua block tren: khi "Tinh tu du lieu tai len" (r.totalNetComputed)
  // NHIEU HON "Ngan hang" (r.bankAmount) -- tuc la co dong du lieu QR khong
  // duoc ngan hang xac nhan thuc su ve ngay do -- Ngan hang duoc lay lam
  // CHUAN, phan du BEN DU LIEU bi tru bot thay vi co gang do tim CHINH XAC
  // dong QR nao la rac bang tham chieu ngan hang nhu bidv7702/77021 (Chi
  // Nhan xac nhan KHONG can lam vay cho kenh nay).
  //
  // Uu tien 1: neu co dong nao KHAC gian mac dinh MA DA CO HOA DON rieng
  // nhung gross dang VUOT invoiceTotal (diff am -- "gross du so hoa don cua
  // CHINH gian do"), tru truoc tu chinh dong do (theo thu tu du nhieu nhat
  // truoc) -- vua khop dung "Ngan hang<->Du lieu" ngay, vua tinh co khop
  // luon "Du lieu<->Hoa don" cua dong do (vd SB VINH PHN ngay 20/07: gross du
  // dung 20.000d so hoa don, khop tuyet doi voi phan du ca ngay). Chi sau khi
  // het cac dong nay ma van con du thi moi don phan con lai vao gian mac
  // dinh (nhu truoc), tranh lam xau di 1 dong dang khop dung khi co dong
  // khac giai thich dung hon.
  if (cfg.bankExcessDefaultCode) {
    reconciled.forEach((r) => {
      if (r.pendingBank) return;
      let deficit = r.totalNetComputed - r.bankAmount;
      if (deficit <= 1000) return;
      const applyDeduction = (line, applied) => {
        line.gross -= applied;
        line.net = line.gross;
        line.diff = line.invoiceTotal - line.gross;
        line.matched = line.invoiceNumbers.length > 0 && Math.abs(line.diff) < 1;
        line.bankDeficitApplied = (line.bankDeficitApplied || 0) + applied;
        line.manualNote = [
          line.manualNote,
          `Da tu dong tru ${applied.toLocaleString("vi-VN")}d khoi gian nay (du lieu QR tai len nhieu hon so ngan hang thuc nhan trong ngay, lay ngan hang lam chuan).`,
        ]
          .filter(Boolean)
          .join(" -- ");
        r.totalNetComputed -= applied;
        deficit -= applied;
      };
      const candidates = r.lines
        .filter((l) => l.code !== cfg.bankExcessDefaultCode && l.invoiceNumbers.length > 0 && l.gross - l.invoiceTotal > 0)
        .sort((a, b) => b.gross - b.invoiceTotal - (a.gross - a.invoiceTotal));
      candidates.forEach((l) => {
        if (deficit <= 1000) return;
        const applied = Math.min(deficit, l.gross - l.invoiceTotal);
        if (applied > 0) applyDeduction(l, applied);
      });
      if (deficit > 1000) {
        const line = r.lines.find((l) => l.code === cfg.bankExcessDefaultCode);
        if (line && line.gross > 0) {
          const applied = Math.min(deficit, line.gross);
          if (applied > 0) applyDeduction(line, applied);
        }
      }
      r.diffVsBank = r.totalNetComputed - r.bankAmount;
      r.lines.sort((a, b) => b.gross - a.gross);
    });
    // Chi Nhan, 2026-07-29: "mấy cái lệch lẻ lẻ này nè là gộp lại xuất 2
    // ngày khớp mà đừng có để lệch cho tôi chớ" -- 1 hoa don gop nhieu ngay
    // (vd 1824 ngay 4-5) co the CHI can bang dung SAU KHI 2 block tren dieu
    // chinh lai gross theo ngan hang (xem ghi chu applyMultiDayGroupConsolidation
    // trong utils/vietqrReconcile.js) -- goi lai lan 2 voi gross MOI de bat
    // duoc ca truong hop nay, khong chi cac cap da can bang tu dau.
    applyMultiDayGroupConsolidation(reconciled);
  }

  // Luyen, 2026-07-27: "các giao dịch không phải của vietqr thì trừ ra nhá
  // cái nào có mã tham chiếu á" -- tu ngay cutover (refMatchFrom), 1 giao
  // dich "thu" tren sao ke KHONG khop duoc So tham chieu nao voi bat ky
  // gian nao (vd "Thanh toan lai thang 07/2026" tu ngan hang, chuyen khoan
  // IBFT ca nhan...) khong phai tien VietQR that su -- da duoc gom san
  // trong refUnmatchedBankTx (canh bao) nhung truoc gio CHUA tru khoi tong
  // "Ngan hang" cua doi soat, lam Lech gia (ngan hang > du lieu tai len chi
  // vi cong nham cac dong khong lien quan). Tru dung so tien nay khoi
  // bankAmount cua NGAY tuong ung roi tinh lai diffVsBank.
  //
  // Chi Nhan, 2026-07-30 (lan 1): "chênh lệch tk 7702 thì đưa vô tân phú còn
  // 77021 thì đưa vô hải phòng nhá" -- xac nhan tu dong CONG LUON phan du con
  // lai (tien VietQR that nhung chua gian nao nhan) vao gian mac dinh, CHI
  // cho 2 kenh nay.
  // Chi Nhan, 2026-07-30 (lan 2, sau khi thay 07-29 cua mb02865168 cung bi
  // tu dong don vao AE HP PHN): "còn lại các tài khoản khác nếu lệch giữa
  // ngân hàng và dữ liệu nếu hk có số tham chiếu á dư tiền thì để lệch đi
  // tôi cấn sao á chớ đừng lúc nào cx đưa vô Hải Phòng chèn" -- CHINH LAI:
  // co che tu-dong-cong-vao-gian-mac-dinh o day CHI ap dung cho DUNG 2 kenh
  // bidv7702/bidv77021 (TAN_PHU_AUTO_APPLY_CHANNELS ben duoi), KHONG ap dung
  // chung cho moi kenh co refMatchFrom nua (truoc day dung cfg.defaultBlankCode
  // nen vo tinh ap luon ca mb02865168/bidv8613600999 -- ca 2 co defaultBlankCode
  // cung la "AE HP PHN" nhu bidv77021 -- Chi Nhan khong muon vay, cac kenh do
  // de nguyen "Lech Ngan hang-Du lieu" hien ra, tu chi can trung tay). Phan
  // "tru giao dich khong phai VietQR" (excludedByDate) van ap dung cho MOI
  // kenh co refMatchFrom nhu cu -- day la loai bo rac, khong phai gan tien
  // that vao gian nao ca, khong nam trong yeu cau gioi han nay.
  // Luyen, 2026-08-05: danh sach cac lan tu dong gop phan du vao gian mac
  // dinh (Tan Phu/Hai Phong) NGAY LAN NAY, de lam banner canh bao rieng o
  // dau trang -- xem ghi chu "tanPhuAutoApplied" o cuoi ham nay va route
  // /doi-soat/vietqr/tanphu-ack/:channel.
  const tanPhuAutoAppliedRaw = [];
  if (cfg.refMatchFrom) {
    const excludedByDate = {};
    refUnmatchedBankTx.forEach((tx) => {
      excludedByDate[tx.date] = (excludedByDate[tx.date] || 0) + tx.amount;
    });
    const tanPhuTarget = cfg.defaultBlankCode || "AM TP PHCM";
    const autoApplyTanPhu = TAN_PHU_AUTO_APPLY_CHANNELS.has(channelKey);
    // Luyen, 2026-08-03: "7702 á đối chiếu viet qr á nếu đối chiếu theo mã
    // tham chiếu đã khớp dư này là do ngân hàng thì đưa vô gian AM TP PHCM
    // nhá ... số ngân hàng chính xác giúp tôi nhá" (kem anh man hinh cac ngay
    // 01-24/06/2026, "Lech: Ngan hang <-> Du lieu" nho le -50.000d/-60.000d/
    // -20.000d... chua duoc gom). Truoc day block nay CHI chay cho ngay >=
    // cfg.refMatchFrom (22/07) -- xem "if (r.settlementDate < cfg.refMatchFrom)
    // return;" da bo o day. Bo dieu kien do de ap dung CHO MOI ngay (ke ca
    // truoc cutover) cho 2 kenh TAN_PHU_AUTO_APPLY_CHANNELS: phan du con lai
    // giua "Ngan hang" (chuan, xem extractVietQrSettlements) va "Tinh tu du
    // lieu tai len" duoc tu dong don vao gian mac dinh (AM TP PHCM cho 7702),
    // dung nhu Luyen yeu cau -- ngan hang luon la so CHINH XAC, dong AM TP
    // PHCM la noi hung phan du do. An toan cho ngay < refMatchFrom vi
    // excludedByDate/refUnmatchedBankTx CHI co du lieu tu ngay >= refMatchFrom
    // (xem bankThuTxs.filter(t => t.date >= cfg.refMatchFrom) o tren) nen
    // excluded luon = 0 cho cac ngay cu, khong anh huong gi them; ngay nao da
    // co dieu chinh thu cong (viet_qr_manual_matches) truoc do trong
    // reconcileVietQr roi thi leftover da <= 1.000d nen khong bi cong them
    // lan nua.
    reconciled.forEach((r) => {
      const excluded = excludedByDate[r.settlementDate] || 0;
      if (excluded > 0) {
        r.bankAmount -= excluded;
        r.bankAmountExcluded = excluded;
      }
      if (!r.pendingBank) {
        r.diffVsBank = r.totalNetComputed - r.bankAmount;
        if (!autoApplyTanPhu) return;
        if (TAN_PHU_SUPPRESS[channelKey]?.has(r.settlementDate)) return;
        const leftover = r.bankAmount - r.totalNetComputed;
        if (leftover > 1000) {
          let line = r.lines.find((l) => l.code === tanPhuTarget);
          if (!line) {
            line = {
              code: tanPhuTarget,
              maCongTrinh: tanPhuTarget,
              tkCo: store.gian_mapping[tanPhuTarget] || "131",
              gross: 0,
              net: 0,
              invoiceNumbers: [],
              invoiceTotal: 0,
              diff: 0,
              matched: false,
              manualOverride: false,
            };
            r.lines.push(line);
          }
          line.gross += leftover;
          line.net = line.gross;
          line.diff = line.invoiceTotal - line.gross;
          line.matched = line.invoiceNumbers.length > 0 && Math.abs(line.diff) < 1;
          line.bankExcessApplied = (line.bankExcessApplied || 0) + leftover;
          line.manualNote = [
            line.manualNote,
            `Da tu dong cong ${leftover.toLocaleString("vi-VN")}d phan du ngan hang (tien VietQR ve nhung khong co du lieu QR tai len khop) vao gian nay.`,
          ]
            .filter(Boolean)
            .join(" -- ");
          r.totalNetComputed += leftover;
          r.diffVsBank = r.totalNetComputed - r.bankAmount;
          r.lines.sort((a, b) => b.gross - a.gross);
          r.tanPhuSuggestion = { amount: leftover, targetCode: tanPhuTarget, applied: true };
          tanPhuAutoAppliedRaw.push({ date: r.settlementDate, amount: leftover, targetCode: tanPhuTarget });
        }
      }
    });
  }
  // Luyen, 2026-08-05: "không cần hiển thị lại cái nào cũ, có gì mới thì báo
  // chỗ đó thôi" -- chi giu lai cac lan gop TAN Phu ma Luyen CHUA bam "Da
  // xem" (key = ngay|so tien, xem store.viet_qr_tanphu_ack[channel] va route
  // /doi-soat/vietqr/tanphu-ack/:channel). Neu sau nay so tien du doi khac
  // (vd nap them file lam leftover thay doi), key moi se tu dong hien lai.
  const tanPhuAck = store.viet_qr_tanphu_ack[channelKey] || {};
  const tanPhuAutoApplied = tanPhuAutoAppliedRaw
    .filter((ev) => !tanPhuAck[`${ev.date}|${ev.amount}`])
    .sort((a, b) => a.date.localeCompare(b.date));

  // Luyen, 2026-08-05: "hóa đơn đây soa lại kh lưu dc á gán cho tôi luôn đi"
  // -- BIDV77020, gian "SB CAM RANH PHN" ngay 01/08 hien "Chưa có HĐ" du da
  // luu "Số HĐ bù" qua form/route manual-match. Tim ra goc re: dong "SB CAM
  // RANH PHN" hom do KHONG co doanh thu QR THAT nao khop (khong nam trong
  // gianLines cua reconcileVietQr) -- toan bo 100.000đ hien thi la do block
  // "cfg.bankExcessDefaultCode" o tren TU TAO MOI dong nay (phan du Ngan
  // hang-Du lieu ca ngay, dung y "cho vô gian SB CAM RANH PHN"). reconcileVietQr
  // la noi DUY NHAT doc store.viet_qr_manual_matches truoc gio, nhung no chi
  // xu ly cac dong DA CO san tu grossData that -- dong do cfg.bankExcessDefaultCode
  // (hoac tanPhuTarget o tren) tu tao SAU do khong bao gio duoc doi chieu lai
  // voi manualMatches, nen moi lan Luyen luu deu bi "mat" nhu chua luu gi.
  // Cung 1 loi tiem an cho ca gian tanPhuTarget (AM TP PHCM/AE HP PHN) khi no
  // cung phai TU TAO MOI (chua tung co doanh thu that ngay do). Ap dung lai
  // manualMatches 1 lan CUOI CUNG cho MOI dong -- CHI voi dong nao reconcileVietQr
  // CHUA xu ly (manualOverride con la false) de khong cong grossAdjustment 2
  // lan cho dong DA duoc xu ly dung tu dau.
  reconciled.forEach((r) => {
    r.lines.forEach((l) => {
      if (l.manualOverride) return;
      const mm = manualMatches[`${r.settlementDate}|${l.code}`];
      if (!mm) return;
      if (mm.grossAdjustment) {
        l.gross += mm.grossAdjustment;
        l.net = l.gross;
        if (!r.pendingBank) r.totalNetComputed += mm.grossAdjustment;
      }
      l.invoiceNumbers = mm.invoiceNumbers || [];
      l.invoiceTotal = mm.amount != null ? mm.amount : l.gross;
      l.diff = l.invoiceTotal - l.gross;
      l.matched = Math.abs(l.diff) < 1;
      l.manualOverride = true;
      l.manualNote = mm.note || "";
    });
    if (!r.pendingBank) r.diffVsBank = r.totalNetComputed - r.bankAmount;
  });

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
    crossMatchSuggestions: (() => {
      const rawSuggestions = findCrossMatchSuggestions(reconciled, channelKey);
      const { remaining, changed } = applyApprovedCrossMatches(store, channelKey, reconciled, rawSuggestions);
      if (changed) save(store);
      return remaining;
    })(),
    invoiceDiemAlias,
    unmatchedInvoiceCodes: Array.from(unmatchedInvoiceCodesSet).sort(),
    // Canh bao rieng cho co che khop theo So tham chieu ngan hang (xem
    // CHANNELS.bidv7702.refMatchFrom o tren) -- rong ([]) voi cac kenh khong
    // bat tinh nang nay.
    refUnmatchedBankTx,
    refLateMatches,
    refUnmappedStoreCodes,
    refUnmappedTenDiem,
    // Luyen, 2026-08-05: cac lan tu dong gop phan du Ngan hang-Du lieu vao
    // gian mac dinh (Tan Phu/Hai Phong) CHUA duoc Luyen bam "Da xem" -- xem
    // ghi chu tanPhuAutoAppliedRaw o tren.
    tanPhuAutoApplied,
  };
}

// Chi Nhan, 2026-07-29: "cái đối soát vn paykh mới này cho nó riêng 1 trang
// ... cho nó chung với payoo á" -- tach phan than trang (dung chung cho ca
// /doi-soat/vietqr va /doi-soat/vnpay-khmoi ben duoi) thanh 1 ham rieng,
// nhan activeKeys/pageTitle/pageSubtitle tu 2 route khac nhau -- ban than
// logic doi soat/hien thi giu NGUYEN, chi khac danh sach kenh nao duoc hien.
function renderVietQrPage(req, res, activeKeys, pageTitle, pageSubtitle) {
  const store = load();
  ensureChannelShape(store);
  const activeCompany = getCompany(req);

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

  // Luyen, 2026-08-05: banner canh bao rieng cho cac lan tu dong gop phan du
  // Ngan hang-Du lieu vao gian mac dinh (Tan Phu/Hai Phong) CHUA duoc bam "Da
  // xem" -- xem ghi chu tanPhuAutoApplied trong buildChannelReconciliation.
  const allTanPhuAutoApplied = [];
  activeKeys.forEach((ch) => {
    (built[ch].tanPhuAutoApplied || []).forEach((r) => allTanPhuAutoApplied.push({ channel: ch, channelLabel: CHANNELS[ch].label, ...r }));
  });
  allTanPhuAutoApplied.sort((a, b) => (a.date || "").localeCompare(b.date || ""));

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
    exportTuDefault: monthBounds(selectedMonth).first,
    exportDenDefault: monthBounds(selectedMonth).last,
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
    allTanPhuAutoApplied,
    maCongTrinhOptions: (store.ma_cong_trinh_master && store.ma_cong_trinh_master[activeCompany] && store.ma_cong_trinh_master[activeCompany].rows) || [],
    partnerBankId: store.viet_qr_partner_bank_id || {},
    error: req.query.error || null,
    success: req.query.success || null,
    pageTitle,
    pageSubtitle,
  });
}

router.get("/doi-soat/vietqr", (req, res) => {
  // Chi hien cac ngan hang thuoc cong ty dang chon (nut chuyen cong ty tren
  // topbar) -- KH Cu van thay ca 3 ngan hang nhu truoc gio, KH Moi thay
  // BIDV 7702. Cac route upload/xoa/xuat file van nhan moi channelKey hop
  // le (khong gioi han theo cong ty) vi Luyen la nguoi dung duy nhat. Cac
  // kenh co pageGroup hien o trang RIENG (khong con kenh nao dung pageGroup
  // nua sau khi vnpayKhMoi chuyen sang routes/doisoat-vnpay-khmoi.js, nhung
  // giu dieu kien nay phong khi co kenh moi dung lai co che nay).
  const activeCompany = getCompany(req);
  const activeKeys = CHANNEL_KEYS.filter((ch) => CHANNELS[ch].company === activeCompany && !CHANNELS[ch].pageGroup);
  renderVietQrPage(req, res, activeKeys);
});

// Chi Nhan, 2026-07-29: trang "/doi-soat/vnpay-khmoi" (VNPay/Payoo KH Moi)
// tung dung kien truc VietQR o day da CHUYEN HAN sang routes/doisoat-vnpay-khmoi.js
// (kien truc ZVP dung ngan hang VTB982) -- xem ghi chu day du tai CHANNELS o
// tren. Route GET/POST rieng cho trang do khong con o file nay nua.

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

// Chi Nhan, 2026-07-30: "nếu chưa có á thì cho tôi 1 chỗ mã công trình nx để
// tô nhập pass vô á" -- thay vi bat phai sua file "Ten diem- Ma cong trinh"
// roi tai lai (nhu ghi chu trong view), cho gan truc tiep ngay tren web,
// cung kieu voi gan-ma-cua-hang o duoi. Ghi thang vao viet_qr_ten_diem_master
// (key da chuan hoa qua normText(tenDiem), dung KHOP voi cach resolveGianGrossByBankRef
// tra cuu -- xem utils/vietqrReconcile.js).
router.post("/doi-soat/vietqr/gan-ten-diem/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { tenDiemBan, maCongTrinh } = req.body;
    if (!tenDiemBan || !maCongTrinh) throw new Error("Thieu ten diem ban hoac ma cong trinh de gan.");
    if (!store.viet_qr_ten_diem_master[channelKey]) store.viet_qr_ten_diem_master[channelKey] = {};
    store.viet_qr_ten_diem_master[channelKey][normText(tenDiemBan)] = maCongTrinh;
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da gan ten diem "${tenDiemBan}" -> "${maCongTrinh}". Ket qua doi soat da tu cap nhat.`)
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
    // Nhan, 2026-08-06: nho lai CAP nay da duoc dong y -- ngay khac sau nay co
    // cung cap se tu dong ap dung, khong hoi lai (xem applyApprovedCrossMatches).
    if (!store.viet_qr_cross_match_approved_pairs[channelKey]) store.viet_qr_cross_match_approved_pairs[channelKey] = {};
    store.viet_qr_cross_match_approved_pairs[channelKey][`${fromCode}|${toCode}`] = true;
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(
          `Da doi tru "${displayCode(fromCode)}" <-> "${displayCode(toCode)}" ngay ${settlementDate}. Cac ngay khac co cung cap nay se tu dong ap dung, khong hoi lai.`
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
      // Nhan, 2026-08-06: nho lai CAP nay da duoc dong y -- ngay khac sau nay
      // co cung cap se tu dong ap dung, khong hien lai thanh goi y nua.
      if (!store.viet_qr_cross_match_approved_pairs[channelKey]) store.viet_qr_cross_match_approved_pairs[channelKey] = {};
      store.viet_qr_cross_match_approved_pairs[channelKey][`${fromCode}|${toCode}`] = true;
      applied++;
    }
    save(store);
    let msg = `Da doi tru hang loat ${applied} cap. Cac cap nay se tu dong ap dung cho ngay khac sau nay, khong hoi lai.`;
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
    // Chi Nhan, 2026-07-29: kenh vnpayKhMoi (02865168) dung file xuat truc
    // tiep tu portal VNPay ("DanhSachGiaoDich...xlsx"), dinh dang khac han 2
    // kieu tren -- xem parseVnpayPortalWorkbook trong utils/vietqrReconcile.js.
    // File nay da tra ve LUON storeMap (Diem thu la ten mo ta day du san,
    // khong can sheet "Cua hang" rieng), nen xu ly rieng ca 2 buoc trong 1
    // if thay vi ghep chung logic voi 2 kieu con lai ben duoi.
    const isVnpayPortal = CHANNELS[channelKey].parseMode === "vnpayPortal";
    let parsed;
    let storeMap = {};
    if (isVnpayPortal) {
      parsed = parseVnpayPortalWorkbook(req.file.buffer);
      storeMap = parsed.storeMap || {};
    } else {
      parsed = isMn ? parseVietQrMnRawWorkbook(req.file.buffer) : parseVietQrRawWorkbook(req.file.buffer);
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
      try {
        storeMap = isMn ? parseMaCuaHangAppSheet(req.file.buffer) : parseCuaHangSheet(req.file.buffer);
      } catch (eStore) {
        storeMap = {};
      }
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
// ---------- Danh dau "Da xem" 1 lan tu dong gop phan du Ngan hang-Du lieu
// vao gian mac dinh (Tan Phu/Hai Phong) -- Luyen, 2026-08-05: "cho hẳn vào
// Tân phú luôn là không được nhá bạn nhớ báo tôi nhá ... không cần hiển thị
// lại cái nào cũ, có gì mới thì báo tôi chỗ đó thôi". Giu nguyen co che tu
// dong gop (Luyen xac nhan qua AskUserQuestion la van muon giu, chi can
// canh bao ro hon), nhung luu lai (ngay|so tien) da xac nhan de KHONG hien
// lai banner nay nua o lan tai trang sau -- neu leftover ngay do sau nay
// DOI KHAC (vd tai them file lam thay doi so du), key moi se tu dong hien
// lai (xem tanPhuAutoApplied trong buildChannelReconciliation).
router.post("/doi-soat/vietqr/tanphu-ack/:channel", requireDataEntry, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    const { date, amount } = req.body;
    if (!date || !amount) throw new Error("Thieu ngay hoac so tien de danh dau da xem.");
    store.viet_qr_tanphu_ack[channelKey][`${date}|${amount}`] = {
      ackedAt: new Date().toISOString(),
    };
    save(store);
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da danh dau da xem, se khong hien lai canh bao nay."));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

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

// ---------- Xoa hoa don theo so ngay (trong dayList) + soHd range ----------
// Dung de go trung lap khi 2 file invoice cung co giao dich day X nhung soHD khac nhau.
// VietQR invoice matching dua vao dayList (so ngay trich tu cot DichVuThuHo), KHONG phai ngayHd.
// Body: { dayNum: "3", soHdMin: "12000" }  (soHdMin optional)
// Xoa tat ca invoice co dayList.includes(dayNum) AND soHd >= soHdMin (neu co) tren tat ca CHANNEL_KEYS.
router.post("/doi-soat/vietqr/invoices/remove-by-date", requireAdmin, (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    const { dayNum, soHdMin } = req.body || {};
    if (!dayNum) throw new Error("Thieu dayNum (so ngay, vd: 3).");
    const day = parseInt(dayNum, 10);
    if (isNaN(day)) throw new Error("dayNum khong hop le.");
    const minNum = soHdMin ? parseInt(soHdMin, 10) : null;
    let removedTotal = 0;
    for (const ch of CHANNEL_KEYS) {
      const before = store.viet_qr_invoices[ch].length;
      store.viet_qr_invoices[ch] = store.viet_qr_invoices[ch].filter((inv) => {
        // Giu lai neu invoice nay khong co ngay day trong dayList
        if (!Array.isArray(inv.days) || !inv.days.includes(day)) return true;
        // Giu lai neu soHd < nguong (de chi xoa series moi hon)
        if (minNum !== null && parseInt(inv.soHd, 10) < minNum) return true;
        return false; // xoa
      });
      removedTotal += before - store.viet_qr_invoices[ch].length;
    }
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(
          `Da xoa ${removedTotal} hoa don ngay ${day}${minNum ? ` co so HD >= ${minNum}` : ""} (tat ca kenh).`
        )
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
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

  // Luyen, 2026-07-31: "tu ngay may toi ngay may" -- loc theo khoang ngay
  // (query "tu"/"den", ISO yyyy-mm-dd) truoc khi xuat, giong Momo. Trang nay
  // truoc gio KHONG loc theo thang/ngay o export nen mac dinh (khong truyen
  // tu/den) van xuat toan bo nhu cu, tuong thich nguoc voi cach dung hien tai.
  const tuFilter = req.query.tu || "";
  const denFilter = req.query.den || "";
  let reconciledForExport = built.reconciled;
  if (tuFilter) reconciledForExport = reconciledForExport.filter((r) => r.settlementDate >= tuFilter);
  if (denFilter) reconciledForExport = reconciledForExport.filter((r) => r.settlementDate <= denFilter);

  // Luyen, 2026-07-21: "lý do thu là Thu tiền khách hàng (không theo hóa đơn)
  // đổi hết các file xuất misa nhá" -- dung 1 cau CO DINH giong het Momo, bo
  // cau rieng theo tung kenh nhu truoc (vd "...qua Viet QR (BIDV 7702)").
  const { rows } = buildExportRows(
    reconciledForExport,
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
  // Sample a few recent invoices (last 5) to inspect ngayHd/soHd format
  const sampleInvoices = invoices.slice(-5).map((i) => ({
    soHd: i.soHd,
    soHdType: typeof i.soHd,
    ngayHd: i.ngayHd,
    maDiem: i.maDiem,
  }));
  res.json({
    channel: channelKey,
    rawUploads: raw,
    storeUploads: storeUps,
    invoiceCount: invoices.length,
    sampleInvoices,
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
