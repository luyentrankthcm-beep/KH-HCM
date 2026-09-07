const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const AdmZip = require("adm-zip");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const {
  parseTongMomoWorkbook,
  parseInvoiceWorkbook,
  parseRawMomoPortalWorkbook,
  resolveRawPortalGross,
  parseKhMoiFeeTransactionWorkbook,
  resolveKhMoiFeeTransactionGross,
  resolveKhCuFixedFeeTransactionGross,
  reconcileMomo,
  extractMomoSettlements,
  isoToDmy,
} = require("../utils/momoReconcile");
const { parseSharedInvoiceWorkbook } = require("../utils/zvpReconcile");
const { getCompany } = require("../utils/companies");
const { BANK_COMPANY } = require("../utils/bankCompany");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

// Luyen, 2026-07-31: "mấy cái chỗ xuất ra chứng từ này nè từ ngày mấy tới
// ngày mấy nhá" -- cac nut "Xuat Excel (dinh dang AMIS)" (Momo/VietQR/ZVP)
// truoc day chi loc theo CA THANG (dropdown "Chon thang" tren trang xem);
// them tuy chon loc chinh xac theo NGAY (tu ngay - den ngay) cho cac nut nay.
// Mac dinh khi vao trang: tu = ngay dau, den = ngay cuoi cua thang dang chon,
// de KHONG doi hanh vi xuat hien tai (van dung dung 1 thang) tru khi Luyen tu
// tay sua lai 2 o ngay.
function monthBounds(m) {
  if (!m) return { first: "", last: "" };
  const [y, mo] = m.split("-").map(Number);
  const lastDay = new Date(y, mo, 0).getDate(); // ngay 0 cua thang sau = ngay cuoi thang nay
  return { first: `${m}-01`, last: `${m}-${String(lastDay).padStart(2, "0")}` };
}

const MOMO_BANK_NAME = "BIDV123456"; // TK 8699123456, da co san trong he thong
const MOMO_BANK_ACCOUNT = 8699123456;
const MOMO_BANK_FULLNAME = "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam";

// Cong ty "KH Moi" (TNHH GIAI TRI K&H) dung TK BIDV7701 lam kenh Momo rieng
// -- phap nhan/ngan hang KHAC KH Cu nen dung 1 cap store key rieng
// (momo_moi_gross_uploads/momo_moi_invoices, xem store.js) de khong lan du
// lieu 2 cong ty. gian_mapping/invoice_diem_alias van dung CHUNG (xem
// utils/companies.js) vi la cung 1 danh muc gian/mat bang.
//
// buildMomoReconciliation/cac route ben duoi nhan companyKey ("kh_cu"/
// "kh_moi") de biet dung cau hinh nao -- macDinh la "kh_cu" (tham so co
// default) de router.buildMomoReconciliation(store) ma routes/dashboard.js
// va routes/congno.js dang goi (KHONG sua 2 file do) tiep tuc chay dung y
// nhu truoc gio, khong bi anh huong boi thay doi nay.
const MOMO_CHANNELS = {
  kh_cu: {
    company: "kh_cu",
    bankName: MOMO_BANK_NAME,
    bankAccount: MOMO_BANK_ACCOUNT,
    bankFullName: MOMO_BANK_FULLNAME,
    grossKey: "momo_gross_uploads",
    invoicesKey: "momo_invoices",
    // Luyen, 2026-08-03: hoa don "CHT nop tien" (tien mat CHT nop truc tiep
    // NH) tag rieng trong sheet danh sach hoa don dung chung -- luu tach
    // khoi invoicesKey vi day KHONG phai doanh thu momo, dung cho tinh nang
    // xuat Excel AMIS rieng (xem buildMomoChtDeposits).
    chtInvoicesKey: "momo_cht_invoices",
    label: "BIDV 123456",
    exportSheet: "MISATHUE123456",
  },
  kh_moi: {
    company: "kh_moi",
    bankName: "BIDV7701",
    bankAccount: 8620107701,
    bankFullName: "Ngân hàng TMCP Đầu tư và Phát triển Việt Nam",
    grossKey: "momo_moi_gross_uploads",
    invoicesKey: "momo_moi_invoices",
    chtInvoicesKey: "momo_moi_cht_invoices",
    label: "BIDV 7701",
    exportSheet: "MISATHUE7701",
  },
};

// Luyen, 2026-07-21: "may cai lech 1d hay may dong do ban them vao phi neu
// lam tron .5 nha con khong thi cong vao phi gian nao do cong len nha roi
// tru ra" -- xem dinh nghia goc/ly do day du tai GET /doi-soat/momo ben duoi
// (noi ap dung cho hien thi). Hoist ra module-scope (2026-08-03) de
// export.xlsx dung LAI CHINH XAC nguong nay, tranh dinh nghia trung lap/lech
// nhau giua 2 noi.
const ROUNDING_ABSORB_THRESHOLD = 2000;

// A gian/ma cong trinh can be mapped to "131" (thu thuong), "1388" (doanh thu
// chia se, vd FF SC VIVO), or "SKIP" -- meaning: van hien thi tren trang doi
// soat de theo doi, nhung KHONG dua vao file xuat MISA. Dung cho khach hang
// moi chua duoc setup ke toan/MISA (vd KVC Estella, Farm Lotte Phan Thiet,
// Farm Lotte Nha Trang, AE Tan An KVC khi moi bat dau) -- cac gian nay cung
// duoc biet la settle vao 1 TK ngan hang KHAC, khong phai TK 123456, nen
// duoc loai khoi tong tien dung de doi chieu voi TK 123456 (xem reconcileMomo).
const TKCO_VALUES = ["131", "1388", "SKIP"];

// Different uploads can legitimately overlap in date range -- e.g. the
// hand-maintained "Tong Momo Gop" sheet for a whole month AND a raw MoMo
// portal zip for a few days within that same month both being uploaded.
// Summing overlapping (ngay, gian) values would silently double-count that
// revenue. Instead, sort uploads oldest-to-newest and let the MOST RECENT
// upload win for any (ngay, gian) key it covers, so re-uploading a
// corrected/refreshed file for a period always supersedes older data for
// those same days rather than stacking on top of it.
function mergeGross(uploads) {
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  let hasAnyNetByCode = false;
  const sorted = [...uploads].sort(
    (a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at)
  );
  for (const u of sorted) {
    (u.codes || []).forEach((c) => codes.add(c));
    for (const [k, v] of Object.entries(u.grossByCode || {})) {
      grossByCode[k] = v; // newer upload overwrites, does not add
    }
    // Luyen, 2026-07-21: gio "Bao cao doanh thu momo.xlsm" (Tong Momo KH Moi)
    // di kem san netByCode (so tien THAT ve ngan hang, xem utils/momoReconcile.js
    // reconcileMomo) -- truyen tiep qua day de khong bi mat khi merge nhieu
    // upload voi nhau (giong cach grossByCode duoc merge o tren).
    if (u.netByCode) {
      hasAnyNetByCode = true;
      for (const [k, v] of Object.entries(u.netByCode)) {
        netByCode[k] = v;
      }
    }
  }
  return { codes: Array.from(codes), grossByCode, netByCode: hasAnyNetByCode ? netByCode : undefined };
}

const CHUA_MAP_PREFIX = "CHUA MAP: ";

// "CHUA MAP: X" la 1 ma cua hang MOI xuat hien trong file zip cong MoMo ma
// store.cua_hang_mapping chua tung biet toi (xem resolveRawPortalGross). Ap
// dung anh xa TAI DAY (luc doc du lieu de hien thi/doi soat, KHONG phai luc
// parse file) nen ke ca nhung ban ghi gross DA tai len TU TRUOC (voi ma
// "CHUA MAP: X" cung) cung tu dong duoc gan lai dung Ma Cong Trinh ngay khi
// Luyen dien anh xa 1 lan, khong can tai lai file zip cu. Dung CHUNG
// store.cua_hang_mapping cho ca 2 cong ty (Luyen, 2026-07-17: "map o day cho
// ca 2 KH") -- mot lan dien ap dung ngay cho ca trang KH Cu va KH Moi.
function applyCuaHangAlias(grossData, cuaHangMapping) {
  const resolve = (rawCode) => {
    if (!rawCode.startsWith(CHUA_MAP_PREFIX)) return rawCode;
    const rawMaCuaHang = rawCode.slice(CHUA_MAP_PREFIX.length);
    const known = cuaHangMapping[rawMaCuaHang];
    return known && known.code ? known.code : rawCode;
  };
  const codesOut = new Set();
  const grossByCodeOut = {};
  for (const rawCode of grossData.codes) codesOut.add(resolve(rawCode));
  for (const [key, val] of Object.entries(grossData.grossByCode)) {
    const sep = key.indexOf("|");
    const date = key.slice(0, sep);
    const targetCode = resolve(key.slice(sep + 1));
    const newKey = `${date}|${targetCode}`;
    grossByCodeOut[newKey] = (grossByCodeOut[newKey] || 0) + val;
  }
  // Chi Nhan, 2026-07-22: netByCode (so tien NET, xem parseKhMoiFeeTransactionWorkbook)
  // phai duoc anh xa lai "CHUA MAP: X" -> ma cong trinh THAT giong het
  // grossByCode o tren -- neu khong, sau khi Luyen dien anh xa 1 lan, cac
  // key netByCode van con nam duoi ten "CHUA MAP: X" cu, khong bao gio khop
  // duoc voi grossByCode (da doi ten) nua, khien reconcileMomo tim netByCode
  // luon that bai (roi lai am tham fallback ve cong thuc phi co dinh 0,989).
  let netByCodeOut;
  if (grossData.netByCode) {
    netByCodeOut = {};
    for (const [key, val] of Object.entries(grossData.netByCode)) {
      const sep = key.indexOf("|");
      const date = key.slice(0, sep);
      const targetCode = resolve(key.slice(sep + 1));
      const newKey = `${date}|${targetCode}`;
      netByCodeOut[newKey] = (netByCodeOut[newKey] || 0) + val;
    }
  }
  return {
    ...grossData,
    codes: Array.from(codesOut),
    grossByCode: grossByCodeOut,
    netByCode: netByCodeOut,
  };
}

// Detect a .zip upload (either by extension or by the zip magic bytes "PK"),
// unzip it in memory, and return the buffer of the first .xlsx/.xls entry
// found inside. This is how the raw MoMo merchant-portal "daily_report" is
// distributed (a zip containing one xlsx).
//
// BUG FIXED 2026-07-23 (Luyen: "tai khoong dduocjw nos khoong nhaanj dangj",
// loi "Khong tim thay file Excel nao trong file .zip da tai len." khi tai
// dung file .xlsx "Transaction_report" -- KHONG phai zip): .xlsx/.xlsm ban
// than no CUNG LA 1 file zip (dinh dang Office Open XML), nen luon khop voi
// magic bytes "PK" ben duoi -- truoc day moi file .xlsx deu bi nham la file
// zip "daily_report", roi extractFirstExcelFromZip khong tim thay file Excel
// NAM LONG BEN TRONG (vi ban than no da la file dich, khong phai zip bao
// ngoai 1 file khac) nen bao loi. Sua: uu tien check DUOI FILE truoc (an
// toan, dut khoat cho ca 2 chieu); chi fallback ve magic-byte + kiem tra cau
// truc noi bo (co "[Content_Types].xml" o goc hay khong -- day la dau hieu
// rieng cua file Office, khong co trong file zip "daily_report" thong
// thuong) cho truong hop ten file khong ro duoi.
function isZipFile(file) {
  const name = file.originalname || "";
  if (/\.zip$/i.test(name)) return true;
  if (/\.(xlsx|xlsm|xls)$/i.test(name)) return false;
  const buf = file.buffer;
  if (!(buf && buf.length > 2 && buf[0] === 0x50 && buf[1] === 0x4b)) return false; // "PK"
  try {
    const zip = new AdmZip(buf);
    const isOfficeDoc = zip.getEntries().some((e) => /^\[Content_Types\]\.xml$/i.test(e.entryName));
    return !isOfficeDoc;
  } catch (e) {
    return false;
  }
}

function extractFirstExcelFromZip(buffer) {
  const zip = new AdmZip(buffer);
  const entry = zip
    .getEntries()
    .find((e) => !e.isDirectory && /\.(xlsx|xls|xlsm)$/i.test(e.entryName));
  if (!entry) {
    throw new Error("Khong tim thay file Excel nao trong file .zip da tai len.");
  }
  return entry.getData();
}

// Extracted so /cong-no and the Tong quan dashboard can reuse the EXACT same
// reconciliation this page shows, instead of recomputing it separately (2
// implementations of the same match logic would drift apart over time).
// Tra ve true neu co thay doi thuc su (de goi noi con lai tu quyet dinh co
// can save(store) hay khong).
function ensureGianHidden(store) {
  if (!store.gian_hidden) store.gian_hidden = {};
  if (!store.gian_hidden.kh_cu) store.gian_hidden.kh_cu = [];
  if (!store.gian_hidden.kh_moi) store.gian_hidden.kh_moi = [];
  let changed = false;
  // Bat buoc dong bo, moi lan: bat ky gian nao dang bi an khoi trang KH Cu
  // (nghia la doanh thu do THUC RA la cua KH Moi -- xem ghi chu o route
  // /doi-soat/momo/gian-hidden ben duoi) phai luon co TK Co = SKIP trong
  // store.gian_mapping, bat ke Luyen co nho vao muc 3 tu chinh lai hay khong.
  // Neu khong dong bo, gian se "mat tich" khoi CA HAI trang (an ben Cu, nhung
  // khong hien ben Moi vi tkCo != SKIP) -- day chinh la loi Luyen bao gap lai
  // nhieu lan (2026-07-17: "1 loi ma sua nhieu lan"). Tu dong sua o day MOI
  // LAN trang duoc mo nen tu dong ap dung ca cho nhung gian da bi an tu truoc
  // (KVC ESTELLA, FARM LOTTE NHA TRANG, FARM LOTTE PHAN THIET, AE TAN AN KVC,
  // CHUA MAP: KHEVENT2...), khong can Luyen tu vao muc 3 sua lai tung ma.
  if (store.gian_mapping) {
    (store.gian_hidden.kh_cu || []).forEach((code) => {
      if (store.gian_mapping[code] !== "SKIP") {
        store.gian_mapping[code] = "SKIP";
        changed = true;
      }
    });
  }
  return changed;
}

// Ten diem ghi tren hoa don (sheet "ke ds xuat HD MTT - 705", tab KH Moi)
// KHONG trung voi Ma Cong Trinh dung ben doanh thu cho 3 trong 4 gian KH Moi
// dang biet (Luyen, 2026-07-17: xac nhan qua file MTT 17.07.xlsx) -- vi du
// hoa don ghi "DIY ESTELLA KVC" nhung ben doanh thu/gian_mapping dung ma
// "KVC ESTELLA". Day CHINH XAC la 4 gian rieng cua KH Moi (xem gian_hidden.
// kh_cu o tren) -- ghi ro o day de khong bi lan voi ma nao ben KH Cu. Seed
// san alias nay (chi khi chua co, khong ghi de neu Luyen da tu sua khac) de
// hoa don tu dong khop ma khong can vao muc 3b tu tay them, dung tinh than
// "tu hieu, tu luu" Luyen yeu cau (2026-07-17).
const KNOWN_INVOICE_DIEM_ALIASES = {
  "DIY ESTELLA KVC": "KVC ESTELLA",
  "LM NHA TRANG KVC": "FARM LOTTE NHA TRANG",
  "LM PHAN THIẾT KVC": "FARM LOTTE PHAN THIET",
  // Luyen, 2026-08-03: "đây hóa đơn chỗ lệch đây á thêm vô cho tôi nhá" --
  // gian "AM TP KVCM" (khoan ve 2026-06-08, doanh thu 05-07/06) hien "Lệch
  // -104.994.000đ": doanh thu gop 155.539.000 vs tong tien HD khop chi
  // 50.545.000 (HD 1050, 1079). Hoa don 1082 (ngay HD 08/06, 104.994.000,
  // "momo 6,7") dung KHOP CHINH XAC so tien con thieu, nhung maDiem luu tren
  // he thong lai la "AE BT KVCM" (ten CU/khac cua cung 1 gian nay -- da co
  // san alias "SNOWFUN TÂN PHÚ" -> "AM TP KVCM" tu truoc, xac nhan day cung
  // la 1 ten khac cua CUNG gian Tan Phu). Them alias nay de hoa don 1082 tu
  // dong duoc tinh vao "AM TP KVCM", het "Lệch".
  "AE BT KVCM": "AM TP KVCM",
  // Luyen, 2026-08-10: "hóa đơn momo kh cũ của vũng tàu đây nhá map vô cho tôi
  // đi" -- hoa don KVC LOTTE VUNG TAU (seed o store.js, HĐ 2512/2547 t08.2026)
  // bi reconcileMomo khong khop vi INVOICE_DIEM_ALIAS_DEFAULTS ben doisoat-
  // vietqr.js co entry "KVC LOTTE VUNG TAU" -> "VUNG TAU PHCM" (dung cho VietQR
  // BIDV7702), entry nay duoc seed vao store.invoice_diem_alias (bang chung) va
  // momoEffectiveDiemAlias() ke thua -- effectiveMaDiem thanh "VUNG TAU PHCM"
  // trong khi gross Momo KH Cu dung dung "KVC LOTTE VUNG TAU" nen khong bao gio
  // khop. Bao ve bang KNOWN_INVOICE_DIEM_ALIASES: Momo ghi de lai "KVC LOTTE
  // VUNG TAU" -> "KVC LOTTE VUNG TAU" (tu tham chieu = no-op cho ghi de), giu
  // nguyen ma gian chinh xac cho reconcileMomo tim duoc invoice 2512/2547.
  "KVC LOTTE VUNG TAU": "KVC LOTTE VUNG TAU",
};

function ensureKnownInvoiceDiemAliases(store) {
  if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
  let changed = false;
  for (const [raw, target] of Object.entries(KNOWN_INVOICE_DIEM_ALIASES)) {
    if (!store.invoice_diem_alias[raw]) {
      store.invoice_diem_alias[raw] = target;
      changed = true;
    }
  }
  return changed;
}

// Luyen, 2026-08-01 (lan 7): "đây hóa đơn gian Farm Phan thiết đây á" -- gian
// "FARM LOTTE PHAN THIET" hien "Chưa có HĐ" du hoa don 7609 that su co (Mã
// điểm ghi chú HT Misa = "LM PHAN THIẾT KVC"). Nguyen nhan: store.invoice_diem_
// alias la 1 bang DUNG CHUNG giua Momo/ZVP (xem chu thich store.js), nhung
// ngay 2026-07-31 tinh nang doi ma cong trinh RIENG cua ZVP (xem
// ZVP_GIAN_CODE_RENAMES/seedZvpGianCodeRenames trong routes/doisoat-zvp.js)
// da GHI DE LEN entry "LM PHAN THIẾT KVC" tu "FARM LOTTE PHAN THIET" (dung
// cho Momo) sang "LM PHAN THIET KVC" (dung cho ZVP sau khi doi ma) -- 2 tinh
// nang can 2 gia tri KHAC NHAU cho CUNG 1 khoa nen khong the dung chung mai
// duoc (sua lai tay qua UI se bi ZVP ghi de lai lan sau tai trang ZVP). Thay
// vi tach han 2 bang rieng (rui ro dung cham nhieu noi), Momo tu BAO VE 3 alias
// da biet chac chan cua rieng minh (KNOWN_INVOICE_DIEM_ALIASES) bang cach de
// chung LUON THANG (ghi de) khi doi soat -- khong quan tam bang dung chung dang
// bi ZVP doi thanh gi, khong anh huong ZVP (ZVP van doc thang store.invoice_
// diem_alias binh thuong, khong sua o day).
function momoEffectiveDiemAlias(store) {
  return { ...(store.invoice_diem_alias || {}), ...KNOWN_INVOICE_DIEM_ALIASES };
}

// Luyen, 2026-08-03: "ngân hàng có các giao dịch của Cửa Hàng trưởng nộp tiền
// vào á... như 7701 thì có gian nội dung là Lotte Phan Thiết á hay Nha trang
// dựa vào nội dung có mã nộp tiền á" -- ma nop tien (vd "KH705KVCMN0002")
// nam san trong mo ta giao dich ngan hang (cung quy uoc voi tinh nang CHT nop
// tien cua Hoa Don Ban Ra, xem routes/baocao.js CHT_CODE_RE/collectChtNopTienLines
// -- KHONG import lai file do vi cac ham do khong duoc export rieng, chi co
// router; dinh nghia lai 1 ban gon o day, dung CHUNG bang tra store.cht_nop_
// tien_map da co san (seedChtNopTienMapFromKvcMtdFiles, store.js).
//
// Luyen, 2026-08-03 (lan 2): "Ngân hàng có khoản này với gian này đâu sao lại
// có trên xuất ra z" -- phat hien qua vi du that: dong "MM MARKET DA NANG
// MTD" 11.040.000đ (02/06, VP58888, ma KH705MTDMN0027) bi dua NHAM vao file
// xuat MISA cua Momo. "MTD" (May Tu Dong -- vending/game may rieng) la 1
// LOAI HINH/NGHIEP VU HOAN TOAN KHAC voi "KVC" (Khu Vui Choi -- doanh thu
// theo gian entertainment park ma Momo dang xu ly), co pipeline doi soat
// RIENG (xem sheet MISAMTD/skill cap-nhat-misa-mtd-kvc, sao ke ngan hang
// rieng MTD40222/KVC40111) -- KHONG lien quan gi den settlement Momo. Ham
// collectChtNopTienLines (routes/baocao.js, cho tinh nang Xuat Hoa Don Ban
// Ra) dung ca 2 loai vi tinh nang do gom TAT CA nghiep vu, nhung
// buildMomoChtDeposits (ben duoi, RIENG cho Momo) chi duoc lay CHT nop tien
// cua gian KVC -- dung regex rieng, loai han MTDMB/MTDMN.
const CHT_CODE_RE = /KH(?:705|989)(?:KVCMB|KVCMN|MTDMB|MTDMN)\d{3,4}/;
const CHT_CODE_RE_KVC_ONLY = /KH(?:705|989)(?:KVCMB|KVCMN)\d{3,4}/;

function companyForBankRow(b) {
  if (!b) return null;
  return b.company || BANK_COMPANY[b.name] || null;
}

// Quet TOAN BO giao dich ngan hang (khong phai 1 kenh doi soat rieng, xem ghi
// chu CHT_CODE_RE o tren) tim cac khoan CHT nop tien mat, cong don theo (ngay
// ngan hang nhan tien, gian) roi khop voi hoa don "CHT nộp tiền N" da nap qua
// upload-hoadon (store[momoCfg.chtInvoicesKey], xem parseInvoiceWorkbook/
// momoReconcile.js) theo CUNG ngay + gian. Tra ve 1 dong / (ngay, gian) --
// dung cho export.xlsx (thay the hoan toan store.momo_kl_deposits cu, hand-
// populated 1 lan, chi co 17 dong 2026-07-03->07-20).
function buildMomoChtDeposits(store, activeCompany, chtInvoicesKey) {
  const diemAlias = momoEffectiveDiemAlias(store);
  const depositsByKey = new Map(); // "date|gian" -> tong tien
  (store.transactions || []).forEach((t) => {
    if (t.type !== "thu" || !t.date) return;
    const desc = String(t.description || "").toUpperCase();
    const match = desc.match(CHT_CODE_RE_KVC_ONLY);
    if (!match) return;
    const bank = store.banks.find((b) => b.id === t.bank_id);
    if (companyForBankRow(bank) !== activeCompany) return;
    const mapped = (store.cht_nop_tien_map || {})[match[0]];
    if (!mapped || !mapped.maCongTrinh) return;
    const key = `${t.date}|${mapped.maCongTrinh}`;
    depositsByKey.set(key, (depositsByKey.get(key) || 0) + (t.amount || 0));
  });

  const invoicesByKey = new Map(); // cung key -> [{soHd, ngayHd, tongTt}]
  (store[chtInvoicesKey] || []).forEach((inv) => {
    if (!inv.ngayNopTien) return;
    const gian = diemAlias[inv.maDiem] || inv.maDiem;
    const key = `${inv.ngayNopTien}|${gian}`;
    if (!invoicesByKey.has(key)) invoicesByKey.set(key, []);
    invoicesByKey.get(key).push(inv);
  });

  const results = [];
  for (const [key, amount] of depositsByKey) {
    const [date, gian] = key.split("|");
    const candidates = invoicesByKey.get(key) || [];
    const invoiceTotal = candidates.reduce((s, c) => s + (c.tongTt || 0), 0);
    const soHd = candidates.map((c) => c.soHd).filter(Boolean).join(", ");
    results.push({
      date,
      gian,
      amount,
      soHd,
      invoiceTotal,
      matched: candidates.length > 0 && Math.abs(invoiceTotal - amount) <= 1000,
    });
  }
  return results.sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : 0));
}

// KH Moi CHUA co tai khoan Momo rieng thuc su nhan tien ve (BIDV7701 con 0
// giao dich thuc te tren he thong tinh den 2026-07-16) -- toan bo sao ke
// ngan hang VA doanh thu "Tong Momo" (ca 2 cong ty) van dang chay CHUNG qua
// 1 tai khoan cua KH Cu (BIDV123456). Theo yeu cau Luyen 2026-07-16: trang
// KH Moi van doi soat tren CUNG nguon sao ke/doanh thu do (muon
// MOMO_CHANNELS.kh_cu), nhung khop voi HOA DON RIENG cua KH Moi
// (momo_moi_invoices) -- de xem trong CUNG 1 khoan tien ve do, gian nao da
// co hoa don xuat duoi ten KH Moi.
//
// Luyen, 2026-07-17: thu doi sang uu tien BIDV7701 theo TUNG NGAY (ngay nao
// co giao dich that o do thi dung, fallback ve 123456 cho ngay con thieu) --
// nhung phat hien hong ngay khi test: "Tinh tu Tong Momo" la 1 TONG GOP
// CHUNG ca 2 cong ty (khong the tach theo tung ngan hang duoc), nen khi lay
// so tien nho cua rieng BIDV7701 di so voi tong Tong Momo (ca cong ty, lon
// hon nhieu) thi ra chenh lech gia rat lon tren hau het cac ngay (vd 07-02:
// Ngan hang 5,6tr vs Tong Momo 40,4tr; 07-06: 31,5tr vs 283tr...). Da
// ROLLBACK ve lai dung 1 nguon KH Cu tam thoi, CHO toi khi co cach tach
// doanh thu "Tong Momo" theo tung ngan hang/cong ty mot cach dang tin cay.
//
// Luyen, 2026-07-21: gio DA co file rieng ("BÁO CÁO DOANH THU momo.xlsm",
// sheet "Tổng Momo T7") tach dung doanh thu-theo-gian CUA RIENG KH Moi (voi
// netByCode = so tien THAT ve BIDV7701, xem reconcileMomo/mergeGross), da
// nap vao store.momo_moi_gross_uploads va da KIEM CHUNG khop dung ngan hang
// cho tung ngay. Vay nen KHONG con can muon nguon KH Cu nua MOI KHI KH Moi
// da co du lieu rieng: neu store.momo_moi_gross_uploads con RONG (chua tung
// tai file rieng) thi VAN fallback ve nguon chung KH Cu nhu truoc (an toan,
// khong lam mat du lieu cu); ngay khi co it nhat 1 upload rieng, dung THANG
// nguon rieng do.
function momoSourceCfg(companyKey, store) {
  if (companyKey === "kh_moi") {
    const hasOwnUploads = store && (store.momo_moi_gross_uploads || []).length > 0;
    return hasOwnUploads ? MOMO_CHANNELS.kh_moi : MOMO_CHANNELS.kh_cu;
  }
  return MOMO_CHANNELS[companyKey] || MOMO_CHANNELS.kh_cu;
}

function buildMomoReconciliation(store, companyKey) {
  const cfg = MOMO_CHANNELS[companyKey] || MOMO_CHANNELS.kh_cu;
  const sourceCfg = momoSourceCfg(companyKey, store);
  const usesSharedSource = sourceCfg !== cfg;
  const grossUploads = store[sourceCfg.grossKey] || [];
  // Luyen, 2026-08-10: Momo va Zalo App la 2 kenh RIENG BIET, xuat hoa don rieng.
  // Khong duoc tron HĐ Zalo vao pool Momo. Chi dung dung invoicesKey cua tung kenh.
  const invoices = store[cfg.invoicesKey] || [];
  // Luyen, 2026-07-21: "check so ngan hang 7701" -- BIDV7701 (TK rieng cua
  // KH Moi) gio DA co giao dich REM Momo that (tu 01/07/2026 den nay, xem
  // REM 9901CI2607... "MoMo TT GIAI TRI KvaH"), khac voi luc 07-16/17 khi TK
  // nay con 0 giao dich (xem ghi chu momoSourceCfg o tren). Nen SO TIEN NGAN
  // HANG/settlement date phai luon lay tu TK RIENG cua tung cong ty (cfg.bankName),
  // KHONG con muon cua KH Cu (sourceCfg.bankName) nua -- rieng du lieu DOANH
  // THU THEO GIAN (grossUploads o tren) van phai dung nguon CHUNG vi KH Moi
  // chua tung tai len 1 file "Tong Momo" rieng (momo_moi_gross_uploads con
  // rong). Tach 2 nguon nay ra: bankAmount dung so that cua tung TK, con
  // "Tinh tu Tong Momo"/Chenh lech se tu nhien phan anh dung thuc te la du
  // lieu doanh thu-theo-gian cua KH Moi con thieu (chua du 100% so voi TK),
  // thay vi so sanh nham voi tong CA 2 cong ty nhu truoc day.
  const bank = store.banks.find((b) => b.name === cfg.bankName);

  // Dong bo TRUOC khi tinh doi soat (khong phai sau) de ban ghi vua sua/seed
  // co hieu luc NGAY trong lan render nay, khong phai doi lan tai trang sau.
  let needsSave = ensureGianHidden(store);
  if (ensureKnownInvoiceDiemAliases(store)) needsSave = true;
  if (ensureNo1388(store)) needsSave = true;
  if (needsSave) save(store);

  let reconciledAll = [];
  let error = null;
  let allCodes = new Set();

  if (bank) {
    const txs = store.transactions.filter((t) => t.bank_id === bank.id);
    const settlements = extractMomoSettlements(txs);
    const grossData = applyCuaHangAlias(mergeGross(grossUploads), store.cua_hang_mapping || {});
    const invoiceData = { invoices };
    if (grossData.codes.length > 0) {
      reconciledAll = reconcileMomo(settlements, grossData, invoiceData, store.gian_mapping, momoEffectiveDiemAlias(store));
      // Chi Nhan, 2026-07-30: "các đối soát tất cả các trang điều xếp theo
      // ngày cho tôi nhá" -- reconcileMomo tra ve ket qua theo THU TU giao
      // dich ngan hang trong store.transactions (thu tu tai/nhap lieu, KHONG
      // phai thu tu ngay thang), nen trang co the hien VD ngay 30 roi toi
      // 23 roi 24 (dung nhu Chi Nhan gap). Sap lai TANG DAN theo settlementDate
      // NGAY SAU KHI TINH XONG (giong cach doisoat-vnpay-khmoi.js da lam) de
      // moi cho dung ben duoi (monthSet/months, reconciledMonth, reconciled)
      // deu tu dong ke thua dung thu tu, khong phai sua rai rac nhieu cho.
      reconciledAll.sort((a, b) => (a.settlementDate < b.settlementDate ? -1 : a.settlementDate > b.settlementDate ? 1 : 0));
      // An gian theo YEU CAU RIENG cua tung cong ty (vd 1 gian duoc gop nham
      // vao file tai len cua cong ty nay, nhung se duoc xuat HD ben cong ty
      // KIA -- Luyen, 2026-07-16: "KVC ESTELLA", "FARM LOTTE NHA TRANG",
      // "FARM LOTTE PHAN THIET", "AE TAN AN KVC" o KH Moi). CHI an khoi
      // HIEN THI cua cong ty dang xem (store.gian_hidden theo companyKey) --
      // KHONG dong cham store.gian_mapping (van dung chung), nen cong ty KIA
      // neu co cung ma van thay binh thuong. Loc TRUOC khi tinh allCodes de
      // gian bi an cung bien mat khoi bang "TK Co" muc 3 luon, khong chi
      // khoi bang ket qua muc 4.
      const hiddenSet = new Set(store.gian_hidden[companyKey] || []);
      if (hiddenSet.size > 0) {
        reconciledAll = reconciledAll.map((r) => ({ ...r, lines: r.lines.filter((l) => !hiddenSet.has(l.code)) }));
      }
      reconciledAll.forEach((r) => r.lines.forEach((l) => allCodes.add(l.code)));
    }
  } else {
    error = `Chua co ngan hang "${sourceCfg.bankName}" trong he thong.`;
  }

  // Any invoice "Ma diem" that never lines up with a known gian code (and
  // isn't already aliased) is surfaced here -- these invoices exist and have
  // real money on them, they're just filed under a name the revenue side
  // doesn't recognize (e.g. a sub-brand/corner name like "SNOWFUN TAN PHU"
  // instead of the Ma Cong Trinh "AM TP KVCM"), so the reconciliation above
  // silently can't count them yet. See mục 3b on the page: cho phép Luyen
  // ánh xạ tên này về đúng Ma Cong Trinh, áp dụng ngay không cần tải lại HĐ.
  const knownCodes = new Set([...allCodes, ...Object.keys(store.gian_mapping || {})]);
  const invoiceDiemAlias = momoEffectiveDiemAlias(store);
  const unmatchedInvoiceCodesSet = new Set();
  invoices.forEach((inv) => {
    if (!inv.maDiem) return;
    if (knownCodes.has(inv.maDiem)) return;
    if (invoiceDiemAlias[inv.maDiem]) return;
    unmatchedInvoiceCodesSet.add(inv.maDiem);
  });

  // "Khoa so" -- Luyen, 2026-07-21: "TẤT CẢ CÁC TRANG ĐIỀU CÓ KHÓA SỔ CHO TÔI
  // NHÁ", ap dung cung 1 co che da lam cho VietQR (routes/doisoat-vietqr.js):
  // luu 1 ngay "khoa den het ngay X" THEO TUNG CONG TY (store.momo_lock_date),
  // moi ngay settlementDate <= ngay do duoc coi la "locked" -- dua diff ve 0
  // (khong con cong don vao Chenh lech), giu nguyen gross/invoiceTotal de tra
  // cuu, bao view hien badge "Da khoa so" thay vi "Chua co HD"/"Lech".
  const lockDate = (store.momo_lock_date && store.momo_lock_date[companyKey]) || "";
  if (lockDate) {
    reconciledAll.forEach((r) => {
      if (r.settlementDate <= lockDate) {
        r.locked = true;
        r.lines.forEach((l) => {
          l.locked = true;
          l.diff = 0;
        });
      }
    });
  }

  return {
    reconciledAll,
    lockDate,
    allCodes,
    error,
    invoiceDiemAlias,
    unmatchedInvoiceCodes: Array.from(unmatchedInvoiceCodesSet).sort(),
    usesSharedSource,
    sourceLabel: sourceCfg.label,
  };
}

router.get("/doi-soat/momo", (req, res) => {
  const store = load();
  ensureGianHidden(store);
  const activeCompany = getCompany(req);
  const momoCfg = MOMO_CHANNELS[activeCompany];
  const built = buildMomoReconciliation(store, activeCompany);
  const reconciledAll = built.reconciledAll;
  const allCodes = built.allCodes;
  const error = req.query.error || built.error;

  // Chon theo thang: mac dinh la thang gan nhat de trang khong bi dai/roi
  // ("nhieu roi qua" -- Luyen), nhung van chon "Tat ca" duoc qua dropdown.
  const monthSet = new Set(reconciledAll.map((r) => r.settlementDate.slice(0, 7)));
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";
  const reconciledMonth = selectedMonth
    ? reconciledAll.filter((r) => r.settlementDate.slice(0, 7) === selectedMonth)
    : reconciledAll;

  // Loc dong hien thi o muc 4 theo DUNG cong ty dang xem cho de nhin (Luyen,
  // 2026-07-17: "trang kh moi bo cac ma check tien ve kh cu di nguoc lai de
  // no dep trang"). Trang KH Moi dang tam dung CHUNG sao ke/doanh thu voi KH
  // Cu (xem momoSourceCfg) nen 1 khoan tien ve dang liet ke ca cac gian binh
  // thuong (131/1388) cua KH Cu -- khong lien quan KH Moi. Chi AN bot cac
  // dong KHONG thuoc cong ty dang xem (dua vao TK Co: SKIP = KH Moi, con lai
  // = KH Cu).
  //
  // Luyen, 2026-07-21: "hien so ben momo kh moi thoi cac gian do cong lai" --
  // doi lai quyet dinh 07-17 o tren ("KHONG dong den tong Ngan hang/Chenh
  // lech, tinh tren toan bo du lieu"): gio TK BIDV7701 rieng cua KH Moi DA co
  // giao dich that (xem buildMomoReconciliation, bank = cfg.bankName), nen
  // "Tinh tu Tong Momo"/Chenh lech o trang KH Moi phai la TONG CUA RIENG cac
  // dong dang hien thi (cac gian SKIP = cua KH Moi) cong lai, KHONG con la
  // tong CA 10 ma (ca cua KH Cu) nhu truoc -- so voi bankAmount that cua
  // BIDV7701 se cho thay dung khoang trong du lieu doanh thu-theo-gian con
  // thieu cua KH Moi (con KH Cu thi khong doi gi, van tinh tren toan bo nhu
  // cu vi TK 123456 cua KH Cu van la nguon dung chung goc).
  // Luyen, 2026-07-21: "may cai lech 1d hay may dong do ban them vao phi neu
  // lam tron .5 nha con khong thi cong vao phi gian nao do cong len nha roi
  // tru ra" -- phan lech vai dong con lai (07-06: 693d, va rai rac 1-2d o
  // nhieu ngay khac) la do LAM TRON: file "Tong Momo T7" cua chi tinh phi
  // tung giao dich le (co phan thap phan), roi lam tron TUNG DONG/gian rieng
  // -- cong tat ca dong da lam tron lai KHONG LUON khop tuyet doi voi dong
  // "Tien ve Ngan hang" tong hop (cung lam tron nhung tu tong khac). Tu dong
  // don phan chenh lech nay vao dong gian LON NHAT trong khoan do (giong dung
  // "them vo gian bat ky di" chi noi lan 2) de Chenh lech luon ve dung 0.
  // Nguong 2.000d (Luyen xac nhan khoan 693d/07-06 cung la lam tron, khong
  // phai thieu du lieu that) -- van du thap de neu sau nay co khoan thieu
  // that su lon (hang trieu tro len, nhu vu 200tr/50tr da gap truoc do) thi
  // van hien ra chu khong bi am tham nuot mat. (ROUNDING_ABSORB_THRESHOLD
  // dinh nghia o module-scope phia tren, dung chung voi export.xlsx.)
  // Luyen, 2026-08-14: "them cho chinh sua cho toi nha" -- ap dung override
  // thu cong (momo_manual_matches) len tung dong truoc khi hien thi. Nguoi
  // dung bam "Sua HD" de nhap so HD + so tien dung, luu vao store, trang tu
  // hien lai voi so lieu da chinh.
  const momoManual = store.momo_manual_matches || {};
  const reconciled = reconciledMonth
    .map((r) => {
      const lines = r.lines.filter((l) => (activeCompany === "kh_moi" ? l.tkCo === "SKIP" : l.tkCo !== "SKIP"));
      // Apply manual invoice overrides BEFORE rounding-absorb step
      lines.forEach((l) => {
        const mm = momoManual[`${r.settlementDate}|${l.code}`];
        if (!mm) return;
        if (mm.grossAdjustment) {
          l.grossOriginal = l.gross;
          l.gross += mm.grossAdjustment;
          l.grossAdjusted = true;
        }
        if (mm.netAdjustment !== null && mm.netAdjustment !== undefined) {
          l.netOriginal = l.net;
          l.net = mm.netAdjustment;
          l.netAdjusted = true;
        }
        if (mm.invoiceNumbers) l.invoiceNumbers = mm.invoiceNumbers;
        if (mm.amount !== null && mm.amount !== undefined) l.invoiceTotal = mm.amount;
        l.diff = l.invoiceTotal - l.gross;
        l.matched = l.invoiceNumbers.length > 0 && Math.abs(l.diff) <= 1000;
        l.manualOverride = true;
        l.grossAdjustmentVal = mm.grossAdjustment || 0;
        l.netAdjustmentVal = mm.netAdjustment;
      });
      if (activeCompany !== "kh_moi") return { ...r, lines };
      // pendingBank (chua co giao dich ngan hang that cho ngay nay -- xem
      // buildPendingDaySettlements) -- giu nguyen diffVsBank = null, KHONG
      // tinh/hap thu lam tron o day (khong co r.bankAmount that de so sanh).
      if (r.pendingBank) {
        const totalNetComputed = lines.reduce((sum, l) => sum + l.net, 0);
        return { ...r, lines, totalNetComputed, diffVsBank: null };
      }
      let totalNetComputed = lines.reduce((sum, l) => sum + l.net, 0);
      let diffVsBank = totalNetComputed - r.bankAmount;
      if (diffVsBank !== 0 && Math.abs(diffVsBank) <= ROUNDING_ABSORB_THRESHOLD && lines.length > 0) {
        const biggest = lines.reduce((a, b) => (b.gross > a.gross ? b : a), lines[0]);
        biggest.net -= diffVsBank;
        totalNetComputed = lines.reduce((sum, l) => sum + l.net, 0);
        diffVsBank = totalNetComputed - r.bankAmount;
      }
      return { ...r, lines, totalNetComputed, diffVsBank };
    })
    .filter((r) => r.lines.length > 0);

  const exportRange = monthBounds(selectedMonth);

  res.render("doisoat-momo", {
    userName: req.session.userName,
    bankLabel: momoCfg.label,
    uploads: store[momoCfg.grossKey] || [],
    invoiceCount: (store[momoCfg.invoicesKey] || []).length,
    reconciled,
    months,
    selectedMonth,
    exportTuDefault: exportRange.first,
    exportDenDefault: exportRange.last,
    gianMapping: store.gian_mapping,
    allCodes: Array.from(allCodes).sort(),
    invoiceDiemAlias: built.invoiceDiemAlias,
    unmatchedInvoiceCodes: built.unmatchedInvoiceCodes,
    gianHidden: (store.gian_hidden && store.gian_hidden[activeCompany]) || [],
    usesSharedSource: built.usesSharedSource,
    sourceLabel: built.sourceLabel,
    lockDate: built.lockDate || "",
    error,
    success: req.query.success || null,
    isAdmin: req.session.role === "admin",
  });
});

// ---------- Khoa so (giong VietQR) -- gui lockDate rong de mo khoa lai. ----------
router.post("/doi-soat/momo/khoa-so", requireAdmin, (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  try {
    const lockDate = (req.body.lockDate || "").trim();
    if (lockDate && !/^\d{4}-\d{2}-\d{2}$/.test(lockDate)) throw new Error("Ngay khoa khong hop le (dang YYYY-MM-DD).");
    if (!store.momo_lock_date) store.momo_lock_date = {};
    store.momo_lock_date[activeCompany] = lockDate;
    save(store);
    const msg = lockDate
      ? `Da khoa so den het ngay ${lockDate}. Cac ngay tu do tro ve truoc se khong con hien canh bao lech nua.`
      : "Da mo khoa so.";
    res.redirect("/doi-soat/momo?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

// Guard against the same file being submitted twice in quick succession
// (observed in practice: a slow connection or double-click on "Tai len" can
// cause the browser to fire the upload request twice, creating an exact
// duplicate momo_gross_uploads entry which then double-counts revenue).
// If the most recent upload has the identical file name + gross data and
// was created within the last 2 minutes, treat this as a duplicate submit
// and skip adding a second copy.
function isDuplicateRecentUpload(store, grossKey, fileName, grossByCode) {
  const uploads = store[grossKey] || [];
  const recent = uploads[uploads.length - 1];
  if (!recent) return false;
  if (recent.file_name !== fileName) return false;
  const ageMs = Date.now() - new Date(recent.uploaded_at).getTime();
  if (ageMs > 2 * 60 * 1000) return false;
  return JSON.stringify(recent.grossByCode) === JSON.stringify(grossByCode);
}

function seedGianMappingDefaults(store, codes) {
  codes.forEach((c) => {
    if (!(c in store.gian_mapping)) {
      // Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- TK Co 1388
      // (doanh thu chia se/CSE) khong con duoc dung nua, moi gian moi (ke ca
      // gian FF/chia se truoc day se la 1388) deu mac dinh 131.
      store.gian_mapping[c] = "131";
    }
  });
}

// Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- ap dung cho CA HAI
// cong ty (gian_mapping dung chung KH Cu/KH Moi). Tu dong sua BAT KY gian
// nao con dang la 1388 (du la du lieu cu da co tu truoc) ve 131 moi lan
// trang doi soat duoc mo, khong can Luyen tu vao tung dong sua tay.
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

router.post("/doi-soat/momo/upload-tong", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const momoCfg = MOMO_CHANNELS[activeCompany];
  if (!store[momoCfg.grossKey]) store[momoCfg.grossKey] = [];
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");

    // Case 1: a .zip containing the raw MoMo merchant-portal daily export
    // ("data"/"bnpl"/"account_to_account" sheets, per-transaction detail).
    // Case 2: a normal .xlsx/.xlsm (the hand-maintained "Tong Momo Gop" sheet).
    let successMsg;
    if (isZipFile(req.file)) {
      const xlsxBuffer = extractFirstExcelFromZip(req.file.buffer);
      const { transactions } = parseRawMomoPortalWorkbook(xlsxBuffer);
      const resolved = resolveRawPortalGross(transactions, store.cua_hang_mapping);
      if (isDuplicateRecentUpload(store, momoCfg.grossKey, req.file.originalname, resolved.grossByCode)) {
        return res.redirect(
          "/doi-soat/momo?success=" +
            encodeURIComponent(`File "${req.file.originalname}" vua duoc tai len roi (bo qua ban trung lap).`)
        );
      }
      store[momoCfg.grossKey].push({
        id: nextId(store, "momo_gross_uploads_seq") || Date.now(),
        uploaded_at: new Date().toISOString(),
        file_name: req.file.originalname,
        sheetName: "MoMo portal (zip)",
        dates: resolved.dates,
        codes: resolved.codes,
        grossByCode: resolved.grossByCode,
      });
      seedGianMappingDefaults(store, resolved.codes);
      save(store);
      successMsg = `Da nap file zip bao cao MoMo (${transactions.length} giao dich, ${resolved.dates[0]} - ${resolved.dates[resolved.dates.length - 1]}). Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.`;
      if (resolved.unmapped.length > 0) {
        successMsg += ` CANH BAO: ${resolved.unmapped.length} ma cua hang chua map duoc Ma Cong Trinh (${resolved.unmapped.join(", ")}) -- doanh thu cua cac ma nay dang hien o dong "CHUA MAP: ..." trong bang duoi, hay tai lai 1 file "Tong Momo Gop" co chua cac ma cua hang nay de he thong tu hoc mapping.`;
      }
    } else {
      // Chi Nhan, 2026-07-22: "tôi mới tải file bạn xem nó trả qua ví trả
      // sao hay gì đó lấy ra tiền sao phí cho tôi được không" -- file .xlsx
      // (khong nen zip) tai truc tiep tu cong MoMo, cot KHONG co tien to
      // "MS." (Thời gian/Số tiền/Nguồn tiền/Mã cửa hàng...) chua du thong
      // tin de tinh PHI THAT theo tung giao dich (KH Moi dung 3 muc phi khac
      // nhau tuy Nguon tien: Vi MoMo 1%, Vi tra sau 1.2%, con lai 0.3% --
      // xem utils/momoReconcile.js parseKhMoiFeeTransactionWorkbook). Thu
      // parser nay TRUOC (chi khop dung khi file THAT SU co du cac cot can
      // thiet, nem loi va roi ve parser "Tong Momo Gop/flat" cu neu khong).
      // Luyen, 2026-07-24: "check tại sao lệch 250k" -- phat hien khoan ve
      // 23/07 (doanh thu 22/07) cua KH Cu bi lech dung bang chenh lech giua
      // phi co dinh 1,1% (dung thuc te cua KH Cu, xac nhan qua toan bo cac
      // ngay khac trong thang deu khop tuyet doi) va phi uoc tinh theo Nguon
      // tien cua parser ben duoi (chi dung cho KH Moi). Goc re: parser nay
      // duoc thu truoc cho MOI file .xlsx bat ke dang xem cong ty nao, nen 1
      // file "Transaction report" (co cot Nguon tien) tai len luc dang xem KH
      // Cu van bi parse va sinh netByCode, ghi de nham len phi co dinh dung
      // cua KH Cu.
      //
      // Luyen, 2026-07-25: "sao tôi không gửi đối soát momo kh cũ file này
      // lên đc vậy bạn dựa vào mã cửa hàng để đưa vô còn phí kh cũ cố định á
      // cộng lại số tiền là ra tổng á" -- file "Transaction report" thô KH Cu
      // co CUNG cau truc cot voi KH Moi (Thoi gian/So tien/Ma cua hang/Trang
      // thai/Nguon tien), nen dung LAI parseKhMoiFeeTransactionWorkbook de
      // doc giao dich, nhung resolve bang resolveKhCuFixedFeeTransactionGross
      // (phi CO DINH 1,1%, KHONG doi theo Nguon tien) thay vi
      // resolveKhMoiFeeTransactionGross. Tach ro 2 nhanh theo activeCompany
      // de KHONG bao gio ap phi bien doi cho KH Cu hay phi co dinh cho KH
      // Moi.
      let usedFeeParser = false;
      let feeUnmapped = [];
      let parsed;
      try {
        const { transactions } = parseKhMoiFeeTransactionWorkbook(req.file.buffer);
        const resolved =
          activeCompany === "kh_moi"
            ? resolveKhMoiFeeTransactionGross(transactions, store.cua_hang_mapping)
            : resolveKhCuFixedFeeTransactionGross(transactions, store.cua_hang_mapping);
        parsed = {
          sheetName:
            activeCompany === "kh_moi"
              ? "MoMo Transaction report (co phi theo nguon tien)"
              : "MoMo Transaction report (phi co dinh 1,1%)",
          dates: resolved.dates,
          codes: resolved.codes,
          grossByCode: resolved.grossByCode,
          netByCode: resolved.netByCode,
          cuaHangMap: {},
        };
        feeUnmapped = resolved.unmapped;
        usedFeeParser = true;
      } catch (feeParseErr) {
        parsed = parseTongMomoWorkbook(req.file.buffer);
      }

      store[momoCfg.grossKey].push({
        id: nextId(store, "momo_gross_uploads_seq") || Date.now(),
        uploaded_at: new Date().toISOString(),
        file_name: req.file.originalname,
        sheetName: parsed.sheetName,
        dates: parsed.dates,
        codes: parsed.codes,
        grossByCode: parsed.grossByCode,
        netByCode: parsed.netByCode,
      });
      seedGianMappingDefaults(store, parsed.codes);
      // Learn/refresh the Ma cua hang -> Ma Cong trinh mapping from this
      // upload so future raw MoMo portal zip exports can be resolved
      // automatically without needing this sheet uploaded again.
      Object.assign(store.cua_hang_mapping, parsed.cuaHangMap || {});
      save(store);
      if (usedFeeParser) {
        const feeDesc =
          activeCompany === "kh_moi"
            ? "tu dong tinh phi THAT theo nguon tien (Vi MoMo 1%, Vi tra sau 1.2%, con lai 0.3%)"
            : "tu dong tinh phi co dinh 1,1%";
        successMsg = `Da nap "${req.file.originalname}" (bao cao giao dich MoMo, ${parsed.dates[0]} - ${parsed.dates[parsed.dates.length - 1]}), ${feeDesc} cho ${parsed.codes.length} ma. Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.`;
        if (feeUnmapped.length > 0) {
          successMsg += ` CANH BAO: ${feeUnmapped.length} ma cua hang chua map duoc Ma Cong Trinh (${feeUnmapped.join(", ")}) -- doanh thu/phi cua cac ma nay dang hien o dong "CHUA MAP: ..." trong bang duoi, hay tai 1 file "Tong Momo Gop" co chua cac ma cua hang nay de he thong tu hoc mapping.`;
        }
      } else {
        successMsg = `Da nap "${parsed.sheetName}" (${parsed.dates[0]} - ${parsed.dates[parsed.dates.length - 1]}), ${parsed.codes.length} ma cong trinh/gian. Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.`;
      }
    }

    res.redirect("/doi-soat/momo?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-24: "sao vẫn bị cộng nhầm hóa đơn vậy" -- 5 hoa don
// 9977-9981 (khoan ve 2026-07-20, KVC ESTELLA/AE TAN AN KVC/FARM LOTTE PHAN
// THIET/FARM LOTTE NHA TRANG/TUTU MN AEON MALL TAN AN) da bi xoa 1 lan nhung
// TAI XUAT HIEN LAI sau khi Luyen tai len lai 1 file MTT cu hon (upload la
// CONG DON, khong biet ban da xoa la co chu y nen tai lai la them lai y het).
// Xac nhan qua file "MTT 24.07.xlsx" (moi nhat, 2026-07-24): ca 5 so HD nay
// THAT RA khong con gan tag "momo" nua (vd 9977 la "VietQR POSH MB 18,19",
// diem "GO BAC GIANG PHN", khong lien quan ESTELLA) -- day la tag CU/SAI tu 1
// phien ban MTT truoc do. Chan VINH VIEN 5 dong nay o day (khong chi xoa 1
// lan) de bat ky lan tai file MTT cu nao sau nay cung khong the tai them lai.
const MOMO_INVOICE_BLOCKLIST = new Set([
  "9977|2026-07-20|DIY ESTELLA KVC",
  "9978|2026-07-20|LM PHAN THIẾT KVC",
  "9979|2026-07-20|LM NHA TRANG KVC",
  "9980|2026-07-20|TUTU MN AEON MALL TÂN AN",
  "9981|2026-07-20|AE TAN AN KVC",
]);

// Upload 1 file danh sach hoa don (MTT) tren trang Momo cung cap nhat luon ca
// 3 danh sach hoa don Zalo/VNPay/Payoo (dung chung parseSharedInvoiceWorkbook
// voi trang /doi-soat/zvp) -- khong can upload lai file nay tren trang kia.
router.post("/doi-soat/momo/upload-hoadon", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const momoCfg = MOMO_CHANNELS[activeCompany];
  if (!store[momoCfg.invoicesKey]) store[momoCfg.invoicesKey] = [];
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const shared = parseSharedInvoiceWorkbook(req.file.buffer, activeCompany);

    const existingKeysMomo = new Set(store[momoCfg.invoicesKey].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
    // Luyen, 2026-08-14: dedup thu cap theo ngayHd|maDiem -- giong VietQR,
    // ngan upload file MTT luy ke moi (soHd cap lai) khoi tao invoice 2x.
    const existingDayDiemMomo = new Set(store[momoCfg.invoicesKey].map((i) => `${i.ngayHd}|${i.maDiem}`));
    let addedMomo = 0;
    let blockedMomo = 0;
    for (const inv of shared.momo) {
      const key = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
      if (existingKeysMomo.has(key)) continue;
      const key2 = `${inv.ngayHd}|${inv.maDiem}`;
      if (existingDayDiemMomo.has(key2)) continue;
      if (MOMO_INVOICE_BLOCKLIST.has(key)) {
        blockedMomo++;
        continue;
      }
      existingKeysMomo.add(key);
      existingDayDiemMomo.add(key2);
      store[momoCfg.invoicesKey].push(inv);
      addedMomo++;
    }

    // Luyen, 2026-08-03: hoa don "CHT nop tien" tag rieng trong cot ghi chu
    // (khong phai "Dich vu thu ho") -- luu tach rieng khoi invoicesKey, dung
    // cho tinh nang xuat Excel AMIS "CHT nop tien" (buildMomoChtDeposits).
    if (!store[momoCfg.chtInvoicesKey]) store[momoCfg.chtInvoicesKey] = [];
    const existingKeysCht = new Set(
      store[momoCfg.chtInvoicesKey].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`)
    );
    let addedCht = 0;
    for (const inv of shared.momoCht || []) {
      const key = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
      if (existingKeysCht.has(key)) continue;
      existingKeysCht.add(key);
      store[momoCfg.chtInvoicesKey].push(inv);
      addedCht++;
    }

    // Zalo/VNPay/Payoo (store.zvp_invoices) la cua rieng cong ty KH Cu (chua
    // co tai khoan ZVP nao cho KH Moi) -- chi gop vao cac kenh nay khi dang
    // xem KH Cu, tranh lay hoa don cua KH Moi gan nham vao so sach KH Cu.
    const addedCounts = { zalo: 0, vnpay: 0, payoo: 0 };
    if (activeCompany === "kh_cu") {
      if (!store.zvp_invoices) store.zvp_invoices = { zalo: [], vnpay: [], payoo: [] };
      for (const chKey of ["zalo", "vnpay", "payoo"]) {
        const existingKeys = new Set(store.zvp_invoices[chKey].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
        let added = 0;
        for (const inv of shared[chKey]) {
          const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
          if (existingKeys.has(k)) continue;
          existingKeys.add(k);
          store.zvp_invoices[chKey].push(inv);
          added++;
        }
        addedCounts[chKey] = added;
      }
    }

    save(store);
    let msg = `Da nap sheet "${shared.sheetName}": them moi ${addedMomo} HD momo (${momoCfg.label}).`;
    if (blockedMomo > 0) {
      msg += ` (Bo qua ${blockedMomo} HD da xac dinh la gan nham tag momo tu truoc, khong tinh lai.)`;
    }
    if (addedCht > 0) {
      msg += ` ${addedCht} HD "CHT nộp tiền" (dung cho xuat Excel AMIS tien mat CHT).`;
    }
    if (activeCompany === "kh_cu") {
      msg += ` ${addedCounts.zalo} HD zalo, ${addedCounts.vnpay} HD vnpay, ${addedCounts.payoo} HD payoo (da cap nhat cho ca 2 trang Doi soat Momo va Zalo/VNPay/Payoo).`;
    }
    msg += " Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.";
    res.redirect("/doi-soat/momo?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/momo/mapping", requireAdmin, (req, res) => {
  const store = load();
  const body = req.body || {};
  for (const [key, val] of Object.entries(body)) {
    if (key.startsWith("tkco_")) {
      const code = key.slice("tkco_".length);
      if (TKCO_VALUES.includes(val)) store.gian_mapping[code] = val;
    }
  }
  save(store);
  // Nhan, 2026-07-22: "nhấn vô dưới thì lại đưa lên trang đầu, muốn ở lại màn
  // hình đó luôn" -- bang TK Co khong lam thay doi du lieu doi soat hien tren
  // trang (chi luu tuy chinh xuat MISA), nen luu bang AJAX la an toan, khong
  // can load lai ca trang / mat vi tri cuon.
  if (req.get("X-Requested-With") === "XMLHttpRequest") {
    return res.json({ success: true });
  }
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent("Da luu bang TK Co theo gian."));
});

// ---------- Alias: "Ma diem" tren hoa don ghi ten khac (vd "SNOWFUN TAN
// PHU") nhung thuc chat cung 1 Ma Cong Trinh voi doanh thu (vd "AM TP KVCM")
// -- ap dung ngay luc doi soat, khong can tai lai file hoa don. Bang nay
// dung CHUNG voi trang doi-soat/zvp (store.invoice_diem_alias). ----------
router.post("/doi-soat/momo/diem-alias", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { sourceCode, targetCode } = req.body;
    if (!sourceCode || !targetCode) throw new Error("Thieu ma diem tren hoa don hoac ma cong trinh de anh xa.");
    if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
    store.invoice_diem_alias[sourceCode] = targetCode;
    save(store);
    res.redirect(
      "/doi-soat/momo?success=" + encodeURIComponent(`Da anh xa "${sourceCode}" -> "${targetCode}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/momo/diem-alias/delete", requireAdmin, (req, res) => {
  const store = load();
  const { sourceCode } = req.body;
  if (store.invoice_diem_alias) delete store.invoice_diem_alias[sourceCode];
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent("Da xoa anh xa ma diem."));
});

// ---------- Anh xa 1 ma cua hang MOI (chua tung xuat hien trong file "Tong
// Momo Gop" nao, nen he thong khong biet Ma Cong Trinh tuong ung -- hien len
// nhu "CHUA MAP: X" o muc 3a) ve dung Ma Cong Trinh. Luu vao
// store.cua_hang_mapping (KHONG theo companyKey) vi day la 1 danh muc
// gian/mat bang CHUNG, khong phai rieng cong ty nao -- ap dung ngay cho ca
// trang KH Cu va KH Moi, khong can tai lai file zip (xem applyCuaHangAlias).
router.post("/doi-soat/momo/cuahang-map", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { rawCode, targetCode } = req.body;
    if (!rawCode || !targetCode || !targetCode.trim()) {
      throw new Error("Thieu ma cua hang hoac ma cong trinh de anh xa.");
    }
    if (!store.cua_hang_mapping) store.cua_hang_mapping = {};
    const target = targetCode.trim();
    store.cua_hang_mapping[rawCode] = { maCongTrinh: target, code: target, gian: target };
    save(store);
    res.redirect(
      "/doi-soat/momo?success=" +
        encodeURIComponent(`Da anh xa ma cua hang "${rawCode}" -> "${target}" (ap dung cho ca 2 cong ty).`)
    );
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

// ---------- An gian khoi trang doi soat cua CONG TY DANG XEM (vd doanh thu
// cua 1 gian bi gop nham vao file "Tong Momo" cua cong ty nay, nhung gian do
// se duoc xuat HD ben cong ty KIA -- Luyen, 2026-07-16: "KVC ESTELLA",
// "FARM LOTTE NHA TRANG", "FARM LOTTE PHAN THIET", "AE TAN AN KVC" o KH
// Moi). CHI an khoi HIEN THI cua cong ty dang xem (store.gian_hidden theo
// company key) -- KHONG dong den store.gian_mapping (van dung chung), nen
// cong ty KIA neu dung chung ma nay van thay binh thuong. ----------
router.post("/doi-soat/momo/gian-hidden", requireAdmin, (req, res) => {
  const store = load();
  ensureGianHidden(store);
  const companyKey = getCompany(req);
  const { code } = req.body;
  if (code && !store.gian_hidden[companyKey].includes(code)) {
    store.gian_hidden[companyKey].push(code);
  }
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent(`Da an gian "${code}" khoi trang nay.`));
});

router.post("/doi-soat/momo/gian-hidden/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureGianHidden(store);
  const companyKey = getCompany(req);
  const { code } = req.body;
  store.gian_hidden[companyKey] = (store.gian_hidden[companyKey] || []).filter((c) => c !== code);
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent(`Da hien lai gian "${code}".`));
});

router.delete("/doi-soat/momo/upload-tong/:id", (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  store[momoCfg.grossKey] = (store[momoCfg.grossKey] || []).filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/momo");
});

router.post("/doi-soat/momo/upload-tong/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  store[momoCfg.grossKey] = (store[momoCfg.grossKey] || []).filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/momo");
});

router.post("/doi-soat/momo/invoices/clear", requireAdmin, (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  store[momoCfg.invoicesKey] = [];
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent(`Da xoa toan bo hoa don momo da nap (${momoCfg.label}).`));
});

// Luyen, 2026-08-14: xoa HD momo theo ngayHd chinh xac -- go trung lap khi
// file MTT tai len 2 lan voi soHd khac nhau cho cung ngay (invoice 2x).
// Body: { ngayHd: "2026-08-10" }  -- xoa cho cong ty dang xem (req company).
router.post("/doi-soat/momo/invoices/clear-by-ngayhd", requireAdmin, (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  try {
    const { ngayHd } = req.body || {};
    if (!ngayHd) throw new Error("Thieu ngayHd (vd: 2026-08-10).");
    const before = (store[momoCfg.invoicesKey] || []).length;
    store[momoCfg.invoicesKey] = (store[momoCfg.invoicesKey] || []).filter((i) => i.ngayHd !== ngayHd);
    const removed = before - store[momoCfg.invoicesKey].length;
    save(store);
    res.redirect(
      "/doi-soat/momo?success=" +
        encodeURIComponent(
          `Da xoa ${removed} hoa don momo ngay ${ngayHd} (${momoCfg.label}). Hay tai lai file MTT de nap lai dung.`
        )
    );
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

// Chi Nhan (2026-07-23): "đối chiếu từng ngày đi bị trùng á" -- Luyen phat
// hien 1 truong hop hoa don CU (vd so 9977-9981, xuat rieng cho tung ngay
// 18) van con luu trong he thong SAU KHI NCC xuat lai hoa don GOP thay the
// (vd so 10008-10012, "momo 18,19") -- file MTT moi nhat KHONG con nhung so
// cu nay nua (NCC da huy/thay), nhung upload chi CONG THEM dong moi, khong
// bao gio tu xoa dong cu, nen ban ghi cu nam lai vinh vien va bi cong TRUNG
// vao doi soat (dung logic loai hoa don gop trung o utils/momoReconcile.js
// deu ly ra la trung -- nhung ro rang van con hien Lech tren trang, co the
// vi ban ghi cu nay duoc tao TRUOC khi co logic loc do, hoac 1 truong hop
// bien the khac chua bat duoc). Thay vi phai sua tay file store.json, them
// nut xoa dung 1 so hoa don theo So HD, ap dung cho dung cong ty dang xem --
// dung khi phat hien 1 so HD cu/trung khong con dung nua (NCC da xuat lai).
router.post("/doi-soat/momo/invoices/xoa-theo-so", requireAdmin, (req, res) => {
    const store = load();
    const momoCfg = MOMO_CHANNELS[getCompany(req)];
    const raw = (req.body.soHdList || "").trim();
    if (!raw) {
          return res.redirect("/doi-soat/momo?error=" + encodeURIComponent("Chưa nhập số HĐ cần xoá."));
    }
    const targets = raw
      .split(/[,\n]+/)
      .map((s) => s.trim())
      .filter(Boolean)
      .map((s) => {
              const idx = s.indexOf("@");
              if (idx === -1) return { soHd: s, maDiem: null };
              return { soHd: s.slice(0, idx).trim(), maDiem: s.slice(idx + 1).trim() };
      });
    const before = (store[momoCfg.invoicesKey] || []).length;
    store[momoCfg.invoicesKey] = (store[momoCfg.invoicesKey] || []).filter(
          (i) =>
                  !targets.some(
                            (t) => String(i.soHd) === t.soHd && (t.maDiem === null || i.maDiem === t.maDiem)
                                    )
              );
    const removed = before - store[momoCfg.invoicesKey].length;
    save(store);
    const labels = targets.map((t) => (t.maDiem ? `${t.soHd}@${t.maDiem}` : t.soHd));
    res.redirect(
          "/doi-soat/momo?success=" +
            encodeURIComponent(`Đã xoá ${removed} hoá đơn (${momoCfg.label}) theo số HĐ: ${labels.join(", ")}.`)
        );
});

// Luyen, 2026-08-14: Sua HD thu cong cho momo -- luu override vao
// store.momo_manual_matches["settlementDate|code"] = {invoiceNumbers, amount}
// Tuong tu viet_qr_manual_matches nhung don gian hon (khong can grossAdjustment).
router.post("/doi-soat/momo/manual-match", requireDataEntry, (req, res) => {
  const store = load();
  if (!store.momo_manual_matches) store.momo_manual_matches = {};
  try {
    const { settlementDate, code, invoiceNumbers, amount, grossAdjustment, netAdjustment } = req.body;
    if (!settlementDate || !code) throw new Error("Thiếu settlementDate hoặc code.");
    const key = `${settlementDate}|${code}`;
    const invoiceList = (invoiceNumbers || "")
      .split(/[,\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const amt = amount ? Number(String(amount).replace(/[^\d]/g, "")) : null;
    const existing = store.momo_manual_matches[key] || {};
    const grossAdj = grossAdjustment !== undefined && grossAdjustment !== ""
      ? Number(String(grossAdjustment).replace(/[^\d\-]/g, "")) || 0
      : (existing.grossAdjustment || 0);
    const netAdj = netAdjustment !== undefined && netAdjustment !== ""
      ? Number(String(netAdjustment).replace(/[^\d]/g, "")) || null
      : (existing.netAdjustment !== undefined ? existing.netAdjustment : null);
    store.momo_manual_matches[key] = {
      invoiceNumbers: invoiceList,
      amount: amt,
      grossAdjustment: grossAdj,
      netAdjustment: netAdj,
      created_at: new Date().toISOString(),
    };
    save(store);
    res.redirect(
      "/doi-soat/momo?success=" + encodeURIComponent(`Đã sửa HĐ thủ công cho "${code}" ngày ${settlementDate}.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/momo/manual-match/delete", requireAdmin, (req, res) => {
  const store = load();
  if (!store.momo_manual_matches) store.momo_manual_matches = {};
  try {
    const { settlementDate, code } = req.body;
    delete store.momo_manual_matches[`${settlementDate}|${code}`];
    save(store);
    res.redirect("/doi-soat/momo?success=" + encodeURIComponent("Đã xóa sửa HĐ thủ công."));
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

// Export in the EXACT "Mau phieu thu tien gui de nhap vao AMIS Accounting"
// layout Luyen uses (sheet "MISATHUE123456" of DULIEUMOMO.xlsm): 28 named
// columns, one row per gian per settlement. Extra QC columns (invoice match
// status) are appended after column 28 for her own review and are safe to
// ignore/delete before importing into AMIS.
//
// Gian mapped to "SKIP" (khach hang moi chua len MISA) are excluded entirely
// from this export -- they still show up on the /doi-soat/momo review page
// so Luyen can keep an eye on them, just not written into the file she
// imports into AMIS.
router.get("/doi-soat/momo/export.xlsx", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const momoCfg = MOMO_CHANNELS[activeCompany];
  // Luyen, 2026-07-21: dong bo voi buildMomoReconciliation o tren -- so tien
  // ngan hang/settlement luon lay TK RIENG cua tung cong ty (momoCfg.bankName,
  // KHONG con muon cua KH Cu), con nguon doanh thu-theo-gian (grossKey) van
  // uu tien du lieu rieng cua KH Moi neu da co (xem momoSourceCfg).
  const sourceCfg = momoSourceCfg(activeCompany, store);
  const bank = store.banks.find((b) => b.name === momoCfg.bankName);
  if (!bank) return res.status(400).send(`Chua co ngan hang ${momoCfg.bankName}.`);

  const txs = store.transactions.filter((t) => t.bank_id === bank.id);
  const settlements = extractMomoSettlements(txs);
  const grossData = applyCuaHangAlias(mergeGross(store[sourceCfg.grossKey] || []), store.cua_hang_mapping || {});
  const invoiceData = { invoices: store[momoCfg.invoicesKey] || [] };
  let reconciled = reconcileMomo(settlements, grossData, invoiceData, store.gian_mapping, momoEffectiveDiemAlias(store));
  // An gian theo yeu cau rieng cua cong ty dang xuat (dung 1 danh sach voi
  // trang xem -- xem ensureGianHidden o tren), de khong xuat nham gian
  // thuoc ve cong ty kia vao file MISA cua cong ty nay.
  ensureGianHidden(store);
  const hiddenSet = new Set(store.gian_hidden[activeCompany] || []);
  if (hiddenSet.size > 0) {
    reconciled = reconciled.map((r) => ({ ...r, lines: r.lines.filter((l) => !hiddenSet.has(l.code)) }));
  }

  // Luyen, 2026-07-20: "chon t7 á xuat ra thang 7 cho toi thoi, xuat ra co
  // them thang 6 nua" -- truoc day file xuat LUON gom toan bo lich su (khong
  // loc theo thang dang xem tren trang), nay xuat dung THEO THANG dang chon
  // (query "month", cung 1 dropdown voi trang xem) neu co chon; khong chon
  // thang nao ("Tat ca") thi van xuat het nhu cu.
  //
  // Luyen, 2026-07-31: "tu ngay may toi ngay may" -- them loc chinh xac theo
  // ngay (query "tu"/"den", ISO yyyy-mm-dd) uu tien hon "month" neu co truyen
  // vao; "month" van duoc giu de tuong thich nguoc voi link/bookmark cu.
  const monthFilter = req.query.month || "";
  const tuFilter = req.query.tu || "";
  const denFilter = req.query.den || "";
  if (tuFilter || denFilter) {
    if (tuFilter) reconciled = reconciled.filter((r) => r.settlementDate >= tuFilter);
    if (denFilter) reconciled = reconciled.filter((r) => r.settlementDate <= denFilter);
  } else if (monthFilter) {
    reconciled = reconciled.filter((r) => r.settlementDate.slice(0, 7) === monthFilter);
  }

  // Luyen, 2026-08-15: "tan an co 2 hoa don ma xuat ra co 1 hoa don vay" --
  // phat hien export route KHONG ap dung momo_manual_matches (ghi de so HD thu
  // cong) trong khi trang xem /doi-soat/momo da ap dung (lines 620-633 trong
  // GET route phia tren). Ket qua: override invoiceNumbers cua AE TAN AN KVC
  // (["13409","13410"]) hien dung tren trang nhung file xuat chi ra ["13409"]
  // (raw tu reconcileMomo sau buoc loc AE-TAN-AN-KVC). Sua bang cach ap dung
  // cung 1 logic override o day truoc khi build rows.
  const momoManual = store.momo_manual_matches || {};
  reconciled = reconciled.map((r) => {
    const lines = r.lines.map((l) => {
      const mm = momoManual[`${r.settlementDate}|${l.code}`];
      if (!mm) return l;
      const updated = { ...l };
      if (mm.invoiceNumbers) updated.invoiceNumbers = mm.invoiceNumbers;
      if (mm.amount !== null && mm.amount !== undefined) {
        updated.invoiceTotal = mm.amount;
        updated.diff = updated.invoiceTotal - updated.gross;
        updated.matched = updated.invoiceNumbers.length > 0 && Math.abs(updated.diff) <= 1000;
      }
      updated.manualOverride = true;
      return updated;
    });
    return { ...r, lines };
  });

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;
  let seq = startNo;

  const rows = [];
  reconciled
    // pendingBank (Luyen, 2026-07-24): dong nay CHUA co giao dich ngan hang
    // that khop ngay -- chi hien de xem/theo doi truoc, KHONG dua vao file
    // xuat Misa cho den khi tien that ve ngan hang (tranh hach toan "thu tien
    // gui" truoc khi tien thuc su vao TK).
    .filter((r) => !r.pendingBank)
    .sort((a, b) => (a.settlementDate > b.settlementDate ? 1 : -1))
    .forEach((r) => {
      // BUG (phat hien 2026-07-20, Luyen: "dien giai luc nao cung vay, khong
      // co theo HD"): dong luc nay CHI loc theo tkCo !== "SKIP" bat ke dang
      // xuat cho cong ty nao, nen file xuat cua KH Moi lai xuat NHAM cac gian
      // cua KH Cu (tkCo thuong 131/1388) thay vi dung gian rieng cua KH Moi
      // (tkCo = "SKIP") -- KH Moi khong co hoa don nao khop voi ma cong trinh
      // cua KH Cu nen dong nao cung hien "Chua co HD", khong bao gio co
      // "theo HD ...". Sua giong dung logic trang xem /doi-soat/momo (dong
      // ~332): KH Moi xuat dong tkCo === "SKIP", KH Cu xuat dong con lai.
      const exportableLines = r.lines.filter((l) =>
        activeCompany === "kh_moi" ? l.tkCo === "SKIP" : l.tkCo !== "SKIP"
      );
      if (exportableLines.length === 0) return; // khong co dong nao thuoc cong ty dang xuat

      // Luyen, 2026-08-03: "Số tiền ngân hàng trả về á sao lại khi xuất ra nó
      // không khớp với nhau" -- phat hien qua vi du that: gian KVC ESTELLA
      // (HD 4081) hien 14.699.136đ tren trang xem nhung xuat Excel lai ra
      // 14.700.134đ (lech 998đ, dung bang phan lam tron ma trang xem da tu
      // dong hap thu vao gian LON NHAT trong khoan -- xem ROUNDING_ABSORB_
      // THRESHOLD/GET "/doi-soat/momo" o tren). Route xuat file truoc gio
      // KHONG ap dung buoc hap thu nay nen so tien xuat ra la so THO chua
      // dieu chinh, khong khop voi so tien NGAN HANG THUC NHAN (va khong
      // khop voi trang xem). Ap dung LAI dung 1 logic nay o day (chi kh_moi,
      // giong het dieu kien/nguong ben trang xem) de file xuat luon dung
      // bang tien thuc te ve ngan hang.
      if (activeCompany === "kh_moi" && !r.pendingBank && typeof r.bankAmount === "number") {
        let totalNetComputed = exportableLines.reduce((sum, l) => sum + l.net, 0);
        let diffVsBank = totalNetComputed - r.bankAmount;
        if (diffVsBank !== 0 && Math.abs(diffVsBank) <= ROUNDING_ABSORB_THRESHOLD && exportableLines.length > 0) {
          const biggest = exportableLines.reduce((a, b) => (b.gross > a.gross ? b : a), exportableLines[0]);
          biggest.net -= diffVsBank;
        }
      }

      const soCt = "NTTK" + String(seq).padStart(7, "0") + "/26";
      seq++;
      const ngayDmy = isoToDmy(r.settlementDate);
      exportableLines.forEach((l) => {
        const hdText = l.invoiceNumbers.length > 0 ? l.invoiceNumbers.join(", ") : "";
        // Hau to kenh (Luyen, 2026-07-16): "- MM" de phan biet doanh thu Momo
        // voi cac kenh khac (VNPay/Zalo/Payoo la "- VNP"/"- PAYOO", Viet QR la
        // "- QR") ngay tren file xuat Misa, dat truoc "theo HD ..." neu co.
        const dienGiai = hdText
          ? `Thu tiền dịch vụ vui chơi giải trí - MM theo HĐ ${hdText}`
          : "Thu tiền dịch vụ vui chơi giải trí - MM";
        rows.push({
          "Ngày hạch toán (*)": ngayDmy,
          "Ngày chứng từ (*)": ngayDmy,
          "Số chứng từ (*)": soCt,
          // Luyen, 2026-07-20: "ma doi tuong cho momo la TRUC TUYEN0305289153"
          // -- ap dung cho ca KH Cu va KH Moi (khac voi cac kenh khac van de
          // "KL"/Khach le nhu cu).
          "Mã đối tượng": "TRỰC TUYẾN0305289153",
          "Tên đối tượng": "",
          "Địa chỉ": "",
          "Nộp vào TK": momoCfg.bankAccount,
          "Mở tại ngân hàng": momoCfg.bankFullName,
          "Lý do thu": "Thu tiền khách hàng (không theo hóa đơn)",
          "Diễn giải lý do thu": dienGiai,
          "Mã nhân viên thu": "",
          "Diễn giải (hạch toán)": dienGiai,
          "TK Nợ (*)": 1121,
          // Luyen, 2026-08-03: "tk có là 131 hết á" -- "SKIP" chi la 1 co
          // hieu NOI BO danh dau "gian nay thuoc rieng KH Moi" (xem loc
          // exportableLines o tren), KHONG phai 1 tai khoan ke toan that --
          // truoc gio bi ghi THANG chu "SKIP" vao chinh cot TK Co cua file
          // xuat MISA (sai, AMIS khong hieu duoc). Doi lai thanh "131" (dung
          // nhu cac dong con lai) khi ghi ra file, chi giu "SKIP" cho logic
          // loc/an noi bo.
          "TK Có (*)": l.tkCo === "SKIP" ? "131" : l.tkCo,
          "Số tiền": l.net,
          "Mã đối tượng (hạch toán)": "TRỰC TUYẾN0305289153",
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
          "Doanh thu gộp (trước phí)": l.gross,
          "Chênh lệch HĐ vs doanh thu": l.diff,
          "Trạng thái": l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : l.matched ? "Khớp" : "Lệch",
        });
      });
    });

  // Luyen, 2026-08-03: "ngân hàng có các giao dịch của Cửa Hàng trưởng nộp
  // tiền vào... đối tượng là khách lẻ số tiền là số tiền nợ 1121 có lúc nào
  // cx 131 cho tôi nhá mã công trình là gian đó mỗi 1 ngày là 1 chứng từ có
  // số hóa đơn nữa hôm sau sẽ xuất cho hôm trước gắn vô mẫu á để chỗ diễn
  // giải á là Thu tiền dịch vụ vui chơi giải trí - KL theo HĐ rồi găn số hóa
  // đơn vô nhá" -- thay the hoan toan store.momo_kl_deposits cu (hand-
  // populated 1 lan, chi co 17 dong 2026-07-03->07-20, xem __add_kl.js) bang
  // buildMomoChtDeposits: quet TRUC TIEP tu giao dich ngan hang (ma nop tien
  // trong mo ta -> gian qua store.cht_nop_tien_map) roi tu khop voi hoa don
  // "CHT nộp tiền N" (tag rieng trong cot ghi chu cua sheet danh sach hoa don
  // dung chung, xem parseInvoiceWorkbook/momoReconcile.js) theo CUNG ngay
  // ngan hang nhan tien + gian -- khong con phai tay them tay tung dong nua,
  // tu dong cap nhat khi Luyen nap them sao ke/hoa don moi.
  const chtDeposits = buildMomoChtDeposits(store, activeCompany, momoCfg.chtInvoicesKey).filter((d) => {
    if (tuFilter || denFilter) {
      if (tuFilter && d.date < tuFilter) return false;
      if (denFilter && d.date > denFilter) return false;
      return true;
    }
    return !monthFilter || d.date.slice(0, 7) === monthFilter;
  });
  chtDeposits.forEach((d) => {
    const soCt = "NTTK" + String(seq).padStart(7, "0") + "/26";
    seq++;
    const ngayDmy = isoToDmy(d.date);
    const dienGiai = d.soHd
      ? `Thu tiền dịch vụ vui chơi giải trí - KL theo HĐ ${d.soHd}`
      : "Thu tiền dịch vụ vui chơi giải trí - KL";
    rows.push({
      "Ngày hạch toán (*)": ngayDmy,
      "Ngày chứng từ (*)": ngayDmy,
      "Số chứng từ (*)": soCt,
      // Luyen, 2026-07-21: "7701 á nếu là KL cửa hàng trưởng nạp vô á thì
      // là đối tượng KL luôn nhá" -- cac khoan nay la CHT (cua hang truong)
      // nop tien mat truc tiep, KHONG phai settlement Momo that, nen dung
      // ma doi tuong "KL" (Khach le, mac dinh cho cac kenh khac -- xem ghi
      // chu tai buildExportRows/routes/doisoat-zvp.js), KHONG dung
      // "TRUC TUYEN0305289153" (chi danh rieng cho doanh thu Momo settlement).
      "Mã đối tượng": "KL",
      "Tên đối tượng": "",
      "Địa chỉ": "",
      "Nộp vào TK": momoCfg.bankAccount,
      "Mở tại ngân hàng": momoCfg.bankFullName,
      "Lý do thu": "Thu tiền khách hàng (không theo hóa đơn)",
      "Diễn giải lý do thu": dienGiai,
      "Mã nhân viên thu": "",
      "Diễn giải (hạch toán)": dienGiai,
      "TK Nợ (*)": 1121,
      "TK Có (*)": "131",
      "Số tiền": d.amount,
      "Mã đối tượng (hạch toán)": "KL",
      "Số khế ước đi vay": "",
      "Số khế ước cho vay": "",
      "Mã khoản mục chi phí": "",
      "Mã đơn vị": "",
      "Mã đối tượng THCP": "",
      "Mã công trình": d.gian,
      "Số đơn đặt hàng": "",
      "Số đơn mua hàng": "",
      "Số hợp đồng mua": "",
      "Số hợp đồng bán": "",
      "Mã thống kê": "",
      "CP không hợp lý": "",
      "Số HĐ khớp": d.soHd,
      "Tổng tiền HĐ khớp": d.invoiceTotal,
      "Doanh thu gộp (trước phí)": d.amount,
      "Chênh lệch HĐ vs doanh thu": d.invoiceTotal - d.amount,
      "Trạng thái": !d.soHd ? "Chưa có HĐ (CHT nộp tiền)" : d.matched ? "Khớp (CHT nộp tiền)" : "Lệch (CHT nộp tiền)",
    });
  });

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, momoCfg.exportSheet);
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", `attachment; filename=doi-soat-momo-${momoCfg.exportSheet}.xlsx`);
  res.send(buf);
});

// Exposed so routes/dashboard.js (Tong quan / Cong no) can reuse the exact
// same reconciliation this page shows, without a second implementation.
router.buildMomoReconciliation = buildMomoReconciliation;
router.MOMO_BANK_NAME = MOMO_BANK_NAME;

module.exports = router;
