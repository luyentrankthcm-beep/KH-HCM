const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const {
  extractZvpSettlements,
  parseInvoiceWorkbookByTag,
  parseOfflineVnpayWorkbook,
  parseDiemMappingSheet,
  parsePayooWorkbook,
  parsePayooRawReport,
  parsePayooDiemMapping,
  parseOnlineVnpayWorkbook,
  parseGianXuatHdSheet,
  parseProductCatalogSheet,
  resolveProductCatalogGross,
  applyGianRedirectToInvoices,
  applyGianRedirectToResolvedGross,
  learnGianFromInvoices,
  parseGianMasterSheet,
  mergeGianListWithMaster,
  mergeDiemMapWithMaster,
  parseOrderDetailsWorkbook,
  parseFeeReportWorkbook,
  parseVnpayOfflineFeeReport,
  parseSharedInvoiceWorkbook,
  resolveDiemGross,
  resolveOnlineGross,
  parseOnlineProductMapSheet,
  mergeOnlineProductMap,
  resolveOnlineGrossByProductMap,
  reconcileZvp,
  isoToDmy,
  normText,
  normCode,
  FF_SUFFIX,
} = require("../utils/zvpReconcile");
const { migrateVnpayKhMoiInvoices } = require("../utils/vietqrReconcile");

const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

// Luyen, 2026-07-31: mac dinh 2 o "Tu ngay/Den ngay" cua nut Xuat file MISA
// (export-online.xlsx/export-offline.xlsx) = dau/cuoi thang dang chon o
// dropdown "Chon thang" tren trang xem, giong het ben Momo/VietQR.
function monthBounds(m) {
  if (!m) return { first: "", last: "" };
  const [y, mo] = m.split("-").map(Number);
  const lastDay = new Date(y, mo, 0).getDate();
  return { first: `${m}-01`, last: `${m}-${String(lastDay).padStart(2, "0")}` };
}

// Doi soat Zalo App/VNPay/Payoo hien chi ap dung cho KH Cu (KH Moi chua co
// tai khoan nhan tien Zalo/VNPay/Payoo) -- chan truy cap truc tiep (vd bookmark
// hoac go URL tay) khi dang xem KH Moi, dieu huong ve trang chu.
// QUAN TRONG: phai gan middleware nay voi tien to "/doi-soat/zvp" (khong phai
// router.use(fn) khong co duong dan) -- router nay duoc mount o "/" trong
// server.js, nen mot middleware khong loc duong dan se chan NHAM ca cac
// route khac duoc mount SAU no (vd /doi-soat/vietqr, /bao-cao, /cong-no),
// giong loi da gap voi companyRoutes/requireLogin truoc day.
router.use("/doi-soat/zvp", (req, res, next) => {
  if (getCompany(req) === "kh_moi") {
    return res.redirect("/");
  }
  next();
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

const ZVP_BANK_NAME = "ACB31268"; // TK 12131268, tai khoan nhan tien Zalo App / VNPay / Payoo
const ZVP_BANK_ACCOUNT = 12131268;
const ZVP_BANK_FULLNAME = "Ngân hàng TMCP Á Châu";

const TKCO_VALUES = ["131", "1388", "SKIP"];

const UPDATED_NOTE = " Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.";

// Same overlap-safe merge as Momo's mergeGross: sort uploads oldest-to-newest
// and let the newest upload win for any (ngay, code) key it covers, so
// re-uploading a corrected file for a period supersedes older data instead
// of stacking on top of it and double-counting revenue.
function mergeResolvedGross(uploads) {
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  const sorted = [...uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  for (const u of sorted) {
    (u.codes || []).forEach((c) => { if (c) codes.add(c); });
    for (const [k, v] of Object.entries(u.grossByCode || {})) {
      // Luyen, 2026-08-10: bo qua entry co ma cong trinh TRONG hoac whitespace
      // -- xay ra khi san pham/diem chua duoc map lan dau, sau do duoc seed/map
      // dung vao upload moi hon. Entry cu van con key "ngay|" (rong) trong
      // store; neu khong bo qua, no hien thanh dong blank tren trang doi soat
      // va lam tang tong sai ngay ca khi da upload lai du dung.
      const code = k.includes("|") ? k.slice(k.indexOf("|") + 1) : k;
      if (!code || !code.trim()) continue;
      grossByCode[k] = v;
    }
    for (const [k, v] of Object.entries(u.netByCode || {})) {
      const code = k.includes("|") ? k.slice(k.indexOf("|") + 1) : k;
      if (!code || !code.trim()) continue;
      netByCode[k] = v;
    }
  }
  return { codes: Array.from(codes), grossByCode, netByCode };
}

// Chi Nhan, 2026-07-31: "cứ hiện thị cái ch xuất phải chuyển qua 131 rồi lưu
// lất qua lại vẫn hiện của kh mưới là sao" -- store.gian_mapping (Momo) va
// store.zvp_gian_mapping (day) TUNG la 1 bang duy nhat dung chung, khien
// ensureGianHidden ben Momo (ep 4 gian KH Moi ve SKIP MOI LAN mo trang Momo)
// de luon len ca ZVP cho cung ten gian -- Chi Nhan doi lai 131 ben ZVP xong
// quay lai la bi ghi de mat. Tach hoan toan tu day: zvp_gian_mapping la bang
// RIENG cua trang nay, khoi tao 1 LAN DUY NHAT bang cach sao chep gian_mapping
// hien co (giu lai cac TK Co dang dung dung), roi hoan toan doc lap ve sau --
// Momo sua gian_mapping KHONG con anh huong ZVP nua va nguoc lai.
function ensureZvpGianMapping(store) {
  if (!store.zvp_gian_mapping) {
    store.zvp_gian_mapping = Object.assign({}, store.gian_mapping || {});
    return true; // doi -- caller nen save(store)
  }
  return false;
}

// New codes seen for the first time on any channel default to "1388" if
// they're the CSE/chia-se half (the "__FF" suffixed code), otherwise "131" --
// exactly the same convention as Momo's seedGianMappingDefaults. Rieng bang
// zvp_gian_mapping cua trang nay (xem ensureZvpGianMapping o tren).
function seedGianMappingDefaults(store, codes) {
  ensureZvpGianMapping(store);
  codes.forEach((c) => {
    if (!(c in store.zvp_gian_mapping)) {
      // Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- TK Co 1388
      // (doanh thu chia se/CSE) khong con duoc dung nua, moi gian moi deu
      // mac dinh 131.
      store.zvp_gian_mapping[c] = "131";
    }
  });
}

// Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het".
function ensureNo1388(store) {
  ensureZvpGianMapping(store);
  let changed = false;
  for (const code of Object.keys(store.zvp_gian_mapping)) {
    if (store.zvp_gian_mapping[code] === "1388") {
      store.zvp_gian_mapping[code] = "131";
      changed = true;
    }
  }
  return changed;
}

// Rows from the daily master "gian " sheet (store.zvp_gian_master), filtered
// to the ones belonging to a given channel via its "Thuoc" column -- e.g.
// "zalo mini app" for Online, "vnpay co so" for Offline, "payoo" for Payoo
// (matches both "Payoo QR ..." and "Payoo the ..."). Returns [] if no master
// has been uploaded yet or nothing matches, so callers can merge safely
// without any extra null-checking.
function getMasterRowsForChannel(store, channelNeedle) {
  const rows = (store.zvp_gian_master && store.zvp_gian_master.rows) || [];
  return rows.filter((r) => normText(r.thuoc).includes(channelNeedle));
}

// Guard against the same file being submitted twice in quick succession
// (same double-submit issue observed and fixed for Momo uploads).
function isDuplicateRecentUpload(uploadsList, fileName, grossByCode) {
  const recent = uploadsList[uploadsList.length - 1];
  if (!recent) return false;
  if (recent.file_name !== fileName) return false;
  const ageMs = Date.now() - new Date(recent.uploaded_at).getTime();
  if (ageMs > 2 * 60 * 1000) return false;
  return JSON.stringify(recent.grossByCode) === JSON.stringify(grossByCode);
}

// Chi Nhan, 2026-07-29: "hóa đơn có 2 cái này thôi ... sao lại cộng hóa đơn
// 11143 vậy lấy ra đi kh phải của VNPAY offline á" -- zvp_gian_list (bang
// "gian" chi Nhan tu tai len) co dong "KVC AE HUE" -> "AE HUE KVCN" khien MOI
// hoa don cua 1 chuoi hoa don HOAN TOAN KHAC (dat ten "KVC AE HUE"/"Tàu AE
// Huế", so hoa don rieng 7700-11000+, vd 10413/11143) bi gop nham vao gian
// "AE HUE KVCN" that -- vi khong co doanh thu ngan hang nao khop voi ten "KVC
// AE HUE" ca, gop hoa don vao lam MOI ngay co phat sinh loai hoa don nay deu
// bi "Lệch" dung bang so tien hoa don do (xac nhan qua nhieu ngay: 07-23
// (2.114.000), 07-24 (1.620.000), 07-28/29 (11143 = 1.428.000)). Da xoa dong
// nay 1 lan truc tiep nhung bi MAT lai (server cua chi Nhan tu ghi de
// store.json bang ban cu dang giu trong bo nho no, xem ghi chu day du tai
// INVOICE_DIEM_ALIAS_DEFAULTS trong routes/doisoat-vietqr.js) -- sua han qua
// code o day (chay lai + tu xoa MOI LAN load(), giong co che GIAN_MERGE_DEFAULTS/
// TEN_DIEM_MASTER_DEFAULTS ben VietQR) de KHONG BI MAT nua, chi can chi Nhan
// restart server 1 lan de nap code moi la vinh vien khong con bug nay.
const ZVP_GIAN_LIST_BAD_REDIRECTS = [{ tenDiem: "KVC AE HUE", maCongTrinh: "AE HUE KVCN" }];

// Luyen, 2026-08-08: file Payoo KH Cu "PY-GiaoDichBanHangPayoo-..." chi chua 2
// ma cua hang: DVGIAITRIKH_FZ_IPH (da co) va DVGIAITRIKH_FARM_VCTIMECITY (moi).
// Seed o day de Railway tu dong co entry khi server restart, khong can tai lai
// file "Payoo thu ho" chu? de cap nhat.
const PAYOO_DIEM_MAP_DEFAULTS = {
  "DVGIAITRIKH_FZ_IPH": { maCongTrinh: "FUNZONE IPH KVCN", isCse: false },
  "DVGIAITRIKH_FARM_VCTIMECITY": { maCongTrinh: "Farm Times City", isCse: false },
};

function seedPayooDiemMapDefaults(store) {
  if (!store.zvp_payoo_diem_map) store.zvp_payoo_diem_map = {};
  let changed = false;
  Object.entries(PAYOO_DIEM_MAP_DEFAULTS).forEach(([chiNhanh, val]) => {
    if (!store.zvp_payoo_diem_map[chiNhanh]) {
      store.zvp_payoo_diem_map[chiNhanh] = val;
      changed = true;
    }
  });
  return changed;
}

// Chi Nhan, 2026-07-30: "là hóa đơn này nè đổi tên á bạn coi lại nha lần sao
// nó note z á" -- hoa don 2458 (450.000d, ngay 29/7, gian KVC TIMES that)
// dan Ma diem "VC TC DIY KVCN" (doi ten tu "GHOST BRIDE" cu, gach bo ngay
// tren file goc) thay vi "KVC TIMES" nen bao "Chua co HD" gia. Ghi thang 1
// lan qua script bi MAT ngay lap tuc (server dang chay cua Chi Nhan tu ghi
// de store.json bang ban cu dang giu trong bo nho, giong het co che da giai
// thich tai INVOICE_DIEM_ALIAS_DEFAULTS trong routes/doisoat-vietqr.js) --
// chuyen thanh seed tu dong CHAY LAI + GHI DE MOI LAN load() (giong het
// stripZvpGianListBadRedirects ngay ben duoi) de khong bao gio mat lai.
const ZVP_INVOICE_DIEM_ALIAS_DEFAULTS = {
  "VC TC DIY KVCN": "KVC TIMES",
};

// Luyen, 2026-08-10: san pham moi tu combo upload OrderDetails 10/08/2026 --
// tu dong them vao zvp_online_product_map voi maCongTrinh dung nen khong con
// hien dong trong/blank row tren trang doi soat ZVP nua.
const ZVP_ONLINE_PRODUCT_MAP_EXTRA_DEFAULTS = {
  // Luyen, 2026-08-10: combo upload OrderDetails 10/08/2026
  "❄️ BÌNH DƯƠNG - SALE 20% (chưa bao gồm tất) - SNOW FUN ❄️": { maCongTrinh: "AM BD KVCM", isCse: false },
  // Luyen, 2026-08-10: xac nhan tu doi soat Zalo App KH Cu 08-08 + 09-08
  "AEON BÌNH DƯƠNG - SALE 20% VÉ NHÀ MA ÂM PHỦ":              { maCongTrinh: "AM BD KVCM", isCse: false },
  "❄️HUẾ - COMBO 3 VÉ TẶNG 2 VÉ - SNOW FUN❄️":               { maCongTrinh: "AE HUE KVCN", isCse: false },
  "🌸 SC VivoCity -  Combo 10 vé - Funzone Adventure 🌸":      { maCongTrinh: "SC VIVO KVCM", isCse: false },
  "🌸 SC VivoCity -  Combo 5 vé - Funzone Adventure 🌸":       { maCongTrinh: "SC VIVO KVCM", isCse: false },
  "🔥 THE LOOP (IPH) - FUNZONE - COMBO 05 VÉ 🔥":             { maCongTrinh: "FUNZONE IPH KVCN", isCse: false },
  "LOTTE NHA TRANG - COMBO 10 VÉ - ECOKIDS FARM":              { maCongTrinh: "LM NHA TRANG KVC", isCse: false },
};

function seedOnlineProductMapDefaults(store) {
  if (!store.zvp_online_product_map) store.zvp_online_product_map = {};
  let changed = false;
  Object.keys(ZVP_ONLINE_PRODUCT_MAP_EXTRA_DEFAULTS).forEach((k) => {
    const def = ZVP_ONLINE_PRODUCT_MAP_EXTRA_DEFAULTS[k];
    // Chi seed neu entry chua co HOAC co nhung maCongTrinh dang de trong
    const existing = store.zvp_online_product_map[k];
    if (!existing || !(existing.maCongTrinh || "").trim()) {
      store.zvp_online_product_map[k] = def;
      changed = true;
    }
  });
  return changed;
}

function seedZvpInvoiceDiemAliasDefaults(store) {
  if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
  let changed = false;
  Object.keys(ZVP_INVOICE_DIEM_ALIAS_DEFAULTS).forEach((k) => {
    if (store.invoice_diem_alias[k] !== ZVP_INVOICE_DIEM_ALIAS_DEFAULTS[k]) {
      store.invoice_diem_alias[k] = ZVP_INVOICE_DIEM_ALIAS_DEFAULTS[k];
      changed = true;
    }
  });
  return changed;
}

// Luyen, 2026-07-31: "mã công trình Kh cũ đổi từ Nhà ma Go BÀ rịa thành AE GO
// BA RIA KVC , FARM LOTTE PHAN THIET thành LM PHAN THIET KVC, KVC TIMES
// thành Farm Times City đối với Kh cũ nhá" -- doi TEN CUOI CUNG (Ma cong
// trinh) cho 3 gian KH Cu, ap dung NGUOC (retroactive) cho CA doanh thu da
// resolve tu truoc (zvp_online/offline/payoo_uploads) lan hoa don da tai,
// dung dung co che redirect da co san: them 1 hop redirect vao zvp_gian_list
// (tenDiem = ma CU, dung nhu 1 "ten dien" duoc nhan dien, maCongTrinh = ma
// MOI) -- xem applyGianRedirectToResolvedGross/applyGianRedirectToInvoices
// trong utils/zvpReconcile.js, giong het cach da vá case "GHOST BRIDE MEGA DA
// NANG" -> "GHOST BRIDE AE HUE" -> "AE HUE KVCN" truoc day (redirect nhieu
// buoc, KHONG xoa dong goc trong zvp_gian_list vi dong goc van dung de nhan
// dien tenDiem THAT tren hoa don/sheet gian). Kem theo don sach
// invoice_diem_alias/zvp_gian_mapping/zvp_manual_matches dang tro vao ma CU
// de khong bi mo coi/vong lap khong can thiet. Seed lai + ghi de MOI LAN
// load() (giong moi hang so *_DEFAULTS khac trong file nay) de khong bi mat
// khi server restart.
// Luyen, 2026-08-17: "AE GO BA RIA KVC doi thanh EVMN GHOST MN GO BA RIA,
// LM PHAN THIET KVC doi thanh FARM LOTTE PHAN THIET, AE BAC GIANG KVCN doi
// thanh FARM LOTTE BAC GIANG -- chi KH cu" -- cap nhat lai cac ma cong trinh
// cu thanh ten moi MISA nhan ra. Luu y FARM LOTTE PHAN THIET tung la ten GIA
// TRANG (old) truoc khi doi thanh LM PHAN THIET KVC; nay doi nguoc lai ve
// dung ten thuc trong MISA.
const ZVP_GIAN_CODE_RENAMES = [
  // NHA MA GO BA RIA -> EVMN GHOST MN GO BA RIA (cap nhat tu trung gian cu)
  { from: "NHA MA GO BA RIA",   to: "EVMN GHOST MN GO BÀ RỊA" },
  { from: "AE GO BA RIA KVC",   to: "EVMN GHOST MN GO BÀ RỊA" },
  // LM PHAN THIET KVC -> FARM LOTTE PHAN THIET (ten dung trong MISA)
  { from: "LM PHAN THIET KVC",  to: "FARM LOTTE PHAN THIET" },
  // AE BAC GIANG KVCN -> FARM LOTTE BAC GIANG
  { from: "AE BAC GIANG KVCN", to: "FARM LOTTE BAC GIANG" },
  { from: "LM NHA TRANG KVC",  to: "FARM LOTTE NHA TRANG" },
  { from: "KVC TIMES",          to: "Farm Times City" },
];

// Cac hop redirect cu da bi thay the boi ZVP_GIAN_CODE_RENAMES moi o tren;
// can xoa khoi store.zvp_gian_list de tranh mau thuan/vong lap 1 buoc.
const ZVP_GIAN_CODE_RENAMES_CLEANUP = [
  { tenDiem: "NHA MA GO BA RIA",    maCongTrinh: "AE GO BA RIA KVC" },
  { tenDiem: "FARM LOTTE PHAN THIET", maCongTrinh: "LM PHAN THIET KVC" },
];

function seedZvpGianCodeRenames(store) {
  let changed = false;
  if (!Array.isArray(store.zvp_gian_list)) store.zvp_gian_list = [];

  // Xoa cac redirect cu truoc khi them moi (tranh chain/vong lap)
  ZVP_GIAN_CODE_RENAMES_CLEANUP.forEach(({ tenDiem, maCongTrinh }) => {
    const before = store.zvp_gian_list.length;
    store.zvp_gian_list = store.zvp_gian_list.filter(
      (g) => !(normText(g.tenDiem) === normText(tenDiem) && g.maCongTrinh === maCongTrinh)
    );
    if (store.zvp_gian_list.length !== before) changed = true;
  });

  ZVP_GIAN_CODE_RENAMES.forEach(({ from, to }) => {
    // 1) Hop redirect chinh: ma CU (dung nhu 1 tenDiem) -> ma MOI.
    const already = store.zvp_gian_list.some((g) => normText(g.tenDiem) === normText(from) && g.maCongTrinh === to);
    if (!already) {
      const oldEntry = store.zvp_gian_list.find((g) => g.maCongTrinh === from);
      store.zvp_gian_list.push({ tenDiem: from, maCongTrinh: to, isCse: oldEntry ? !!oldEntry.isCse : false });
      changed = true;
    }

    // 2) invoice_diem_alias dang tro thang vao ma CU -> tro thang sang ma
    // MOI luon (tranh vong lap 2 buoc khong can thiet qua ma CU).
    if (store.invoice_diem_alias) {
      Object.keys(store.invoice_diem_alias).forEach((k) => {
        if (store.invoice_diem_alias[k] === from) {
          store.invoice_diem_alias[k] = to;
          changed = true;
        }
      });
    }

    // 3) TK Co rieng cua trang ZVP (zvp_gian_mapping) -- giu nguyen TK Co da
    // tung chon cho ma CU thay vi de ma MOI roi tu roi ve mac dinh 131.
    if (store.zvp_gian_mapping && store.zvp_gian_mapping[from] !== undefined && store.zvp_gian_mapping[to] === undefined) {
      store.zvp_gian_mapping[to] = store.zvp_gian_mapping[from];
      changed = true;
    }

    // 4) Xac nhan thu cong "da co HD" (zvp_manual_matches, key "ngay|ma") --
    // doi key sang ma MOI de khong bi mo coi.
    if (store.zvp_manual_matches) {
      ["online", "offline", "payoo"].forEach((ch) => {
        const bucket = store.zvp_manual_matches[ch];
        if (!bucket) return;
        Object.keys(bucket).forEach((key) => {
          const sep = key.indexOf("|");
          if (sep < 0) return;
          const date = key.slice(0, sep);
          const code = key.slice(sep + 1);
          const codeBase = code.endsWith("__FF") ? code.slice(0, -4) : code;
          if (codeBase !== from) return;
          const suffix = code.endsWith("__FF") ? "__FF" : "";
          const newKey = date + "|" + to + suffix;
          if (bucket[newKey] === undefined) {
            bucket[newKey] = bucket[key];
            delete bucket[key];
            changed = true;
          }
        });
      });
    }
  });

  return changed;
}

// Chi Nhan, 2026-07-30: "sao ngân hàng trả ngày 16/07 á trả của ngày 15 á nó
// có 19tr mấy sao bạn cộng lên mấy trăm triệu dữ vậy" -- khoan ve ngay 16/7
// (doanh thu 15/7) kenh Offline (VNPay QR OFFLINE) hien Ngan hang 125.766.363d
// trong khi Tinh tu du lieu tai len (khop dung hoa don) chi 19.957.812d. Dieu
// tra: TK ACB31268 co 2 dong giao dich CUNG mo ta y het "...VNPAY TT
// 829168...DV QR OFFLINE NGAY 15.07.26" nhung khac so tien -- 1 dong nhap
// 17/7 (nguoi nhap "Ga Cute") ghi 105.808.551d (KHONG khop bat ky doanh thu/
// hoa don nao), va 1 dong nhap lai sau do ngay 20/7 (Quan tri vien K&H) ghi
// dung 19.957.812d (khop tuyet doi voi tong hoa don ngay do) -- rat co the
// nhap tay 17/7 bi sai/nhamg so, sau do co nguoi nhap lai dung nhung quen xoa
// dong cu, khien 2 dong cong don lai. Chi Nhan xac nhan xoa dong sai (dua
// theo sao ke: so dung la 19.957.812d).
function removeDuplicateOfflineTx20260716(store) {
  const bank = store.banks.find((b) => b.name === ZVP_BANK_NAME);
  if (!bank) return false;
  const before = store.transactions.length;
  store.transactions = store.transactions.filter(
    (t) =>
      !(
        t.bank_id === bank.id &&
        t.date === "2026-07-16" &&
        t.type === "thu" &&
        t.amount === 105808551 &&
        (t.description || "").includes("DV QR OFFLINE NGAY 15.07.26")
      )
  );
  return store.transactions.length !== before;
}

function stripZvpGianListBadRedirects(store) {
  if (!store.zvp_gian_list || !Array.isArray(store.zvp_gian_list)) return false;
  let changed = false;
  store.zvp_gian_list = store.zvp_gian_list.filter((g) => {
    const isBad = ZVP_GIAN_LIST_BAD_REDIRECTS.some(
      (bad) => normText(g.tenDiem) === normText(bad.tenDiem) && g.maCongTrinh === bad.maCongTrinh
    );
    if (isBad) changed = true;
    return !isBad;
  });
  return changed;
}

function buildReconciliation(store) {
  // Cac save() nay la "housekeeping" (tu dong don dep du lieu), khong phai
  // do user yeu cau -- neu Volume day (ENOSPC) save() co the throw, lam crash
  // trang. Wrap trong try-catch de app tiep tuc hien thi du lieu (trong bo nho
  // van dung) ngay ca khi chua ghi duoc ra dia.
  function trySave() {
    try { save(store); } catch (e) {
      console.error("[zvp] housekeeping save() that bai (co the Volume day):", e.message);
    }
  }
  if (ensureNo1388(store)) trySave();
  if (seedPayooDiemMapDefaults(store)) trySave();
  if (stripZvpGianListBadRedirects(store)) trySave();
  if (seedOnlineProductMapDefaults(store)) trySave();
  if (seedZvpInvoiceDiemAliasDefaults(store)) trySave();
  if (seedZvpGianCodeRenames(store)) trySave();
  if (removeDuplicateOfflineTx20260716(store)) trySave();
  // Chi Nhan, 2026-07-29: xem ghi chu day du tai VNPAY_KHMOI_INVOICE_MADIEM_MAP
  // trong utils/vietqrReconcile.js -- hoa don KVC AE HUE/KVC TIMES/KVC ROYAL/
  // SAVICO PHN thuc ra la doanh thu VNPay KH Moi (02865168), khong phai KH
  // Cu, nen tu dong chuyen ra khoi day moi lan trang nay duoc mo (kem ca
  // trang VietQR, xem doisoat-vietqr.js) -- tu "don" ca hoa don MOI van tiep
  // tuc duoc tai len qua nut "Tai len combo" o trang nay (van gan tag "Vnpay
  // CS MB"/... nhu cu).
  if (migrateVnpayKhMoiInvoices(store)) trySave();
  const bank = store.banks.find((b) => b.name === ZVP_BANK_NAME);
  if (!bank) {
    return { error: `Chua co ngan hang "${ZVP_BANK_NAME}" (TK ${ZVP_BANK_ACCOUNT}) trong he thong.` };
  }
  const txs = store.transactions.filter((t) => t.bank_id === bank.id);
  const settlements = extractZvpSettlements(txs);

  const onlineMergedRaw = mergeResolvedGross(store.zvp_online_uploads);
  const offlineMergedRaw = mergeResolvedGross(store.zvp_offline_uploads);
  const payooMergedRaw = mergeResolvedGross(store.zvp_payoo_uploads);

  // Offline/Payoo each keep their OWN "Chi nhanh -> Ma cong trinh" mapping
  // table, separate from the shared zvp_gian_list -- if that table's target
  // code is itself just a site nickname that zvp_gian_list ALSO redirects
  // further (e.g. Offline map says "GHOST BRIDE MEGA DA NANG" -> "GHOST
  // BRIDE AE HUE", while zvp_gian_list separately redirects "GHOST BRIDE AE
  // HUE" -> "AE HUE KVCN" for invoices), Offline/Payoo revenue would
  // permanently land on the stale intermediate code while invoices (via
  // applyGianRedirectToInvoices below) land on the final one, showing a
  // wrong "Chua co HD"/"Lech" split for money that's actually the same. One
  // more zvp_gian_list hop here keeps every channel pointed at the same
  // final code (see applyGianRedirectToResolvedGross in utils/zvpReconcile.js).
  const onlineMerged = applyGianRedirectToResolvedGross(onlineMergedRaw, store.zvp_gian_list);
  const offlineMerged = applyGianRedirectToResolvedGross(offlineMergedRaw, store.zvp_gian_list);
  const payooMerged = applyGianRedirectToResolvedGross(payooMergedRaw, store.zvp_gian_list);

  const manualMatches = store.zvp_manual_matches || { online: {}, offline: {}, payoo: {} };
  // Chi Nhan (2026-07-27): "ngân hàng đối soát tất cả điều là 131 và không
  // chia theo chia sẻ hay không chia sẻ nữa" -- TK Co da luon la 131 cho moi
  // gian roi (seedGianMappingDefaults), nen KHONG con tach rieng dong CSE
  // (__FF) nua o bat ky buoc nao. cseOverrideCodes truoc day dung de CHU
  // DONG gan them hau to __FF cho 1 gian duoc danh dau CSE tren zvp_gian_list
  // -- gio luon de TRONG (Set rong) de reconcileZvpChannel khong bao gio tu
  // tao/giu lai split nay nua, du du lieu goc (gross/hoa don) co con hau to
  // __FF cu tu truoc khi doi (da duoc don sach 1 lan qua script migrate) hay
  // khong.
  const cseOverrideCodes = new Set();
  // Redirect each invoice's own "Ma diem tren misa thue" through the SAME
  // zvp_gian_list mapping already used for Online revenue (see
  // applyGianRedirectToInvoices in utils/zvpReconcile.js), keyed by the
  // invoice's own "Ten diem xuat hoa don" -- otherwise an invoice whose Ma
  // diem column is just the site's own name (not a real code) never lines
  // up against a settlement line that's already correctly redirected,
  // showing a permanent wrong "Chua co HD"/"Lech" even after Luyen fixes
  // the mapping via muc 1b or the "gian hang xuat HD"/master sheet.
  const zaloInvoicesRedirected = applyGianRedirectToInvoices(store.zvp_invoices.zalo, store.zvp_gian_list);
  const vnpayInvoicesRedirected = applyGianRedirectToInvoices(store.zvp_invoices.vnpay, store.zvp_gian_list);
  const payooInvoicesRedirected = applyGianRedirectToInvoices(store.zvp_invoices.payoo, store.zvp_gian_list);

  // Nhan, 2026-08-06: "zalo app vũng tàu nè map cho tôi đi KVC LOTTE VUNG
  // TAU" -- hoa don maDiem "KVC LOTTE VUNG TAU" (Zalo Online, hoa don 2558)
  // hien "Chưa có HĐ" du zvp_gian_list/zvp_gian_mapping da co san dung ma nay
  // -- nguyen nhan giong het vu "PINBALL VÀ GHẾ LOTTE BAC GIANG": store.
  // invoice_diem_alias la bang GLOBAL dung chung voi kenh VietQR bidv7702
  // (o do "KVC LOTTE VUNG TAU"/"POSH LOTTE MART VUNG TAU" duoc tro ve "VUNG
  // TAU PHCM", xem routes/doisoat-vietqr.js) -- truyen thang alias global vao
  // day khien hoa don Zalo bi doi maDiem nham truoc khi so voi gross cua
  // CHINH kenh nay (dang dung dung ten "KVC LOTTE VUNG TAU"). Loai 2 key nay
  // khoi alias truyen vao reconcileZvp (giu nguyen cho VietQR).
  const zvpScopedAlias = Object.assign({}, store.invoice_diem_alias);
  delete zvpScopedAlias["KVC LOTTE VUNG TAU"];
  delete zvpScopedAlias["POSH LOTTE MART VUNG TAU"];

  const reconciled = reconcileZvp(
    settlements,
    { online: onlineMerged, offline: offlineMerged, payoo: payooMerged },
    {
      online: { invoices: zaloInvoicesRedirected },
      offline: { invoices: vnpayInvoicesRedirected },
      payoo: { invoices: payooInvoicesRedirected },
    },
    store.zvp_gian_mapping,
    manualMatches,
    zvpScopedAlias,
    cseOverrideCodes
  );

  // Chi Nhan, 2026-07-30: "các đối soát tất cả các trang điều xếp theo ngày
  // cho tôi nhá" -- reconcileZvp tra ve ket qua theo thu tu giao dich ngan
  // hang trong store.transactions (thu tu tai/nhap lieu), khong phai thu tu
  // ngay thang, nen trang co the hien lon xon (vd ngay 30 roi 23 roi 24).
  // Sap lai TANG DAN theo settlementDate cho CA 3 kenh NGAY SAU KHI TINH
  // XONG (giong cach doisoat-vnpay-khmoi.js da lam) de moi cho dung ben duoi
  // deu tu dong ke thua dung thu tu.
  ["online", "offline", "payoo"].forEach((ch) => {
    (reconciled[ch] || []).sort((a, b) => (a.settlementDate < b.settlementDate ? -1 : a.settlementDate > b.settlementDate ? 1 : 0));
  });

  const allCodes = new Set();
  ["online", "offline", "payoo"].forEach((ch) => {
    reconciled[ch].forEach((r) => r.lines.forEach((l) => allCodes.add(l.code)));
  });

  // Same "invoice filed under a different site name" surface as Momo's
  // /doi-soat/momo route -- any Ma diem seen on the shared zalo/vnpay/payoo
  // invoice list that doesn't line up with a known gian code (and has no
  // alias yet) is shown so Luyen can map it once, applied immediately.
  // Uses the SAME gian-redirected invoices as reconciliation above, so this
  // list only shows what's STILL unmatched after the zvp_gian_list redirect
  // -- not names that are already fixed there.
  const knownCodesForAlias = new Set([...allCodes, ...Object.keys(store.zvp_gian_mapping || {})]);
  const invoiceDiemAlias = store.invoice_diem_alias || {};
  const unmatchedInvoiceCodesSet = new Set();
  [zaloInvoicesRedirected, vnpayInvoicesRedirected, payooInvoicesRedirected].forEach((list) => {
    (list || []).forEach((inv) => {
      if (!inv.maDiem) return;
      if (knownCodesForAlias.has(inv.maDiem)) return;
      if (invoiceDiemAlias[inv.maDiem]) return;
      unmatchedInvoiceCodesSet.add(inv.maDiem);
    });
  });

  // Any diem name/product title that couldn't be matched to the "gian hang
  // xuat HD" list is surfaced here rather than silently dropped -- caller
  // shows this as a warning banner so Luyen knows to fix the mapping sheet.
  // NOTE: Online is NOT included here anymore -- an unmatched Online product
  // is now auto-learned as its own new Ma cong trinh (see resolveProductCatalogGross's
  // learnedGian) so its revenue IS already being counted, just possibly under
  // the wrong code -- that's surfaced instead as "pendingOnlineGian" below,
  // with an actionable form to redirect it, rather than a scary "revenue not
  // counted" warning that's no longer true.
  const unmappedWarnings = [];
  store.zvp_offline_uploads.forEach((u) => {
    if (u.unmapped && u.unmapped.length) unmappedWarnings.push(`Offline "${u.file_name}": chua khop diem ${u.unmapped.join(", ")}`);
  });
  store.zvp_payoo_uploads.forEach((u) => {
    if (u.unmapped && u.unmapped.length) unmappedWarnings.push(`Payoo "${u.file_name}": chua khop diem ${u.unmapped.join(", ")}`);
  });
  // Luyen, 2026-07-27: canh bao TON TAI (khong chi flash 1 lan) cho cac GD
  // Online bi bo qua vi khong tim thay so don hang trong file OrderDetails --
  // day la DOANH THU THAT bi mat, khac voi "unmapped ten san pham" (van con
  // duoc auto-learn/dem). Chi hien tong so + vai so don hang mau, khong lam
  // day trang neu co qua nhieu.
  store.zvp_online_uploads.forEach((u) => {
    if (u.unmatchedOrders && u.unmatchedOrders.length) {
      unmappedWarnings.push(
        `Online "${u.file_name}": ${u.unmatchedOrders.length} giao dich KHONG tim thay so don hang trong file OrderDetails (vd: ${u.unmatchedOrders.slice(0, 5).join(", ")}) -- doanh thu cac GD nay CHUA duoc tinh vao Doanh thu gop, can tai lai OrderDetails moi hon + file phi nay qua muc "Tai len combo".`
      );
    }
  });

  // Chi Nhan (2026-07-27): "cảnh báo trên đầu tên sản phẩm mới chưa có" --
  // ten san pham moi duoc tu dong them vao zvp_online_product_map (o route
  // upload-combo) voi Ma cong trinh de TRONG -- canh bao TON TAI o day (khong
  // chi flash 1 lan) cho toi khi chi dien xong, de khong bi quen mat sau khi
  // dong thong bao. Bang o muc 1c cung tu day cac dong con trong Ma cong
  // trinh len dau danh sach (xem view) de de tim.
  const productMapMissingCode = Object.keys(store.zvp_online_product_map || {}).filter(
    (k) => !(store.zvp_online_product_map[k].maCongTrinh || "").trim()
  );
  if (productMapMissingCode.length > 0) {
    unmappedWarnings.push(
      `Bảng tra sản phẩm Online (mục 1c): ${productMapMissingCode.length} sản phẩm MỚI chưa có Mã công trình (${productMapMissingCode.slice(0, 8).join(" | ")}${productMapMissingCode.length > 8 ? "..." : ""}) -- chị điền Mã công trình vào bảng ở mục 1c rồi tải lại combo 2 file 1 lần nữa.`
    );
  }

  // Gian tren "Danh muc ten san pham" ma auto-learn da tu tao ma cong trinh
  // rieng (ten gian == ma cong trinh, vi chua khop duoc voi gian nao co san
  // luc tai. Luyen tu quyet dinh qua form "Sua ma gian" o muc 1b.
  const pendingOnlineGian = (store.zvp_gian_list || []).filter(
    (g) => normText(g.tenDiem) === normText(g.maCongTrinh)
  );

  // "Khoa so" -- Luyen, 2026-07-21: "TẤT CẢ CÁC TRANG ĐIỀU CÓ KHÓA SỔ CHO TÔI
  // NHÁ", ap dung cung 1 co che da lam cho VietQR/Momo: 1 ngay khoa duy nhat
  // cho ca trang (chi 1 cong ty/1 tai khoan ngan hang), ap dung cho CA 3 kenh
  // online/offline/payoo (chung 1 mo hinh dong tien "tien ve ngan hang" gop 1
  // TK). Moi ngay settlementDate <= ngay khoa duoc coi la "locked" -- dua diff
  // ve 0, giu nguyen gross/invoiceTotal de tra cuu.
  const lockDate = store.zvp_lock_date || "";
  if (lockDate) {
    ["online", "offline", "payoo"].forEach((ch) => {
      (reconciled[ch] || []).forEach((r) => {
        if (r.settlementDate <= lockDate) {
          r.locked = true;
          r.lines.forEach((l) => {
            l.locked = true;
            l.diff = 0;
          });
        }
      });
    });
  }

  return {
    reconciled,
    lockDate,
    allCodes: Array.from(allCodes).sort(),
    unmappedWarnings,
    invoiceDiemAlias,
    unmatchedInvoiceCodes: Array.from(unmatchedInvoiceCodesSet).sort(),
    pendingOnlineGian,
  };
}

router.get("/doi-soat/zvp", (req, res) => {
  const store = load();
  const built = buildReconciliation(store);
  const reconciledAll = built.reconciled || { online: [], offline: [], payoo: [] };

  // Chon theo thang: mac dinh la thang gan nhat (tinh tren ca 3 kenh gop lai)
  // de trang khong bi dai/roi ("nhieu roi qua" -- Luyen), van chon "Tat ca"
  // duoc qua dropdown. Bang TK Co (allCodes) va canh bao chua khop van tinh
  // tren TOAN BO du lieu, khong bi anh huong boi thang dang loc xem.
  const monthSet = new Set();
  ["online", "offline", "payoo"].forEach((ch) => {
    (reconciledAll[ch] || []).forEach((r) => monthSet.add(r.settlementDate.slice(0, 7)));
  });
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";

  const monthFiltered = { online: [], offline: [], payoo: [] };
  ["online", "offline", "payoo"].forEach((ch) => {
    monthFiltered[ch] = selectedMonth
      ? (reconciledAll[ch] || []).filter((r) => r.settlementDate.slice(0, 7) === selectedMonth)
      : reconciledAll[ch] || [];
  });

  // Luyen, 2026-07-20: "chon ngay cua no nha" -- ngoai loc theo thang, them
  // loc theo TUNG NGAY cu the de nhay thang toi dung ngay can xem, khong phai
  // cuon qua het ca thang. Danh sach ngay chi tinh trong pham vi thang dang
  // chon (rong hon neu dang "Tat ca") de dropdown ngay luon khop voi du lieu
  // dang hien.
  const daySet = new Set();
  ["online", "offline", "payoo"].forEach((ch) => {
    monthFiltered[ch].forEach((r) => daySet.add(r.settlementDate));
  });
  const days = Array.from(daySet).sort().reverse();
  const selectedDay = req.query.day || "";

  const reconciled = { online: [], offline: [], payoo: [] };
  ["online", "offline", "payoo"].forEach((ch) => {
    reconciled[ch] = selectedDay
      ? monthFiltered[ch].filter((r) => r.settlementDate === selectedDay)
      : monthFiltered[ch];
  });

  res.render("doisoat-zvp", {
    userName: req.session.userName,
    onlineUploads: store.zvp_online_uploads,
    offlineUploads: store.zvp_offline_uploads,
    payooUploads: store.zvp_payoo_uploads,
    invoiceCounts: {
      zalo: store.zvp_invoices.zalo.length,
      vnpay: store.zvp_invoices.vnpay.length,
      payoo: store.zvp_invoices.payoo.length,
    },
    hasGianList: store.zvp_gian_list && store.zvp_gian_list.length > 0,
    hasOfflineDiemMap: store.zvp_offline_diem_map && Object.keys(store.zvp_offline_diem_map).length > 0,
    hasOnlineProductMap: store.zvp_online_product_map && Object.keys(store.zvp_online_product_map).length > 0,
    onlineProductMap: store.zvp_online_product_map || {},
    gianMaster: store.zvp_gian_master || null,
    reconciled,
    months,
    selectedMonth,
    exportTuDefault: monthBounds(selectedMonth).first,
    exportDenDefault: monthBounds(selectedMonth).last,
    days,
    selectedDay,
    gianMapping: store.zvp_gian_mapping,
    allCodes: built.allCodes || [],
    unmappedWarnings: built.unmappedWarnings || [],
    invoiceDiemAlias: built.invoiceDiemAlias || {},
    unmatchedInvoiceCodes: built.unmatchedInvoiceCodes || [],
    pendingOnlineGian: built.pendingOnlineGian || [],
    lockDate: built.lockDate || "",
    error: built.error || req.query.error || null,
    success: req.query.success || null,
  });
});

// ---------- Khoa so (giong VietQR/Momo) -- gui lockDate rong de mo khoa lai. ----------
router.post("/doi-soat/zvp/khoa-so", requireAdmin, (req, res) => {
  const store = load();
  try {
    const lockDate = (req.body.lockDate || "").trim();
    if (lockDate && !/^\d{4}-\d{2}-\d{2}$/.test(lockDate)) throw new Error("Ngay khoa khong hop le (dang YYYY-MM-DD).");
    store.zvp_lock_date = lockDate;
    save(store);
    const msg = lockDate
      ? `Da khoa so den het ngay ${lockDate}. Cac ngay tu do tro ve truoc se khong con hien canh bao lech nua.`
      : "Da mo khoa so.";
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});
// ---------- DEBUG TAM THOI: xem raw invoice records theo soHd (xoa sau khi dung xong) ----------
router.get("/doi-soat/zvp/debug-invoices", requireAdmin, (req, res) => {
  const store = load();
  const raw = (req.query.soHd || "").trim();
  const targets = new Set(raw.split(/[,\s]+/).map((s) => s.trim()).filter(Boolean));
  const out = { zalo: [], vnpay: [], payoo: [], momo: [] };
  ["zalo", "vnpay", "payoo"].forEach((ch) => {
    out[ch] = (store.zvp_invoices[ch] || []).filter((i) => targets.has(String(i.soHd)));
  });
  out.momo = (store.momo_invoices || []).filter((i) => targets.has(String(i.soHd)));
  res.type("json").send(JSON.stringify(out, null, 1));
});


// ---------- DEBUG TAM THOI: xem raw invoice records theo soHd (xoa sau khi dung xong) ----------
// Luyen, 2026-08-13: dung de dieu tra trung hoa don ZVP (giong da lam ben
// Momo) -- can xem THAT store co nhung ban ghi nao cho 1 so HD (co the trung
// nhieu ban ghi khac maDiem/ngayHd do composite key), thay vi doan tu file
// excel local (co the khong khop 100% voi du lieu that tren server).
router.get("/doi-soat/zvp/debug-invoices", requireAdmin, (req, res) => {
  const store = load();
  const raw = (req.query.soHd || "").trim();
  const targets = new Set(raw.split(/[,\s]+/).map((s) => s.trim()).filter(Boolean));
  const out = { zalo: [], vnpay: [], payoo: [], momo: [] };
  ["zalo", "vnpay", "payoo"].forEach((ch) => {
    out[ch] = (store.zvp_invoices[ch] || []).filter((i) => targets.has(String(i.soHd)));
  });
  out.momo = (store.momo_invoices || []).filter((i) => targets.has(String(i.soHd)));
  res.type("json").send(JSON.stringify(out, null, 1));
});

// ---------- Sua ma gian tren "Danh muc ten san pham" (Online) ----------
// Danh cho truong hop 1 san pham moi thuc ra thuoc VE 1 gian DA CO (vd
// "SNOWFUN TAN PHU" thuc ra la 1 san pham cua "AM TP KVCM") nhung auto-learn
// da tam ghi nhan no nhu 1 ma cong trinh rieng (tenDiem == maCongTrinh) vi
// khong khop duoc voi gian nao co san luc tai. Luu y: sua o day CHI anh huong
// cho lan tai file Online TIEP THEO -- doanh thu da tai truoc do van giu
// nguyen ma cu, can tai lai file "Tong hop Zalo App" (muc 1) sau khi sua.
router.post("/doi-soat/zvp/online-gian-fix", requireAdmin, (req, res) => {
  const store = load();
  try {
    const { tenDiem, maCongTrinh, isCse } = req.body;
    if (!tenDiem || !maCongTrinh) throw new Error("Thieu ten gian hoac ma cong trinh de gan lai.");
    if (!store.zvp_gian_list) store.zvp_gian_list = [];
    const key = normText(tenDiem);
    const entry = { tenDiem, maCongTrinh: normCode(maCongTrinh), isCse: !!isCse };
    const idx = store.zvp_gian_list.findIndex((g) => normText(g.tenDiem) === key);
    if (idx >= 0) store.zvp_gian_list[idx] = entry;
    else store.zvp_gian_list.push(entry);
    save(store);
    res.redirect(
      "/doi-soat/zvp?success=" +
        encodeURIComponent(
          `Da gan "${tenDiem}" -> "${entry.maCongTrinh}"${entry.isCse ? " (CSE)" : ""}. Hay tai lai file "Tong hop Zalo App" o muc 1 de ap dung cho doanh thu da tai truoc do.`
        )
    );
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: file "Tong hop Zalo App" (gian hang xuat HD + Doi soat Vnpay Online) ----------
router.post("/doi-soat/zvp/upload-online", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    let gianList = parseGianXuatHdSheet(req.file.buffer);
    if (gianList.length === 0) {
      throw new Error('Khong tim thay sheet "gian hang xuat HD" trong file nay.');
    }
    // Merge in the daily master "gian " sheet's Online rows (if any has been
    // uploaded) so more-complete/corrected gian entries there (e.g. distinct
    // non-CSE products that share a Ma cong trinh with a CSE product) are
    // never lost just because this file's own "gian hang xuat HD" list is
    // uploaded again later -- master always wins on a name collision.
    gianList = mergeGianListWithMaster(gianList, getMasterRowsForChannel(store, "zalo mini app"));
    // Preserve gian Luyen already corrected by hand via "Sua ma gian" (muc
    // 1b) -- otherwise re-uploading this SAME file wipes those fixes out:
    // parseGianXuatHdSheet/master rebuild gianList from scratch every time,
    // so without this merge, a name like "SNOWFUN TAN PHU" would resolve
    // back to its OWN self-referential code again on the very next upload,
    // undoing her fix and putting it right back in the "1b" review list.
    // Only entries that are NOT self-referential (tenDiem !== maCongTrinh --
    // i.e. actually redirected to a real code) count as "fixed"; plain
    // auto-learned placeholders are left to be freshly re-evaluated against
    // this upload's own data.
    const previouslyFixedGian = (store.zvp_gian_list || []).filter(
      (g) => normText(g.tenDiem) !== normText(g.maCongTrinh)
    );
    if (previouslyFixedGian.length > 0) {
      gianList = mergeGianListWithMaster(
        gianList,
        previouslyFixedGian.map((g) => ({ raw: g.tenDiem, maCongTrinh: g.maCongTrinh, isCse: g.isCse }))
      );
    }
    // Online gross now comes SOLELY from "Danh muc ten san pham" -- a per-
    // product, per-day sheet inside the same file that already carries the
    // exact Ma cong trinh for every product row (no fuzzy keyword matching
    // against product titles needed, which used to misattribute revenue --
    // e.g. KVC ROYAL showing gross on days it had zero real sales).
    const parsedOnline = parseProductCatalogSheet(req.file.buffer);
    if (parsedOnline.rows.length === 0) {
      throw new Error('Khong tim thay sheet "Danh muc ten san pham" (hoac khong doc duoc cau truc) trong file nay.');
    }
    const resolved = resolveProductCatalogGross(parsedOnline.rows, gianList);
    // Auto-remember any brand-new gian this file introduced (not on "gian
    // hang xuat HD" or the master sheet yet) so Luyen never has to hand-add
    // them -- next upload won't flag them as unmapped anymore. Defaulted to
    // non-CSE (131); she can still flip a code to CSE later via the normal
    // TK Co mapping screen if one of these ever turns out to be CSE.
    if (resolved.learnedGian && resolved.learnedGian.length > 0) {
      const existingKeys = new Set(gianList.map((g) => normText(g.tenDiem)));
      for (const g of resolved.learnedGian) {
        if (!existingKeys.has(normText(g.tenDiem))) {
          gianList.push(g);
          existingKeys.add(normText(g.tenDiem));
        }
      }
    }

    if (isDuplicateRecentUpload(store.zvp_online_uploads, req.file.originalname, resolved.grossByCode)) {
      return res.redirect(
        "/doi-soat/zvp?success=" + encodeURIComponent(`File "${req.file.originalname}" vua duoc tai len roi (bo qua ban trung lap).`)
      );
    }

    store.zvp_online_uploads.push({
      id: nextId(store, "zvp_online_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsedOnline.sheetName,
      dates: resolved.dates,
      codes: resolved.codes,
      grossByCode: resolved.grossByCode,
      netByCode: resolved.netByCode,
      unmapped: resolved.unmapped,
    });
    seedGianMappingDefaults(store, resolved.codes);
    // Persist the raw gian list (Ten diem/Ma cong trinh/CSE) -- now including
    // any newly-learned gian above -- so a later raw-portal combo upload (or
    // the next "Tong hop Zalo App" upload) can resolve product -> gian
    // without this file being re-uploaded every time, and without the new
    // gian showing up as unmapped again.
    store.zvp_gian_list = gianList;
    save(store);

    let successMsg = `Da nap "${parsedOnline.sheetName}" (${resolved.dates[0]} - ${resolved.dates[resolved.dates.length - 1]}), ${resolved.codes.length} ma cong trinh (${gianList.length} gian tren danh sach, ${gianList.filter((g) => g.isCse).length} gian CSE).${UPDATED_NOTE}`;
    if (resolved.learnedGian && resolved.learnedGian.length > 0) {
      successMsg += ` Da tu dong ghi nho ${resolved.learnedGian.length} gian moi chua co tren danh sach (${resolved.learnedGian.map((g) => g.tenDiem).slice(0, 5).join(", ")}${resolved.learnedGian.length > 5 ? "..." : ""}), mac dinh khong CSE (TK 131) -- kiem tra lai neu gian nao trong so nay thuc ra la CSE.`;
    }
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: bang tra CHINH XAC "Ten san pham -> Ma cong trinh" cho Online (sheet "noi") ----------
// Luyen, 2026-07-24: "giờ tôi sẽ thiết lập lại chính sát hơn ... dựa vào
// file hehehehehe sheet nối rồi gắn mã công trình vô ... còn cái nào sau
// này có tên sản phẩm mới bạn cảnh báo tên sản phẩm đó cho tôi" -- file
// rieng chi tu duy tri (sheet ten "noi"/"nối", 2 cot Ten san pham + Ma cong
// trinh), thay the buildOnlineProductMatcher (fuzzy) trong luong
// upload-combo cho ket qua chinh xac hon (khong doan nham khi 2 san pham
// tinh co trung tu khoa). Merge (khong xoa dong cu) vao store.zvp_online_product_map.
router.post("/doi-soat/zvp/upload-online-product-map", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const { sheetName, rows } = parseOnlineProductMapSheet(req.file.buffer);
    if (!sheetName) {
      throw new Error('Khong tim thay sheet "noi" (hoac "nối") trong file nay.');
    }
    if (rows.length === 0) {
      throw new Error('Sheet "noi" khong doc duoc du lieu (can cot "Ten san pham" va "Ma cong trinh").');
    }
    const { map, added, updated } = mergeOnlineProductMap(store.zvp_online_product_map, rows, store.zvp_gian_list);
    store.zvp_online_product_map = map;
    save(store);
    let successMsg = `Da nap sheet "${sheetName}": ${rows.length} dong (${added} moi, ${updated} cap nhat) -- tong ${Object.keys(map).length} san pham dang co trong bang tra Online.`;
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// Them nhanh 1 dong tra rieng le (khong can tai lai ca file) -- dung khi
// canh bao "san pham moi" hien ra sau 1 lan tai combo, Luyen go luon Ma cong
// trinh dung cho san pham do de lan sau tu khop.
router.post("/doi-soat/zvp/online-product-map/add", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { tenSanPham, maCongTrinh, isCse } = req.body;
    if (!tenSanPham || !tenSanPham.trim()) throw new Error("Thieu ten san pham.");
    if (!maCongTrinh || !maCongTrinh.trim()) throw new Error("Thieu ma cong trinh.");
    if (!store.zvp_online_product_map) store.zvp_online_product_map = {};
    const key = String(tenSanPham).trim().replace(/\s+/g, " ");
    store.zvp_online_product_map[key] = {
      maCongTrinh: normCode(maCongTrinh),
      isCse: isCse === "1" || isCse === "on" || isCse === true,
    };
    save(store);
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(`Da them anh xa san pham "${key}" -> ${normCode(maCongTrinh)}.`));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: file VNPay Offline ("du lieu VNpay co so KHxxx" + "gian hang VNpay co so KHxxx") ----------
router.post("/doi-soat/zvp/upload-offline", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const parsed = parseOfflineVnpayWorkbook(req.file.buffer);
    let diemMap = parseDiemMappingSheet(req.file.buffer, "gian hang VNpay co so");
    if (Object.keys(diemMap).length === 0) {
      throw new Error('Khong tim thay sheet "gian hang VNpay co so ..." (bang mapping Chi nhanh -> Ma cong trinh) trong file nay.');
    }
    // Master "gian " sheet's VNPay Co so rows win on a Chi nhanh collision.
    diemMap = mergeDiemMapWithMaster(diemMap, getMasterRowsForChannel(store, "vnpay co so"));
    const resolved = resolveDiemGross(parsed, diemMap);

    if (isDuplicateRecentUpload(store.zvp_offline_uploads, req.file.originalname, resolved.grossByCode)) {
      return res.redirect(
        "/doi-soat/zvp?success=" + encodeURIComponent(`File "${req.file.originalname}" vua duoc tai len roi (bo qua ban trung lap).`)
      );
    }

    store.zvp_offline_uploads.push({
      id: nextId(store, "zvp_offline_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      dates: resolved.dates,
      codes: resolved.codes,
      grossByCode: resolved.grossByCode,
      netByCode: resolved.netByCode,
      unmapped: resolved.unmapped,
    });
    seedGianMappingDefaults(store, resolved.codes);
    // Persist the raw Chi nhanh -> Ma cong trinh mapping so a later combo
    // upload of the raw fee report can resolve Offline gian without this
    // file being re-uploaded every time.
    store.zvp_offline_diem_map = diemMap;
    save(store);

    let successMsg = `Da nap "${parsed.sheetName}" (${resolved.dates[0]} - ${resolved.dates[resolved.dates.length - 1]}), ${resolved.codes.length} ma cong trinh.${UPDATED_NOTE}`;
    if (resolved.unmapped.length > 0) {
      successMsg += ` CANH BAO: ${resolved.unmapped.length} ten diem chua khop duoc mapping (${resolved.unmapped.join(", ")}).`;
    }
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: file THO VNPay Offline (chi can 1 file "Du lieu bao cao
// phi theo GD thanh toan", khong can OrderDetails) ----------
// Luyen, 2026-07-24: "chỗ offline á có thể tôi tải đối soát này bạn lên đối
// soát vnpay offline cho tôi nhá á dựa vào điểm thu để lấy ra gian và ngày
// giao dịch phí hay là tổng tiền á lấy các giao dịch thành công nhá và bỏ qua
// điểm thu FUNZONE MINI APP nhá và thêm trên wed úp cái dữ liệu này lên nha
// thêm dạng này á" -- khac voi muc 2b (upload-combo, can CA 2 file vi con xu
// ly ca phan Online), route nay CHI can 1 file duy nhat vi Offline khong can
// tra ten san pham (OrderDetails chi de dung cho Online). Dung lai
// store.zvp_offline_diem_map da co san (tu it nhat 1 lan tai file o muc 2)
// de tra Ma cong trinh theo "Chi nhanh", khu trung theo TUNG giao dich (Ma
// giao dich) qua store.zvp_offline_raw_tx -- an toan khi tai chong file cho
// cung khoang ngay, giong het pattern upload-payoo-raw o duoi.
router.post("/doi-soat/zvp/upload-offline-raw", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    if (!store.zvp_offline_diem_map || Object.keys(store.zvp_offline_diem_map).length === 0) {
      throw new Error(
        'Chua co bang mapping "Chi nhanh -> Ma cong trinh" cua Offline -- hay tai len file "VNPay thu ho co so" (sheet "gian hang VNpay co so ...") o muc 2 truoc, it nhat 1 lan.'
      );
    }
    if (!store.zvp_offline_raw_tx) store.zvp_offline_raw_tx = {};

    const parsed = parseVnpayOfflineFeeReport(req.file.buffer);
    let addedCount = 0;
    let dupCount = 0;
    let upgradedCount = 0;
    for (const tx of parsed.transactions) {
      const existing = store.zvp_offline_raw_tx[tx.txKey];
      if (existing) {
        // Luyen 2026-08-06: ngoai nang cap phi, con cap nhat khi date thay doi
        // (vi du: parser doi tu "Ngay hach toan" sang "Thoi gian GD" -- re-upload
        // se tu dong sua toan bo ngay ma khong can xoa thu cong).
        const needsUpdate = (existing.fee === 0 && tx.fee > 0) || (existing.date !== tx.date);
        if (needsUpdate) {
          store.zvp_offline_raw_tx[tx.txKey] = { date: tx.date, chiNhanh: tx.chiNhanh, gross: tx.gross, fee: tx.fee, net: tx.net };
          upgradedCount++;
        } else {
          dupCount++;
        }
        continue;
      }
      store.zvp_offline_raw_tx[tx.txKey] = { date: tx.date, chiNhanh: tx.chiNhanh, gross: tx.gross, fee: tx.fee, net: tx.net };
      addedCount++;
    }

    if (addedCount === 0 && upgradedCount === 0) {
      return res.redirect(
        "/doi-soat/zvp?success=" +
          encodeURIComponent(`File "${req.file.originalname}": ca ${dupCount} giao dich deu da co roi (bo qua toan bo, khong trung lap).`)
      );
    }

    // Tinh lai TOAN BO grossByCode/netByCode tu TOAN BO giao dich da tich luy
    // (khong chi file vua tai) -- moi lan tai la 1 "phien ban day du" moi
    // nhat, dung khop pattern upload-payoo-raw. Chi nhanh khop CHINH XAC voi
    // store.zvp_offline_diem_map truoc (da xac nhan khop 100% tren file thuc
    // te), roi moi thu khop khong dau lam du phong cho file tuong lai co the
    // ghi Chi nhanh khac dau/hoa thuong.
    const diemMap = store.zvp_offline_diem_map;
    const normalizedDiemIndex = new Map();
    Object.keys(diemMap).forEach((k) => {
      const nk = normText(k);
      if (!normalizedDiemIndex.has(nk)) normalizedDiemIndex.set(nk, k);
    });
    const grossByCode = {};
    const netByCode = {};
    const codes = new Set();
    const dates = new Set();
    const unmapped = new Set();
    for (const tx of Object.values(store.zvp_offline_raw_tx)) {
      let mapped = diemMap[tx.chiNhanh];
      if (!mapped) {
        const origKey = normalizedDiemIndex.get(normText(tx.chiNhanh));
        if (origKey) mapped = diemMap[origKey];
      }
      if (!mapped) {
        unmapped.add(tx.chiNhanh);
        continue;
      }
      // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong rieng.
      const code = mapped.maCongTrinh;
      codes.add(code);
      dates.add(tx.date);
      const key = `${tx.date}|${code}`;
      grossByCode[key] = (grossByCode[key] || 0) + tx.gross;
      netByCode[key] = (netByCode[key] || 0) + tx.net;
    }

    // Xoa cac "[File tho]" cu truoc khi push entry moi. GrossByCode cua entry
    // moi da duoc tinh lai tu TOAN BO zvp_offline_raw_tx (voi GD-date moi) nen
    // hoan toan thay the duoc tat ca entry cu. Neu giu ca 2, keys settlement-
    // date (cu) lan GD-date (moi) se cung duoc cong vao -- double-count.
    // Upload "combo" (khong phai "[File tho]") van duoc giu nguyen.
    store.zvp_offline_uploads = store.zvp_offline_uploads.filter(u => !u.file_name.startsWith("[File thô]"));

    store.zvp_offline_uploads.push({
      id: nextId(store, "zvp_offline_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: `[File thô] ${req.file.originalname}`,
      sheetName: parsed.sheetName,
      dates: Array.from(dates).sort(),
      codes: Array.from(codes),
      grossByCode,
      netByCode,
      unmapped: Array.from(unmapped),
    });
    seedGianMappingDefaults(store, Array.from(codes));
    save(store);

    let successMsg =
      `Da nap file tho "${req.file.originalname}": ${addedCount} giao dich moi` +
      (upgradedCount > 0 ? `, bo sung phi cho ${upgradedCount} giao dich da co truoc do` : "") +
      (dupCount > 0 ? `, bo qua ${dupCount} giao dich da co (trung lap)` : "") +
      ` (bo qua ${parsed.excludedFunzone} giao dich FUNZONE MINI APP - Online).${UPDATED_NOTE}`;
    if (parsed.excludedFailedStatus > 0) {
      successMsg += ` Bo qua ${parsed.excludedFailedStatus} giao dich khong "thanh cong".`;
    }
    if (unmapped.size > 0) {
      successMsg += ` CANH BAO: ${unmapped.size} Chi nhanh chua khop mapping (${Array.from(unmapped).join(", ")}) -- can tai lai file "VNPay thu ho co so" (muc 2) neu day la co so moi.`;
    }
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Tra cuu nhanh VNPay Offline raw theo gian + khoang ngay ----------
// Luyen, 2026-08-06: "chon cac gian vnpay offline roi chon ngay la ra dung so
// dung bỏ qua cai gi do nhu cua onl tach rieng 2 kieu lay du lieu" -- cong cu
// tra cuu truc tiep tu zvp_offline_raw_tx (khong qua reconcileZvp), loc theo
// khoang ngay hach toan + tuy chon theo gian, tra ve JSON tong gross/net theo
// ma cong trinh de hien thi tren trang ma khong reload toan bo.
router.get("/doi-soat/zvp/offline-raw-detail", requireLogin, (req, res) => {
  const store = load();
  const rawTx = store.zvp_offline_raw_tx || {};
  const diemMap = store.zvp_offline_diem_map || {};
  const gianList = store.zvp_gian_list || [];

  const tu = req.query.tu || "";
  const den = req.query.den || "";
  const giansFilter = req.query.gians ? req.query.gians.split(",").map((s) => s.trim()).filter(Boolean) : [];

  // Build normalized diem index (for case/diacritics-insensitive lookup)
  function normT(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim(); }
  const normalizedDiemIndex = new Map();
  Object.keys(diemMap).forEach((k) => {
    const nk = normT(k);
    if (!normalizedDiemIndex.has(nk)) normalizedDiemIndex.set(nk, k);
  });

  // Build gian_list redirect map (same as applyGianRedirectToResolvedGross)
  const gianRedirect = {};
  (gianList || []).forEach((g) => {
    if (g.tenDiem && g.maCongTrinh) gianRedirect[g.tenDiem] = g.maCongTrinh;
  });
  function resolveCode(code) {
    let c = code;
    const visited = new Set();
    while (c && gianRedirect[c] && gianRedirect[c] !== c && !visited.has(c)) {
      visited.add(c);
      c = gianRedirect[c];
    }
    return c || code;
  }

  const byCode = {};
  const byCodeDate = {}; // for per-date breakdown

  for (const tx of Object.values(rawTx)) {
    const date = tx.date || "";
    if (tu && date < tu) continue;
    if (den && date > den) continue;

    let mapped = diemMap[tx.chiNhanh];
    if (!mapped) {
      const origKey = normalizedDiemIndex.get(normT(tx.chiNhanh));
      if (origKey) mapped = diemMap[origKey];
    }
    if (!mapped) continue;
    const code = resolveCode(mapped.maCongTrinh);

    if (giansFilter.length > 0 && !giansFilter.includes(code)) continue;

    if (!byCode[code]) byCode[code] = { code, chiNhanh: tx.chiNhanh, gross: 0, net: 0, count: 0 };
    byCode[code].gross += tx.gross;
    byCode[code].net += tx.net;
    byCode[code].count++;

    const dk = `${date}|${code}`;
    if (!byCodeDate[dk]) byCodeDate[dk] = { date, code, gross: 0, net: 0, count: 0 };
    byCodeDate[dk].gross += tx.gross;
    byCodeDate[dk].net += tx.net;
    byCodeDate[dk].count++;
  }

  const rows = Object.values(byCode).sort((a, b) => b.gross - a.gross);
  const rowsByDate = Object.values(byCodeDate).sort((a, b) => a.date.localeCompare(b.date) || b.gross - a.gross);

  // List of all available gian (for the filter dropdown)
  const allGians = [];
  const seenGian = new Set();
  for (const tx of Object.values(rawTx)) {
    let mapped = diemMap[tx.chiNhanh];
    if (!mapped) {
      const origKey = normalizedDiemIndex.get(normT(tx.chiNhanh));
      if (origKey) mapped = diemMap[origKey];
    }
    if (!mapped) continue;
    const code = resolveCode(mapped.maCongTrinh);
    if (!seenGian.has(code)) { seenGian.add(code); allGians.push(code); }
  }
  allGians.sort();

  res.json({
    tu, den,
    totalGross: rows.reduce((s, r) => s + r.gross, 0),
    totalNet: rows.reduce((s, r) => s + r.net, 0),
    totalCount: rows.reduce((s, r) => s + r.count, 0),
    rows,
    rowsByDate,
    allGians,
  });
});

// ---------- Upload: file Payoo ("Du lieu Payoo co so KHxxx" + "Danh muc ten diem") ----------
router.post("/doi-soat/zvp/upload-payoo", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const parsed = parsePayooWorkbook(req.file.buffer);
    let diemMap = parsePayooDiemMapping(req.file.buffer);
    if (Object.keys(diemMap).length === 0) {
      throw new Error('Khong tim thay sheet "Danh muc ten diem" (bang mapping Chi nhanh -> Ma cong trinh) trong file Payoo nay.');
    }
    // Master "gian " sheet's Payoo QR + Payoo the rows win on a Chi nhanh
    // collision; persisted separately so it can be reused/inspected later.
    diemMap = mergeDiemMapWithMaster(diemMap, getMasterRowsForChannel(store, "payoo"));
    store.zvp_payoo_diem_map = diemMap;
    const resolved = resolveDiemGross(parsed, diemMap);

    if (isDuplicateRecentUpload(store.zvp_payoo_uploads, req.file.originalname, resolved.grossByCode)) {
      return res.redirect(
        "/doi-soat/zvp?success=" + encodeURIComponent(`File "${req.file.originalname}" vua duoc tai len roi (bo qua ban trung lap).`)
      );
    }

    store.zvp_payoo_uploads.push({
      id: nextId(store, "zvp_payoo_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      dates: resolved.dates,
      codes: resolved.codes,
      grossByCode: resolved.grossByCode,
      netByCode: resolved.netByCode,
      unmapped: resolved.unmapped,
    });
    seedGianMappingDefaults(store, resolved.codes);
    save(store);

    let successMsg = `Da nap "${parsed.sheetName}" (${resolved.dates[0]} - ${resolved.dates[resolved.dates.length - 1]}), ${resolved.codes.length} ma cong trinh.${UPDATED_NOTE}`;
    if (resolved.unmapped.length > 0) {
      successMsg += ` CANH BAO: ${resolved.unmapped.length} ten diem chua khop duoc mapping (${resolved.unmapped.join(", ")}).`;
    }
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: file THO Payoo/VNPay (khong can sheet "Danh muc ten
// diem" moi lan) ----------
// Chi Nhan, 2026-07-24: Luyen tai thang file export goc tu cong ("BÁO CÁO
// GIAO DỊCH BÁN HÀNG HỢP TÁC VỚI PAYOO", hoac "BÁO CÁO GIAO DỊCH QR" --
// ca 2 deu duoc, tu nhan dien qua parsePayooRawReport). Dung LAI
// store.zvp_payoo_diem_map da co san (tu lan tai file day du "Payoo thu ho"
// it nhat 1 lan truoc do qua muc 3) de tra Ma cong trinh, khong bat tai lai
// sheet mapping moi lan. Khu trung THEO TUNG GIAO DICH (store.zvp_payoo_raw_tx,
// khoa la txKey -- xem parsePayooRawReport) nen tai chong file (vd tai ca bao
// cao "Giao dich ban hang" LAN "Giao dich QR" rieng cho CUNG khoang ngay) se
// TU DONG bo qua giao dich da thay, khong cong don/trung doanh thu.
router.post("/doi-soat/zvp/upload-payoo-raw", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    if (!store.zvp_payoo_diem_map || Object.keys(store.zvp_payoo_diem_map).length === 0) {
      throw new Error(
        'Chua co bang mapping "Chi nhanh -> Ma cong trinh" cua Payoo -- hay tai len file "Payoo thu ho" (sheet "Danh muc ten diem") o muc 3 truoc, it nhat 1 lan.'
      );
    }
    if (!store.zvp_payoo_raw_tx) store.zvp_payoo_raw_tx = {};

    const parsed = parsePayooRawReport(req.file.buffer);
    let addedCount = 0;
    let dupCount = 0;
    let upgradedCount = 0;
    for (const tx of parsed.transactions) {
      const existing = store.zvp_payoo_raw_tx[tx.txKey];
      if (existing) {
        // Chi Nhan, 2026-07-24: bao cao "Giao dich QR" (rieng) khong co cot
        // phi, nen 1 giao dich thay TRUOC qua bao cao do se bi ghi fee=0 --
        // neu sau nay CUNG giao dich do xuat hien lai qua bao cao "Giao dich
        // ban hang" (co phi that), NANG CAP len ban co phi thay vi giu mai
        // ban thieu phi (tranh Tien ve NH bi tinh du hon thuc te mai mai).
        if (existing.fee === 0 && tx.fee > 0) {
          store.zvp_payoo_raw_tx[tx.txKey] = { date: tx.date, gian: tx.gian, gross: tx.gross, fee: tx.fee, net: tx.net };
          upgradedCount++;
        } else {
          dupCount++;
        }
        continue;
      }
      store.zvp_payoo_raw_tx[tx.txKey] = { date: tx.date, gian: tx.gian, gross: tx.gross, fee: tx.fee, net: tx.net };
      addedCount++;
    }

    if (addedCount === 0 && upgradedCount === 0) {
      return res.redirect(
        "/doi-soat/zvp?success=" +
          encodeURIComponent(`File "${req.file.originalname}": ca ${dupCount} giao dich deu da co roi (bo qua toan bo, khong trung lap).`)
      );
    }

    // Tinh lai TOAN BO grossByCode/netByCode tu TOAN BO giao dich da tich luy
    // (khong chi file vua tai) -- moi lan tai la 1 "phien ban day du" moi
    // nhat, luon dung du da tai bao nhieu lan/dinh dang khac nhau truoc do.
    const diemMap = store.zvp_payoo_diem_map;
    const grossByCode = {};
    const netByCode = {};
    const codes = new Set();
    const dates = new Set();
    const unmapped = new Set();
    for (const tx of Object.values(store.zvp_payoo_raw_tx)) {
      const mapped = diemMap[tx.gian];
      if (!mapped) {
        unmapped.add(tx.gian);
        continue;
      }
      // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong rieng.
      const code = mapped.maCongTrinh;
      codes.add(code);
      dates.add(tx.date);
      const key = `${tx.date}|${code}`;
      grossByCode[key] = (grossByCode[key] || 0) + tx.gross;
      netByCode[key] = (netByCode[key] || 0) + tx.net;
    }

    store.zvp_payoo_uploads.push({
      id: nextId(store, "zvp_payoo_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: `[File thô] ${req.file.originalname}`,
      sheetName: parsed.sheetName,
      dates: Array.from(dates).sort(),
      codes: Array.from(codes),
      grossByCode,
      netByCode,
      unmapped: Array.from(unmapped),
    });
    seedGianMappingDefaults(store, Array.from(codes));
    save(store);

    let successMsg =
      `Da nap file tho "${req.file.originalname}": ${addedCount} giao dich moi` +
      (upgradedCount > 0 ? `, bo sung phi cho ${upgradedCount} giao dich da co truoc do` : "") +
      (dupCount > 0 ? `, bo qua ${dupCount} giao dich da co (trung lap)` : "") +
      `.${UPDATED_NOTE}`;
    if (unmapped.size > 0) {
      successMsg += ` CANH BAO: ${unmapped.size} Chi nhanh chua khop mapping (${Array.from(unmapped).join(", ")}) -- can tai lai file "Payoo thu ho" (muc 3) neu day la cua hang moi.`;
    }
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: bang gian tong hop hang ngay (sheet "gian ") ----------
// Mot bang duy nhat gom ca Momo, Viet QR, Zalo Mini App, VNPay Co so, Payoo
// QR/the: raw text -> Ma cong trinh -> "Thuoc" (kenh) -> co CSE hay khong.
// Day len file nay MOI NGAY se cap nhat lai zvp_gian_list / zvp_offline_diem_map
// / zvp_payoo_diem_map (uu tien du lieu moi khi trung ten) va invoice_diem_alias
// (cho moi dong co raw khac Ma cong trinh, bat ke thuoc kenh nao) -- KHONG
// dung lai lich su cu (moi lan tai la thay the toan bo danh sach goc, nen
// khong lo bi trung dong khi tai lai file da cap nhat).
router.post("/doi-soat/zvp/upload-gian-master", requireAdmin, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const { sheetName, rows } = parseGianMasterSheet(req.file.buffer);
    if (!sheetName) {
      throw new Error('Khong tim thay sheet "gian" trong file nay.');
    }
    if (rows.length === 0) {
      throw new Error('Sheet "gian" khong doc duoc du lieu (can cot "Ma cong trinh" va "Thuoc").');
    }

    store.zvp_gian_master = {
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName,
      rows,
    };

    const onlineRows = rows.filter((r) => normText(r.thuoc).includes("zalo mini app"));
    const offlineRows = rows.filter((r) => normText(r.thuoc).includes("vnpay co so"));
    const payooRows = rows.filter((r) => normText(r.thuoc).includes("payoo"));

    store.zvp_gian_list = mergeGianListWithMaster(store.zvp_gian_list, onlineRows);
    store.zvp_offline_diem_map = mergeDiemMapWithMaster(store.zvp_offline_diem_map, offlineRows);
    store.zvp_payoo_diem_map = mergeDiemMapWithMaster(store.zvp_payoo_diem_map, payooRows);

    // Any row whose raw text differs from its Ma cong trinh is effectively a
    // "invoice ghi ten khac" alias -- shared table, benefits Momo/ZVP/VietQR
    // invoice matching immediately without needing invoices re-uploaded.
    if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
    let aliasAdded = 0;
    rows.forEach((r) => {
      if (normText(r.raw) !== normText(r.maCongTrinh) && store.invoice_diem_alias[r.raw] !== r.maCongTrinh) {
        store.invoice_diem_alias[r.raw] = r.maCongTrinh;
        aliasAdded++;
      }
    });

    // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong rieng.
    seedGianMappingDefaults(
      store,
      rows.map((r) => r.maCongTrinh)
    );

    save(store);

    const cseCount = rows.filter((r) => r.isCse).length;
    let successMsg =
      `Da nap sheet "${sheetName}": ${rows.length} dong (${cseCount} dong CSE) -- ` +
      `Online ${onlineRows.length}, VNPay co so ${offlineRows.length}, Payoo ${payooRows.length}, ` +
      `${aliasAdded} anh xa ten hoa don moi/cap nhat.${UPDATED_NOTE} ` +
      `LUU Y: gian cho Online/Offline/Payoo da UPLOAD TRUOC DO se KHONG tu tach lai theo du lieu moi ` +
      `(so lieu gross da tinh san luc tai) -- neu can tach lai chinh xac, hay tai lai file doanh thu ` +
      `Online/Offline/Payoo tuong ung sau khi nap bang gian nay.`;
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload gop 1 lan: OrderDetails (.xls) + Du lieu bao cao phi theo GD thanh toan (.xlsx) ----------
// Day la 2 file tai thang tu cong VNPay/Zalo -- KHONG can tu tay gop vao
// sheet "Tong hop Zalo App" nua. "Diem thu" = "FUNZONE MINI APP" -> Online
// (Zalo Mini App), tat ca "Diem thu" con lai -> Offline, tu dong tach va
// cap nhat vao ca 2 danh sach upload cung luc. Can da tai len it nhat 1 lan
// file "Tong hop Zalo App" (de co danh sach gian) va 1 lan file "VNpay thu
// ho co so" (de co bang mapping Chi nhanh Offline) truoc do.
router.post(
  "/doi-soat/zvp/upload-combo",
  requireDataEntry, upload.fields([
    { name: "fileOrders", maxCount: 1 },
    { name: "fileFee", maxCount: 1 },
  ]),
  (req, res) => {
    const store = load();
    try {
      const fileOrders = req.files && req.files.fileOrders && req.files.fileOrders[0];
      const fileFee = req.files && req.files.fileFee && req.files.fileFee[0];
      if (!fileOrders || !fileFee) {
        throw new Error('Vui long chon ca 2 file: "OrderDetails..." va "DuLieuBaoCaoPhiTheoGDThanhToan...".');
      }
      if (!store.zvp_online_product_map || Object.keys(store.zvp_online_product_map).length === 0) {
        throw new Error(
          'Chua co bang tra "Ten san pham -> Ma cong trinh" cho Online -- hay tai len file co sheet "noi" o muc 1b truoc, it nhat 1 lan.'
        );
      }
      if (!store.zvp_offline_diem_map || Object.keys(store.zvp_offline_diem_map).length === 0) {
        throw new Error(
          'Chua co bang mapping diem Offline -- hay tai len file "VNPay thu ho co so" o muc 2 truoc, it nhat 1 lan.'
        );
      }

      const orderMap = parseOrderDetailsWorkbook(fileOrders.buffer);
      const parsed = parseFeeReportWorkbook(fileFee.buffer, orderMap);

      const onlineResolved = resolveOnlineGrossByProductMap(parsed.online, store.zvp_online_product_map);
      const offlineResolved = resolveDiemGross(parsed.offline, store.zvp_offline_diem_map);

      // Chi Nhan (2026-07-27): "nếu có tên sản phẩm mới trong file này các
      // ngày mới á thì bạn thêm tên sản phẩm cho tôi vào mục 1c cho tôi đi
      // nha và tôi sẽ điền mã công trình" -- truoc day ten san pham moi
      // (chua co trong bang tra) chi bi CANH BAO 1 lan qua flash message roi
      // mat, khong luu lai vao bang o muc 1c, nen moi lan tai file moi lai
      // phai doi cho Claude/nguoi khac bao lai ten san pham. Tu dong them
      // ngay vao zvp_online_product_map voi Ma cong trinh de TRONG (chi tu
      // dien vao, khong doan) -- dong trong nay se duoc hien canh bao rieng +
      // day len dau bang o muc 1c (xem GET /doi-soat/zvp va view) de de tim.
      let addedNewProductMapEntries = false;
      (onlineResolved.unmappedProducts || []).forEach((prod) => {
        const key = String(prod).trim().replace(/\s+/g, " ");
        if (!store.zvp_online_product_map[key]) {
          store.zvp_online_product_map[key] = { maCongTrinh: "", isCse: false };
          addedNewProductMapEntries = true;
        }
      });

      const comboName = `${fileOrders.originalname} + ${fileFee.originalname}`;
      // Chi Nhan (2026-07-27): dam bao van luu (save) cac dong san pham moi
      // vua tu dong them o tren ngay ca khi ca 2 file nay khong co ma cong
      // trinh nao duoc ghi nhan (vd toan bo la san pham moi, chua khop duoc
      // gi) -- neu khong se roi vao nhanh "khong co gi moi de luu" ben duoi
      // va mat luon cac dong moi vua them (khong duoc ghi xuong dia).
      let addedAny = addedNewProductMapEntries;

      if (onlineResolved.codes.length > 0) {
        if (!isDuplicateRecentUpload(store.zvp_online_uploads, comboName, onlineResolved.grossByCode)) {
          store.zvp_online_uploads.push({
            id: nextId(store, "zvp_online_uploads_seq") || Date.now(),
            uploaded_at: new Date().toISOString(),
            file_name: comboName,
            sheetName: parsed.online.sheetName + " (Online, loc theo Diem thu = FUNZONE MINI APP)",
            dates: onlineResolved.dates,
            codes: onlineResolved.codes,
            grossByCode: onlineResolved.grossByCode,
            netByCode: onlineResolved.netByCode,
            unmapped: onlineResolved.unmappedProducts,
            // Luyen, 2026-07-27: "sao gian AMTP ... nó có 14tr thôi mà" (khong
            // phai cai nay, day la ZVP Online) -- van de thuc te phat hien:
            // giao dich FUNZONE MINI APP khong tim thay so don hang tuong ung
            // trong file OrderDetails bi BO QUA HOAN TOAN (khong con ten san
            // pham de auto-learn Ma cong trinh nhu truong hop unmappedProducts
            // o tren), chi bao 1 LAN qua flash message roi mat -- luu lai day
            // de hien lai thanh canh bao TON TAI tren trang, tranh mat doanh
            // thu ma khong ai biet (verified: file Fee Report 27/07 co 449/484
            // GD Online khong khop vi OrderDetails dung la ban cu 24/07).
            unmatchedOrders: parsed.online.unmatchedOrders,
          });
          seedGianMappingDefaults(store, onlineResolved.codes);
          addedAny = true;
        }
      }

      if (offlineResolved.codes.length > 0) {
        if (!isDuplicateRecentUpload(store.zvp_offline_uploads, comboName, offlineResolved.grossByCode)) {
          store.zvp_offline_uploads.push({
            id: nextId(store, "zvp_offline_uploads_seq") || Date.now(),
            uploaded_at: new Date().toISOString(),
            file_name: comboName,
            sheetName: parsed.offline.sheetName + " (Offline, cac diem con lai)",
            dates: offlineResolved.dates,
            codes: offlineResolved.codes,
            grossByCode: offlineResolved.grossByCode,
            netByCode: offlineResolved.netByCode,
            unmapped: offlineResolved.unmapped,
          });
          seedGianMappingDefaults(store, offlineResolved.codes);
          addedAny = true;
        }
      }

      if (!addedAny) {
        return res.redirect(
          "/doi-soat/zvp?success=" + encodeURIComponent("2 file nay vua duoc tai len roi (bo qua ban trung lap).")
        );
      }

      save(store);

      let successMsg =
        `Da nap 2 file: Online ${parsed.online.rowsMatched} giao dich (${onlineResolved.codes.length} ma cong trinh), ` +
        `Offline ${parsed.offline.rowsMatched} giao dich (${offlineResolved.codes.length} ma cong trinh).${UPDATED_NOTE}`;
      if (parsed.online.unmatchedOrders.length > 0) {
        successMsg += ` CANH BAO: ${parsed.online.unmatchedOrders.length} giao dich Online khong tim thay don hang tuong ung trong file OrderDetails (${parsed.online.unmatchedOrders.slice(0, 5).join(", ")}) -- co the do file OrderDetails chua du ngay.`;
      }
      if (onlineResolved.unmappedProducts.length > 0) {
        successMsg += ` CANH BAO Online: ${onlineResolved.unmappedProducts.length} TEN SAN PHAM MOI chua co Ma cong trinh (${onlineResolved.unmappedProducts.join(" | ")}) -- da tu dong them vao bang o muc 1c (dong de trong o dau bang), chi dien Ma cong trinh dung vao do roi tai lai 2 file nay 1 lan nua de ap dung cho doanh thu.`;
      }
      if (offlineResolved.unmapped.length > 0) {
        successMsg += ` CANH BAO Offline: ${offlineResolved.unmapped.length} diem chua khop gian (${offlineResolved.unmapped.join(", ")}).`;
      }
      res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
    } catch (e) {
      res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
    }
  }
);

// ---------- Upload: danh sach hoa don dung chung (file MTT) ----------
// Upload 1 file tren TRANG NAY cung cap nhat luon ca hoa don Momo (dung chung
// parseSharedInvoiceWorkbook voi trang /doi-soat/momo) -- khong can upload lai
// file nay tren trang kia.
router.post("/doi-soat/zvp/upload-hoadon", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const shared = parseSharedInvoiceWorkbook(req.file.buffer, getCompany(req));

    const addedCounts = {};
    for (const key of ["zalo", "vnpay", "payoo"]) {
      const existingKeys = new Set(store.zvp_invoices[key].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
      let added = 0;
      for (const inv of shared[key]) {
        const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
        if (existingKeys.has(k)) continue;
        existingKeys.add(k);
        store.zvp_invoices[key].push(inv);
        added++;
      }
      addedCounts[key] = added;
    }

    if (!store.momo_invoices) store.momo_invoices = [];
    const existingKeysMomo = new Set(store.momo_invoices.map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
    let addedMomo = 0;
    for (const inv of shared.momo) {
      const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
      if (existingKeysMomo.has(k)) continue;
      existingKeysMomo.add(k);
      store.momo_invoices.push(inv);
      addedMomo++;
    }

    // Auto-learn any invoice-only site name ("Ten diem xuat hoa don") that
    // zvp_gian_list doesn't already know about -- same convention as
    // resolveProductCatalogGross's Online learnedGian, so a brand-new gian
    // seen ONLY on an invoice (never on the "Danh muc ten san pham" sheet)
    // also shows up in "1b. Ra soat gian moi" for a one-time review instead
    // of silently staying unmatched invoice after invoice.
    if (!store.zvp_gian_list) store.zvp_gian_list = [];
    const learnedFromInvoices = learnGianFromInvoices(
      [...store.zvp_invoices.zalo, ...store.zvp_invoices.vnpay, ...store.zvp_invoices.payoo],
      store.zvp_gian_list
    );
    if (learnedFromInvoices.length > 0) {
      const existingGianKeys = new Set(store.zvp_gian_list.map((g) => normText(g.tenDiem)));
      for (const g of learnedFromInvoices) {
        if (!existingGianKeys.has(normText(g.tenDiem))) {
          store.zvp_gian_list.push(g);
          existingGianKeys.add(normText(g.tenDiem));
        }
      }
    }

    save(store);
    let successMsg = `Da nap sheet "${shared.sheetName}": them moi ${addedCounts.zalo} HD zalo, ${addedCounts.vnpay} HD vnpay, ${addedCounts.payoo} HD payoo, ${addedMomo} HD momo (da cap nhat cho ca 2 trang Doi soat Momo va Zalo/VNPay/Payoo).${UPDATED_NOTE}`;
    if (learnedFromInvoices.length > 0) {
      successMsg += ` Da tu dong ghi nho ${learnedFromInvoices.length} ten diem moi tren hoa don chua co trong danh sach gian (${learnedFromInvoices.map((g) => g.tenDiem).slice(0, 5).join(", ")}${learnedFromInvoices.length > 5 ? "..." : ""}) -- kiem tra o muc 1b neu can gan lai ma cong trinh.`;
    }
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/zvp/mapping", requireAdmin, (req, res) => {
  const store = load();
  ensureZvpGianMapping(store);
  const body = req.body || {};
  for (const [key, val] of Object.entries(body)) {
    if (key.startsWith("tkco_")) {
      const code = key.slice("tkco_".length);
      if (TKCO_VALUES.includes(val)) store.zvp_gian_mapping[code] = val;
    }
  }
  save(store);
  res.redirect("/doi-soat/zvp?success=" + encodeURIComponent("Da luu bang TK Co theo gian."));
});

// ---------- Alias: "Ma diem" tren hoa don ghi ten khac (vd "SNOWFUN TAN
// PHU") nhung thuc chat cung 1 Ma Cong Trinh voi doanh thu (vd "AM TP KVCM")
// -- ap dung ngay luc doi soat, khong can tai lai file hoa don. Bang nay
// dung CHUNG voi trang doi-soat/momo (store.invoice_diem_alias). ----------
router.post("/doi-soat/zvp/diem-alias", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { sourceCode, targetCode } = req.body;
    if (!sourceCode || !targetCode) throw new Error("Thieu ma diem tren hoa don hoac ma cong trinh de anh xa.");
    if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
    store.invoice_diem_alias[sourceCode] = targetCode;
    save(store);
    res.redirect(
      "/doi-soat/zvp?success=" + encodeURIComponent(`Da anh xa "${sourceCode}" -> "${targetCode}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/zvp/diem-alias/delete", requireAdmin, (req, res) => {
  const store = load();
  const { sourceCode } = req.body;
  if (store.invoice_diem_alias) delete store.invoice_diem_alias[sourceCode];
  save(store);
  res.redirect("/doi-soat/zvp?success=" + encodeURIComponent("Da xoa anh xa ma diem."));
});

router.post("/doi-soat/zvp/upload-online/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  store.zvp_online_uploads = store.zvp_online_uploads.filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/zvp");
});

router.post("/doi-soat/zvp/upload-offline/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  store.zvp_offline_uploads = store.zvp_offline_uploads.filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/zvp");
});

router.post("/doi-soat/zvp/upload-payoo/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  store.zvp_payoo_uploads = store.zvp_payoo_uploads.filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/zvp");
});

router.post("/doi-soat/zvp/invoices/clear", requireAdmin, (req, res) => {
  const store = load();
  store.zvp_invoices = { zalo: [], vnpay: [], payoo: [] };
  save(store);
  res.redirect("/doi-soat/zvp?success=" + encodeURIComponent("Da xoa toan bo hoa don zalo/vnpay/payoo da nap."));
});
// ---------- Xoa hoa don cu/trung theo so HD, ho tro @Ma cong trinh de tranh xoa nham HD trung so o gian/kenh khac (giong Momo) ----------
router.post("/doi-soat/zvp/invoices/xoa-theo-so", requireAdmin, (req, res) => {
  const store = load();
  const channel = ["zalo", "vnpay", "payoo"].includes(req.body.channel) ? req.body.channel : null;
  if (!channel) {
    return res.redirect("/doi-soat/zvp?error=" + encodeURIComponent("Thieu kenh (zalo/vnpay/payoo) de xoa hoa don."));
  }
  const raw = (req.body.soHdList || "").trim();
  if (!raw) {
    return res.redirect("/doi-soat/zvp?error=" + encodeURIComponent("Chua nhap so HD can xoa."));
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
  if (!store.zvp_invoices[channel]) store.zvp_invoices[channel] = [];
  const before = store.zvp_invoices[channel].length;
  store.zvp_invoices[channel] = store.zvp_invoices[channel].filter(
    (i) => !targets.some((t) => String(i.soHd) === t.soHd && (t.maDiem === null || i.maDiem === t.maDiem))
      );
  const removed = before - store.zvp_invoices[channel].length;
  save(store);
  const labels = targets.map((t) => (t.maDiem ? `${t.soHd}@${t.maDiem}` : t.soHd));
  res.redirect(
    "/doi-soat/zvp?success=" +
    encodeURIComponent(`Da xoa ${removed} hoa don (${channel}) theo so HD: ${labels.join(", ")}.`)
    );
});


// ---------- Xoa hoa don cu/trung theo so HD, ho tro @Ma cong trinh de tranh
// xoa nham HD trung so o gian/kenh khac -- giong het co che da lam ben Momo
// (routes/doisoat.js, "/doi-soat/momo/invoices/xoa-theo-so"). channel = zalo/
// vnpay/payoo, chon dung kenh dang xem tren trang de xoa dung mang. ----------
router.post("/doi-soat/zvp/invoices/xoa-theo-so", requireAdmin, (req, res) => {
  const store = load();
  const channel = ["zalo", "vnpay", "payoo"].includes(req.body.channel) ? req.body.channel : null;
  if (!channel) {
    return res.redirect("/doi-soat/zvp?error=" + encodeURIComponent("Thieu kenh (zalo/vnpay/payoo) de xoa hoa don."));
  }
  const raw = (req.body.soHdList || "").trim();
  if (!raw) {
    return res.redirect("/doi-soat/zvp?error=" + encodeURIComponent("Chua nhap so HD can xoa."));
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
  if (!store.zvp_invoices[channel]) store.zvp_invoices[channel] = [];
  const before = store.zvp_invoices[channel].length;
  store.zvp_invoices[channel] = store.zvp_invoices[channel].filter(
    (i) => !targets.some((t) => String(i.soHd) === t.soHd && (t.maDiem === null || i.maDiem === t.maDiem))
  );
  const removed = before - store.zvp_invoices[channel].length;
  save(store);
  const labels = targets.map((t) => (t.maDiem ? `${t.soHd}@${t.maDiem}` : t.soHd));
  res.redirect(
    "/doi-soat/zvp?success=" +
      encodeURIComponent(`Da xoa ${removed} hoa don (${channel}) theo so HD: ${labels.join(", ")}.`)
  );
});

// ---------- Manual match: dong "Chua co HD" ma Luyen da xac nhan la co HD bu (thuong la ngay hom sau) ----------
// Dung cho truong hop 1 gian khong di qua dò hoa don tu dong cho ky doi soat
// nay (vd hoa don chi duoc xuat ngay hom sau, ngoai khoang ngay cua ky nay),
// nhung Luyen da tu kiem tra va biet chac hoa don nao bu cho khoan tien nay.
router.post("/doi-soat/zvp/manual-match", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { channel, settlementDate, code, invoiceNumbers, amount, grossAdjustment, note } = req.body;
    if (!["online", "offline", "payoo"].includes(channel)) throw new Error("Kenh khong hop le.");
    if (!settlementDate || !code) throw new Error("Thieu thong tin dong can danh dau.");
    if (!store.zvp_manual_matches) store.zvp_manual_matches = { online: {}, offline: {}, payoo: {} };
    if (!store.zvp_manual_matches[channel]) store.zvp_manual_matches[channel] = {};
    const key = `${settlementDate}|${code}`;
    const invoiceList = (invoiceNumbers || "")
      .split(/[,\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const amt = amount ? Number(String(amount).replace(/[^\d-]/g, "")) : null;
    const grossAdj = grossAdjustment ? Number(String(grossAdjustment).replace(/[^\d-]/g, "")) : 0;
    store.zvp_manual_matches[channel][key] = {
      invoiceNumbers: invoiceList,
      amount: amt,
      grossAdjustment: grossAdj,
      note: note || "",
      created_at: new Date().toISOString(),
    };
    save(store);
    res.redirect(
      "/doi-soat/zvp?success=" + encodeURIComponent(`Da danh dau thu cong dong "${code}" ngay ${settlementDate}.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/zvp/manual-match/delete", requireAdmin, (req, res) => {
  const store = load();
  try {
    const { channel, settlementDate, code } = req.body;
    if (store.zvp_manual_matches && store.zvp_manual_matches[channel]) {
      delete store.zvp_manual_matches[channel][`${settlementDate}|${code}`];
    }
    save(store);
    res.redirect("/doi-soat/zvp?success=" + encodeURIComponent("Da xoa danh dau thu cong."));
  } catch (e) {
    res.redirect("/doi-soat/zvp?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Export: "MISATHuế" (Online) va "MISAThue OFFLINE" (Offline + Payoo) ----------
// Same "Mau phieu thu tien gui de nhap vao AMIS Accounting" 28-column layout
// as Momo's export. Gian mapped to "SKIP" are excluded from both files.
// suffixKenh (Luyen, 2026-07-16): hau to phan biet kenh ngay tren Dien giai
// cua file xuat Misa -- "VNP" cho Zalo App (Online) va VNPay offline, "PAYOO"
// rieng cho Payoo (Momo dung "MM", Viet QR dung "QR" o 2 file route khac).
// maDoiTuong (Luyen, 2026-07-20, cap nhat 2026-07-31: "mã đối tượng zalo app
// với vn pay offline điều là này ... VN PAY0102182292 ... còn của Payoo là
// DONGVIET0305458683"): Zalo App (Online) VA VNPay Offline dung CHUNG 1 ma
// "VN PAY0102182292" (CTY CP GIAI PHAP THANH TOAN VIET NAM), rieng Payoo
// dung "DONGVIET0305458683" (CTY CP DICH VU TRUC TUYEN CONG DONG...) --
// KHONG con dung "KL" (Khach le) mac dinh cho 2 kenh nay nua. "Tên đối tượng"
// van de trong ("") nhu truoc, de MISA tu tra theo Mã đối tượng trong danh
// muc khach hang cua no. Xem 3 lan goi ham nay o duoi (export-online/
// export-offline) de biet kenh nao truyen gi.
function buildExportRows(reconciledList, startNo, lyDoThu, suffixKenh, maDoiTuong) {
  const doiTuong = maDoiTuong || "KL";
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
        const dienGiai = hdText
          ? `Thu tiền dịch vụ vui chơi giải trí - ${suffixKenh} theo HĐ ${hdText}`
          : `Thu tiền dịch vụ vui chơi giải trí - ${suffixKenh}`;
        rows.push({
          "Ngày hạch toán (*)": ngayDmy,
          "Ngày chứng từ (*)": ngayDmy,
          "Số chứng từ (*)": soCt,
          "Mã đối tượng": doiTuong,
          "Tên đối tượng": "",
          "Địa chỉ": "",
          "Nộp vào TK": ZVP_BANK_ACCOUNT,
          "Mở tại ngân hàng": ZVP_BANK_FULLNAME,
          "Lý do thu": lyDoThu,
          "Diễn giải lý do thu": dienGiai,
          "Mã nhân viên thu": "",
          "Diễn giải (hạch toán)": dienGiai,
          "TK Nợ (*)": 1121,
          "TK Có (*)": l.tkCo,
          "Số tiền": l.net,
          "Mã đối tượng (hạch toán)": doiTuong,
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
  return { rows, nextSeq: seq };
}

// Luyen, 2026-07-31: "tu ngay may toi ngay may" -- loc theo khoang ngay
// (query "tu"/"den", ISO yyyy-mm-dd) truoc khi xuat, dung chung cho ca 2
// export (online/offline) ben duoi. Khong truyen tu/den thi van xuat toan bo
// nhu cu (tuong thich nguoc).
function filterByDateRange(list, tuFilter, denFilter) {
  let out = list;
  if (tuFilter) out = out.filter((r) => r.settlementDate >= tuFilter);
  if (denFilter) out = out.filter((r) => r.settlementDate <= denFilter);
  return out;
}

router.get("/doi-soat/zvp/export-online.xlsx", (req, res) => {
  const store = load();
  const built = buildReconciliation(store);
  if (built.error) return res.status(400).send(built.error);

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;
  const onlineForExport = filterByDateRange(built.reconciled.online, req.query.tu || "", req.query.den || "");
  // Luyen, 2026-07-21: "lý do thu là Thu tiền khách hàng (không theo hóa đơn)
  // đổi hết các file xuất misa nhá" -- dung 1 cau CO DINH giong het Momo, bo
  // cau rieng theo tung kenh nhu truoc.
  const { rows } = buildExportRows(onlineForExport, startNo, "Thu tiền khách hàng (không theo hóa đơn)", "VNP", "VN PAY0102182292");

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, "MISATHue Online");
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=MISATHue-online.xlsx");
  res.send(buf);
});

router.get("/doi-soat/zvp/export-offline.xlsx", (req, res) => {
  const store = load();
  const built = buildReconciliation(store);
  if (built.error) return res.status(400).send(built.error);

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;
  const tuFilter = req.query.tu || "";
  const denFilter = req.query.den || "";
  // Offline VNPay + Payoo folded into the SAME file (per Luyen: Payoo duoc
  // xu ly don gian giong nhu Offline), continuing the same document-number
  // sequence across both channels.
  const offlinePart = buildExportRows(
    filterByDateRange(built.reconciled.offline, tuFilter, denFilter),
    startNo,
    "Thu tiền khách hàng (không theo hóa đơn)",
    "VNP",
    "VN PAY0102182292"
  );
  const payooPart = buildExportRows(
    filterByDateRange(built.reconciled.payoo, tuFilter, denFilter),
    offlinePart.nextSeq,
    "Thu tiền khách hàng (không theo hóa đơn)",
    "PAYOO",
    "DONGVIET0305458683"
  );
  const rows = [...offlinePart.rows, ...payooPart.rows];

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, "MISAThue OFFLINE");
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=MISATHue-offline.xlsx");
  res.send(buf);
});

// Exposed so routes/dashboard.js (Tong quan / Cong no) can reuse the exact
// same reconciliation this page shows, without a second implementation.
router.buildZvpReconciliation = buildReconciliation;
router.ZVP_BANK_NAME = ZVP_BANK_NAME;

module.exports = router;
