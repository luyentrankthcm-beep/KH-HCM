const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const AdmZip = require("adm-zip");
const { load, save, nextId } = require("../store");
const { requireLogin } = require("../middleware/auth");
const {
  parseTongMomoWorkbook,
  parseInvoiceWorkbook,
  parseRawMomoPortalWorkbook,
  resolveRawPortalGross,
  reconcileMomo,
  extractMomoSettlements,
  isoToDmy,
} = require("../utils/momoReconcile");
const { parseSharedInvoiceWorkbook } = require("../utils/zvpReconcile");
const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

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
    label: "BIDV 7701",
    exportSheet: "MISATHUE7701",
  },
};

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
  const sorted = [...uploads].sort(
    (a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at)
  );
  for (const u of sorted) {
    (u.codes || []).forEach((c) => codes.add(c));
    for (const [k, v] of Object.entries(u.grossByCode || {})) {
      grossByCode[k] = v; // newer upload overwrites, does not add
    }
  }
  return { codes: Array.from(codes), grossByCode };
}

// Detect a .zip upload (either by extension or by the zip magic bytes "PK"),
// unzip it in memory, and return the buffer of the first .xlsx/.xls entry
// found inside. This is how the raw MoMo merchant-portal "daily_report" is
// distributed (a zip containing one xlsx).
function isZipFile(file) {
  if (/\.zip$/i.test(file.originalname || "")) return true;
  const buf = file.buffer;
  return buf && buf.length > 2 && buf[0] === 0x50 && buf[1] === 0x4b; // "PK"
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
function buildMomoReconciliation(store, companyKey) {
  const cfg = MOMO_CHANNELS[companyKey] || MOMO_CHANNELS.kh_cu;
  const grossUploads = store[cfg.grossKey] || [];
  const invoices = store[cfg.invoicesKey] || [];
  const bank = store.banks.find((b) => b.name === cfg.bankName);

  let reconciledAll = [];
  let error = null;
  let allCodes = new Set();

  if (bank) {
    const txs = store.transactions.filter((t) => t.bank_id === bank.id);
    const settlements = extractMomoSettlements(txs);
    const grossData = mergeGross(grossUploads);
    const invoiceData = { invoices };
    if (grossData.codes.length > 0) {
      reconciledAll = reconcileMomo(settlements, grossData, invoiceData, store.gian_mapping, store.invoice_diem_alias);
      reconciledAll.forEach((r) => r.lines.forEach((l) => allCodes.add(l.code)));
    }
  } else {
    error = `Chua co ngan hang "${cfg.bankName}" trong he thong.`;
  }

  // Any invoice "Ma diem" that never lines up with a known gian code (and
  // isn't already aliased) is surfaced here -- these invoices exist and have
  // real money on them, they're just filed under a name the revenue side
  // doesn't recognize (e.g. a sub-brand/corner name like "SNOWFUN TAN PHU"
  // instead of the Ma Cong Trinh "AM TP KVCM"), so the reconciliation above
  // silently can't count them yet. See mục 3b on the page: cho phép Luyen
  // ánh xạ tên này về đúng Ma Cong Trinh, áp dụng ngay không cần tải lại HĐ.
  const knownCodes = new Set([...allCodes, ...Object.keys(store.gian_mapping || {})]);
  const invoiceDiemAlias = store.invoice_diem_alias || {};
  const unmatchedInvoiceCodesSet = new Set();
  invoices.forEach((inv) => {
    if (!inv.maDiem) return;
    if (knownCodes.has(inv.maDiem)) return;
    if (invoiceDiemAlias[inv.maDiem]) return;
    unmatchedInvoiceCodesSet.add(inv.maDiem);
  });

  return {
    reconciledAll,
    allCodes,
    error,
    invoiceDiemAlias,
    unmatchedInvoiceCodes: Array.from(unmatchedInvoiceCodesSet).sort(),
  };
}

router.get("/doi-soat/momo", (req, res) => {
  const store = load();
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
  const reconciled = selectedMonth
    ? reconciledAll.filter((r) => r.settlementDate.slice(0, 7) === selectedMonth)
    : reconciledAll;

  res.render("doisoat-momo", {
    userName: req.session.userName,
    bankLabel: momoCfg.label,
    uploads: store[momoCfg.grossKey] || [],
    invoiceCount: (store[momoCfg.invoicesKey] || []).length,
    reconciled,
    months,
    selectedMonth,
    gianMapping: store.gian_mapping,
    allCodes: Array.from(allCodes).sort(),
    invoiceDiemAlias: built.invoiceDiemAlias,
    unmatchedInvoiceCodes: built.unmatchedInvoiceCodes,
    error,
    success: req.query.success || null,
  });
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
      // The "SC VIVO KVCM" FF (chia se) gian defaults to 1388, everything
      // else (including the normal SC VIVO KVCM line) defaults to 131.
      store.gian_mapping[c] = c.endsWith("__FF") ? "1388" : "131";
    }
  });
}

router.post("/doi-soat/momo/upload-tong", upload.single("file"), (req, res) => {
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
      const parsed = parseTongMomoWorkbook(req.file.buffer);
      store[momoCfg.grossKey].push({
        id: nextId(store, "momo_gross_uploads_seq") || Date.now(),
        uploaded_at: new Date().toISOString(),
        file_name: req.file.originalname,
        sheetName: parsed.sheetName,
        dates: parsed.dates,
        codes: parsed.codes,
        grossByCode: parsed.grossByCode,
      });
      seedGianMappingDefaults(store, parsed.codes);
      // Learn/refresh the Ma cua hang -> Ma Cong trinh mapping from this
      // upload so future raw MoMo portal zip exports can be resolved
      // automatically without needing this sheet uploaded again.
      Object.assign(store.cua_hang_mapping, parsed.cuaHangMap || {});
      save(store);
      successMsg = `Da nap "${parsed.sheetName}" (${parsed.dates[0]} - ${parsed.dates[parsed.dates.length - 1]}), ${parsed.codes.length} ma cong trinh/gian. Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.`;
    }

    res.redirect("/doi-soat/momo?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

// Upload 1 file danh sach hoa don (MTT) tren trang Momo cung cap nhat luon ca
// 3 danh sach hoa don Zalo/VNPay/Payoo (dung chung parseSharedInvoiceWorkbook
// voi trang /doi-soat/zvp) -- khong can upload lai file nay tren trang kia.
router.post("/doi-soat/momo/upload-hoadon", upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const momoCfg = MOMO_CHANNELS[activeCompany];
  if (!store[momoCfg.invoicesKey]) store[momoCfg.invoicesKey] = [];
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const shared = parseSharedInvoiceWorkbook(req.file.buffer);

    const existingKeysMomo = new Set(store[momoCfg.invoicesKey].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
    let addedMomo = 0;
    for (const inv of shared.momo) {
      const key = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
      if (existingKeysMomo.has(key)) continue;
      existingKeysMomo.add(key);
      store[momoCfg.invoicesKey].push(inv);
      addedMomo++;
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
    if (activeCompany === "kh_cu") {
      msg += ` ${addedCounts.zalo} HD zalo, ${addedCounts.vnpay} HD vnpay, ${addedCounts.payoo} HD payoo (da cap nhat cho ca 2 trang Doi soat Momo va Zalo/VNPay/Payoo).`;
    }
    msg += " Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.";
    res.redirect("/doi-soat/momo?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/momo?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/momo/mapping", (req, res) => {
  const store = load();
  const body = req.body || {};
  for (const [key, val] of Object.entries(body)) {
    if (key.startsWith("tkco_")) {
      const code = key.slice("tkco_".length);
      if (TKCO_VALUES.includes(val)) store.gian_mapping[code] = val;
    }
  }
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent("Da luu bang TK Co theo gian."));
});

// ---------- Alias: "Ma diem" tren hoa don ghi ten khac (vd "SNOWFUN TAN
// PHU") nhung thuc chat cung 1 Ma Cong Trinh voi doanh thu (vd "AM TP KVCM")
// -- ap dung ngay luc doi soat, khong can tai lai file hoa don. Bang nay
// dung CHUNG voi trang doi-soat/zvp (store.invoice_diem_alias). ----------
router.post("/doi-soat/momo/diem-alias", (req, res) => {
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

router.post("/doi-soat/momo/diem-alias/delete", (req, res) => {
  const store = load();
  const { sourceCode } = req.body;
  if (store.invoice_diem_alias) delete store.invoice_diem_alias[sourceCode];
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent("Da xoa anh xa ma diem."));
});

router.delete("/doi-soat/momo/upload-tong/:id", (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  store[momoCfg.grossKey] = (store[momoCfg.grossKey] || []).filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/momo");
});

router.post("/doi-soat/momo/upload-tong/:id/delete", (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  store[momoCfg.grossKey] = (store[momoCfg.grossKey] || []).filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/momo");
});

router.post("/doi-soat/momo/invoices/clear", (req, res) => {
  const store = load();
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  store[momoCfg.invoicesKey] = [];
  save(store);
  res.redirect("/doi-soat/momo?success=" + encodeURIComponent(`Da xoa toan bo hoa don momo da nap (${momoCfg.label}).`));
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
  const momoCfg = MOMO_CHANNELS[getCompany(req)];
  const bank = store.banks.find((b) => b.name === momoCfg.bankName);
  if (!bank) return res.status(400).send(`Chua co ngan hang ${momoCfg.bankName}.`);

  const txs = store.transactions.filter((t) => t.bank_id === bank.id);
  const settlements = extractMomoSettlements(txs);
  const grossData = mergeGross(store[momoCfg.grossKey] || []);
  const invoiceData = { invoices: store[momoCfg.invoicesKey] || [] };
  const reconciled = reconcileMomo(settlements, grossData, invoiceData, store.gian_mapping, store.invoice_diem_alias);

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;
  let seq = startNo;

  const rows = [];
  reconciled
    .sort((a, b) => (a.settlementDate > b.settlementDate ? 1 : -1))
    .forEach((r) => {
      const exportableLines = r.lines.filter((l) => l.tkCo !== "SKIP");
      if (exportableLines.length === 0) return; // toan bo la KH moi -- khong tao chung tu rong

      const soCt = "NTTK" + String(seq).padStart(7, "0") + "/26";
      seq++;
      const ngayDmy = isoToDmy(r.settlementDate);
      exportableLines.forEach((l) => {
        const hdText = l.invoiceNumbers.length > 0 ? l.invoiceNumbers.join(", ") : "";
        const dienGiai = hdText
          ? `Thu tiền dịch vụ vui chơi giải trí theo HĐ ${hdText}`
          : "Thu tiền dịch vụ vui chơi giải trí";
        rows.push({
          "Ngày hạch toán (*)": ngayDmy,
          "Ngày chứng từ (*)": ngayDmy,
          "Số chứng từ (*)": soCt,
          "Mã đối tượng": "KL",
          "Tên đối tượng": "",
          "Địa chỉ": "",
          "Nộp vào TK": momoCfg.bankAccount,
          "Mở tại ngân hàng": momoCfg.bankFullName,
          "Lý do thu": "Thu tiền khách hàng (không theo hóa đơn)",
          "Diễn giải lý do thu": dienGiai,
          "Mã nhân viên thu": "",
          "Diễn giải (hạch toán)": dienGiai,
          "TK Nợ (*)": 112,
          "TK Có (*)": l.tkCo,
          "Số tiền": l.net,
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
          "Doanh thu gộp (trước phí)": l.gross,
          "Chênh lệch HĐ vs doanh thu": l.diff,
          "Trạng thái": l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : l.matched ? "Khớp" : "Lệch",
        });
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
