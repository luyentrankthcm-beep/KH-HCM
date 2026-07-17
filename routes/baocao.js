const express = require("express");
const XLSX = require("xlsx");
const { load } = require("../store");
const { requireLogin } = require("../middleware/auth");
const momo = require("../utils/momoReconcile");
const zvp = require("../utils/zvpReconcile");
const overviewAggregate = require("../utils/overviewAggregate");
const { COMPANIES } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

const MOMO_BANK_NAME = "BIDV123456";
const ZVP_BANK_NAME = "ACB31268";

// ---------- Shared helpers ----------

// Same "newest upload wins per (ngay, code) key" merge used by both the Momo
// and Zalo/VNPay/Payoo reconciliation pages -- duplicated here (rather than
// imported) because the originals are private helpers inside routes/doisoat.js
// and routes/doisoat-zvp.js, not exported from those files.
function mergeGrossNet(uploads) {
  const grossByCode = {};
  const netByCode = {};
  const sorted = [...(uploads || [])].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  for (const u of sorted) {
    for (const [k, v] of Object.entries(u.grossByCode || {})) grossByCode[k] = v;
    for (const [k, v] of Object.entries(u.netByCode || {})) netByCode[k] = v;
  }
  return { grossByCode, netByCode };
}

function monthOf(key) {
  return key.slice(0, 7); // "YYYY-MM-DD|code" -> "YYYY-MM"
}

// hasRealNet=false means this channel has no real per-transaction fee data
// (currently only Momo) -- net is estimated using the SAME flat 1.1% rate
// already used elsewhere in the app's Momo reconciliation display, clearly
// labeled as an estimate rather than presented as real fee data.
function monthlyFeeFromMerged(merged, hasRealNet) {
  const byMonth = {};
  for (const [key, gross] of Object.entries(merged.grossByCode)) {
    if (!gross) continue;
    const month = monthOf(key);
    if (!byMonth[month]) byMonth[month] = { gross: 0, net: 0 };
    byMonth[month].gross += gross;
    const net = hasRealNet ? (merged.netByCode[key] != null ? merged.netByCode[key] : gross) : Math.round(gross * 0.989);
    byMonth[month].net += net;
  }
  return byMonth;
}

function buildFeeRows(store) {
  const onlineMerged = mergeGrossNet(store.zvp_online_uploads);
  const offlineMerged = mergeGrossNet(store.zvp_offline_uploads);
  const payooMerged = mergeGrossNet(store.zvp_payoo_uploads);
  const momoMerged = mergeGrossNet(store.momo_gross_uploads);

  const channels = [
    { label: "VNPay Online (Zalo App)", merged: onlineMerged, hasRealNet: true },
    { label: "VNPay Offline (QR)", merged: offlineMerged, hasRealNet: true },
    { label: "Payoo", merged: payooMerged, hasRealNet: true },
    { label: "Momo (ước tính phí 1.1%)", merged: momoMerged, hasRealNet: false },
  ];

  const rows = [];
  for (const ch of channels) {
    const byMonth = monthlyFeeFromMerged(ch.merged, ch.hasRealNet);
    for (const [month, v] of Object.entries(byMonth)) {
      const fee = v.gross - v.net;
      rows.push({
        month,
        channel: ch.label,
        gross: Math.round(v.gross),
        net: Math.round(v.net),
        fee: Math.round(fee),
        feePct: v.gross > 0 ? (fee / v.gross) * 100 : 0,
        estimated: !ch.hasRealNet,
      });
    }
  }
  rows.sort((a, b) => (a.month === b.month ? a.channel.localeCompare(b.channel) : a.month < b.month ? -1 : 1));
  return rows;
}

router.get("/bao-cao", (req, res) => {
  res.render("baocao-index", { userName: req.session.userName });
});

// ---------- 1) Bang phi VNPay/Momo theo thang ----------
router.get("/bao-cao/phi", (req, res) => {
  const store = load();
  const rows = buildFeeRows(store);
  res.render("baocao-phi", { userName: req.session.userName, rows });
});

router.get("/bao-cao/phi/export.xlsx", (req, res) => {
  const store = load();
  const rows = buildFeeRows(store).map((r) => ({
    "Tháng": r.month,
    "Kênh": r.channel + (r.estimated ? " (ước tính)" : ""),
    "Tổng doanh thu gộp": r.gross,
    "Tổng phí": r.fee,
    "% phí": Number(r.feePct.toFixed(2)),
    "Net (sau phí)": r.net,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Bang phi theo thang");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=bang-phi-theo-thang.xlsx");
  res.send(buf);
});

// ---------- 2) Chuyen tien noi bo giua cac NH ----------
// Detected via the "CTNB <SRC>-<DST>" / "Chuyen tien noi bo" tag that BIDV/ACB
// statements already stamp on internal-transfer transactions. The short
// tokens after CTNB (e.g. "BIDV456", "ACB268", "651") are matched against the
// trailing digits of each of our OWN bank account numbers to identify which
// leg is which; a token that doesn't match any of our own accounts is kept
// as "khac" (an external account -- e.g. a loan or a partner's bank, not a
// transfer between our own accounts).
const CTNB_RE = /CTNB|chuy.n\s*ti.n\s*n.i\s*b./i;

function findBankByToken(banks, token) {
  if (!token) return null;
  const digits = (token.match(/\d+$/) || [""])[0];
  if (!digits || digits.length < 3) return null;
  return banks.find((b) => String(b.account_number).endsWith(digits)) || null;
}

function detectInternalTransfers(store) {
  const bankById = {};
  store.banks.forEach((b) => (bankById[b.id] = b));

  const rows = [];
  for (const t of store.transactions) {
    const desc = t.description || "";
    if (!CTNB_RE.test(desc)) continue;
    const ownBank = bankById[t.bank_id];
    if (!ownBank) continue;

    const m = desc.match(/CTNB[^A-Za-z0-9]{0,6}([A-Za-z]*\d{3,10})\s*-\s*([A-Za-z]*\d{3,10})/i);
    let otherBank = null;
    if (m) {
      const candA = findBankByToken(store.banks, m[1]);
      const candB = findBankByToken(store.banks, m[2]);
      if (candA && candA.id !== ownBank.id) otherBank = candA;
      if (candB && candB.id !== ownBank.id) otherBank = candB;
    }

    rows.push({
      id: t.id,
      bankName: ownBank.name,
      date: t.date,
      type: t.type,
      amount: t.amount,
      description: desc.replace(/\s+/g, " ").trim(),
      otherBankName: otherBank ? otherBank.name : "(không xác định — có thể là TK ngoài công ty)",
      isFullyInternal: !!otherBank,
    });
  }
  rows.sort((a, b) => (a.date === b.date ? 0 : a.date < b.date ? -1 : 1));
  return rows;
}

router.get("/bao-cao/chuyen-tien-noi-bo", (req, res) => {
  const store = load();
  const rows = detectInternalTransfers(store);
  const totalInternal = rows.filter((r) => r.isFullyInternal).reduce((s, r) => s + r.amount, 0);
  res.render("baocao-noibo", { userName: req.session.userName, rows, totalInternal });
});

router.get("/bao-cao/chuyen-tien-noi-bo/export.xlsx", (req, res) => {
  const store = load();
  const rows = detectInternalTransfers(store).map((r) => ({
    "Ngày": r.date,
    "Ngân hàng": r.bankName,
    "Loại": r.type === "chi" ? "Chi (đi)" : "Thu (đến)",
    "Số tiền": r.amount,
    "NH đối ứng": r.otherBankName,
    "Diễn giải": r.description,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Chuyen tien noi bo");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=chuyen-tien-noi-bo.xlsx");
  res.send(buf);
});

// ---------- 3) Thu/chi theo gian, tung ngan hang ----------
// Flattens the existing per-channel reconciliation (Momo + Zalo/VNPay/Payoo)
// into ONE row per (khoan ve NH, gian) pair, across every bank/channel, so
// Luyen can see every individual settlement broken down by gian in one
// place instead of switching between the Momo and Zalo/VNPay/Payoo pages.
function buildThuChiTheoGianRows(store) {
  const rows = [];

  const momoBank = store.banks.find((b) => b.name === MOMO_BANK_NAME);
  if (momoBank) {
    const txs = store.transactions.filter((t) => t.bank_id === momoBank.id);
    const settlements = momo.extractMomoSettlements(txs);
    const grossData = { codes: [], grossByCode: {} };
    const sorted = [...store.momo_gross_uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
    const codesSet = new Set();
    for (const u of sorted) {
      (u.codes || []).forEach((c) => codesSet.add(c));
      for (const [k, v] of Object.entries(u.grossByCode || {})) grossData.grossByCode[k] = v;
    }
    grossData.codes = Array.from(codesSet);
    if (grossData.codes.length > 0) {
      const reconciled = momo.reconcileMomo(settlements, grossData, { invoices: store.momo_invoices }, store.gian_mapping);
      reconciled.forEach((r) => {
        r.lines.forEach((l) => {
          rows.push({
            bankName: momoBank.name,
            channel: "Momo",
            settlementDate: r.settlementDate,
            maCongTrinh: l.maCongTrinh,
            tkCo: l.tkCo,
            gross: l.gross,
            net: l.net,
            invoiceNumbers: l.invoiceNumbers.join(", "),
            trangThai: l.tkCo === "SKIP" ? "KH mới - chưa xuất MISA" : l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : l.matched ? "Khớp" : "Lệch",
          });
        });
      });
    }
  }

  const zvpBank = store.banks.find((b) => b.name === ZVP_BANK_NAME);
  if (zvpBank) {
    const txs = store.transactions.filter((t) => t.bank_id === zvpBank.id);
    const settlements = zvp.extractZvpSettlements(txs);
    const onlineMerged = mergeGrossNet(store.zvp_online_uploads);
    const offlineMerged = mergeGrossNet(store.zvp_offline_uploads);
    const payooMerged = mergeGrossNet(store.zvp_payoo_uploads);
    const codesOf = (m) => Array.from(new Set(Object.keys(m.grossByCode).map((k) => k.slice(k.indexOf("|") + 1))));
    const manualMatches = store.zvp_manual_matches || { online: {}, offline: {}, payoo: {} };
    const reconciled = zvp.reconcileZvp(
      settlements,
      {
        online: { codes: codesOf(onlineMerged), grossByCode: onlineMerged.grossByCode, netByCode: onlineMerged.netByCode },
        offline: { codes: codesOf(offlineMerged), grossByCode: offlineMerged.grossByCode, netByCode: offlineMerged.netByCode },
        payoo: { codes: codesOf(payooMerged), grossByCode: payooMerged.grossByCode, netByCode: payooMerged.netByCode },
      },
      {
        online: { invoices: store.zvp_invoices.zalo },
        offline: { invoices: store.zvp_invoices.vnpay },
        payoo: { invoices: store.zvp_invoices.payoo },
      },
      store.gian_mapping,
      manualMatches
    );
    const channelLabel = { online: "VNPay Online (Zalo App)", offline: "VNPay Offline (QR)", payoo: "Payoo" };
    ["online", "offline", "payoo"].forEach((ch) => {
      reconciled[ch].forEach((r) => {
        r.lines.forEach((l) => {
          rows.push({
            bankName: zvpBank.name,
            channel: channelLabel[ch],
            settlementDate: r.settlementDate,
            maCongTrinh: l.maCongTrinh,
            tkCo: l.tkCo,
            gross: l.gross,
            net: l.net,
            invoiceNumbers: l.invoiceNumbers.join(", "),
            trangThai:
              l.tkCo === "SKIP"
                ? "KH mới - chưa xuất MISA"
                : l.invoiceNumbers.length === 0
                ? "Chưa có HĐ"
                : l.matched
                ? l.manualOverride
                  ? "Khớp (bù thủ công)"
                  : "Khớp"
                : "Lệch",
          });
        });
      });
    });
  }

  rows.sort((a, b) => (a.settlementDate === b.settlementDate ? a.bankName.localeCompare(b.bankName) : a.settlementDate < b.settlementDate ? -1 : 1));
  return rows;
}

router.get("/bao-cao/thu-chi-theo-gian", (req, res) => {
  const store = load();
  const allRows = buildThuChiTheoGianRows(store);
  const bankFilter = req.query.bank || "";

  // Chon theo thang: mac dinh thang gan nhat de bang khong bi qua dai
  // ("nhieu roi qua" -- Luyen), van chon "Tat ca" duoc qua dropdown.
  const monthSet = new Set(allRows.map((r) => r.settlementDate.slice(0, 7)));
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";

  const banks = Array.from(new Set(allRows.map((r) => r.bankName)));
  let filtered = bankFilter ? allRows.filter((r) => r.bankName === bankFilter) : allRows;
  if (selectedMonth) filtered = filtered.filter((r) => r.settlementDate.slice(0, 7) === selectedMonth);

  res.render("baocao-thuchi", {
    userName: req.session.userName,
    rows: filtered,
    banks,
    bankFilter,
    months,
    selectedMonth,
  });
});

router.get("/bao-cao/thu-chi-theo-gian/export.xlsx", (req, res) => {
  const store = load();
  const bankFilter = req.query.bank || "";
  const monthFilter = req.query.month || "";
  let allRows = buildThuChiTheoGianRows(store);
  if (bankFilter) allRows = allRows.filter((r) => r.bankName === bankFilter);
  if (monthFilter) allRows = allRows.filter((r) => r.settlementDate.slice(0, 7) === monthFilter);
  const rows = allRows.map((r) => ({
    "Ngân hàng": r.bankName,
    "Kênh": r.channel,
    "Ngày về NH": r.settlementDate,
    "Mã công trình (gian)": r.maCongTrinh,
    "TK Có": r.tkCo,
    "Doanh thu gộp": r.gross,
    "Tiền về NH (net)": r.net,
    "Số HĐ khớp": r.invoiceNumbers,
    "Trạng thái": r.trangThai,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Thu chi theo gian");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", "attachment; filename=thu-chi-theo-gian.xlsx");
  res.send(buf);
});

// ---------- 4) Trang thai doi soat theo cong ty (bieu do tron) ----------
// Luyen yeu cau 2026-07-17: bao cao rieng cho KH Cu / KH Moi, dang bieu do
// tron/thong ke. Dung LAI ket qua doi soat cua tung kenh qua
// utils/overviewAggregate.buildAllFlatLines (giong /cong-no) roi phan loai
// tung dong theo cong ty dua vao truong "company" da gan san cho tung kenh
// Viet QR trong routes/doisoat-vietqr.js (CHANNELS). Momo va Zalo/VNPay/
// Payoo hien tai CHUA duoc tach rieng theo cong ty trong du lieu (van la
// KH Cu, xem ghi chu trong views/partials/nav.ejs) nen mac dinh xep vao
// "kh_cu" -- khi nao du lieu duoc tach that thi chi can sua ham
// companyOfChannelKey nay, khong dung lai logic doi soat.
function companyOfChannelKey(channelKey, vietqrChannels) {
  if (channelKey.startsWith("vietqr_")) {
    const rawKey = channelKey.slice("vietqr_".length);
    const ch = vietqrChannels[rawKey];
    return (ch && ch.company) || "kh_cu";
  }
  return "kh_cu";
}

function buildTrangThaiDoiSoat(store) {
  // Require ngay trong ham (khong o dau file) de tranh vong lap require --
  // giong cach utils/overviewAggregate.js da lam voi 3 module doi soat.
  const vietqrRouter = require("./doisoat-vietqr");
  const flat = overviewAggregate.buildAllFlatLines(store);

  const blankStatus = () => ({
    "Khớp": { count: 0, amount: 0 },
    "Lệch": { count: 0, amount: 0 },
    "Chưa có HĐ": { count: 0, amount: 0 },
  });
  const byCompany = { kh_cu: blankStatus(), kh_moi: blankStatus() };

  flat.forEach((l) => {
    const company = companyOfChannelKey(l.channelKey, vietqrRouter.VIETQR_CHANNELS);
    const status = l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : overviewAggregate.isResolved(l) ? "Khớp" : "Lệch";
    byCompany[company][status].count++;
    byCompany[company][status].amount += l.gross;
  });

  return byCompany;
}

router.get("/bao-cao/trang-thai-doi-soat", (req, res) => {
  const store = load();
  let error = null;
  let stats = null;
  try {
    stats = buildTrangThaiDoiSoat(store);
  } catch (e) {
    error = e.message;
    console.error("Loi tinh trang thai doi soat:", e);
  }
  res.render("baocao-trangthai", { userName: req.session.userName, stats, COMPANIES, error });
});

module.exports = router;
