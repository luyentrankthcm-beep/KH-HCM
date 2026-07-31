const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save } = require("../store");
const { requireLogin, requireDataEntry } = require("../middleware/auth");
const momo = require("../utils/momoReconcile");
const zvp = require("../utils/zvpReconcile");
const vietqr = require("../utils/vietqrReconcile");
const overviewAggregate = require("../utils/overviewAggregate");
const { COMPANIES, getCompany } = require("../utils/companies");
const { BANK_COMPANY } = require("../utils/bankCompany");

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

const router = express.Router();
router.use(requireLogin);

const MOMO_BANK_NAME = "BIDV123456";
const ZVP_BANK_NAME = "ACB31268";
const VNPAY_KHMOI_BANK_NAME = "VTB982";
// Cac ngan hang dang chay qua kien truc VietQR (routes/doisoat-vietqr.js
// CHANNELS) -- dung chung 1 kieu loc giao dich "thu" hop le (khong phai
// Momo/ngoai le) cua utils/vietqrReconcile.js de xac dinh giao dich nao DA
// duoc dua vao doi soat.
const VIETQR_BANK_NAMES = new Set([
  "BIDV7704",
  "BIDV77020",
  "MB11521268",
  "BIDV7702",
  "BIDV77021",
  "MB02865168",
  "BIDV8613600999",
]);

// Luyen, 2026-07-31: "cái nào trả của bên đối soát thì tô màu xanh nhạt" --
// xac dinh 1 giao dich ngan hang co nam trong danh sach da duoc 1 trong cac
// engine doi soat (Momo / Zalo-VNPay-Payoo / VietQR / VNPay KH Moi) "gom vao"
// hay chua, DUNG LAI chinh cac ham trich xuat da co (khong tu viet lai logic
// nhan dien de tranh lech voi trang doi soat that). Ngan hang KHONG thuoc bat
// ky kenh doi soat nao (vd tai khoan vay, tai khoan noi bo thuan) thi khong
// co khai niem "da doi soat" nen luon tra ve set rong (khong to mau).
function matchedTxIdsForBank(bank, txs) {
  const ids = new Set();
  if (!bank) return ids;
  if (bank.name === MOMO_BANK_NAME) {
    momo.extractMomoSettlements(txs).forEach((s) => ids.add(s.id));
  } else if (bank.name === ZVP_BANK_NAME || bank.name === VNPAY_KHMOI_BANK_NAME) {
    const ext = zvp.extractZvpSettlements(txs);
    ["online", "offline", "payoo"].forEach((k) => (ext[k] || []).forEach((s) => (s.txIds || []).forEach((id) => ids.add(id))));
  } else if (VIETQR_BANK_NAMES.has(bank.name)) {
    vietqr.extractVietQrThuTransactions(txs).forEach((t) => ids.add(t.id));
  }
  return ids;
}

function companyOfBankRow(b) {
  return b.company || BANK_COMPANY[b.name] || "kh_cu";
}

// Sao ke chi tiet tung giao dich, gop tat ca ngan hang, tach tab theo tung
// ngan hang (giong cac sheet rieng trong file Excel Luyen gui). Danh dau
// giao dich nao da duoc dua vao 1 kenh doi soat (matchedTxIdsForBank).
function buildBankStatementTabs(store) {
  const tabs = [];
  const banksSorted = [...store.banks].sort((a, b) => a.name.localeCompare(b.name));
  for (const bank of banksSorted) {
    const txs = store.transactions
      .filter((t) => t.bank_id === bank.id)
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.id - b.id));
    if (txs.length === 0) continue;
    const matchedIds = matchedTxIdsForBank(bank, txs);
    let bal = bank.opening_balance || 0;
    const rows = txs.map((t) => {
      bal += t.type === "thu" ? t.amount : -t.amount;
      return {
        date: t.date,
        type: t.type,
        amount: t.amount,
        description: t.description || "",
        tenDoiUng: t.tenDoiUng || "",
        balance: bal,
        matched: matchedIds.has(t.id),
      };
    });
    tabs.push({
      bankId: bank.id,
      bankName: bank.name,
      company: companyOfBankRow(bank),
      openingBalance: bank.opening_balance || 0,
      rows,
      totalThu: rows.reduce((s, r) => (r.type === "thu" ? s + r.amount : s), 0),
      totalChi: rows.reduce((s, r) => (r.type === "chi" ? s + r.amount : s), 0),
    });
  }
  return tabs;
}

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

// Loc theo cong ty dang chon -- cung quy uoc voi routes/banks.js's
// companyBanks() (bank khong co field "company" thi mac dinh coi la kh_cu).
// Luyen bao 2026-07-20: chon "KH Moi" nhung trang nay van hien TK123456/1268
// (tai khoan cua KH Cu) vi truoc gio chua loc gi ca -- fix bang cach chi xet
// giao dich cua cac ngan hang THUOC cong ty dang xem, va chi doi chieu voi
// cac ngan hang KHAC cung cong ty do (khong con bat cheo qua TK cong ty kia).
function companyBanks(store, company) {
  return store.banks.filter((b) => (b.company || "kh_cu") === company);
}

function detectInternalTransfers(store, company) {
  const banks = companyBanks(store, company);
  const bankById = {};
  banks.forEach((b) => (bankById[b.id] = b));

  const rows = [];
  for (const t of store.transactions) {
    const desc = t.description || "";
    if (!CTNB_RE.test(desc)) continue;
    const ownBank = bankById[t.bank_id];
    if (!ownBank) continue;

    const m = desc.match(/CTNB[^A-Za-z0-9]{0,6}([A-Za-z]*\d{3,10})\s*-\s*([A-Za-z]*\d{3,10})/i);
    let otherBank = null;
    if (m) {
      const candA = findBankByToken(banks, m[1]);
      const candB = findBankByToken(banks, m[2]);
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

// Luyen bao 2026-07-31: "cho chỗ chọn thời gian đi ưu tiên hiển thị thnág
// hiện tại nhá" -- them bo loc thang, mac dinh uu tien THANG HIEN TAI (neu co
// giao dich), khong con thang hien tai thi lui ve thang gan nhat co du lieu,
// van giu tuy chon "Tat ca" de xem full lich su nhu truoc gio.
function pickDefaultMonth(availableMonths) {
  const currentMonth = new Date().toISOString().slice(0, 7);
  if (availableMonths.includes(currentMonth)) return currentMonth;
  return availableMonths[0] || "";
}

router.get("/bao-cao/chuyen-tien-noi-bo", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const allRows = detectInternalTransfers(store, activeCompany);

  const monthSet = new Set();
  allRows.forEach((r) => { const m = (r.date || "").slice(0, 7); if (m) monthSet.add(m); });
  const availableMonths = [...monthSet].sort().reverse();
  const thangFilter = req.query.thang !== undefined ? req.query.thang : pickDefaultMonth(availableMonths);
  const rows = thangFilter ? allRows.filter((r) => (r.date || "").slice(0, 7) === thangFilter) : allRows;

  const totalInternal = rows.filter((r) => r.isFullyInternal).reduce((s, r) => s + r.amount, 0);
  res.render("baocao-noibo", { userName: req.session.userName, rows, totalInternal, availableMonths, thangFilter });
});

router.get("/bao-cao/chuyen-tien-noi-bo/export.xlsx", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const allRows = detectInternalTransfers(store, activeCompany);
  const monthSet = new Set();
  allRows.forEach((r) => { const m = (r.date || "").slice(0, 7); if (m) monthSet.add(m); });
  const availableMonths = [...monthSet].sort().reverse();
  const thangFilter = req.query.thang !== undefined ? req.query.thang : pickDefaultMonth(availableMonths);
  const rows = (thangFilter ? allRows.filter((r) => (r.date || "").slice(0, 7) === thangFilter) : allRows).map((r) => ({
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
  res.setHeader("Content-Disposition", `attachment; filename=chuyen-tien-noi-bo-${thangFilter || "TatCa"}.xlsx`);
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

  // Sao ke chi tiet tung giao dich, tach tab theo ngan hang -- xem
  // buildBankStatementTabs() o tren. Loc theo cung "thang" da chon o bang tren
  // (neu co) de 2 bang khop nhau; rieng bo loc thang cua bang sao ke van co
  // dropdown tach doc lap (stmtMonth) phong khi Luyen muon xem thang khac voi
  // bang doi soat theo gian o tren.
  const stmtTabsAll = buildBankStatementTabs(store);
  const stmtMonthSet = new Set();
  stmtTabsAll.forEach((tab) => tab.rows.forEach((r) => stmtMonthSet.add(r.date.slice(0, 7))));
  const stmtMonths = Array.from(stmtMonthSet).sort().reverse();
  // Chi dung selectedMonth (thang cua bang doi soat theo gian o tren) lam mac
  // dinh NEU thang do thuc su co giao dich sao ke -- tranh truong hop bang
  // doi soat co dong ngay trong tuong lai (vd khoan ve du kien) nhung sao ke
  // ngan hang thuc te chua toi ngay do, se lam bang sao ke moi trong rong.
  const stmtMonth =
    req.query.stmtMonth !== undefined
      ? req.query.stmtMonth
      : stmtMonths.includes(selectedMonth)
      ? selectedMonth
      : stmtMonths[0] || "";
  const stmtTabsAllWithRows = stmtTabsAll
    .map((tab) => {
      const rows = stmtMonth ? tab.rows.filter((r) => r.date.slice(0, 7) === stmtMonth) : tab.rows;
      return {
        ...tab,
        rows,
        totalThu: rows.reduce((s, r) => (r.type === "thu" ? s + r.amount : s), 0),
        totalChi: rows.reduce((s, r) => (r.type === "chi" ? s + r.amount : s), 0),
      };
    })
    .filter((tab) => tab.rows.length > 0);

  // Tong theo cong ty tinh tren CA 2 cong ty (khong phu thuoc dang xem cong
  // ty nao) de van doi chieu duoc nhanh giua KH Cu / KH Moi cung luc.
  const stmtCompanyTotals = { kh_cu: { thu: 0, chi: 0 }, kh_moi: { thu: 0, chi: 0 } };
  stmtTabsAllWithRows.forEach((tab) => {
    const bucket = stmtCompanyTotals[tab.company] || stmtCompanyTotals.kh_cu;
    bucket.thu += tab.totalThu;
    bucket.chi += tab.totalChi;
  });

  // Luyen, 2026-07-31: "tôi chọn bên kh cũ á thì hiển thị kh cũ thôi" -- danh
  // sach tab/panel hien thi CHI theo cong ty dang chon o thanh tren (giong
  // moi trang khac trong app), dung getCompany(req) nhu thuong le.
  const activeCompany = getCompany(req);
  const stmtTabs = stmtTabsAllWithRows.filter((tab) => tab.company === activeCompany);

  res.render("baocao-thuchi", {
    userName: req.session.userName,
    rows: filtered,
    banks,
    bankFilter,
    months,
    selectedMonth,
    COMPANIES,
    stmtTabs,
    stmtMonths,
    stmtMonth,
    stmtCompanyTotals,
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

// ---------- 5) Xuat "Hoa don ban ra" (mau VietInvoice) gop nhieu doi soat ----------
// Chi Nhan, 2026-07-30: "trong báo cáo này bạn thêm cho tôi thêm 1 cái khung
// nữa là xuất ra mẫu Hóa Đơn Bán Ra nhá cái mẫu lúc trc tôi thêm trong 7702"
// -- kenh BIDV7702 (routes/doisoat-vietqr.js, route /xuat-hoa-don-dau-ra) da
// co san xuat file mau VietInvoice CHO 1 NGAN HANG, 1 dong/gian/ngay. Bao cao
// nay MOI, gop CA NHIEU kenh doi soat cua 1 cong ty lai lam 1 file: chon 1
// khoang Tu ngay-Den ngay, CONG DON tat ca ngay trong khoang do lai, xuat
// DUNG 1 dong cho moi cap (Ma cong trinh, Thuoc doi soat) -- khac voi ban
// BIDV7702 la 1 dong/gian/MOI ngay rieng. Them 2 cot cuoi "Thuộc đối soát"
// (Viet QR <NH> / VN Pay <NH> / Payoo <NH> / Zalo app <NH>) va "Mã công
// trình" de Chi Nhan tu loc/doi chieu lai sau khi xuat.
//
// Chi Nhan xac nhan qua AskUserQuestion (2026-07-30): cot "Ngày hóa đơn" lay
// NGAY XUAT FILE (hom nay, luc bam nut) cho MOI dong, KHONG phai tu/den ngay
// cua khoang loc doanh thu -- khac voi ban BIDV7702 (lay dung ngay giao dich
// vi do la 1 dong/ngay rieng, con ban nay la 1 dong GOP NHIEU NGAY nen khong
// co "1 ngay" duy nhat de dien).
//
// KH Moi lam TRUOC (Chi Nhan xac nhan): gom VietQR (bidv7702/bidv77021/
// mb02865168/bidv8613600999, moi kenh 1 nhan "Viet QR <ten TK ngan hang>")
// va VNPay/Payoo KH Moi (routes/doisoat-vnpay-khmoi.js, ca 2 kenh cung ve TK
// VTB982 -- "VN Pay VTB982" va "Payoo VTB982"). KH Cu (Momo/ZVP) Chi Nhan se
// gui them thong tin sau ("KH cũ tôi gửi sau"), tam thoi bao loi ro rang neu
// chon KH Cu thay vi doan bua nhan/kenh nao dung.
const HOA_DON_BAN_RA_HEADER_INFO = [
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
const HOA_DON_BAN_RA_HEADER_ROW = [
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
  "Thuộc đối soát",
  "Mã công trình",
];
const HOA_DON_BAN_RA_VAT_RATE = 0.08;

// Chi Nhan, 2026-07-30: gui file mau "Copy of EasyInvoice.xlsx" ("y chan vậy
// với thêm 2 cột y như kh mới để cho kh cũ nha") -- KH Cu dung phan mem xuat
// hoa don TEN "EasyInvoice", mau HOAN TOAN KHAC voi "VietInvoice" cua KH Moi
// (48 cot: 46 cot chuan EasyInvoice + 2 cot them "Thuộc đối soát"/"Mã công
// trình" giong KH Moi). Header lay Y CHANG dong 1 sheet "Hóa đơn" trong file
// mau (khong co cac dong huong dan phia tren nhu VietInvoice). Doi chieu voi
// vi du co san trong file de suy ra dung cong thuc: dong 2 mau co DonGia =
// ThanhTien = 7162037, TienThue = 572963. Neu 7162037 la SO GOC (gross, da
// gom thue) thi tach thue 8% ra phai duoc 7162037/1.08 = 6631515 (KHONG khop
// TienThue mau). Nguoc lai neu 7162037 CHINH LA phan da tru thue (net) thi
// gross that = 7162037*1.08 = 7735000, TienThue = gross - 7162037 = 572963 --
// KHOP CHINH XAC. Vay cot "ThanhTien"/"DonGia" trong mau EasyInvoice DA LA SO
// SAU KHI TACH THUE (net), giong y nghia cot "Thành tiền(*)" cua VietInvoice
// -- CUNG 1 cong thuc da dung cho KH Moi: goi gross la so tien thu duoc thuc
// te (Chi Nhan xac nhan "số tiền bạn lấy phải là số chẵn số chưa trừ phí" +
// "đã bao gồm thuế nên bạn trừ cho thuế nha"), thanhTien = round(gross/1.08),
// tienThue = gross - thanhTien.
const HOA_DON_BAN_RA_EASYINVOICE_HEADER_ROW = [
  "MaHD(*)",
  "NgayHoaDon(*)",
  "MaKhachHang",
  "TenNguoiMua",
  "TenDonVi",
  "MaSoThue",
  "DiaChiKhachHang",
  "SoDienThoai",
  "SoBangKe",
  "NgayBangKe",
  "SOTKKHACH",
  "TENNHKHACH",
  "HinhThucThanhToan(*)",
  "ThueSuat(*)",
  "ThueSuatKhac",
  "MaHang",
  "TenHangHoa(*)",
  "DVT",
  "SoLuong",
  "DonGia",
  "ThanhTien",
  "TienTe",
  "SoTT",
  "TinhChat(*)",
  "Email",
  "Ghichu",
  "TyGia",
  "GiamTruHoaDon",
  "GiamTruTungDongHangHoa",
  "TienGiamTru",
  "TyLe%ChietKhau",
  "TienChietKhau",
  "TienThue",
  "MaDonViQuanHeNganSach",
  "CanCuocCongDan",
  "SoHoChieu",
  "SoKhung",
  "SoMay",
  "BienKiemSoatPhuongTienVanchuyen",
  "TenNguoiGuiHang",
  "DiaChiNguoiGuiHang",
  "MaSoThueNguoiGuiHang",
  "SoDinhDanhNguoiGuiHang",
  "Madiadiemkinhdoanh",
  "Tendiadiemkinhdoanh",
  "Diachidiadiemkinhdoanh",
  "Thuộc đối soát",
  "Mã công trình",
];

// { key: { bankName, label } } -- moi kenh VietQR cua 1 cong ty, dung LAI
// dung cau hinh CHANNELS trong routes/doisoat-vietqr.js (khong doan lai ten
// ngan hang) qua vietqrRouter.VIETQR_CHANNELS.
const HOA_DON_BAN_RA_VIETQR_KEYS = {
  kh_moi: ["bidv7702", "bidv77021", "mb02865168", "bidv8613600999"],
  kh_cu: ["bidv7704", "bidv77020", "mb11521268"],
};

// Luyen, 2026-07-31: "ngoài các giao dịch đối soát còn có doanh thu khách
// thu bằng tiền mặt cửa hàng trưởng sẽ thu về á rồi nộp sale bạn cộng vô
// nếu hôm đó có tiền về nhá mỗi gian điều có 1 mã nộp tiền á" -- "Mã nội
// dung nộp tiền" luon co dang <KH705|KH989><KVCMB|KVCMN|MTDMB|MTDMN><4 so>
// (vd "KH989MTDMB0003") theo 4 file tham khao Luyen gui, nam san trong noi
// dung giao dich cua store.transactions (sao ke ngan hang chung, KHONG phai
// 1 kenh doi soat rieng -- an toan quet toan bo, khong trung voi Momo/ZVP/
// VietQR vi cac kenh do dung bang upload rieng, khong dung store.transactions).
const CHT_CODE_RE = /KH(?:705|989)(?:KVCMB|KVCMN|MTDMB|MTDMN)\d{3,4}/;

function bankNameForId(store, bankId) {
  const b = (store.banks || []).find((x) => x.id === bankId);
  return b ? b.name : "";
}

// Luyen, 2026-07-31: BANK_COMPANY (utils/bankCompany.js) la danh sach VIET
// TAY, thieu vai tai khoan tao SAU nay (vd VP58888, BIDV8681) du chinh ban
// ghi ngan hang do (store.banks) DA CO SAN truong `company` dung (gan luc
// tao/seed) -- neu chi dung BANK_COMPANY se bi tra ve "khong xac dinh cong
// ty" cho cac tai khoan nay, khien doanh thu CHT nop tien cua chung khong
// bao gio duoc cong vao ban xuat nao ca. Dung LAI companyOfBankRow (dinh
// nghia o tren, cung ham dashboard.js/congno.js da dung on dinh) thay vi tu
// viet lai -- da uu tien doc truong `company` truoc, fallback BANK_COMPANY.
function companyForBankId(store, bankId) {
  const b = (store.banks || []).find((x) => x.id === bankId);
  return b ? companyOfBankRow(b) : null;
}

// Luyen, 2026-07-31: "nếu xuất đối soát ngày 30 mà ngày 31 xuất thì lấy 31
// nếu giao dịch vô tiền trc lúc tôi xuất của ngày 31 thì cứ thêm vào cho lúc
// tôi xuất của ngày 30 có của ngày 31 CHT nộp tiền đó ln" -- tien mat cua
// hang truong thu trong ngay X thuong duoc mang di nop ngan hang vao SANG
// HOM SAU (ngay X+1), giong het do tre T+1 da thay o Momo/ZVP settlement --
// khac voi cac kenh do (co san truong `from`/`to` rieng cho tung dot), CHT
// nop tien dung THANG giao dich ngan hang tho (khong co khai niem "dot"), nen
// tu them 1 ngay "nhin truoc" (toDate + 1) khi loc, roi tinh het vao dung
// khoang dang xuat (khong tach dong rieng theo ngay thuc nhan tien). CANH
// BAO: neu Luyen xuat 2 lan CHONG NGAY (vd xuat rieng ngay 30 XONG ROI xuat
// rieng ca ngay 31), giao dich CHT ngay 31 se bi tinh 2 LAN (1 lan qua do tre
// cua ban ngay 30, 1 lan qua chinh ban than no o ban ngay 31) -- chi an toan
// khi xuat TUNG khoang ngay KHONG CHONG LAP nhau.
function addOneDay(dateStr) {
  const d = new Date(dateStr + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

// Quet TOAN BO store.transactions (giao dich "thu" trong khoang ngay), tim
// giao dich nao co nhung 1 "Ma noi dung nop tien" trong mo ta -- cong ty xac
// dinh qua CHINH tai khoan ngan hang nhan tien (BANK_COMPANY), KHONG doc cot
// "Pháp nhân" cua file tham khao (tai khoan da lam dung viec nay on dinh cho
// moi kenh khac trong file nay roi, tranh phai tin cay 1 nguon du lieu ngoai
// them). Ma nao KHOP mau nhung CHUA co trong bang tra (store.cht_nop_tien_map)
// duoc bao ve rieng (unmapped) thay vi bo qua doanh thu do trong im lang.
function collectChtNopTienLines(store, company, fromDate, toDate) {
  const lines = [];
  const unmapped = new Map();
  const toDatePlus1 = addOneDay(toDate);
  (store.transactions || []).forEach((t) => {
    if (t.type !== "thu") return;
    if (!t.date || t.date < fromDate || t.date > toDatePlus1) return;
    const desc = String(t.description || "").toUpperCase();
    const match = desc.match(CHT_CODE_RE);
    if (!match) return;
    const code = match[0];
    const txCompany = companyForBankId(store, t.bank_id);
    if (txCompany !== company) return; // giao dich nay khong thuoc cong ty dang xuat
    const mapped = (store.cht_nop_tien_map || {})[code];
    if (!mapped || !mapped.maCongTrinh) {
      if (!unmapped.has(code)) {
        unmapped.set(code, { code, bankName: bankNameForId(store, t.bank_id), date: t.date, amount: t.amount, description: t.description });
      }
      return;
    }
    lines.push({ maCongTrinh: mapped.maCongTrinh, gross: t.amount });
  });
  return { lines, unmapped: Array.from(unmapped.values()) };
}

// Gop tat ca kenh doi soat cua 1 cong ty trong 1 khoang ngay thanh danh sach
// { thuocDoiSoat, maCongTrinh, gross } DA CONG DON (khong con tach theo
// ngay/thang nua) -- moi dong DUY NHAT 1 cap (thuocDoiSoat, maCongTrinh).
function buildHoaDonBanRaGroups(store, company, fromDate, toDate) {
  const groups = new Map();
  function addLine(thuocDoiSoat, maCongTrinh, gross) {
    const roundedGross = Math.round(gross || 0);
    if (!roundedGross) return;
    const key = thuocDoiSoat + "||" + maCongTrinh;
    if (!groups.has(key)) groups.set(key, { thuocDoiSoat, maCongTrinh, gross: 0 });
    groups.get(key).gross += roundedGross;
  }
  // `skipMeansBelongsToKhMoi`: rieng kenh Momo cua KH Moi (BIDV7701) dung TK
  // Co = "SKIP" theo 1 nghia HOAN TOAN KHAC voi moi kenh con lai -- do la di
  // san tu thoi KH Moi chua co tai khoan Momo rieng, phai doi soat chung
  // nguon voi KH Cu, nen SKIP duoc dung lam CO HIEU PHAN BIET "gian nay la
  // cua KH Moi" (xem ensureGianHidden/route /doi-soat/momo, dong loc
  // `l.tkCo === "SKIP"` cho activeCompany === "kh_moi"), KHONG phai "chua
  // xuat MISA" nhu moi noi khac. Chi Nhan xac nhan (2026-07-30): "4 gian đó
  // của kh mới á" (FARM LOTTE PHAN THIET/NHA TRANG, AE TAN AN KVC, KVC
  // ESTELLA) -- nen kenh nay phai LAY CA cac dong SKIP thay vi loai bo,
  // nguoc lai voi VietQR/VNPay (noi SKIP dung dung nghia "chua xuat MISA",
  // van phai loai nhu cu).
  // Chi Nhan, 2026-07-30: "momo có 3 tr mấy mà với payoo có 2 tr mấy lận mà
  // bạn lấy số đối soát chê vậy" -- loc theo `day.settlementDate` (ngay TIEN
  // VE NGAN HANG) la SAI cho Momo/ZVP/VNPay-KhMoi: nhung kenh nay dung
  // reconcileMomo/reconcileZvpChannel, TIEN VE NGAN HANG bi LECH ngay (thuong
  // T+1) so voi NGAY BAN HANG THUC TE -- xem label "Khoản về ngày <settlementDate>
  // (doanh thu <from> → <to>)" tren chinh trang doi soat. Vi du hom Chi Nhan
  // xuat 29/7-29/7: doanh thu THAT su cua ngay 29/7 lai settle vao ngan hang
  // NGAY 30/7, nen loc theo settlementDate=='2026-07-29' vo tinh lay nham
  // doanh thu cua ngay 28/7 (settle vao 29/7) thay vi dung ngay 29/7 Chi Nhan
  // muon. Rieng VietQR KHONG bi loi nay (tien ve tung giao dich mot, cung
  // ngay, khong co do tre) nen khong co truong `from`/`to` -- fallback ve
  // settlementDate cho kenh do. Dung phep GIAO KHOANG NGAY (thay vi so sanh 1
  // ngay duy nhat) vi 1 "ngay" doi soat co the gom nhieu ngay doanh thu lai
  // (vd tien cuoi tuan gop lai settle chung 1 hom, from != to).
  function collectDays(reconciledDays, thuocDoiSoat, skipMeansBelongsToKhMoi) {
    (reconciledDays || []).forEach((day) => {
      const rangeStart = day.from || day.settlementDate;
      const rangeEnd = day.to || day.settlementDate;
      if (!rangeStart || !rangeEnd) return;
      if (rangeEnd < fromDate || rangeStart > toDate) return; // khong giao voi khoang loc
      (day.lines || []).forEach((l) => {
        if (l.tkCo === "SKIP" && !skipMeansBelongsToKhMoi) return; // chua xuat MISA -- khong xuat hoa don
        addLine(thuocDoiSoat, l.maCongTrinh, l.gross);
      });
    });
  }

  if (company === "kh_moi") {
    const vietqrRouter = require("./doisoat-vietqr");
    HOA_DON_BAN_RA_VIETQR_KEYS.kh_moi.forEach((chKey) => {
      const built = vietqrRouter.buildChannelReconciliation(store, chKey);
      if (built && !built.error && built.reconciled) {
        const bankName = vietqrRouter.VIETQR_CHANNELS[chKey].bankName;
        collectDays(built.reconciled, `Viet QR ${bankName}`);
      }
    });
    const vnpayKhMoiRouter = require("./doisoat-vnpay-khmoi");
    const vnpayBuilt = vnpayKhMoiRouter.buildReconciliation(store);
    if (vnpayBuilt && !vnpayBuilt.error && vnpayBuilt.reconciled) {
      collectDays(vnpayBuilt.reconciled.offline, "VN Pay VTB982");
      collectDays(vnpayBuilt.reconciled.payoo, "Payoo VTB982");
    }
    // Chi Nhan, 2026-07-30: "sao xuất ra thử ngày 29 á không có đối soát momo
    // kh mới" -- thieu sot: Momo CUNG co rieng 1 TK cho KH Moi (BIDV7701, xem
    // MOMO_CHANNELS.kh_moi trong routes/doisoat.js), khong chi KH Cu, nen phai
    // gop them kenh nay vao bao cao (nhan "Momo BIDV7701" giong cach dat ten
    // "Viet QR <NH>"/"VN Pay <NH>" o tren).
    const momoRouter = require("./doisoat");
    const momoBuilt = momoRouter.buildMomoReconciliation(store, "kh_moi");
    if (momoBuilt && !momoBuilt.error && momoBuilt.reconciledAll) {
      collectDays(momoBuilt.reconciledAll, "Momo BIDV7701", true);
    }
  } else {
    // Chi Nhan, 2026-07-30: gui file mau "EasyInvoice" rieng cho KH Cu, xac
    // nhan lam theo cung kien truc voi KH Moi -- gom VietQR (bidv7704/
    // bidv77020/mb11521268) + Momo (BIDV123456, TK Co SKIP dung DUNG nghia
    // "chua xuat MISA" o day, khong phai co hieu phan cong ty nhu ben KH Moi,
    // nen KHONG bat skipMeansBelongsToKhMoi) + Zalo App/VNPay/Payoo (routes/
    // doisoat-zvp.js, ca 3 kenh cung ve TK ACB31268 -- "online"=Zalo App,
    // "offline"=VNPay tai co so, "payoo"=Payoo, dat nhan giong quy uoc KH Moi:
    // "Zalo app <NH>"/"VN Pay <NH>"/"Payoo <NH>").
    const vietqrRouter = require("./doisoat-vietqr");
    HOA_DON_BAN_RA_VIETQR_KEYS.kh_cu.forEach((chKey) => {
      const built = vietqrRouter.buildChannelReconciliation(store, chKey);
      if (built && !built.error && built.reconciled) {
        const bankName = vietqrRouter.VIETQR_CHANNELS[chKey].bankName;
        collectDays(built.reconciled, `Viet QR ${bankName}`);
      }
    });
    const momoRouter = require("./doisoat");
    const momoBuilt = momoRouter.buildMomoReconciliation(store, "kh_cu");
    if (momoBuilt && !momoBuilt.error && momoBuilt.reconciledAll) {
      collectDays(momoBuilt.reconciledAll, `Momo ${MOMO_BANK_NAME}`);
    }
    const zvpRouter = require("./doisoat-zvp");
    const zvpBuilt = zvpRouter.buildZvpReconciliation(store);
    if (zvpBuilt && !zvpBuilt.error && zvpBuilt.reconciled) {
      const zvpBankName = zvpRouter.ZVP_BANK_NAME || ZVP_BANK_NAME;
      collectDays(zvpBuilt.reconciled.online, `Zalo app ${zvpBankName}`);
      collectDays(zvpBuilt.reconciled.offline, `VN Pay ${zvpBankName}`);
      collectDays(zvpBuilt.reconciled.payoo, `Payoo ${zvpBankName}`);
    }
  }

  // "CHT nộp tiền" -- tien mat cua hang truong thu roi nop lai ngan hang,
  // ap dung cho CA 2 cong ty (moi cong ty co tai khoan nhan tien rieng cua
  // no, xem BANK_COMPANY) -- Luyen xac nhan them 1 dong rieng ten "CHT nộp
  // tiền" trong cot "Thuộc đối soát", giong cach dat ten "Viet QR <NH>"/
  // "VN Pay <NH>" o tren.
  const chtResult = collectChtNopTienLines(store, company, fromDate, toDate);
  chtResult.lines.forEach((l) => addLine("CHT nộp tiền", l.maCongTrinh, l.gross));

  const result = Array.from(groups.values())
    .filter((r) => r.gross !== 0)
    .sort((a, b) => a.thuocDoiSoat.localeCompare(b.thuocDoiSoat) || a.maCongTrinh.localeCompare(b.maCongTrinh));
  result.chtUnmapped = chtResult.unmapped; // dinh kem canh bao (khong pha vo cac cho dang dung result nhu 1 mang thuong)
  return result;
}

// Canh bao (khong phu thuoc khoang ngay) -- quet TOAN BO store.transactions
// 1 lan cho ca 2 cong ty, liet ke moi "Ma noi dung nop tien" KHOP mau nhung
// CHUA co trong bang tra, de Luyen thay ngay tren trang (khong can bam Xuat
// file thu roi moi biet thieu mapping).
function findAllUnmappedChtCodes(store) {
  const byCode = new Map();
  (store.transactions || []).forEach((t) => {
    if (t.type !== "thu") return;
    const desc = String(t.description || "").toUpperCase();
    const match = desc.match(CHT_CODE_RE);
    if (!match) return;
    const code = match[0];
    if ((store.cht_nop_tien_map || {})[code]) return; // da co mapping
    const bankName = bankNameForId(store, t.bank_id);
    if (!byCode.has(code)) {
      byCode.set(code, { code, bankName, company: companyForBankId(store, t.bank_id) || "?", count: 0, lastDate: t.date });
    }
    const rec = byCode.get(code);
    rec.count++;
    if (t.date > rec.lastDate) rec.lastDate = t.date;
  });
  return Array.from(byCode.values()).sort((a, b) => a.code.localeCompare(b.code));
}

router.get("/bao-cao/xuat-hoa-don-ban-ra", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const chtMapCount = Object.keys(store.cht_nop_tien_map || {}).length;
  const chtUnmapped = findAllUnmappedChtCodes(store);
  res.render("baocao-hoadonbanra", {
    userName: req.session.userName,
    activeCompany,
    COMPANIES,
    error: req.query.error || null,
    success: req.query.success || null,
    chtMapCount,
    chtUnmapped,
  });
});

router.post("/bao-cao/xuat-hoa-don-ban-ra/upload-cht-nop-tien", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { sheetsParsed, rows } = zvp.parseChtNopTienMasterSheet(req.file.buffer);
    if (rows.length === 0) {
      throw new Error(
        'Không đọc được dòng nào -- cần sheet có cả cột "Nội dung nộp tiền" và "Mã công trình" (misa thuế).'
      );
    }
    const { map, added, updated } = zvp.mergeChtNopTienMap(store.cht_nop_tien_map, rows);
    store.cht_nop_tien_map = map;
    store.cht_nop_tien_uploads.push({
      id: (store.cht_nop_tien_uploads.length || 0) + 1,
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetsParsed,
      rowCount: rows.length,
    });
    save(store);
    const successMsg =
      `Đã nạp file "${req.file.originalname}" (${sheetsParsed.map((s) => `${s.sheetName}: ${s.rows} dòng`).join(", ")}) -- ` +
      `${added} mã mới, ${updated} mã cập nhật. Đang có ${Object.keys(store.cht_nop_tien_map).length} mã trong bảng tra.`;
    res.redirect("/bao-cao/xuat-hoa-don-ban-ra?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/bao-cao/xuat-hoa-don-ban-ra?error=" + encodeURIComponent(e.message));
  }
});

router.get("/bao-cao/xuat-hoa-don-ban-ra/xuat.xlsx", (req, res) => {
  const store = load();
  const company = req.query.company === "kh_cu" ? "kh_cu" : "kh_moi";
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

  let groups;
  try {
    groups = buildHoaDonBanRaGroups(store, company, fromDate, toDate);
  } catch (e) {
    return res.status(400).send(e.message);
  }
  if (groups.length === 0) {
    return res.status(400).send(`Từ ${fromDate} đến ${toDate} không có mã công trình nào có doanh thu (khác 0) để xuất hóa đơn.`);
  }

  const today = new Date();
  const ngayHoaDon = [today.getDate(), today.getMonth() + 1, today.getFullYear()]
    .map((n) => String(n).padStart(2, "0"))
    .join("/");

  const dataRows = [];
  let seq = startNo;
  let aoa;
  let sheetName;

  if (company === "kh_moi") {
    // KH Moi -- mau "VietInvoice" (giu nguyen nhu Chi Nhan da xac nhan truoc).
    groups.forEach((g) => {
      const grossTotal = g.gross;
      const thanhTien = Math.round(grossTotal / (1 + HOA_DON_BAN_RA_VAT_RATE));
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
        g.thuocDoiSoat,
        g.maCongTrinh,
      ]);
      seq++;
    });
    aoa = [...HOA_DON_BAN_RA_HEADER_INFO, HOA_DON_BAN_RA_HEADER_ROW, ...dataRows];
    sheetName = "Hóa đơn";
  } else {
    // KH Cu -- mau "EasyInvoice" (file Chi Nhan gui 2026-07-30), 48 cot (46
    // cot chuan + 2 cot "Thuộc đối soát"/"Mã công trình" giong KH Moi). Xem
    // ghi chu day du o dinh nghia HOA_DON_BAN_RA_EASYINVOICE_HEADER_ROW ve
    // cong thuc ThanhTien/TienThue (giong VietInvoice: tach 8% thue tu gross).
    groups.forEach((g) => {
      const grossTotal = g.gross;
      const thanhTien = Math.round(grossTotal / (1 + HOA_DON_BAN_RA_VAT_RATE));
      const tienThue = grossTotal - thanhTien;
      dataRows.push([
        `HD${seq}`, // MaHD(*)
        ngayHoaDon, // NgayHoaDon(*)
        "", // MaKhachHang
        "Bán cho người tiêu dùng ", // TenNguoiMua
        "", // TenDonVi
        "", // MaSoThue
        "", // DiaChiKhachHang
        "", // SoDienThoai
        "", // SoBangKe
        "", // NgayBangKe
        "", // SOTKKHACH
        "", // TENNHKHACH
        "Tiền mặt/Chuyển khoản", // HinhThucThanhToan(*)
        8, // ThueSuat(*)
        "", // ThueSuatKhac
        "", // MaHang
        "Dịch vụ vui chơi giải trí", // TenHangHoa(*)
        "Kỳ ", // DVT
        1, // SoLuong
        thanhTien, // DonGia
        thanhTien, // ThanhTien
        "VND", // TienTe
        "", // SoTT
        1, // TinhChat(*)
        "", // Email
        "", // Ghichu
        "", // TyGia
        "", // GiamTruHoaDon
        "", // GiamTruTungDongHangHoa
        "", // TienGiamTru
        "", // TyLe%ChietKhau
        "", // TienChietKhau
        tienThue, // TienThue
        "", // MaDonViQuanHeNganSach
        "", // CanCuocCongDan
        "", // SoHoChieu
        "", // SoKhung
        "", // SoMay
        "", // BienKiemSoatPhuongTienVanchuyen
        "", // TenNguoiGuiHang
        "", // DiaChiNguoiGuiHang
        "", // MaSoThueNguoiGuiHang
        "", // SoDinhDanhNguoiGuiHang
        "", // Madiadiemkinhdoanh
        "", // Tendiadiemkinhdoanh
        "", // Diachidiadiemkinhdoanh
        g.thuocDoiSoat, // Thuộc đối soát
        g.maCongTrinh, // Mã công trình
      ]);
      seq++;
    });
    aoa = [HOA_DON_BAN_RA_EASYINVOICE_HEADER_ROW, ...dataRows];
    sheetName = "Hóa đơn";
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, sheetName);
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=HoaDonBanRa-${company}-${fromDate}_${toDate}.xlsx`);
  res.send(buf);
});

// ---------- 6) To khai thue GTGT (uoc tinh) theo cong ty ----------
// Chi Nhan, 2026-07-30: "dưới kinh nghiệm kế toán 80 năm kế toán trưởng tôi
// có đầu vào đầu ra rồi thêm cho tôi tờ khai thuế giá trị gia tăng đi nhá xem
// số thuế phải nộp hay cần nộp là bao nhiêu ... 2 công ty cho tôi luôn nhá" --
// bao cao MOI, tinh UOC TINH thue GTGT phai nop/duoc khau tru trong 1 thang,
// theo dung phuong phap khau tru (Thue phai nop = Thue GTGT dau ra - Thue
// GTGT dau vao duoc khau tru), cho CA 2 cong ty cung 1 luc (giong cach trang
// "Trạng thái đối soát" da lam, 2 khung canh nhau).
//
// Dau ra: dung LAI CHINH XAC buildHoaDonBanRaGroups() o tren (da kiem chung
// qua tinh nang "Xuất Hóa Đơn Bán Ra") de lay TONG DOANH THU GOP (da gom
// thue) cua CA cong ty trong thang, roi tach thue 8% ra giong cong thuc dang
// dung (Thanh tien = gross/1.08, Thue = gross - Thanh tien) -- dam bao KHONG
// tinh sai lech so voi bao cao Xuat Hoa Don Ban Ra da co.
// Dau vao: dung truc tiep store.hoa_don_dau_vao (routes/hoa-don-dau-vao.js),
// cong don cot `tienThue` (thue GTGT tren hoa don mua vao, da co san tu luc
// nhap/import) cho dung cong ty + dung thang.
//
// QUAN TRONG: day la SO UOC TINH tu du lieu app dang co (co the thieu hoa don
// dau vao chua nhap, hoac doanh thu chua chot so het thang), KHONG thay the
// to khai chinh thuc nop co quan thue -- Chi Nhan/ke toan van can doi chieu
// lai truoc khi nop thuc te.
function buildToKhaiThueGtgt(store, month) {
  const fromDate = `${month}-01`;
  const lastDay = new Date(Number(month.slice(0, 4)), Number(month.slice(5, 7)), 0).getDate();
  const toDate = `${month}-${String(lastDay).padStart(2, "0")}`;

  const result = {};
  ["kh_cu", "kh_moi"].forEach((company) => {
    let groups = [];
    let error = null;
    try {
      groups = buildHoaDonBanRaGroups(store, company, fromDate, toDate);
    } catch (e) {
      error = e.message;
    }
    const doanhThuGop = groups.reduce((s, g) => s + g.gross, 0);
    const doanhThuTinhThue = Math.round(doanhThuGop / (1 + HOA_DON_BAN_RA_VAT_RATE));
    const thueDauRa = doanhThuGop - doanhThuTinhThue;

    const hoaDonDauVaoRows = (store.hoa_don_dau_vao || []).filter(
      (r) => r.congTy === company && (r.ngayHD || "").slice(0, 7) === month
    );
    const tongTienMuaVao = hoaDonDauVaoRows.reduce((s, r) => s + (r.soTien || 0), 0);
    const thueDauVao = hoaDonDauVaoRows.reduce((s, r) => s + (r.tienThue || 0), 0);

    const thuePhaiNop = thueDauRa - thueDauVao;

    result[company] = {
      error,
      soLuongDoiSoat: groups.length,
      doanhThuGop,
      doanhThuTinhThue,
      thueDauRa,
      soHoaDonDauVao: hoaDonDauVaoRows.length,
      tongTienMuaVao,
      thueDauVao,
      thuePhaiNop,
    };
  });
  return result;
}

router.get("/bao-cao/to-khai-thue-gtgt", (req, res) => {
  const store = load();
  const today = new Date();
  const defaultMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
  const selectedMonth = /^\d{4}-\d{2}$/.test(req.query.month || "") ? req.query.month : defaultMonth;
  let data = null;
  let error = null;
  try {
    data = buildToKhaiThueGtgt(store, selectedMonth);
  } catch (e) {
    error = e.message;
    console.error("Loi tinh to khai thue GTGT:", e);
  }
  res.render("baocao-tokhaigtgt", {
    userName: req.session.userName,
    COMPANIES,
    selectedMonth,
    data,
    error,
    vatRate: HOA_DON_BAN_RA_VAT_RATE,
  });
});

router.buildHoaDonBanRaGroups = buildHoaDonBanRaGroups;

module.exports = router;
