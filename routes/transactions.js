const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");
const { parsePastedTransactions, parseAmount, parseDate } = require("../utils/parse");
const { parseBankStatement, computeThuChi } = require("../utils/bankStatementParser");
const { getCompany } = require("../utils/companies");
const { parseMaCongTrinhSheet } = require("../utils/maCongTrinh");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

function maCongTrinhMasterFor(store, company) {
  return (store.ma_cong_trinh_master && store.ma_cong_trinh_master[company]) || null;
}

function withBankLabel(store, tx) {
  const bank = store.banks.find((b) => b.id === tx.bank_id);
  return { ...tx, bank_label: bank ? bank.name : "(da xoa)" };
}

// Cac ngan hang cu (KH Cu) duoc tao truoc khi co tinh nang da cong ty nen
// khong co field "company" -- coi nhu mac dinh la "kh_cu" de tuong thich nguoc.
function companyBanks(store, company) {
  return store.banks
    .filter((b) => (b.company || "kh_cu") === company)
    .sort((a, b) => a.name.localeCompare(b.name));
}

function companyBankIds(store, company) {
  return new Set(companyBanks(store, company).map((b) => b.id));
}

function filterTransactions(store, { bank_id, from, to, type, bankIds }) {
  let rows = [...store.transactions];
  if (bankIds) rows = rows.filter((t) => bankIds.has(t.bank_id));
  if (bank_id) rows = rows.filter((t) => t.bank_id === Number(bank_id));
  if (from) rows = rows.filter((t) => t.date >= from);
  if (to) rows = rows.filter((t) => t.date <= to);
  if (type === "thu" || type === "chi") rows = rows.filter((t) => t.type === type);
  rows.sort((a, b) => (b.date > a.date ? 1 : b.date < a.date ? -1 : b.id - a.id));
  return rows.map((t) => withBankLabel(store, t));
}

function bankBalanceBefore(store, bankId, beforeDate) {
  const bank = store.banks.find((b) => b.id === bankId);
  if (!bank) return 0;
  let bal = bank.opening_balance;
  for (const t of store.transactions) {
    if (t.bank_id !== bankId) continue;
    if (beforeDate !== null && t.date >= beforeDate) continue;
    bal += t.type === "thu" ? t.amount : -t.amount;
  }
  return bal;
}

// Luyen, 2026-07-21: "ngân hàng hk có xóa đâu, thay vào đó hiện số dư cuối kì
// cho tôi đi" -- moi dong giao dich hien so du LUY KE (giong sao ke ngan
// hang that, cot "Số dư tham chiếu"/"Running Balance" -- xem file
// AccountStmt Luyen gui). Tinh theo THU TU THOI GIAN THAT cua tung ngan hang
// rieng (ngay tang dan, cung ngay thi theo id tang dan -- id la thu tu nhap/
// import, gan dung thu tu thuc te trong pham vi 1 ngay), bat dau tu
// opening_balance -- HOAN TOAN doc lap voi thu tu hien thi tren bang (thuong
// la moi nhat truoc). Tra ve Map<transactionId, soDuSauGiaoDichDo>, tinh 1
// lan cho MOI ngan hang trong bankIds (khong phai toan bo store.transactions)
// de khong tinh du lieu cua ngan hang khong lien quan.
function buildBalanceMap(store, bankIds) {
  const map = new Map();
  const ids = bankIds instanceof Set ? bankIds : new Set(bankIds || []);
  ids.forEach((bankId) => {
    const bank = store.banks.find((b) => b.id === bankId);
    if (!bank) return;
    const txs = store.transactions
      .filter((t) => t.bank_id === bankId)
      .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.id - b.id));
    let bal = bank.opening_balance || 0;
    for (const t of txs) {
      bal += t.type === "thu" ? t.amount : -t.amount;
      map.set(t.id, bal);
    }
  });
  return map;
}

// filterTransactions + gan them so du luy ke (t.balance) cho tung dong, dung
// chung cho moi cho render danh sach giao dich (thay vi lap lai buildBalanceMap
// o tung route).
function filterTransactionsWithBalance(store, opts) {
  const rows = filterTransactions(store, opts);
  const balanceMap = buildBalanceMap(store, opts.bankIds);
  return rows.map((t) => ({ ...t, balance: balanceMap.has(t.id) ? balanceMap.get(t.id) : null }));
}

router.get("/transactions", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const { bank_id, from, to, type } = req.query;

  // Chi Nhan (2026-07-22): "tách thu chi thành 2 cột ... cộng tổng thu chi
  // đang hiển thị ... lọc thu chi ở đây theo ngày tháng" -- tinh tong Thu/Chi
  // tren TOAN BO danh sach da loc (khong chi 500 dong hien thi), de dung voi
  // bo loc dang chon, roi moi cat con 500 dong de hien thi bang.
  const allFiltered = filterTransactionsWithBalance(store, {
    bank_id,
    from,
    to,
    type,
    bankIds: companyBankIds(store, activeCompany),
  });
  const rows = allFiltered.slice(0, 500);
  const totalThu = allFiltered.filter((t) => t.type === "thu").reduce((s, t) => s + Number(t.amount || 0), 0);
  const totalChi = allFiltered.filter((t) => t.type === "chi").reduce((s, t) => s + Number(t.amount || 0), 0);

  res.render("transactions", {
    banks,
    rows,
    totalThu,
    totalChi,
    filters: { bank_id: bank_id || "", from: from || "", to: to || "", type: type || "" },
    userName: req.session.userName,
    pasteResult: null,
    uploadResult: null,
    maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
    maCongTrinhResult: null,
    error: null,
  });
});

router.post("/transactions", requireAdmin, (req, res) => {
  const { bank_id, date, description, amount, type } = req.body;
  const parsedDate = parseDate(date) || date;
  const parsedAmount = Math.abs(parseAmount(amount));
  if (!bank_id || !parsedDate || isNaN(parsedAmount) || !["thu", "chi"].includes(type)) {
    return res.redirect("/transactions?error=1");
  }
  const store = load();
  store.transactions.push({
    id: nextId(store, "transactions"),
    bank_id: Number(bank_id),
    date: parsedDate,
    description: description || "",
    amount: parsedAmount,
    type,
    created_at: new Date().toISOString(),
    created_by: req.session.userName || "",
  });
  save(store);
  res.redirect("/transactions?bank_id=" + encodeURIComponent(bank_id));
});

router.post("/transactions/paste", requireAdmin, (req, res) => {
  const { bank_id, paste_text } = req.body;
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);

  if (!bank_id) {
    return res.render("transactions", {
      banks,
      rows: [],
      filters: { bank_id: "", from: "", to: "" },
      userName: req.session.userName,
      pasteResult: null,
      uploadResult: null,
      maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
      maCongTrinhResult: null,
      error: "Vui long chon ngan hang truoc khi dan sao ke.",
    });
  }

  const { rows, errors } = parsePastedTransactions(paste_text);

  for (const r of rows) {
    store.transactions.push({
      id: nextId(store, "transactions"),
      bank_id: Number(bank_id),
      date: r.date,
      description: r.description,
      amount: r.amount,
      type: r.type,
      created_at: new Date().toISOString(),
      created_by: req.session.userName || "",
    });
  }
  if (rows.length > 0) save(store);

  const currentRows = filterTransactionsWithBalance(store, {
    bank_id,
    bankIds: companyBankIds(store, activeCompany),
  }).slice(0, 500);

  res.render("transactions", {
    banks,
    rows: currentRows,
    filters: { bank_id, from: "", to: "" },
    userName: req.session.userName,
    pasteResult: { inserted: rows.length, errors },
    uploadResult: null,
    maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
    maCongTrinhResult: null,
    error: null,
  });
});

// Upload a raw statement file exported directly from the bank (.xlsx/.xls).
// Auto-detects the "Ngay giao dich" + "So du" columns, derives Thu/Chi from
// the balance delta, and skips any row that already exists for this bank so
// re-uploading an overlapping date range is safe.
//
// Dedup key: prefer the bank's own per-transaction reference number ("So
// tham chieu" / "So chung tu"), when the export exposes one. This is
// required because many same-day, same-amount transactions from DIFFERENT
// customers are completely normal for VietQR fixed-price ticket sales
// (20.000d / 50.000d / 100.000d recurring hundreds of times a day) -- a key
// of just date+amount+type would collapse all of them into "duplicates" and
// silently drop the rest, which is exactly what happened before this fix.
// When no reference column is detected (older/other bank formats), fall
// back to date+amount+type as before -- description is intentionally still
// excluded from that fallback key, since different statement exports can
// render slightly different description text for the same transaction.
router.post("/transactions/upload-statement", requireAdmin, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const { bank_id } = req.body;

  const renderError = (message) =>
    res.render("transactions", {
      banks,
      rows: filterTransactionsWithBalance(store, {
        bank_id: bank_id || "",
        bankIds: companyBankIds(store, activeCompany),
      }).slice(0, 500),
      filters: { bank_id: bank_id || "", from: "", to: "" },
      userName: req.session.userName,
      pasteResult: null,
      uploadResult: null,
      maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
      maCongTrinhResult: null,
      error: message,
    });

  if (!bank_id) return renderError("Vui long chon ngan hang truoc khi tai file sao ke.");
  if (!req.file) return renderError("Vui long chon 1 file sao ke de tai len.");

  const bankIdNum = Number(bank_id);
  const bank = store.banks.find((b) => b.id === bankIdNum);
  if (!bank) return renderError("Khong tim thay ngan hang da chon.");

  let parsed;
  try {
    parsed = parseBankStatement(req.file.buffer);
  } catch (e) {
    return renderError(e.message);
  }

  const firstDate = parsed.rows[0].date;
  const lastDate = parsed.rows[parsed.rows.length - 1].date;
  const priorBalance = bankBalanceBefore(store, bankIdNum, firstDate);
  const candidates = computeThuChi(parsed.rows, priorBalance);

  // Self-heal: if this statement format exposes a real per-row reference
  // number (only known for formats where the header-based detection in
  // bankStatementParser finds one -- e.g. BIDV's "So chung tu"/"So tham
  // chieu"), then any OLDER rows for this bank+date-range that have no
  // `reference` were inserted by the pre-fix dedup logic (date+amount+type
  // only), which silently collapsed distinct same-day/same-amount
  // transactions into a single row. Since we're about to re-derive the full,
  // authoritative transaction list for this exact range straight from the
  // bank's own file, it is safe to drop those old collapsed rows first so
  // the fresh parse can fully repopulate the range. This is scoped tightly
  // (bank + exact covered date range + only rows lacking a reference) and is
  // a no-op for statement formats without a detected reference column, so it
  // never touches banks/uploads unaffected by this bug.
  const candidatesHaveRef = candidates.some((c) => c.reference);
  let healedRemoved = 0;
  if (candidatesHaveRef) {
    const before = store.transactions.length;
    store.transactions = store.transactions.filter(
      (t) => !(t.bank_id === bankIdNum && !t.reference && t.date >= firstDate && t.date <= lastDate)
    );
    healedRemoved = before - store.transactions.length;
  }

  const bankTx = store.transactions.filter((t) => t.bank_id === bankIdNum);
  const existingRefs = new Set(bankTx.filter((t) => t.reference).map((t) => t.reference));
  const existingKeys = new Set(
    bankTx.filter((t) => !t.reference).map((t) => `${t.date}|${t.amount}|${t.type}`)
  );

  let added = 0;
  let skipped = 0;
  for (const c of candidates) {
    if (c.reference) {
      if (existingRefs.has(c.reference)) {
        skipped++;
        continue;
      }
      existingRefs.add(c.reference);
    } else {
      const key = `${c.date}|${c.amount}|${c.type}`;
      if (existingKeys.has(key)) {
        skipped++;
        continue;
      }
      existingKeys.add(key);
    }
    store.transactions.push({
      id: nextId(store, "transactions"),
      bank_id: bankIdNum,
      date: c.date,
      description: c.description,
      amount: c.amount,
      type: c.type,
      reference: c.reference || "",
      created_at: new Date().toISOString(),
      created_by: req.session.userName || "",
    });
    added++;
  }
  if (added > 0 || healedRemoved > 0) save(store);

  const currentRows = filterTransactionsWithBalance(store, {
    bank_id,
    bankIds: companyBankIds(store, activeCompany),
  }).slice(0, 500);
  res.render("transactions", {
    banks,
    rows: currentRows,
    filters: { bank_id, from: "", to: "" },
    userName: req.session.userName,
    pasteResult: null,
    uploadResult: {
      sheetName: parsed.sheetName,
      totalRows: parsed.rows.length,
      added,
      skipped,
      healedRemoved,
    },
    maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
    maCongTrinhResult: null,
    error: null,
  });
});

// Upload danh sach "Ma cong trinh" chuan (rieng theo tung cong ty dang chon).
// Luyen, 2026-07-20: file "DANH SACH CONG TRINH" tu phan mem quan ly cong
// trinh -- dung lam nguon chuan de sau nay doi chieu/chuan hoa cac ten
// gian/cong trinh xuat hien o cac trang khac (vd Chi Phi). Moi lan tai len
// THAY THE toan bo danh sach cua dung cong ty dang chon (KH Cu / KH Moi
// khong dung chung 1 danh sach vi la 2 phap nhan khac nhau).
router.post("/transactions/upload-ma-cong-trinh", requireAdmin, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const rowsForList = filterTransactionsWithBalance(store, {
    bankIds: companyBankIds(store, activeCompany),
  }).slice(0, 500);

  const renderWith = (maCongTrinhResult, error) =>
    res.render("transactions", {
      banks,
      rows: rowsForList,
      filters: { bank_id: "", from: "", to: "" },
      userName: req.session.userName,
      pasteResult: null,
      uploadResult: null,
      maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
      maCongTrinhResult,
      error: error || null,
    });

  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const { sheetName, rows } = parseMaCongTrinhSheet(req.file.buffer);
    if (!sheetName || rows.length === 0) {
      throw new Error(
        'Khong doc duoc danh sach ma cong trinh tu file nay (can co cot "Ma công trình").'
      );
    }
    if (!store.ma_cong_trinh_master) store.ma_cong_trinh_master = {};
    store.ma_cong_trinh_master[activeCompany] = {
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName,
      rows,
    };
    save(store);
    return renderWith({ sheetName, count: rows.length, fileName: req.file.originalname });
  } catch (e) {
    return renderWith(null, e.message);
  }
});

// Luyen, 2026-07-21: "ngân hàng hk có xóa đâu" -- bo nut/route Xoa giao dich
// (khong con dung tu UI), thay bang cot "So du" (buildBalanceMap o tren).
// Route xuat Excel cung them cot "So du" tuong ung, dong bo voi bang tren
// man hinh (giong sao ke ngan hang that co cot "Số dư tham chiếu").
router.get("/export.xlsx", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const { bank_id, from, to } = req.query;
  const rows = filterTransactionsWithBalance(store, {
    bank_id,
    from,
    to,
    bankIds: companyBankIds(store, activeCompany),
  })
    .slice()
    .reverse() // chronological order for the export
    .map((t) => ({
      "Ngan hang": t.bank_label,
      Ngay: t.date,
      "Dien giai": t.description,
      "So tien": t.amount,
      Loai: t.type,
      "So du": t.balance,
    }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Giao dich");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", "attachment; filename=giao-dich-ngan-hang.xlsx");
  res.send(buf);
});

module.exports = router;
