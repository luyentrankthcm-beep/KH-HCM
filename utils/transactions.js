const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { parsePastedTransactions, parseAmount, parseDate } = require("../utils/parse");
const { parseBankStatement, computeThuChi } = require("../utils/bankStatementParser");
const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

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

function filterTransactions(store, { bank_id, from, to, bankIds }) {
  let rows = [...store.transactions];
  if (bankIds) rows = rows.filter((t) => bankIds.has(t.bank_id));
  if (bank_id) rows = rows.filter((t) => t.bank_id === Number(bank_id));
  if (from) rows = rows.filter((t) => t.date >= from);
  if (to) rows = rows.filter((t) => t.date <= to);
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

router.get("/transactions", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const { bank_id, from, to } = req.query;

  const rows = filterTransactions(store, {
    bank_id,
    from,
    to,
    bankIds: companyBankIds(store, activeCompany),
  }).slice(0, 500);

  res.render("transactions", {
    banks,
    rows,
    filters: { bank_id: bank_id || "", from: from || "", to: to || "" },
    userName: req.session.userName,
    pasteResult: null,
    uploadResult: null,
    error: null,
  });
});

router.post("/transactions", (req, res) => {
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

router.post("/transactions/paste", (req, res) => {
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

  const currentRows = filterTransactions(store, {
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
router.post("/transactions/upload-statement", upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const { bank_id } = req.body;

  const renderError = (message) =>
    res.render("transactions", {
      banks,
      rows: filterTransactions(store, {
        bank_id: bank_id || "",
        bankIds: companyBankIds(store, activeCompany),
      }).slice(0, 500),
      filters: { bank_id: bank_id || "", from: "", to: "" },
      userName: req.session.userName,
      pasteResult: null,
      uploadResult: null,
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

  const currentRows = filterTransactions(store, {
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
    error: null,
  });
});

router.post("/transactions/:id/delete", (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  store.transactions = store.transactions.filter((t) => t.id !== id);
  save(store);
  res.redirect("back");
});

router.get("/export.xlsx", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const { bank_id, from, to } = req.query;
  const rows = filterTransactions(store, {
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
