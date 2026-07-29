const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireDataEntry, requireAdmin } = require("../middleware/auth");
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

// Chi Nhan (2026-07-23): "sao tôi tải ngân hàng kh cũ không được" -- loi
// "totalThu is not defined" khi render lai view "transactions" tu cac route
// /transactions/paste, /transactions/upload-statement, /transactions/upload-
// ma-cong-trinh: view can totalThu/totalChi (them luc lam tinh nang tach cot
// Thu/Chi tren GET /transactions) nhung 3 route con lai o duoi chua tung
// duoc cap nhat de tinh va truyen 2 gia tri nay, nen bi crash (500) MOI LAN
// dung 1 trong 3 route do (vd tai sao ke ngan hang) du GET /transactions binh
// thuong van chay tot. Ham dung chung nay tinh dung 1 lan, dung cho ca 4 noi
// render "transactions" de khong con lech nhau nua.
function computeThuChiTotals(allFilteredRows) {
  const totalThu = allFilteredRows
    .filter((t) => t.type === "thu")
    .reduce((s, t) => s + Number(t.amount || 0), 0);
  const totalChi = allFilteredRows
    .filter((t) => t.type === "chi")
    .reduce((s, t) => s + Number(t.amount || 0), 0);
  return { totalThu, totalChi };
}

// Chi Nhan, 2026-07-24: Luyen phat hien tong "Ngan hang" cua 1 ngay bi cong
// SAI cao gap nhieu lan so tien that (235 trieu thay vi ~21 trieu that cua
// ngay 22/07, BIDV7702) -- goc re: qua nhieu lan tai lai sao ke trong luc do
// bug "So chung tu bi nham la So tham chieu" (xem utils/bankStatementParser.js)
// khien tu-sua (self-heal, xem /transactions/upload-statement) chi don duoc
// CUNG mot lan tai (dung dung 1 khoang ngay), con cac lan tai KHAC (pham vi
// ngay khac nhau moi lan, do Luyen thu nhieu file khac nhau trong luc debug)
// khong nam trong dung khoang [firstDate,lastDate] cua lan tai MOI NHAT nen
// khong duoc don, de lai nhieu ban ghi TRUNG (cung 1 giao dich that nhung
// Reference khac nhau -- 1 ban ghi cu voi "So chung tu" sai, 1 ban ghi moi
// voi "So tham chieu" dung). Vi Reference KHAC nhau nen khong the dedup theo
// Reference; nhung Ngay+So tien+Loai+Dien giai CHAC CHAN giong het nhau cho
// CUNG 1 giao dich that (Dien giai doc tu 1 cot co dinh, khong phu thuoc bug
// Reference) -- dung 4 truong nay lam khoa nhom de tim ban trung.
function findDuplicateGroups(store, bankId) {
  const groups = new Map();
  store.transactions.forEach((t) => {
    if (t.bank_id !== bankId) return;
    const key = `${t.date}|${t.amount}|${t.type}|${t.description}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(t);
  });
  const dupGroups = [];
  groups.forEach((rows) => {
    if (rows.length > 1) dupGroups.push(rows);
  });
  dupGroups.sort((a, b) => (a[0].date < b[0].date ? 1 : a[0].date > b[0].date ? -1 : 0));
  return dupGroups;
}

function summarizeDuplicates(dupGroups) {
  let excessCount = 0;
  let excessAmount = 0;
  dupGroups.forEach((rows) => {
    excessCount += rows.length - 1;
    excessAmount += rows[0].amount * (rows.length - 1);
  });
  return { groupCount: dupGroups.length, excessCount, excessAmount };
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

  // Chi Nhan, 2026-07-24: chi tinh khi da loc dung 1 ngan hang cu the (tranh
  // quet toan bo store.transactions moi lan mo trang /transactions chung).
  let duplicateSummary = null;
  if (bank_id) {
    const dupGroups = findDuplicateGroups(store, Number(bank_id));
    if (dupGroups.length > 0) duplicateSummary = summarizeDuplicates(dupGroups);
  }

  res.render("transactions", {
    banks,
    rows,
    totalThu,
    totalChi,
    filteredCount: allFiltered.length,
    filters: { bank_id: bank_id || "", from: from || "", to: to || "", type: type || "" },
    userName: req.session.userName,
    pasteResult: null,
    uploadResult: null,
    maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
    maCongTrinhResult: null,
    duplicateSummary,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/transactions", requireDataEntry, (req, res) => {
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

// Chi Nhan (2026-07-23): "số tiền có 8tr mấy mà bạn lấy lên mất tỷ dữ vậy" --
// phat hien 1 giao dich nhap tay bi sai so tien rat nang (2.324.723.431d thay
// vi 8.250.000d, chac do dan nham/go nham khi nhap tay). Truoc gio trang nay
// CHUA CO cach sua 1 giao dich da co san (chi co them moi qua form/dan/tai
// file) -- them route sua truc tiep ngay/dien giai/so tien/loai cho 1 dong,
// dung cho khi phat hien nhap sai nhu the nay.
router.post("/transactions/:id/sua", requireDataEntry, (req, res) => {
  const store = load();
  const tx = store.transactions.find((t) => t.id === Number(req.params.id));
  if (!tx) {
    if (req.get("X-Requested-With") === "XMLHttpRequest") {
      return res.status(404).json({ error: "Không tìm thấy giao dịch này." });
    }
    return res.redirect("/transactions?error=" + encodeURIComponent("Không tìm thấy giao dịch này."));
  }
  try {
    const { date, description, amount, type, tenDoiUng } = req.body;
    const parsedDate = parseDate(date) || date;
    const parsedAmount = Math.abs(parseAmount(amount));
    if (!parsedDate || isNaN(parsedAmount) || !["thu", "chi"].includes(type)) {
      throw new Error("Dữ liệu không hợp lệ (kiểm tra lại ngày/số tiền/loại).");
    }
    tx.date = parsedDate;
    tx.description = (description || "").trim();
    // Chi Nhan, 2026-07-29: cho sua tay Ten doi ung tu day (vd dong cu tai
    // truoc khi co tinh nang nay nen dang trong, hoac file sao ke khong co
    // san cot nay).
    tx.tenDoiUng = (tenDoiUng || "").trim();
    tx.amount = parsedAmount;
    tx.type = type;
    tx.edited_at = new Date().toISOString();
    tx.edited_by = req.session.userName || "";
    save(store);
    if (req.get("X-Requested-With") === "XMLHttpRequest") {
      return res.json({ success: true, tx: withBankLabel(store, tx) });
    }
    res.redirect("/transactions?success=" + encodeURIComponent("Đã sửa giao dịch."));
  } catch (e) {
    if (req.get("X-Requested-With") === "XMLHttpRequest") {
      return res.status(400).json({ error: e.message });
    }
    res.redirect("/transactions?error=" + encodeURIComponent(e.message));
  }
});

router.post("/transactions/paste", requireDataEntry, (req, res) => {
  const { bank_id, paste_text } = req.body;
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);

  if (!bank_id) {
    return res.render("transactions", {
      banks,
      rows: [],
      totalThu: 0,
      totalChi: 0,
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

  const allFilteredAfterPaste = filterTransactionsWithBalance(store, {
    bank_id,
    bankIds: companyBankIds(store, activeCompany),
  });
  const currentRows = allFilteredAfterPaste.slice(0, 500);

  res.render("transactions", {
    banks,
    rows: currentRows,
    ...computeThuChiTotals(allFilteredAfterPaste),
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
router.post("/transactions/upload-statement", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const { bank_id } = req.body;

  const renderError = (message) => {
    const allFilteredForError = filterTransactionsWithBalance(store, {
      bank_id: bank_id || "",
      bankIds: companyBankIds(store, activeCompany),
    });
    return res.render("transactions", {
      banks,
      rows: allFilteredForError.slice(0, 500),
      ...computeThuChiTotals(allFilteredForError),
      filters: { bank_id: bank_id || "", from: "", to: "" },
      userName: req.session.userName,
      pasteResult: null,
      uploadResult: null,
      maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
      maCongTrinhResult: null,
      error: message,
    });
  };

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
    // Chi Nhan, 2026-07-24: sua tiep 1 bug rieng vua phat hien -- truoc day
    // "So chung tu" (so phieu tuan tu, vd "75918") bi nham thanh reference
    // thay vi "So tham chieu" that (dung de khop VietQR theo Ma tham chieu,
    // xem utils/vietqrReconcile.js) tren cac file co CA HAI cot nay (BIDV7702).
    // Nhung dong DA LUU truoc khi sua bug se mang gia tri reference SAI (kieu
    // so ngan, thuan so) nen khong the tu nhan lai o buoc "existingRefs" ben
    // duoi -- can coi nhung dong do NHU THE KHONG CO reference (giong het
    // truong hop "!t.reference" da xu ly san o day) de duoc xoa+nap lai voi
    // reference DUNG, nhung CHI khi lan tai nay thuc su doc duoc "So tham
    // chieu" that (refIsPrimary) -- tranh dong cham nham toi cac ngan hang
    // khac von CHI CO "So chung tu" hop le (khong phai loi voi ho).
    const looksLikeStaleNumericRef = (ref) => /^\d{1,10}$/.test(String(ref || ""));
    const shouldHealStaleNumericRef = !!parsed.refIsPrimary;
    const before = store.transactions.length;
    store.transactions = store.transactions.filter((t) => {
      if (t.bank_id !== bankIdNum || t.date < firstDate || t.date > lastDate) return true;
      if (!t.reference) return false;
      if (shouldHealStaleNumericRef && looksLikeStaleNumericRef(t.reference)) return false;
      return true;
    });
    healedRemoved = before - store.transactions.length;
  }

  const bankTx = store.transactions.filter((t) => t.bank_id === bankIdNum);
  // Chi Nhan, 2026-07-29: doi tu Set sang Map (reference -> chinh dong da
  // luu) de co the SUA (backfill) truc tiep dong da co san, khong chi biet
  // "co roi" nhu Set truoc day.
  const existingByRef = new Map(bankTx.filter((t) => t.reference).map((t) => [t.reference, t]));
  const existingKeys = new Set(
    bankTx.filter((t) => !t.reference).map((t) => `${t.date}|${t.amount}|${t.type}`)
  );

  let added = 0;
  let skipped = 0;
  // Chi Nhan, 2026-07-29: "tôi đã tải sao kê bên ngân hàng rồi ... sao lại
  // không có [Tên đối ứng]" -- tai khoan da co san du lieu TU TRUOC (luc
  // chua doc duoc cot Ten doi ung) se bi tinh la "trung" theo reference va bo
  // qua nhu cu (khong tao dong moi), nen KHONG BAO GIO duoc dien Ten doi ung
  // moi neu chi dung logic upsert cu. Backfill THEM o day: dong da co san nao
  // dang TRONG tenDoiUng ma lan doc file nay ra duoc gia tri thi dien vao
  // (khong ghi de neu dong do da co san Ten doi ung, tranh mat du lieu da co).
  let backfilledVendor = 0;
  for (const c of candidates) {
    if (c.reference) {
      const existingRow = existingByRef.get(c.reference);
      if (existingRow) {
        skipped++;
        if (!existingRow.tenDoiUng && c.tenDoiUng) {
          existingRow.tenDoiUng = c.tenDoiUng;
          backfilledVendor++;
        }
        continue;
      }
    } else {
      const key = `${c.date}|${c.amount}|${c.type}`;
      if (existingKeys.has(key)) {
        skipped++;
        continue;
      }
      existingKeys.add(key);
    }
    const newRow = {
      id: nextId(store, "transactions"),
      bank_id: bankIdNum,
      date: c.date,
      description: c.description,
      amount: c.amount,
      type: c.type,
      reference: c.reference || "",
      // Chi Nhan, 2026-07-29: "thêm cho tôi trên cái ngân hàng có hiển thị
      // cái tên đối ứng trên sao kê" -- luu them Ten doi ung (neu file sao ke
      // co cot nay, xem utils/bankStatementParser.js) de hien tren bang Giao
      // dich va dung doi soat Da chi tien ben Hoa Don Dau Vao.
      tenDoiUng: c.tenDoiUng || "",
      created_at: new Date().toISOString(),
      created_by: req.session.userName || "",
    };
    store.transactions.push(newRow);
    if (c.reference) existingByRef.set(c.reference, newRow);
    added++;
  }
  if (added > 0 || healedRemoved > 0 || backfilledVendor > 0) save(store);

  const allFilteredAfterUpload = filterTransactionsWithBalance(store, {
    bank_id,
    bankIds: companyBankIds(store, activeCompany),
  });
  const currentRows = allFilteredAfterUpload.slice(0, 500);
  res.render("transactions", {
    banks,
    rows: currentRows,
    ...computeThuChiTotals(allFilteredAfterUpload),
    filters: { bank_id, from: "", to: "" },
    userName: req.session.userName,
    pasteResult: null,
    uploadResult: {
      sheetName: parsed.sheetName,
      totalRows: parsed.rows.length,
      added,
      skipped,
      healedRemoved,
      backfilledVendor,
    },
    maCongTrinhMaster: maCongTrinhMasterFor(store, activeCompany),
    maCongTrinhResult: null,
    error: null,
  });
});

// Chi Nhan, 2026-07-24: don giao dich TRUNG LAP cho 1 ngan hang (xem
// findDuplicateGroups o tren) -- moi nhom trung giu lai DUNG 1 dong (uu tien
// dong co Reference "trong" hop le (khong phai chuoi thuan so kieu "So chung
// tu" cu -- xem utils/bankStatementParser.js) hon dong Reference thuan so cu;
// neu ca 2 dong deu cung loai, giu dong created_at moi nhat), xoa cac dong con
// lai trong nhom. requireAdmin (thao tac xoa hang loat, khong the hoan tac).
function looksLikeStaleNumericRef(ref) {
  return /^\d{1,10}$/.test(String(ref || ""));
}
router.post("/transactions/dedupe/:bank_id", requireAdmin, (req, res) => {
  const store = load();
  const bankIdNum = Number(req.params.bank_id);
  try {
    const dupGroups = findDuplicateGroups(store, bankIdNum);
    if (dupGroups.length === 0) {
      return res.redirect(`/transactions?bank_id=${bankIdNum}&success=` + encodeURIComponent("Khong tim thay giao dich trung lap nao."));
    }
    const idsToRemove = new Set();
    let removedCount = 0;
    let removedAmount = 0;
    dupGroups.forEach((rows) => {
      const sorted = [...rows].sort((a, b) => {
        const aGood = a.reference && !looksLikeStaleNumericRef(a.reference) ? 1 : 0;
        const bGood = b.reference && !looksLikeStaleNumericRef(b.reference) ? 1 : 0;
        if (aGood !== bGood) return bGood - aGood; // uu tien reference "tot" hon len truoc
        return (b.created_at || "").localeCompare(a.created_at || ""); // moi nhat truoc
      });
      // Giu lai sorted[0], xoa phan con lai.
      sorted.slice(1).forEach((t) => {
        idsToRemove.add(t.id);
        removedCount += 1;
        removedAmount += t.amount;
      });
    });
    store.transactions = store.transactions.filter((t) => !idsToRemove.has(t.id));
    save(store);
    res.redirect(
      `/transactions?bank_id=${bankIdNum}&success=` +
        encodeURIComponent(`Đã xoá ${removedCount} dòng trùng lặp, giảm ${removedAmount.toLocaleString("vi-VN")}đ bị cộng thừa.`)
    );
  } catch (e) {
    res.redirect(`/transactions?bank_id=${bankIdNum}&error=` + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-24: "thêm chỗ xóa các giao dịch đã lọc ra á xóa đồng loạt" --
// sao ke 1 ngan hang/1 ngay bi loi/trung nang (vd 235tr thay vi ~21tr thuc te
// cua BIDV7702 ngay 22/07) can 1 cach xoa SACH toan bo giao dich dang loc de
// tai lai file sao ke DUNG tu dau, thay vi phai bam "Sua" tung dong 1 (khong
// co gioi han so dong). Xoa THEO DUNG bo loc dang ap dung tren trang (ngan
// hang + tu ngay + den ngay + loai Thu/Chi) -- bat buoc phai chon 1 ngan hang
// cu the (khong cho xoa khi dang xem "Tat ca ngan hang") de tranh xoa nham
// toan bo du lieu nhieu ngan hang cung luc. Admin-only vi la thao tac xoa
// hang loat khong the hoan tac.
router.post("/transactions/xoa-loc", requireAdmin, (req, res) => {
  const { bank_id, from, to, type } = req.body;
  if (!bank_id) {
    return res.redirect("/transactions?error=" + encodeURIComponent("Phai chon 1 ngan hang cu the truoc khi xoa hang loat."));
  }
  const store = load();
  const bankIdNum = Number(bank_id);
  const before = store.transactions.length;
  let removedAmount = 0;
  store.transactions = store.transactions.filter((t) => {
    if (t.bank_id !== bankIdNum) return true;
    if (from && t.date < from) return true;
    if (to && t.date > to) return true;
    if ((type === "thu" || type === "chi") && t.type !== type) return true;
    removedAmount += t.amount;
    return false;
  });
  const removedCount = before - store.transactions.length;
  save(store);
  const qs = `bank_id=${encodeURIComponent(bank_id)}&from=${encodeURIComponent(from || "")}&to=${encodeURIComponent(to || "")}&type=${encodeURIComponent(type || "")}`;
  res.redirect(
    `/transactions?${qs}&success=` +
      encodeURIComponent(`Đã xoá ${removedCount} dòng giao dịch theo bộ lọc, tổng ${removedAmount.toLocaleString("vi-VN")}đ.`)
  );
});

// Upload danh sach "Ma cong trinh" chuan (rieng theo tung cong ty dang chon).
// Luyen, 2026-07-20: file "DANH SACH CONG TRINH" tu phan mem quan ly cong
// trinh -- dung lam nguon chuan de sau nay doi chieu/chuan hoa cac ten
// gian/cong trinh xuat hien o cac trang khac (vd Chi Phi). Moi lan tai len
// THAY THE toan bo danh sach cua dung cong ty dang chon (KH Cu / KH Moi
// khong dung chung 1 danh sach vi la 2 phap nhan khac nhau).
router.post("/transactions/upload-ma-cong-trinh", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany);
  const allFilteredForMaCongTrinh = filterTransactionsWithBalance(store, {
    bankIds: companyBankIds(store, activeCompany),
  });
  const rowsForList = allFilteredForMaCongTrinh.slice(0, 500);

  const renderWith = (maCongTrinhResult, error) =>
    res.render("transactions", {
      banks,
      rows: rowsForList,
      ...computeThuChiTotals(allFilteredForMaCongTrinh),
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
      "Ten doi ung": t.tenDoiUng || "",
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
