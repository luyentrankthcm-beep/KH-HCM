const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireDataEntry, requireAdmin } = require("../middleware/auth");
const { parsePastedTransactions, parseAmount, parseDate } = require("../utils/parse");
const { parseBankStatement, computeThuChi } = require("../utils/bankStatementParser");
const { getCompany } = require("../utils/companies");
const { parseMaCongTrinhSheet } = require("../utils/maCongTrinh");
const { extractGianRentText, matchGianRecord } = require("../utils/rentPaymentMatcher");

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

// Luyen, 2026-08-01: "cần để mốt tôi tải nhầm tôi có thể xóa á" -- danh sach
// lich su tai sao ke (store.bank_statement_uploads) de hien tren trang, loc
// theo ngan hang cua cong ty dang xem (giong cach banks/rows da loc), moi
// nhat truoc. Dung chung cho MOI noi render view "transactions", giong
// computeThuChiTotals o tren -- tranh crash "bankStatementUploads is not
// defined" o 1 trong cac route neu quen truyen.
function bankStatementUploadsFor(store, bankIds) {
  return (store.bank_statement_uploads || [])
    .filter((u) => bankIds.has(u.bank_id))
    .slice()
    .sort((a, b) => (a.uploaded_at < b.uploaded_at ? 1 : a.uploaded_at > b.uploaded_at ? -1 : 0))
    .slice(0, 30);
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
    bankStatementUploads: bankStatementUploadsFor(store, companyBankIds(store, activeCompany)),
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
// Lightweight route: only update tenDoiUng (used by inline-edit NCC cell)
router.post("/transactions/:id/update-ncc", requireDataEntry, (req, res) => {
  const store = load();
  const tx = store.transactions.find((t) => t.id === Number(req.params.id));
  if (!tx) return res.status(404).json({ ok: false, message: "Không tìm thấy." });
  tx.tenDoiUng = (req.body.tenDoiUng || "").trim();
  save(store);
  res.json({ ok: true, tenDoiUng: tx.tenDoiUng });
});

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
    const { date, description, amount, type, tenDoiUng, excludeFromVietQrRecon } = req.body;
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
    // Chi Nhan, 2026-07-29: "hải phòng làm gì cộng dữ liệu nhiều vậy check
    // lại xem có bị trùng hk 77021" -- dieu tra ra: KHONG phai trung du lieu,
    // ma 1 giao dich "thu" khong phai tien khach tra qua VietQR (vd tien doi
    // tac chuyen trang toan doanh thu chia se theo ky/thang, khong co Ma tham
    // chieu khop voi file xuat QR) bi Doi soat VietQR gom NHAM vao gian mac
    // dinh (AE HP PHN, xem CHANNELS.bidv77021.defaultBlankCode trong
    // routes/doisoat-vietqr.js) lam sai lech ca ngay. utils/vietqrReconcile.js
    // da co san co che loai tru (t.excludeFromVietQrRecon, dung cho dung
    // truong hop nay -- xem extractVietQrThuTransactions) nhung CHUA co cho
    // nao tren giao dien de tu bat/tat -- them checkbox nay tai day de Chi
    // Nhan tu xu ly duoc cac truong hop tuong tu sau nay (giao dich VAN o lai
    // Giao dich/Sao ke binh thuong, chi khong tinh vao doanh thu QR nua).
    tx.excludeFromVietQrRecon = excludeFromVietQrRecon === "1" || excludeFromVietQrRecon === "on";
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

// Chi Nhan, 2026-08-05: "sao tôi chuyển qua lại giữa các trang KH cũ và KH
// [moi] bị z ta" -- truoc day route nay (va upload-statement/upload-ma-cong-
// trinh ben duoi) xu ly xong roi res.render() THANG trang "transactions",
// khien thanh dia chi TRINH DUYET van dung nguyen "/transactions/paste" (URL
// cua chinh request POST nay). Nut chuyen cong ty tren topbar (POST /chon-
// cong-ty, xem routes/company.js) lay redirectTo = duong dan dang xem LUC DO
// de quay lai sau khi doi cong ty -- nen no lay nham "/transactions/paste"
// (chi co POST, khong co GET) roi redirect toi do bang GET, ra loi 404 "Khong
// tim thay trang". Sua bang cach LUON res.redirect() ve "/transactions" (GET,
// co that) sau khi xu ly xong, dung chung co che success/error o query string
// da co san (xem GET /transactions ben tren) thay vi tu render rieng.
// Luyen, 2026-08-14: them nut xoa 1 giao dich don le (admin only) -- truoc day
// chi co "xoa hang loat theo bo loc" (xoa-loc), khong co xoa don le nen giao
// dich nhap sai (vd 409M phantom) khong co cach xoa nhanh tren UI.
router.post("/transactions/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  const idx = store.transactions.findIndex((t) => t.id === id);
  if (idx < 0) return res.redirect("/transactions?error=" + encodeURIComponent("Không tìm thấy giao dịch."));
  const tx = store.transactions[idx];
  store.transactions.splice(idx, 1);
  save(store);
  res.redirect("/transactions?success=" + encodeURIComponent(`Đã xoá giao dịch ${tx.date} ${tx.amount.toLocaleString("vi-VN")}đ.`));
});

router.post("/transactions/paste", requireDataEntry, (req, res) => {
  const { bank_id, paste_text } = req.body;
  const store = load();

  if (!bank_id) {
    return res.redirect(
      "/transactions?error=" + encodeURIComponent("Vui long chon ngan hang truoc khi dan sao ke.")
    );
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

  let message = `Đã nhập thành công ${rows.length} dòng.`;
  if (errors.length > 0) {
    message += ` Có ${errors.length} dòng bị bỏ qua: ${errors.slice(0, 5).join("; ")}${
      errors.length > 5 ? "…" : ""
    }`;
  }
  res.redirect(
    "/transactions?bank_id=" + encodeURIComponent(bank_id) + "&success=" + encodeURIComponent(message)
  );
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
// Chi Nhan, 2026-08-05: doi tu res.render() sang res.redirect() sau khi xu ly
// xong -- ly do xem comment dai o /transactions/paste ngay phia tren (bug
// "Khong tim thay trang" luc chuyen KH Cu/KH Moi vi thanh dia chi ket o
// "/transactions/upload-statement", 1 URL chi co POST khong co GET).
router.post("/transactions/upload-statement", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  const { bank_id } = req.body;

  const redirectError = (message) =>
    res.redirect(
      "/transactions?bank_id=" + encodeURIComponent(bank_id || "") + "&error=" + encodeURIComponent(message)
    );

  if (!bank_id) return redirectError("Vui long chon ngan hang truoc khi tai file sao ke.");
  if (!req.file) return redirectError("Vui long chon 1 file sao ke de tai len.");

  const bankIdNum = Number(bank_id);
  const bank = store.banks.find((b) => b.id === bankIdNum);
  if (!bank) return redirectError("Khong tim thay ngan hang da chon.");

  let parsed;
  try {
    parsed = parseBankStatement(req.file.buffer);
  } catch (e) {
    return redirectError(e.message);
  }

  const firstDate = parsed.rows[0].date;
  const lastDate = parsed.rows[parsed.rows.length - 1].date;
  const priorBalance = bankBalanceBefore(store, bankIdNum, firstDate);
  const candidates = computeThuChi(parsed.rows, priorBalance);
  // Luyen, 2026-08-15: "tôi cần gắn vào đây mà" -- dung de dien NCC (benChoThue)
  // vao Ten doi ung cua giao dich chi tien thue gian (VPBANK9997...) khi import.
  const gianListForImport = store.phap_danh_hop_dong_thue || [];

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
  const insertedIds = [];
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
        } else if (!existingRow.tenDoiUng && c.type === "chi") {
          const gianText = extractGianRentText(c.description);
          if (gianText) {
            const rec = matchGianRecord(gianText, gianListForImport);
            if (rec && rec.benChoThue) {
              existingRow.tenDoiUng = rec.benChoThue;
              backfilledVendor++;
            }
          }
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
    // Luyen, 2026-08-15: neu giao dich chi tien thue gian ma file sao ke khong
    // co cot Ten doi ung (vd VPBANK9997 file Excel co the de trong), tu dien
    // benChoThue tu hop dong thue gian de hien tren trang sao ke.
    let tenDoiUngForRow = c.tenDoiUng || "";
    if (!tenDoiUngForRow && c.type === "chi") {
      const gianText = extractGianRentText(c.description);
      if (gianText) {
        const rec = matchGianRecord(gianText, gianListForImport);
        if (rec && rec.benChoThue) tenDoiUngForRow = rec.benChoThue;
      }
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
      tenDoiUng: tenDoiUngForRow,
      created_at: new Date().toISOString(),
      created_by: req.session.userName || "",
    };
    store.transactions.push(newRow);
    if (c.reference) existingByRef.set(c.reference, newRow);
    insertedIds.push(newRow.id);
    added++;
  }
  // Luyen, 2026-08-01: "cần để mốt tôi tải nhầm tôi có thể xóa á" -- ghi lai
  // DUNG cac id giao dich vua tao trong lan tai nay (khong phai cac dong da
  // co san bi bo qua/backfill) de co the xoa nguyen 1 dot qua nut rieng neu
  // phat hien tai nham file/nham ngan hang (xem vu MB02865168 nhan nham sao
  // ke BIDV8681, 2026-08-01 -- luc do phai do tay id vi chua co so nay).
  if (insertedIds.length > 0) {
    if (!store.bank_statement_uploads) store.bank_statement_uploads = [];
    store.bank_statement_uploads.push({
      id: nextId(store, "bank_statement_uploads_seq") || Date.now(),
      bank_id: bankIdNum,
      bank_name: bank.name,
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      uploaded_at: new Date().toISOString(),
      uploaded_by: req.session.userName || "",
      transaction_ids: insertedIds,
      rows_inserted: insertedIds.length,
      rows_skipped: skipped,
    });
  }
  if (added > 0 || healedRemoved > 0 || backfilledVendor > 0) save(store);

  let message = `Đọc file (sheet "${parsed.sheetName}"): ${parsed.rows.length} dòng giao dịch. Đã thêm mới ${added} dòng, bỏ qua ${skipped} dòng đã có sẵn (trùng).`;
  if (healedRemoved > 0) {
    message += ` Đã tự động dọn ${healedRemoved} dòng cũ bị gộp nhầm trước khi nạp lại đầy đủ.`;
  }
  if (backfilledVendor > 0) {
    message += ` Đã điền thêm Tên đối ứng cho ${backfilledVendor} dòng cũ đang trống.`;
  }
  res.redirect(
    "/transactions?bank_id=" + encodeURIComponent(bank_id) + "&success=" + encodeURIComponent(message)
  );
});

// Luyen, 2026-08-01: "cần để mốt tôi tải nhầm tôi có thể xóa á" -- xoa nguyen
// 1 dot tai sao ke (dung DUNG cac giao dich do lan tai do tao ra, xem
// transaction_ids da luu san o /transactions/upload-statement), khong dung
// filter bank+ngay+loai chung (se dinh phai cac giao dich khac cung ngay,
// giong vu MB02865168/BIDV8681). requireAdmin giong cac route xoa hang loat
// khac trong file nay (khong the hoan tac).
router.post("/transactions/upload-statement/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  const uploadId = Number(req.params.id);
  const upload = (store.bank_statement_uploads || []).find((u) => u.id === uploadId);
  if (!upload) {
    return res.redirect("/transactions?error=" + encodeURIComponent("Khong tim thay lan tai nay (co the da bi xoa)."));
  }
  const idsToRemove = new Set(upload.transaction_ids || []);
  const before = store.transactions.length;
  store.transactions = store.transactions.filter((t) => !idsToRemove.has(t.id));
  const removed = before - store.transactions.length;
  store.bank_statement_uploads = store.bank_statement_uploads.filter((u) => u.id !== uploadId);
  save(store);
  res.redirect(
    "/transactions?bank_id=" +
      encodeURIComponent(upload.bank_id) +
      "&success=" +
      encodeURIComponent(`Da xoa ${removed} giao dich cua lan tai "${upload.file_name}".`)
  );
});

// temp debug: xem bank ids va mau giao dich VP58888
router.get("/transactions/debug-unc", requireDataEntry, (req, res) => {
  const store = load();
  const banks = store.banks.map(b => ({id: b.id, name: b.name, company: b.company || 'n/a', account: b.accountNumber || ''}));
  const moiIds = store.banks.filter(b => b.company === 'kh_moi').map(b => b.id);
  const cuIds = store.banks.filter(b => b.company === 'kh_cu' || !b.company).map(b => b.id);
  const chiMoi = (store.transactions || []).filter(t => t.type === 'chi' && moiIds.includes(t.bank_id)).slice(0, 10).map(t => ({id: t.id, bank_id: t.bank_id, desc: t.description, tenDoiUng: t.tenDoiUng || ''}));
  const chiCu = (store.transactions || []).filter(t => t.type === 'chi' && cuIds.includes(t.bank_id)).slice(0, 5).map(t => ({id: t.id, bank_id: t.bank_id, desc: t.description, tenDoiUng: t.tenDoiUng || ''}));
  res.json({ banks, moiIds, cuIds, chiMoi, chiCu });
});

// Luyen, 2026-08-15: "bạn dựa vào diễn giải á của tài khoản chi VP5888 với lại
// 9997 á bạn check trên link đi UNC Miền nam ... lấy ra tên ncc của lệnh đi
// tiền đó nhá" -- nhan mapping {cu: {desc: ncc, ...}, moi: {desc: ncc, ...}}
// tu UNC Mien Nam Google Sheet, ap dung vao cac giao dich CHI dang trong
// Ten doi ung cua VP9997 (KH cu) va VP58888 (KH moi).
//
// Matching: normalize ca hai chuoi (lowercase, bo dau cach thua), check
// containment (UNC desc la substring cua bank desc, hoac nguoc lai) -- dung
// cho truong hop bank desc co prefix/suffix ngan hang khong co trong UNC.
router.post("/transactions/apply-unc-ncc", requireDataEntry, (req, res) => {
  const store = load();
  // Doc tu file uncNccMap.json trong utils/ (cap nhat tu UNC Mien Nam GSheet)
  let cu, moi;
  try {
    const mapPath = require("path").join(__dirname, "../utils/uncNccMap.json");
    const mapData = JSON.parse(require("fs").readFileSync(mapPath, "utf-8"));
    cu = mapData.cu;
    moi = mapData.moi;
  } catch (e) {
    return res.status(500).json({ error: "Khong doc duoc uncNccMap.json: " + e.message });
  }
  if (!cu && !moi) return res.status(400).json({ error: "Thieu du lieu mapping" });

  // Tim bank IDs cho VP9997 (kh_cu) va VP58888 (kh_moi)
  const cuBankIds = new Set(
    store.banks.filter((b) => b.company === "kh_cu" || !b.company).map((b) => b.id)
  );
  const moiBankIds = new Set(
    store.banks.filter((b) => b.company === "kh_moi").map((b) => b.id)
  );

  // Normalize: bo dau, lowercase, collapse spaces, sau do bo prefix ten cong ty
  // K&H (UNC dung "CTY K VA H TT", sao ke VP58888 dung "Cty TNHH GIAI TRI K VA H tt",
  // sao ke VP9997 cung co the khac) -- chi giu lai phan noi dung thanh toan thuc su.
  function removeDiac(s) {
    return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/gi, "d");
  }
  const PAYER_PREFIX_RE = /^(?:cong\s+ty\s+tnhh\s+giai\s+tri\s+)?(?:cty\s+tnhh\s+giai\s+tri\s+)?(?:cong\s+ty\s+)?(?:cty\s+)?k(?:\s+va?\s+|\s*&\s*)h(?:\s+tnhh(?:\s+giai\s+tri)?)?\s+tt?\s*/i;

  function normDesc(s) {
    let n = removeDiac(String(s || "")).toLowerCase().replace(/\s+/g, " ").trim();
    // Bo prefix "xcty k va h tt" (co chu X dau nham)
    n = n.replace(/^x/, "").trim();
    // Bo prefix ten cong ty K&H
    n = n.replace(PAYER_PREFIX_RE, "").trim();
    // Bo ky tu & va \ (sao ke ngan hang khong co & trong ten cong ty, UNC GSheet
    // markdown render lai & thanh \&, sau khi clean key van con the khi so sanh)
    n = n.replace(/[&\\]/g, "").replace(/\s+/g, " ").trim();
    // Luyen, 2026-08-17: VP sao ke cat "Thang" thanh "T hang" (xuong dong giua tu)
    // trong o dien giai -- normalize lai thanh "thang" de khop voi UNC map key.
    n = n.replace(/\bt\s+hang\b/g, "thang");
    return n;
  }

  // Build lookup list: [{norm: ..., ncc: ...}] for fast substring matching
  const cuList = Object.entries(cu || {}).map(([d, n]) => ({ norm: normDesc(d), ncc: n }));
  const moiList = Object.entries(moi || {}).map(([d, n]) => ({ norm: normDesc(d), ncc: n }));

  function findNCC(txDesc, list) {
    const txNorm = normDesc(txDesc);
    if (!txNorm || txNorm.length < 8) return null;
    // 1. Exact match
    for (const { norm, ncc } of list) {
      if (txNorm === norm) return ncc;
    }
    // 2. UNC desc contained in bank desc
    for (const { norm, ncc } of list) {
      if (norm.length >= 10 && txNorm.includes(norm)) return ncc;
    }
    // 3. Bank desc contained in UNC desc
    for (const { norm, ncc } of list) {
      if (txNorm.length >= 10 && norm.includes(txNorm)) return ncc;
    }
    return null;
  }

  let updated = 0;
  (store.transactions || []).forEach((t) => {
    // Luyen, 2026-08-17: "sao ke la chi het ma soa ban bo qua thu" -- mot so
    // giao dich la CHI tren sao ke nhung bi parse thanh "thu" do delta so du
    // tinh sai (vi du so du luy ke bi lech 1 dong). Bo dieu kien chi xu ly
    // "chi" de UNC matching ap dung cho ca 2 loai -- NCC trong UNC luon la
    // nguoi NHAN tien tu K&H (chi), neu tim thay khop trong UNC thi chac chan
    // la giao dich chi du parser co the da classify nham.
    if (t.tenDoiUng && t.tenDoiUng.trim()) return;
    const isCu = cuBankIds.has(t.bank_id);
    const isMoi = moiBankIds.has(t.bank_id);
    if (!isCu && !isMoi) return;
    const list = isCu ? cuList : moiList;
    const ncc = findNCC(t.description, list);
    if (ncc) {
      t.tenDoiUng = ncc;
      updated++;
    }
  });

  if (updated > 0) save(store);
  res.json({ success: true, updated });
});

// Luyen, 2026-08-15: "tôi cần gắn vào đây mà" -- backfill Ten doi ung = NCC
// (benChoThue) cho cac giao dich CHI tien thue gian da co san ma dang de trong
// Ten doi ung (giao dich nhap truoc khi co tinh nang nay). Dung extractGianRentText
// + matchGianRecord de nhan biet va ghep hop dong.
router.post("/transactions/backfill-ncc-ten-doi-ung", requireDataEntry, (req, res) => {
  const store = load();
  const gianList = store.phap_danh_hop_dong_thue || [];
  let updated = 0;
  (store.transactions || []).forEach((t) => {
    if (t.type !== "chi") return;
    if (t.tenDoiUng && t.tenDoiUng.trim()) return;
    const gianText = extractGianRentText(t.description);
    if (!gianText) return;
    const rec = matchGianRecord(gianText, gianList);
    if (rec && rec.benChoThue) {
      t.tenDoiUng = rec.benChoThue;
      updated++;
    }
  });
  if (updated > 0) save(store);
  res.json({ success: true, updated });
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
// Chi Nhan, 2026-08-05: doi tu res.render() sang res.redirect() -- ly do xem
// comment dai o /transactions/paste ben tren (bug "Khong tim thay trang" luc
// chuyen KH Cu/KH Moi).
router.post("/transactions/upload-ma-cong-trinh", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);

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
    return res.redirect(
      "/transactions?success=" +
        encodeURIComponent(`Đã nạp "${req.file.originalname}" (sheet "${sheetName}"): ${rows.length} mã công trình.`)
    );
  } catch (e) {
    return res.redirect("/transactions?error=" + encodeURIComponent(e.message));
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

// ---------- MISA So Tien Gui import → backfill tenDoiUng ----------
//
// Luyen, 2026-08-08: "tk VP 8888 với 9777 không có bên đối tác nhưng tôi đã
// hạch toán trên misa rồi bạn dựa vào đó bạn gắn lên báo cáo thu chi gian
// thêm vô cho tôi bên đối tác nhá" -- upload file "Sổ tiền gửi ngân hàng"
// xuất từ MISA để điền tên đối ứng (benChoThue / tenDoiTuong) vào giao dịch
// ngân hàng còn đang trống trường này.
//
// Hỗ trợ 2 định dạng:
//   VP9997 (KH Cũ): header row 3, cols [NgayCT, SoUNC, MaDT_chung,
//     TenDT_chung, DienGiai, TKDuU, Thu, Chi, Ton, MaDT, MaCT, TenCT]
//     → tenDoiUng = TenDT_chung (col 3)
//   VP58888 (KH Mới): header row 3, cols [NgayHT, NgayCT, SoUNC, DienGiai,
//     TKDuU, Thu, Chi, Ton, MaCT]  -- không có cột TenDT
//     → chi TK 331 với MaCT: tenDoiUng = benChoThue từ hop_dong_thue (nếu
//       có), ngược lại "Mã công trình: {MaCT}"
//     → thu TK 131: trích đầu DienGiai làm tenDoiUng
function excelSerialToISO(serial) {
  if (typeof serial !== "number") return null;
  const d = new Date(new Date(1899, 11, 30).getTime() + serial * 86400000);
  const yy = d.getFullYear();
  if (yy < 2000 || yy > 2100) return null;
  return d.toISOString().slice(0, 10);
}

function parseMisaSoTienGui997(sheetRows) {
  // cols: [NgayCT(0), SoUNC(1), MaDT_chung(2), TenDT_chung(3), DienGiai(4),
  //        TKDuU(5), Thu(6), Chi(7), Ton(8), MaDT(9), MaCT(10), TenCT(11)]
  const result = [];
  for (let i = 5; i < sheetRows.length; i++) {
    const r = sheetRows[i];
    if (!r || typeof r[0] !== "number") continue;
    const date = excelSerialToISO(r[0]);
    if (!date) continue;
    const tenDT = r[3] ? String(r[3]).trim() : "";
    const maDT = r[9] ? String(r[9]).trim() : r[2] ? String(r[2]).trim() : "";
    const maCT = r[10] ? String(r[10]).trim() : "";
    const tenCT = r[11] ? String(r[11]).trim() : "";
    const thu = typeof r[6] === "number" && r[6] > 0 ? r[6] : 0;
    const chi = typeof r[7] === "number" && r[7] > 0 ? r[7] : 0;
    if (!tenDT) continue;
    if (thu) result.push({ date, amount: thu, type: "thu", tenDoiUng: tenDT, maDT, maCT, tenCT });
    if (chi) result.push({ date, amount: chi, type: "chi", tenDoiUng: tenDT, maDT, maCT, tenCT });
  }
  return result;
}

function parseMisaSoTienGui888(sheetRows, ctToNCC) {
  // cols: [NgayHT(0), NgayCT(1), SoUNC(2), DienGiai(3), TKDuU(4),
  //        Thu(5), Chi(6), Ton(7), MaCT(8)]
  const result = [];
  for (let i = 5; i < sheetRows.length; i++) {
    const r = sheetRows[i];
    if (!r || typeof r[0] !== "number") continue;
    const dateSerial = typeof r[1] === "number" ? r[1] : r[0];
    const date = excelSerialToISO(dateSerial);
    if (!date) continue;
    const dienGiai = r[3] ? String(r[3]).trim() : "";
    const tkDuU = r[4] ? String(r[4]).trim() : "";
    const maCT = r[8] ? String(r[8]).trim() : "";
    const thu = typeof r[5] === "number" && r[5] > 0 ? r[5] : 0;
    const chi = typeof r[6] === "number" && r[6] > 0 ? r[6] : 0;
    if (!thu && !chi) continue;

    let tenDoiUng = "";
    if (tkDuU === "131" || tkDuU === "1311") {
      // Thu từ đối tác: tên công ty nằm đầu diễn giải
      tenDoiUng = dienGiai
        .replace(/\s+(SAL|HD|HĐ|THÁNG|T\d|t\d|\d{8,}|EV\d+).*/i, "")
        .trim()
        .substring(0, 80);
    } else if (tkDuU === "331" || tkDuU === "3311") {
      if (maCT) {
        tenDoiUng = ctToNCC[maCT] ? ctToNCC[maCT] : `Mã công trình: ${maCT}`;
      }
    }
    if (!tenDoiUng) continue;
    if (thu) result.push({ date, amount: thu, type: "thu", tenDoiUng, maCT });
    if (chi) result.push({ date, amount: chi, type: "chi", tenDoiUng, maCT });
  }
  return result;
}

router.post(
  "/transactions/import-misa-so-tien-gui",
  requireAdmin,
  upload.single("file"),
  (req, res) => {
    if (!req.file)
      return res.status(400).json({ error: "Thiếu file" });
    const bankId = parseInt(req.body.bank_id || "0", 10);
    if (!bankId)
      return res.status(400).json({ error: "Thiếu bank_id" });

    const store = load();
    const bank = store.banks.find((b) => b.id === bankId);
    if (!bank)
      return res.status(400).json({ error: `Không tìm thấy ngân hàng id=${bankId}` });

    // Build maCongTrinh → benChoThue lookup from hop_dong_thue
    const hdThue = store.phap_danh_hop_dong_thue || {};
    const ctToNCC = {};
    Object.values(hdThue).forEach((hd) => {
      const mct = (hd.maCongTrinh || "").trim();
      const ben = (hd.benChoThue || "").trim();
      if (mct && ben && !ctToNCC[mct]) ctToNCC[mct] = ben;
    });

    // Parse XLSX
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

    // Auto-detect format: 997 has 12 cols at header row, 888 has 9 cols
    const headerRow = rows[3] || [];
    const ncols = headerRow.filter((c) => c !== null).length;
    const is997Format = ncols >= 10;

    let misaRows;
    if (is997Format) {
      misaRows = parseMisaSoTienGui997(rows);
    } else {
      misaRows = parseMisaSoTienGui888(rows, ctToNCC);
    }

    // Build lookup: date|amount|type → tenDoiUng (first match wins)
    const lookup = new Map();
    for (const r of misaRows) {
      const key = `${r.date}|${r.amount}|${r.type}`;
      if (!lookup.has(key)) lookup.set(key, r.tenDoiUng);
    }

    // Backfill transactions
    const bankTxs = store.transactions.filter((t) => t.bank_id === bankId);
    let backfilled = 0;
    let upgraded = 0;
    for (const tx of bankTxs) {
      const key = `${tx.date}|${tx.amount}|${tx.type}`;
      const val = lookup.get(key);
      if (!tx.tenDoiUng && val) {
        tx.tenDoiUng = val;
        backfilled++;
      } else if (tx.tenDoiUng && tx.tenDoiUng.startsWith("Mã công trình: ")) {
        // Try to upgrade Mã công trình → real NCC name
        const mct = tx.tenDoiUng.replace("Mã công trình: ", "").trim();
        if (ctToNCC[mct]) {
          tx.tenDoiUng = ctToNCC[mct];
          upgraded++;
        }
      }
    }

    // Update ncc_doi_tuong_master from 997 format rows
    let masterAdded = 0;
    if (is997Format) {
      if (!store.ncc_doi_tuong_master) store.ncc_doi_tuong_master = {};
      const srcLabel = `So_tien_gui_ngan_hang ${bank.name} (upload ${new Date().toISOString().slice(0, 10)})`;
      for (const r of misaRows) {
        if (!r.maDT || !r.tenDoiUng) continue;
        if (!store.ncc_doi_tuong_master[r.maDT]) {
          store.ncc_doi_tuong_master[r.maDT] = {
            tenDoiTuong: r.tenDoiUng,
            maCongTrinh: r.maCT || "",
            tenCongTrinh: r.tenCT || "",
            nguon: srcLabel,
            updated_at: new Date().toISOString(),
          };
          masterAdded++;
        }
      }
    }

    save(store);

    res.json({
      ok: true,
      bank: bank.name,
      format: is997Format ? "VP9997 (12 cột)" : "VP58888 (9 cột)",
      misaRows: misaRows.length,
      backfilled,
      upgraded,
      masterAdded,
      message: `Đã điền bên đối tác: +${backfilled} mới, +${upgraded} nâng cấp từ "Mã công trình". NCC master: +${masterAdded} mục mới.`,
    });
  }
);

// ---------- Backfill tenDoiUng từ chi_phi + hoa_don_dau_vao ----------
//
// Luyen, 2026-08-12: "soa kê nó hk có tên đối ứng vp bank á bạn check từ
// diễn giải với trên chi phí với lại hóa đơn đầu vào xem số tiền đi của ncc
// nào dựa vào số tiền và diễn giải á để lấy ra tên ncc điền vào đây" --
// VPBANK9997 (và các NH khác không in tên đối ứng) hay có UNC đã hạch toán
// trên chi phí MN, nên khớp qua chi_phi (exact bankTxId rồi fuzzy amount ±1k
// + ngay ±7) để điền tenDoiUng = NCC.  Nếu không khớp chi phí thì thử rút
// số HĐ từ diễn giải (pattern "theo hd N" / "hd N") rồi tra hoa_don_dau_vao.
router.post(
  "/transactions/backfill-from-chiphi",
  requireDataEntry,
  (req, res) => {
    const store = load();
    const bankId = Number(req.body.bank_id);
    const bank = (store.banks || []).find((b) => b.id === bankId);
    if (!bank) return res.status(400).json({ error: `Không tìm thấy ngân hàng id=${bankId}` });

    // Xac dinh company tu bank
    const company = bank.company || "kh_cu";

    // Chi phi records cua company nay
    const chiPhiCompany = (store.chi_phi || []).filter(
      (r) => (r.congTy || "kh_cu") === company
    );
    const chiPhiByBankTxId = {};
    chiPhiCompany.forEach((r) => {
      if (r.bankTxId) chiPhiByBankTxId[String(r.bankTxId)] = r;
    });

    // Hoa don dau vao cua company
    const hdList = (store.hoa_don_dau_vao || []).filter(
      (h) => (h.congTy || "kh_cu") === company && h.soHoaDon && h.tenNCC
    );
    // Index soHoaDon -> tenNCC (normalize: lowercase, bo dau cach)
    const hdBySoHD = {};
    hdList.forEach((h) => {
      const k = String(h.soHoaDon).trim().toLowerCase();
      if (k && !hdBySoHD[k]) hdBySoHD[k] = h.tenNCC;
    });

    // Rut so hoa don tu dien giai: "theo hd 513", "hd513", "hd 97634"
    function extractHdFromDesc(desc) {
      if (!desc) return null;
      const m = (desc + "").match(/\bhd\s*(\d+)/i);
      return m ? m[1].trim().toLowerCase() : null;
    }

    // Normalize tieng Viet (bo dau) + lowercase + bo ky tu dac biet
    function normVn(s) {
      return (s || "")
        .toLowerCase()
        .normalize("NFD").replace(/[̀-ͯ]/g, "")
        .replace(/đ/g, "d").replace(/Đ/g, "d")
        .replace(/[^a-z0-9\s]/g, " ")
        .replace(/\s+/g, " ").trim();
    }

    // Tu dung: loai khoi khi tim keyword NCC trong dien giai
    const STOP = new Set([
      "cong","ty","tnhh","cp","cty","va","tt","theo","tien","phi","hang",
      "nuoc","ghe","may","thang","gio","ngoai","dien","thue","mua","ban",
      "cung","cap","dich","vu","hop","tac","kinh","doanh","mat","bang","gia",
      "trong","ngoai","lich","thuc","pham","chi","thanh","toan","unc","kh",
      "nam","bac","mien","thanh","pho","san","bay","cua","hang","nha","moi",
      "gan","sau","truoc","thuc","hien","theo","phi","hang","in","an","bao",
      "bao","ve","van","chuyen","phi","bao","phi","dich","thu",
    ]);

    function descWords(s) {
      return normVn(s).split(/\s+/).filter((w) => w.length >= 3 && !STOP.has(w));
    }

    // Build keyword index: normNcc -> origNcc
    // Lay NCC tu chi_phi + hoa_don_dau_vao (loai trung)
    const allNccSet = new Map(); // normNcc -> origNcc (first seen)
    chiPhiCompany.forEach((r) => {
      if (!r.ncc) return;
      const k = normVn(r.ncc);
      if (k && !allNccSet.has(k)) allNccSet.set(k, r.ncc);
    });
    hdList.forEach((h) => {
      if (!h.tenNCC) return;
      const k = normVn(h.tenNCC);
      if (k && !allNccSet.has(k)) allNccSet.set(k, h.tenNCC);
    });

    // Pre-compute nccWords cho tung NCC (chi giu nhung NCC co >= 1 keyword co nghia)
    const nccEntries = []; // [{words, origNcc}]
    for (const [normNcc, origNcc] of allNccSet) {
      const words = descWords(normNcc);
      if (words.length === 0) continue;
      nccEntries.push({ words, origNcc });
    }

    function findNccByKeyword(desc) {
      const dw = descWords(desc);
      if (dw.length === 0) return "";
      let best = null;
      let bestScore = 0;
      for (const { words: nw, origNcc } of nccEntries) {
        // Dem bao nhieu tu cua NCC xuat hien trong dien giai
        const hit = nw.filter((w) => dw.includes(w)).length;
        if (hit === 0) continue;
        // Yeu cau it nhat >= 50% tu cua NCC phai khop, va toi thieu 1 tu
        const ratio = hit / nw.length;
        if (ratio >= 0.5 && hit > bestScore) {
          bestScore = hit;
          best = origNcc;
        }
      }
      return best || "";
    }

    const AMOUNT_TOL = 1000;
    const DATE_WIN_MS = 7 * 86400000;
    function dateMs(d) { return new Date((d || "").slice(0, 10) + "T00:00:00").getTime(); }

    const bankTxs = (store.transactions || []).filter(
      (t) => t.bank_id === bankId && !t.tenDoiUng
    );

    let filled = 0;
    const matchSources = { bankTxId: 0, fuzzy: 0, hd: 0, keyword: 0 };
    for (const tx of bankTxs) {
      let ncc = "";

      // 1. Exact match qua bankTxId
      const exact = chiPhiByBankTxId[String(tx.id)];
      if (exact && exact.ncc) {
        ncc = exact.ncc;
        matchSources.bankTxId++;
      }

      // 2. Fuzzy: amount ±1000 + ngay ±7
      if (!ncc) {
        const tMs = dateMs(tx.date);
        const fuzzy = chiPhiCompany.find((r) => {
          if (!r.ngay || !r.soTien) return false;
          if (Math.abs((Number(r.soTien) || 0) - tx.amount) > AMOUNT_TOL) return false;
          if (Math.abs(dateMs(r.ngay) - tMs) > DATE_WIN_MS) return false;
          return true;
        });
        if (fuzzy && fuzzy.ncc) { ncc = fuzzy.ncc; matchSources.fuzzy++; }
      }

      // 3. Tra so hoa don tu dien giai -> hoa_don_dau_vao
      if (!ncc) {
        const hdNo = extractHdFromDesc(tx.description);
        if (hdNo && hdBySoHD[hdNo]) { ncc = hdBySoHD[hdNo]; matchSources.hd++; }
      }

      // 4. Keyword: tim ten NCC trong dien giai
      if (!ncc) {
        const kw = findNccByKeyword(tx.description);
        if (kw) { ncc = kw; matchSources.keyword++; }
      }

      if (ncc) {
        tx.tenDoiUng = ncc;
        filled++;
      }
    }

    if (filled > 0) save(store);
    const srcDetail = [
      matchSources.bankTxId && `${matchSources.bankTxId} qua mã GD`,
      matchSources.fuzzy    && `${matchSources.fuzzy} qua tiền+ngày`,
      matchSources.hd       && `${matchSources.hd} qua số HĐ`,
      matchSources.keyword  && `${matchSources.keyword} qua tên NCC trong diễn giải`,
    ].filter(Boolean).join(", ");
    res.json({
      ok: true,
      filled,
      total: bankTxs.length,
      message: `Đã điền Tên đối ứng cho ${filled} / ${bankTxs.length} giao dịch trống.`
        + (srcDetail ? ` (${srcDetail})` : ""),
    });
  }
);

// ---------- Import UNC file → backfill tenDoiUng cho giao dich chi ----------
// Luyen, 2026-08-18: "sao nó hk có tên người thụ hưởng á trên gg sheet Đi ủy
// nhiệm chi có mà bạn làm cho tôi luôn đi ns" -- file UNC Excel (vd "Tạo lệnh
// UNC KVC MB.xlsx") co cot "Nội dung trên UNC" va "Tên đơn vị thụ hưởng".
// Doc file nay, ghep voi giao dich Chi co tenDoiUng trong (dua tren noi dung
// UNC la substring cua dien giai ngan hang + so tien khop), dien tenDoiUng.
// Xu ly TAT CA sheet trong file (moi sheet la 1 tai khoan/thang UNC khac nhau).
function normUncText(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, "d")
    .replace(/[^a-z0-9 ]/gi, " ")
    .replace(/\s+/g, " ")
    .toLowerCase()
    .trim();
}

// Shared logic: parse UNC buffer → match → fill tenDoiUng
async function processUncBuffer(buffer, store) {
  const wb = XLSX.read(buffer, { type: "buffer" });
    const uncRows = []; // { noiDung, tenThuHuong, amount }

    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

      // Find header row
      let hdr = null;
      let hdrIdx = -1;
      for (let r = 0; r < Math.min(grid.length, 10); r++) {
        const row = (grid[r] || []).map((c) => normUncText(String(c || "")));
        const noiDungCol = row.findIndex((h) => h.includes("noi dung") && (h.includes("unc") || h.includes("tren")));
        const tenCol = row.findIndex((h) => h.includes("ten") && (h.includes("thu huong") || h.includes("don vi")));
        const amtCol = row.findIndex((h) => h.includes("so tien") || h === "tien");
        if (noiDungCol >= 0 && tenCol >= 0) {
          hdr = { noiDungCol, tenCol, amtCol };
          hdrIdx = r;
          break;
        }
      }
      if (!hdr) continue;

      for (let r = hdrIdx + 1; r < grid.length; r++) {
        const row = grid[r] || [];
        const noiDung = String(row[hdr.noiDungCol] || "").trim();
        const tenThuHuong = String(row[hdr.tenCol] || "").trim();
        if (!noiDung || !tenThuHuong) continue;
        const amount = hdr.amtCol >= 0 ? Math.abs(Number(String(row[hdr.amtCol]).replace(/[^0-9.]/g, "")) || 0) : 0;
        uncRows.push({ noiDung, tenThuHuong, amount, noiDungNorm: normUncText(noiDung) });
      }
    }

    if (uncRows.length === 0) {
      return { ok: false, message: "Không tìm thấy dữ liệu UNC trong file (cần cột 'Nội dung trên UNC' và 'Tên đơn vị thụ hưởng')." };
    }

    // Match against store.transactions
    let filled = 0;
    const MIN_MATCH_LEN = 10;

    for (const tx of store.transactions) {
      if (tx.type !== "chi") continue;
      if (tx.tenDoiUng && tx.tenDoiUng.trim()) continue;

      const descNorm = normUncText(tx.description);

      for (const unc of uncRows) {
        if (unc.noiDungNorm.length < MIN_MATCH_LEN) continue;
        if (!descNorm.includes(unc.noiDungNorm)) continue;
        if (unc.amount > 0 && tx.amount !== unc.amount) continue;
        tx.tenDoiUng = unc.tenThuHuong;
        filled++;
        break;
      }
    }

    if (filled > 0) save(store);
    return {
      ok: true,
      uncRows: uncRows.length,
      filled,
      sheets: wb.SheetNames.length,
      message: `Đã điền tên thụ hưởng cho ${filled} giao dịch (đọc ${uncRows.length} dòng UNC từ ${wb.SheetNames.length} sheet).`,
    };
}

router.post(
  "/transactions/import-unc",
  requireAdmin,
  upload.single("unc_file"),
  async (req, res) => {
    const store = load();
    let buffer;

    // Support both: ggsheet_url (Google Sheet link) or file upload
    const ggUrl = (req.body && req.body.ggsheet_url) ? req.body.ggsheet_url.trim() : "";
    if (ggUrl) {
      // Extract sheet ID from Google Sheet URL and build export URL
      const m = ggUrl.match(/\/spreadsheets\/d\/([\w-]+)/);
      if (!m) return res.json({ ok: false, message: "Link Google Sheet không hợp lệ." });
      const exportUrl = `https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`;
      try {
        const resp = await fetch(exportUrl, { redirect: "follow" });
        if (!resp.ok) return res.json({ ok: false, message: `Không tải được Google Sheet (${resp.status}). Kiểm tra quyền chia sẻ (phải là "Bất kỳ ai có link").` });
        buffer = Buffer.from(await resp.arrayBuffer());
      } catch (e) {
        return res.json({ ok: false, message: "Lỗi kết nối Google Sheet: " + e.message });
      }
    } else if (req.file) {
      buffer = req.file.buffer;
    } else {
      return res.json({ ok: false, message: "Cần nhập link Google Sheet hoặc chọn file." });
    }

    try {
      const result = await processUncBuffer(buffer, store);
      res.json(result);
    } catch (e) {
      res.json({ ok: false, message: "Lỗi đọc file: " + e.message });
    }
  }
);

// One-time migration: rename KVC ROYAL -> PINBALL DA NANG in vnpay_khmoi
// uploads for dates >= 2026-07-31 (Luyen, 2026-08-08).
// Idempotent -- safe to call multiple times.
router.post("/transactions/migrate-kvc-royal-to-pinball", requireAdmin, (req, res) => {
  const store = load();
  const OLD = "KVC ROYAL";
  const NEW = "PINBALL DA NANG";
  const CUTOFF = "2026-07-31";

  function migrateUploads(uploads) {
    let count = 0;
    for (const u of uploads || []) {
      const gb = u.grossByCode || {};
      const toRename = Object.keys(gb).filter(
        (k) => k.includes("|" + OLD) && k.slice(0, 10) >= CUTOFF
      );
      for (const oldKey of toRename) {
        const newKey = oldKey.replace("|" + OLD, "|" + NEW);
        gb[newKey] = gb[oldKey];
        delete gb[oldKey];
        count++;
      }
      if (u.codes) {
        const hasOld = Object.keys(gb).some((k) => k.includes("|" + OLD));
        if (!hasOld && u.codes.includes(OLD)) {
          u.codes = u.codes.map((c) => (c === OLD ? NEW : c));
        } else if (!u.codes.includes(NEW) && Object.keys(gb).some((k) => k.includes("|" + NEW))) {
          u.codes.push(NEW);
        }
      }
    }
    return count;
  }

  let total = 0;
  total += migrateUploads(store.vnpay_khmoi_uploads);
  total += migrateUploads(store.vnpay_khmoi_payoo_uploads);
  store.gian_mapping["PINBALL DA NANG"] = "131";
  save(store);

  res.json({ ok: true, renamedKeys: total, message: `Migration done: ${total} grossByCode keys renamed.` });
});

// One-time migration: revert Payoo KVC ROYAL -> PINBALL DA NANG back to KVC ROYAL
// for dates >= 2026-07-31 (Luyen, 2026-08-08: Payoo invoices still say "KVC ROYAL",
// so Payoo gross must stay "KVC ROYAL". Only VNPay offline was renamed to PINBALL).
// Idempotent -- safe to call multiple times.
router.post("/transactions/migrate-payoo-pinball-to-kvcroyal", requireAdmin, (req, res) => {
  const store = load();
  const OLD = "PINBALL DA NANG";
  const NEW = "KVC ROYAL";
  const CUTOFF = "2026-07-31";

  let count = 0;
  for (const u of store.vnpay_khmoi_payoo_uploads || []) {
    const gb = u.grossByCode || {};
    const toRename = Object.keys(gb).filter(
      (k) => k.includes("|" + OLD) && k.slice(0, 10) >= CUTOFF
    );
    for (const oldKey of toRename) {
      const newKey = oldKey.replace("|" + OLD, "|" + NEW);
      gb[newKey] = gb[oldKey];
      delete gb[oldKey];
      count++;
    }
    if (u.codes) {
      u.codes = u.codes.map((c) => (c === OLD ? NEW : c));
    }
  }
  save(store);

  res.json({ ok: true, renamedKeys: count, message: `Payoo revert done: ${count} grossByCode keys renamed PINBALL -> KVC ROYAL.` });
});

// Temp debug: find invoices by soHd across zvp_invoices pools + vnpayKhMoi
router.get("/transactions/debug-invoice-madiem", requireAdmin, (req, res) => {
  const store = load();
  const soHDs = (req.query.soHD || "").split(",").map((s) => s.trim()).filter(Boolean);
  const results = [];
  const pools = {
    "viet_qr_invoices.vnpayKhMoi": store.viet_qr_invoices?.vnpayKhMoi || [],
    "zvp_invoices.vnpay": store.zvp_invoices?.vnpay || [],
    "zvp_invoices.payoo": store.zvp_invoices?.payoo || [],
    "zvp_invoices.zalo": store.zvp_invoices?.zalo || [],
  };
  for (const [poolName, invoices] of Object.entries(pools)) {
    const items = soHDs.length > 0
      ? invoices.filter((i) => soHDs.includes(String(i.soHd)) || soHDs.includes(String(i.soHD)))
      : invoices.slice(-5);
    for (const i of items) {
      results.push({ pool: poolName, soHd: i.soHd ?? i.soHD, ngayHd: i.ngayHd ?? i.ngay, maDiem: i.maDiem, tenDiem: i.tenDiem, soTien: i.soTien, raw: i.raw });
    }
  }
  res.json(results);
});

module.exports = router;
