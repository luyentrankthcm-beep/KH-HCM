const express = require("express");
const { load } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { computeBalance } = require("./banks");
const { buildAllFlatLines, buildGianPivot, buildChannelSummary, getCacheDebugInfo } = require("../utils/overviewAggregate");
const { getDataVersion } = require("../store");
const { COMPANIES, getCompany } = require("../utils/companies");
const { BANK_COMPANY } = require("../utils/bankCompany");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan, 2026-07-30: "tách theo công ty cho tôi" -- cong ty cua 1 TK ngan
// hang thuc: uu tien field b.company (cac TK moi tao co san), fallback ve
// BANK_COMPANY (utils/bankCompany.js, ten TK -> cong ty), roi moi ve mac
// dinh "kh_cu" cho cac TK rat cu/khong ro (vd BIDV7555, BIDV7703, BIDV8651)
// -- cung logic BANK_COMPANY[..] || "kh_cu" da dung o routes/congno.js.
function companyOfBank(b) {
  return b.company || BANK_COMPANY[b.name] || "kh_cu";
}

router.get("/", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);

  const banks = [...store.banks]
    .filter((b) => companyOfBank(b) === activeCompany)
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((b) => ({ ...b, balance: computeBalance(store, b.id) }));
  const totalBalance = banks.reduce((s, b) => s + b.balance, 0);
  const bankIdSet = new Set(banks.map((b) => b.id));

  const today = new Date();
  const monthStart = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-01`;
  const currentMonth = monthStart.slice(0, 7);

  const monthSums = { thu: 0, chi: 0 };
  for (const t of store.transactions) {
    if (!bankIdSet.has(t.bank_id)) continue;
    if (t.date >= monthStart) {
      if (t.type === "thu") monthSums.thu += t.amount;
      else monthSums.chi += t.amount;
    }
  }

  const recent = [...store.transactions]
    .filter((t) => bankIdSet.has(t.bank_id))
    .sort((a, b) => (b.date > a.date ? 1 : b.date < a.date ? -1 : b.id - a.id))
    .slice(0, 15)
    .map((t) => {
      const bank = store.banks.find((b) => b.id === t.bank_id);
      return { ...t, bank_label: bank ? bank.name : "(da xoa)" };
    });

  // Dong tien theo gian x kenh doi soat (Momo/Zalo/VNPay/Payoo/3 kenh Viet
  // QR), tong hop tu chinh ket qua doi soat cua tung trang (khong tinh lai
  // rieng) -- xem utils/overviewAggregate.js. Loc rieng theo cong ty dang
  // xem (BANK_COMPANY[bankLabel], giong routes/congno.js) truoc khi pivot.
  // Mac dinh xem thang gan nhat, co the doi qua dropdown rieng voi thang
  // cua bang "So du theo ngan hang".
  let gianMonths = [];
  let selectedGianMonth = "";
  let gianPivot = { channels: [], gianRows: [], totalsByChannel: [], grandTotal: 0 };
  let channelSummary = [];
  let vietQrChannelSummary = [];
  // Chi Nhan, 2026-09-12: debug tam thoi -- xem ghi chu getCacheDebugInfo
  // trong utils/overviewAggregate.js. Xoa cac dong "Debug-*" nay sau khi tim
  // ra ly do cache khong giam duoc thoi gian request lap lai.
  const debugBefore = getCacheDebugInfo();
  const versionBefore = getDataVersion();
  try {
    const flatAll = buildAllFlatLines(store);
    const flat = flatAll.filter((l) => (BANK_COMPANY[l.bankLabel] || "kh_cu") === activeCompany);
    gianMonths = Array.from(new Set(flat.map((l) => l.month))).sort().reverse();
    selectedGianMonth = req.query.gianMonth !== undefined ? req.query.gianMonth : gianMonths[0] || "";
    gianPivot = buildGianPivot(flat, selectedGianMonth);
    channelSummary = buildChannelSummary(flat, selectedGianMonth);
    vietQrChannelSummary = channelSummary.filter((s) => s.channelKey.startsWith("vietqr_"));
  } catch (e) {
    // Khong de 1 loi tinh tong hop lam sap ca trang Tong quan -- cac bang
    // so du/giao dich phia tren van hien binh thuong, chi thieu phan nay.
    console.error("Loi tong hop dong tien theo gian:", e);
  }
  try {
    const debugAfter = getCacheDebugInfo();
    res.set("X-Debug-Pid", String(process.pid));
    res.set("X-Debug-Version-Before", String(versionBefore));
    res.set("X-Debug-Version-After", String(getDataVersion()));
    res.set("X-Debug-Cache-Before", JSON.stringify(debugBefore));
    res.set("X-Debug-Cache-After", JSON.stringify(debugAfter));
  } catch (e) {}

  // Chi phi theo loai, thang nay, cho dung 1 cong ty dang xem -- tra loi
  // "chi phí nào nhiều / hàng hóa" (Chi Nhan, 2026-07-30). Top 8 danh muc
  // theo tien, con lai gop vao "Khac" de bieu do khong qua roi.
  let chiPhiTheoLoai = [];
  try {
    const byLoai = {};
    (store.chi_phi || []).forEach((r) => {
      const congTy = r.congTy || "kh_cu";
      if (congTy !== activeCompany) return;
      if (!r.ngay || r.ngay.slice(0, 7) !== currentMonth) return;
      const loai = r.loaiChiPhi && r.loaiChiPhi.trim() ? r.loaiChiPhi.trim() : "(chưa phân loại)";
      byLoai[loai] = (byLoai[loai] || 0) + (r.soTien || 0);
    });
    const sorted = Object.entries(byLoai)
      .map(([loai, soTien]) => ({ loai, soTien }))
      .sort((a, b) => b.soTien - a.soTien);
    chiPhiTheoLoai = sorted.slice(0, 8);
    if (sorted.length > 8) {
      const khac = sorted.slice(8).reduce((s, r) => s + r.soTien, 0);
      chiPhiTheoLoai.push({ loai: "Khác", soTien: khac });
    }
  } catch (e) {
    console.error("Loi tong hop chi phi theo loai:", e);
  }

  // Top gian theo doanh thu thang dang xem, lay lai tu gianPivot da tinh
  // tren (khong tinh lai) -- tra loi "theo dõi xem gian".
  const topGian = gianPivot.gianRows.slice(0, 8).map((r) => ({ gian: r.gian, gross: r.totalGross }));

  res.render("dashboard", {
    activeCompany,
    COMPANIES,
    banks,
    totalBalance,
    monthSums,
    recent,
    gianMonths,
    selectedGianMonth,
    gianPivot,
    channelSummary,
    vietQrChannelSummary,
    chiPhiTheoLoai,
    topGian,
    currentMonth,
    userName: req.session.userName,
    pwerror: req.query.pwerror || null,
    pwsuccess: req.query.pwsuccess || null,
  });
});

// JSON summary for the chart: daily thu/chi over the last N days. `company`
// (kh_cu/kh_moi) loc theo cac TK ngan hang thuoc cong ty do -- de bieu do
// "dong tien" tren Tong quan cung tach duoc theo cong ty dang xem.
router.get("/api/summary", (req, res) => {
  const days = Math.min(parseInt(req.query.days || "30", 10), 365);
  const bankId = req.query.bank_id;
  const company = req.query.company;

  const since = new Date();
  since.setDate(since.getDate() - days + 1);
  const sinceStr = since.toISOString().slice(0, 10);

  const store = load();
  let allowedBankIds = null;
  if (company && COMPANIES[company]) {
    allowedBankIds = new Set(store.banks.filter((b) => companyOfBank(b) === company).map((b) => b.id));
  }

  const byDate = {};
  for (const t of store.transactions) {
    if (t.date < sinceStr) continue;
    if (bankId && t.bank_id !== Number(bankId)) continue;
    if (allowedBankIds && !allowedBankIds.has(t.bank_id)) continue;
    if (!byDate[t.date]) byDate[t.date] = { date: t.date, thu: 0, chi: 0 };
    if (t.type === "thu") byDate[t.date].thu += t.amount;
    else byDate[t.date].chi += t.amount;
  }
  const rows = Object.values(byDate).sort((a, b) => (a.date > b.date ? 1 : -1));
  res.json(rows);
});

module.exports = router;
