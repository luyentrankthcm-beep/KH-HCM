const express = require("express");
const { load } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { computeBalance } = require("./banks");
const { buildAllFlatLines, buildGianPivot, buildChannelSummary } = require("../utils/overviewAggregate");

const router = express.Router();
router.use(requireLogin);

router.get("/", (req, res) => {
  const store = load();
  const banks = [...store.banks]
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((b) => ({ ...b, balance: computeBalance(store, b.id) }));
  const totalBalance = banks.reduce((s, b) => s + b.balance, 0);

  const today = new Date();
  const monthStart = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-01`;

  const monthSums = { thu: 0, chi: 0 };
  for (const t of store.transactions) {
    if (t.date >= monthStart) {
      if (t.type === "thu") monthSums.thu += t.amount;
      else monthSums.chi += t.amount;
    }
  }

  const recent = [...store.transactions]
    .sort((a, b) => (b.date > a.date ? 1 : b.date < a.date ? -1 : b.id - a.id))
    .slice(0, 15)
    .map((t) => {
      const bank = store.banks.find((b) => b.id === t.bank_id);
      return { ...t, bank_label: bank ? bank.name : "(da xoa)" };
    });

  // Dong tien theo gian x kenh doi soat (Momo/Zalo/VNPay/Payoo/3 kenh Viet
  // QR), tong hop tu chinh ket qua doi soat cua tung trang (khong tinh lai
  // rieng) -- xem utils/overviewAggregate.js. Mac dinh xem thang gan nhat,
  // co the doi qua dropdown rieng voi thang cua bang "So du theo ngan hang".
  let gianMonths = [];
  let selectedGianMonth = "";
  let gianPivot = { channels: [], gianRows: [], totalsByChannel: [], grandTotal: 0 };
  let channelSummary = [];
  try {
    const flat = buildAllFlatLines(store);
    gianMonths = Array.from(new Set(flat.map((l) => l.month))).sort().reverse();
    selectedGianMonth = req.query.gianMonth !== undefined ? req.query.gianMonth : gianMonths[0] || "";
    gianPivot = buildGianPivot(flat, selectedGianMonth);
    channelSummary = buildChannelSummary(flat, selectedGianMonth);
  } catch (e) {
    // Khong de 1 loi tinh tong hop lam sap ca trang Tong quan -- cac bang
    // so du/giao dich phia tren van hien binh thuong, chi thieu phan nay.
    console.error("Loi tong hop dong tien theo gian:", e);
  }

  res.render("dashboard", {
    banks,
    totalBalance,
    monthSums,
    recent,
    gianMonths,
    selectedGianMonth,
    gianPivot,
    channelSummary,
    userName: req.session.userName,
    pwerror: req.query.pwerror || null,
    pwsuccess: req.query.pwsuccess || null,
  });
});

// JSON summary for the chart: daily thu/chi over the last N days
router.get("/api/summary", (req, res) => {
  const days = Math.min(parseInt(req.query.days || "30", 10), 365);
  const bankId = req.query.bank_id;

  const since = new Date();
  since.setDate(since.getDate() - days + 1);
  const sinceStr = since.toISOString().slice(0, 10);

  const store = load();
  const byDate = {};
  for (const t of store.transactions) {
    if (t.date < sinceStr) continue;
    if (bankId && t.bank_id !== Number(bankId)) continue;
    if (!byDate[t.date]) byDate[t.date] = { date: t.date, thu: 0, chi: 0 };
    if (t.type === "thu") byDate[t.date].thu += t.amount;
    else byDate[t.date].chi += t.amount;
  }
  const rows = Object.values(byDate).sort((a, b) => (a.date > b.date ? 1 : -1));
  res.json(rows);
});

module.exports = router;
