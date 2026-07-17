const express = require("express");
const { load, save, nextId } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { buildAllFlatLines, buildAgingRows } = require("../utils/overviewAggregate");

const router = express.Router();
router.use(requireLogin);

// Cong no "nhap tay" (Luyen yeu cau 2026-07-17): nhung dong cong no chi da tu
// doi soat xong TU TRUOC (co san so hoa don theo tung gian) -- khong tinh tu
// dong nhu bang aging o tren (bang do CHI danh cho dong CHUA xong), day chi
// la so ghi lai de tra cuu, khong cong don vao aging/cac tong so tu dong.
function ensureManualShape(store) {
  if (!store.congno_manual_entries) store.congno_manual_entries = [];
}

// "Cong no" o day nghia la: cac dong doi soat CHUA xong -- hoac chua co hoa
// don, hoac co hoa don nhung con lech -- tinh theo so ngay da troi qua ke tu
// ngay giao dich/settlement, giong cach nhin "cong no phai thu" thong thuong
// (tien da qua ngan hang nhung chua "khop so" voi hoa don tuong ung). Dung
// LAI chinh ket qua doi soat cua tung kenh (Momo/Zalo/VNPay/Payoo/3 kenh Viet
// QR) qua utils/overviewAggregate.js, khong tinh toan rieng.
router.get("/cong-no", (req, res) => {
  const store = load();
  ensureManualShape(store);
  let error = req.query.error || null;
  const success = req.query.success || null;
  let aging = { rows: [], buckets: [], grandTotal: 0 };
  let channelOptions = [];

  const manualEntries = [...store.congno_manual_entries].sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  const manualTotal = manualEntries.reduce((s, r) => s + r.amount, 0);

  try {
    const flat = buildAllFlatLines(store);
    channelOptions = Array.from(
      new Map(flat.map((l) => [l.channelKey, l.channelLabel])).entries()
    ).map(([key, label]) => ({ key, label }));

    const today = new Date().toISOString().slice(0, 10);
    const fullAging = buildAgingRows(flat, today);

    const selectedChannel = req.query.channel || "";
    const selectedBucket = req.query.bucket || "";
    let rows = fullAging.rows;
    if (selectedChannel) rows = rows.filter((r) => r.channelKey === selectedChannel);
    if (selectedBucket) rows = rows.filter((r) => r.bucket === selectedBucket);

    aging = {
      rows,
      buckets: fullAging.buckets,
      grandTotal: rows.reduce((s, r) => s + r.amount, 0),
      selectedChannel,
      selectedBucket,
    };
  } catch (e) {
    error = e.message;
    console.error("Loi tinh cong no:", e);
  }

  res.render("congno", {
    userName: req.session.userName,
    aging,
    channelOptions,
    error,
    success,
    manualEntries,
    manualTotal,
  });
});

// ---------- Cong no nhap tay (da doi soat xong tu truoc, co so hoa don theo
// gian) -- 4 cot co ban theo yeu cau Luyen: Gian, Ngay, So tien, So HD. ----------
router.post("/cong-no/thu-cong", (req, res) => {
  const store = load();
  ensureManualShape(store);
  try {
    const { gian, date, amount, invoiceNumbers } = req.body;
    if (!gian || !gian.trim()) throw new Error("Thiếu Gian.");
    if (!date) throw new Error("Thiếu Ngày.");
    const amt = Number(String(amount || "").replace(/[^\d.-]/g, ""));
    if (!amt || amt <= 0) throw new Error("Số tiền không hợp lệ.");
    store.congno_manual_entries.push({
      id: nextId(store, "congno_manual_seq") || Date.now(),
      gian: gian.trim(),
      date,
      amount: amt,
      invoiceNumbers: (invoiceNumbers || "").trim(),
      createdAt: new Date().toISOString(),
    });
    save(store);
    res.redirect("/cong-no?success=" + encodeURIComponent("Đã lưu công nợ (nhập tay)."));
  } catch (e) {
    res.redirect("/cong-no?error=" + encodeURIComponent(e.message));
  }
});

router.post("/cong-no/thu-cong/:id/delete", (req, res) => {
  const store = load();
  ensureManualShape(store);
  store.congno_manual_entries = store.congno_manual_entries.filter((r) => String(r.id) !== req.params.id);
  save(store);
  res.redirect("/cong-no?success=" + encodeURIComponent("Đã xóa dòng công nợ nhập tay."));
});

module.exports = router;
