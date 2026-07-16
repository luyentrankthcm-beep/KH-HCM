const express = require("express");
const { load } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { buildAllFlatLines, buildAgingRows } = require("../utils/overviewAggregate");

const router = express.Router();
router.use(requireLogin);

// "Cong no" o day nghia la: cac dong doi soat CHUA xong -- hoac chua co hoa
// don, hoac co hoa don nhung con lech -- tinh theo so ngay da troi qua ke tu
// ngay giao dich/settlement, giong cach nhin "cong no phai thu" thong thuong
// (tien da qua ngan hang nhung chua "khop so" voi hoa don tuong ung). Dung
// LAI chinh ket qua doi soat cua tung kenh (Momo/Zalo/VNPay/Payoo/3 kenh Viet
// QR) qua utils/overviewAggregate.js, khong tinh toan rieng.
router.get("/cong-no", (req, res) => {
  const store = load();
  let error = null;
  let aging = { rows: [], buckets: [], grandTotal: 0 };
  let channelOptions = [];

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
  });
});

module.exports = router;
