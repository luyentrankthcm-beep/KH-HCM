// Shared aggregation for the "Tong quan" (overview) dashboard and the "Cong
// no" (unresolved-line aging) page. Both pages need to combine data across
// every reconciliation channel (Momo, Zalo/VNPay/Payoo, 3x Viet QR banks),
// which all already produce the SAME per-settlement-day "line" shape:
//   { code, maCongTrinh, tkCo, gross, net, invoiceTotal, diff,
//     invoiceNumbers, matched, manualOverride }
// (see utils/momoReconcile.js, utils/zvpReconcile.js, utils/vietqrReconcile.js
// -- all 3 build their export rows off exactly this shape). Rather than
// re-implement any matching logic here, this module only FLATTENS/AGGREGATES
// numbers the 3 route files already computed via their exported builder
// functions (router.buildMomoReconciliation, router.buildZvpReconciliation,
// router.buildChannelReconciliation) -- so this dashboard can never show a
// number that disagrees with the channel's own detailed page.

// Turn one channel's { settlementDate, lines[] }[] into a flat list, each
// line tagged with which channel/bank it came from + a "drill" URL back to
// that channel's own detail page (pre-filtered to the same month) so a user
// can click any dashboard number and land exactly on the rows behind it.
function flattenChannel(reconciledDays, channelKey, channelLabel, drillUrlForMonth) {
  const out = [];
  (reconciledDays || []).forEach((day) => {
    const month = day.settlementDate.slice(0, 7);
    (day.lines || []).forEach((l) => {
      if (l.tkCo === "SKIP") return; // KH moi chua len MISA -- khong tinh vao tong quan/cong no
      out.push({
        channelKey,
        channelLabel,
        settlementDate: day.settlementDate,
        month,
        gian: l.maCongTrinh,
        tkCo: l.tkCo,
        gross: l.gross || 0,
        invoiceTotal: l.invoiceTotal || 0,
        diff: l.diff || 0,
        invoiceNumbers: l.invoiceNumbers || [],
        matched: !!l.matched,
        manualOverride: !!l.manualOverride,
        drillUrl: drillUrlForMonth(month),
      });
    });
  });
  return out;
}

// isResolved: co HD va da khop (tu dong hoac da danh dau thu cong) -- day la
// dinh nghia "da doi soat xong" dung chung cho ca dashboard va cong no.
function isResolved(l) {
  return l.invoiceNumbers.length > 0 && (l.matched || l.manualOverride);
}

// ---------- Tong quan: gian x kenh pivot (cho 1 thang) ----------
function buildGianPivot(flatLines, month) {
  const rows = month ? flatLines.filter((l) => l.month === month) : flatLines;
  const channelSet = new Map(); // key -> label, giu thu tu xuat hien
  rows.forEach((l) => {
    if (!channelSet.has(l.channelKey)) channelSet.set(l.channelKey, l.channelLabel);
  });
  const channels = Array.from(channelSet.entries()).map(([key, label]) => ({ key, label }));

  const byGian = {};
  rows.forEach((l) => {
    if (!byGian[l.gian]) byGian[l.gian] = {};
    if (!byGian[l.gian][l.channelKey]) {
      byGian[l.gian][l.channelKey] = { gross: 0, invoiceTotal: 0, unresolvedCount: 0, drillUrl: l.drillUrl };
    }
    const cell = byGian[l.gian][l.channelKey];
    cell.gross += l.gross;
    cell.invoiceTotal += l.invoiceTotal;
    if (!isResolved(l)) cell.unresolvedCount++;
  });

  const gianRows = Object.keys(byGian)
    .sort()
    .map((gian) => {
      const cells = channels.map((ch) => byGian[gian][ch.key] || null);
      const totalGross = cells.reduce((s, c) => s + (c ? c.gross : 0), 0);
      const totalInvoiceTotal = cells.reduce((s, c) => s + (c ? c.invoiceTotal : 0), 0);
      const totalUnresolved = cells.reduce((s, c) => s + (c ? c.unresolvedCount : 0), 0);
      return { gian, cells, totalGross, totalInvoiceTotal, totalUnresolved };
    })
    .sort((a, b) => b.totalGross - a.totalGross);

  const totalsByChannel = channels.map((ch) =>
    gianRows.reduce((s, r) => {
      const idx = channels.findIndex((c) => c.key === ch.key);
      const cell = r.cells[idx];
      return s + (cell ? cell.gross : 0);
    }, 0)
  );
  const grandTotal = totalsByChannel.reduce((s, v) => s + v, 0);

  return { channels, gianRows, totalsByChannel, grandTotal };
}

// ---------- Tong quan: tom tat theo kenh (cho 1 thang) ----------
function buildChannelSummary(flatLines, month) {
  const rows = month ? flatLines.filter((l) => l.month === month) : flatLines;
  const byChannel = {};
  rows.forEach((l) => {
    if (!byChannel[l.channelKey]) {
      byChannel[l.channelKey] = {
        channelKey: l.channelKey,
        channelLabel: l.channelLabel,
        totalGross: 0,
        totalMatched: 0,
        totalChuaCoHD: 0,
        totalLech: 0,
        unresolvedCount: 0,
        drillUrl: l.drillUrl,
      };
    }
    const s = byChannel[l.channelKey];
    s.totalGross += l.gross;
    if (l.invoiceNumbers.length === 0) {
      s.totalChuaCoHD += l.gross;
      s.unresolvedCount++;
    } else if (isResolved(l)) {
      s.totalMatched += l.gross;
    } else {
      s.totalLech += Math.abs(l.diff);
      s.unresolvedCount++;
    }
  });
  return Object.values(byChannel).sort((a, b) => b.totalGross - a.totalGross);
}

// ---------- Cong no: cac dong CHUA doi soat xong, xep theo "tuoi" ----------
function bucketFor(days) {
  if (days <= 7) return "0-7";
  if (days <= 30) return "8-30";
  if (days <= 60) return "31-60";
  return "60+";
}

function buildAgingRows(flatLines, todayIso) {
  const today = new Date(todayIso + "T00:00:00");
  const rows = flatLines.filter((l) => !isResolved(l) && l.gross !== 0);
  const aged = rows.map((l) => {
    const settleDate = new Date(l.settlementDate + "T00:00:00");
    const daysAged = Math.max(0, Math.round((today - settleDate) / 86400000));
    const amount = l.invoiceNumbers.length === 0 ? l.gross : Math.abs(l.diff);
    return {
      channelKey: l.channelKey,
      channelLabel: l.channelLabel,
      settlementDate: l.settlementDate,
      gian: l.gian,
      status: l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : "Lệch",
      amount,
      daysAged,
      bucket: bucketFor(daysAged),
      invoiceNumbers: l.invoiceNumbers,
      drillUrl: l.drillUrl,
    };
  });
  aged.sort((a, b) => b.daysAged - a.daysAged);

  const bucketOrder = ["60+", "31-60", "8-30", "0-7"];
  const byBucket = {};
  bucketOrder.forEach((b) => (byBucket[b] = { bucket: b, count: 0, total: 0 }));
  aged.forEach((r) => {
    byBucket[r.bucket].count++;
    byBucket[r.bucket].total += r.amount;
  });

  return {
    rows: aged,
    buckets: bucketOrder.map((b) => byBucket[b]),
    grandTotal: aged.reduce((s, r) => s + r.amount, 0),
  };
}

// ---------- Gop ca 3 he thong doi soat lai thanh 1 danh sach dong phang ----------
// Dung chung boi routes/dashboard.js (Tong quan) va routes/congno.js (Cong
// no) -- ca 2 trang deu can CHINH XAC cung 1 tap du lieu, nen chi viet 1 lan
// o day. Require routes/*.js ngay trong ham (khong o dau file) de tranh loi
// require-cycle: cac file routes/doisoat*.js KHONG require lai module nay,
// nen ve ly thuyet khong co vong lap, nhung require ngay luc goi ham van an
// toan hon va de doc hon la require o dau file khi ban than cac module nay
// cung export qua nhieu thu (router + cac ham dinh kem).
function buildAllFlatLines(store) {
  const momoRouter = require("../routes/doisoat");
  const zvpRouter = require("../routes/doisoat-zvp");
  const vietqrRouter = require("../routes/doisoat-vietqr");

  const flat = [];

  const momoBuilt = momoRouter.buildMomoReconciliation(store);
  flat.push(
    ...flattenChannel(momoBuilt.reconciledAll, "momo", "Momo (BIDV123456)", (m) => `/doi-soat/momo?month=${m}`)
  );

  const zvpBuilt = zvpRouter.buildZvpReconciliation(store);
  if (zvpBuilt && zvpBuilt.reconciled) {
    flat.push(
      ...flattenChannel(zvpBuilt.reconciled.online, "zvp_online", "Zalo Mini App (Online)", (m) => `/doi-soat/zvp?month=${m}`)
    );
    flat.push(
      ...flattenChannel(zvpBuilt.reconciled.offline, "zvp_offline", "VNPay thu hộ (Offline)", (m) => `/doi-soat/zvp?month=${m}`)
    );
    flat.push(
      ...flattenChannel(zvpBuilt.reconciled.payoo, "zvp_payoo", "Payoo", (m) => `/doi-soat/zvp?month=${m}`)
    );
  }

  (vietqrRouter.VIETQR_CHANNEL_KEYS || []).forEach((chKey) => {
    const built = vietqrRouter.buildChannelReconciliation(store, chKey);
    if (built && built.reconciled) {
      const label = vietqrRouter.VIETQR_CHANNELS[chKey].label;
      flat.push(
        ...flattenChannel(built.reconciled, `vietqr_${chKey}`, `Viet QR ${label}`, (m) => `/doi-soat/vietqr?channel=${chKey}&month=${m}`)
      );
    }
  });

  return flat;
}

module.exports = {
  flattenChannel,
  isResolved,
  buildGianPivot,
  buildChannelSummary,
  buildAgingRows,
  buildAllFlatLines,
  bucketFor,
};
