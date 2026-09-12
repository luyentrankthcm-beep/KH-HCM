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
// `bankLabel`: ten TAI KHOAN NGAN HANG THUC (vd "ACB31268") -- khac voi
// channelLabel co the la ten SAN PHAM/kenh (vd "Zalo Mini App (Online)")
// trong khi 2-3 kenh cung don ve 1 tai khoan duy nhat. Them field nay de
// trang Cong No co the nhom "theo tai khoan/ngan hang" (Chi Nhan yeu cau
// 2026-07-30) ma khong lam mat granularity channelKey/channelLabel dang dung
// cho drill-down.
// `keepSkip`: mac dinh false (loai SKIP = "chua xuat MISA", dung cho hau het
// kenh). Rieng Momo KH Moi (TK BIDV7701 rieng) dung TK Co = SKIP theo 1 nghia
// KHAC HOAN TOAN -- co hieu "gian nay la cua KH Moi" (di san tu thoi KH Moi
// chua co TK Momo rieng, xem routes/doisoat.js/ensureGianHidden) -- neu loai
// SKIP nhu thuong thi TOAN BO du lieu Momo KH Moi bien mat (100% dong hien co
// deu la SKIP). Truyen true cho dung 1 truong hop nay.
function flattenChannel(reconciledDays, channelKey, channelLabel, drillUrlForMonth, bankLabel, keepSkip) {
  const out = [];
  (reconciledDays || []).forEach((day) => {
    const month = day.settlementDate.slice(0, 7);
    (day.lines || []).forEach((l) => {
      if (l.tkCo === "SKIP" && !keepSkip) return; // chua xuat MISA -- khong tinh vao tong quan/cong no
      out.push({
        channelKey,
        channelLabel,
        bankLabel: bankLabel || channelLabel,
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
      bankLabel: l.bankLabel,
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
// Chi Nhan, 2026-09-12: Luyen bao "web chậm" -- ham nay chay lai TOAN BO doi
// soat Momo (2 cong ty) + Zalo/VNPay/Payoo + VNPay KH Moi + 3 kenh Viet QR TU
// DAU moi lan goi, va duoc goi tu RAT NHIEU noi (trang Tong quan, Cong No,
// mot so bao cao, phap danh...) tren MOI request -- voi luong giao dich hien
// tai mat ~11 giay MOI LAN, khien "ca trang web" cam giac cham (dung ra chi
// 1 ham nay cham, nhung duoc goi khap noi nen cam giac lan ra toan bo web).
// Luyen xac nhan cach sua: "tính 1 lần thôi đừng tính lại, cái nào đụng dữ
// liệu cũ (đổi) thì tính lại, cái nào khớp rồi thì thôi" -- cache ket qua
// theo dataVersion (store.js, tang moi lan save() ghi thanh cong): con so
// nay GIONG lan truoc nghia la CHUA CO GHI DU LIEU MOI nao xen giua, dung
// luon ket qua cu KHONG tinh lai; khac thi tinh lai 1 lan roi cache lai.
// An toan cho MOI noi goi ham nay (khong phai sua tung route rieng le).
let _flatAllCache = null;
let _flatAllCacheVersion = null;
function buildAllFlatLines(store) {
  const { getDataVersion } = require("../store");
  const currentVersion = getDataVersion();
  if (_flatAllCache !== null && _flatAllCacheVersion === currentVersion) {
    return _flatAllCache;
  }

  const momoRouter = require("../routes/doisoat");
  const zvpRouter = require("../routes/doisoat-zvp");
  const vietqrRouter = require("../routes/doisoat-vietqr");
  const vnpayKhMoiRouter = require("../routes/doisoat-vnpay-khmoi");

  const flat = [];

  const momoBuiltKhCu = momoRouter.buildMomoReconciliation(store, "kh_cu");
  flat.push(
    ...flattenChannel(momoBuiltKhCu.reconciledAll, "momo", "Momo (BIDV123456)", (m) => `/doi-soat/momo?month=${m}`, "BIDV123456")
  );
  // Chi Nhan, 2026-07-30: "chi tiết theo tài khoản cho tôi nhá các ngân hàng
  // á" -- truoc gio trang Cong No CHUA HE co Momo KH Moi (BIDV7701) va
  // VNPay/Payoo KH Moi (VTB982, ben duoi), lam thieu cong no thuc su cua cac
  // TK nay. Them vao day, dung LAI cach xu ly SKIP dac biet cho Momo KH Moi
  // da xac nhan o routes/baocao.js (keepSkip = true).
  const momoBuiltKhMoi = momoRouter.buildMomoReconciliation(store, "kh_moi");
  flat.push(
    ...flattenChannel(momoBuiltKhMoi.reconciledAll, "momo_khmoi", "Momo (BIDV7701)", (m) => `/doi-soat/momo?month=${m}`, "BIDV7701", true)
  );

  const zvpBuilt = zvpRouter.buildZvpReconciliation(store);
  if (zvpBuilt && zvpBuilt.reconciled) {
    flat.push(
      ...flattenChannel(zvpBuilt.reconciled.online, "zvp_online", "Zalo Mini App (Online)", (m) => `/doi-soat/zvp?month=${m}`, "ACB31268")
    );
    flat.push(
      ...flattenChannel(zvpBuilt.reconciled.offline, "zvp_offline", "VNPay thu hộ (Offline)", (m) => `/doi-soat/zvp?month=${m}`, "ACB31268")
    );
    flat.push(
      ...flattenChannel(zvpBuilt.reconciled.payoo, "zvp_payoo", "Payoo", (m) => `/doi-soat/zvp?month=${m}`, "ACB31268")
    );
  }

  const vnpayKhMoiBuilt = vnpayKhMoiRouter.buildReconciliation(store);
  if (vnpayKhMoiBuilt && vnpayKhMoiBuilt.reconciled) {
    flat.push(
      ...flattenChannel(vnpayKhMoiBuilt.reconciled.offline, "vnpay_khmoi_offline", "VNPay (KH Mới)", (m) => `/doi-soat/vnpay-khmoi?month=${m}`, "VTB982")
    );
    flat.push(
      ...flattenChannel(vnpayKhMoiBuilt.reconciled.payoo, "vnpay_khmoi_payoo", "Payoo (KH Mới)", (m) => `/doi-soat/vnpay-khmoi?month=${m}`, "VTB982")
    );
  }

  (vietqrRouter.VIETQR_CHANNEL_KEYS || []).forEach((chKey) => {
    const built = vietqrRouter.buildChannelReconciliation(store, chKey);
    if (built && built.reconciled) {
      const label = vietqrRouter.VIETQR_CHANNELS[chKey].label;
      flat.push(
        ...flattenChannel(built.reconciled, `vietqr_${chKey}`, `Viet QR ${label}`, (m) => `/doi-soat/vietqr?channel=${chKey}&month=${m}`, vietqrRouter.VIETQR_CHANNELS[chKey].bankName)
      );
    }
  });

  _flatAllCache = flat;
  // Chi Nhan, 2026-09-12: dung lai getDataVersion() SAU KHI tinh xong (khong
  // dung "currentVersion" da bat luc BAT DAU o tren) -- cac ham buildXxxReconciliation
  // o tren (Momo/ZVP/VNPay KH Moi/VietQR) tu dong chay vai buoc "don dep du
  // lieu" (vd migrateVnpayKhMoiInvoices) va co the tu goi save() ngay TRONG
  // luc tinh, lam dataVersion tang len giua chung. Neu van dung "currentVersion"
  // cu de luu cache thi lan goi KE TIEP se luon thay version MOI hon version
  // da luu -> luon coi la "cache mat hieu luc" -> tinh lai TU DAU MOI LAN, dung
  // y het bug "cache khong bao gio hit" ma Luyen gap phai. Doc lai gia tri MOI
  // NHAT o day de cache luon dung voi trang thai du lieu THUC TE sau cung.
  _flatAllCacheVersion = getDataVersion();
  return flat;
}

// Chi Nhan, 2026-09-12: ham debug tam thoi -- phuc vu dieu tra vi sao cache
// buildAllFlatLines duong nhu khong giam duoc thoi gian request lap lai. Cho
// phep route "/" gan cac gia tri nay vao response header de so sanh giua
// nhieu request lien tiep (dataVersion co doi khong, cache co duoc dung
// khong, co phai request roi vao process khac khong). Co the xoa sau khi tim
// ra nguyen nhan.
function getCacheDebugInfo() {
  return {
    hasCache: _flatAllCache !== null,
    cacheVersion: _flatAllCacheVersion,
    cacheLen: _flatAllCache ? _flatAllCache.length : null,
  };
}

module.exports = {
  flattenChannel,
  isResolved,
  buildGianPivot,
  buildChannelSummary,
  buildAgingRows,
  buildAllFlatLines,
  bucketFor,
  getCacheDebugInfo,
};
