const express = require("express");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { parseVnpayPortalWorkbook, resolveGianGross, migrateVnpayKhMoiInvoices, normText } = require("../utils/vietqrReconcile");
const { extractZvpSettlements, reconcileZvpChannel, parsePayooWorkbook, parsePayooRawReport, parseVnpayOfflineFeeReport } = require("../utils/zvpReconcile");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan, 2026-07-29: "tôi nhầm rồi cái vn pay này trả vè ngân hàng VTB982
// á đây á với payoo cx về đây á" -- doanh thu VNPay/Payoo cua "TUTU TRAIN"
// (AE HUE/TIMES/ROYAL) + SAVICO ve tai khoan Vietinbank VTB982 (KHONG PHAI
// MB02865168 nhu gia dinh ban dau khi xay trang nay). VTB982 tra tien theo
// DOT (nhieu ngay gop 1 lan, giong het VNPay Offline/Payoo cua KH Cu ben
// routes/doisoat-zvp.js), KHONG PHAI tung giao dich 1 nhu VietQR -- nen dung
// KIEN TRUC ZVP (extractZvpSettlements/reconcileZvpChannel) o day, khong phai
// kien truc VietQR (channel "vnpayKhMoi" cu trong routes/doisoat-vietqr.js da
// bi go bo).
const VTB_BANK_NAME = "VTB982";

// Chi Nhan cung cap bang anh xa "ten tren VNPay -> ma cong trinh" 2026-07-29:
// TUTU TRAIN AM HUE -> AE HUE KVCN, TUTU TRAIN VC TIME(S)/NHA TUYET VC ROYAL
// -> KVC TIMES/KVC ROYAL, FARM SAVICO -> SAVICO KVCN. Cac ma dich nay TRUNG
// KHOP voi VNPAY_KHMOI_INVOICE_MADIEM_MAP (utils/vietqrReconcile.js) dung de
// di hoa don ra store.viet_qr_invoices.vnpayKhMoi, nen 2 ben (doanh thu tu
// file nay + hoa don da di sang) se tu khop dung ma.
const TEN_DIEM_TO_MA_CONG_TRINH = {
  "tutu train am hue": "AE HUE KVCN",
  "tutu train vc time": "KVC TIMES",
  "tutu train vc times": "KVC TIMES",
  "tutu train vc royal": "KVC ROYAL",
  "nha tuyet vc royal": "KVC ROYAL",
  "farm savico": "SAVICO KVCN",
};

// Chi Nhan, 2026-07-29: "payoo có 1 gian bên kh mới thôi á KVC ROYAL" -- khac
// voi VNPay (4 ma cong trinh, phai tra qua Diem thu), Payoo cua KH Moi CHI CO
// DUY NHAT 1 gian ("GIAITRIKH_TUTUVR_VCROYALCITY" tren file, xac nhan qua 2
// file mau chi gui: "Payoo thu hộ KH 705.xlsx" va "PY-GiaoDichBanHangPayoo-
// ...xlsx", ca 2 deu 100% 1 "Cua hang" duy nhat) -- nen KHONG can bang tra
// "Danh muc ten diem" nhu ZVP KH Cu, gan THANG moi dong doc duoc ve "KVC
// ROYAL" bat ke gia tri "Cua hang"/"Gian hang" tren file la gi.
const PAYOO_FIXED_CODE = "KVC ROYAL";

function ensureVnpayKhMoiShape(store) {
  let changed = false;
  if (!store.vnpay_khmoi_uploads) {
    store.vnpay_khmoi_uploads = [];
    changed = true;
  }
  if (!store.vnpay_khmoi_payoo_uploads) {
    store.vnpay_khmoi_payoo_uploads = [];
    changed = true;
  }
  if (!store.vnpay_khmoi_manual_matches) {
    store.vnpay_khmoi_manual_matches = { offline: {}, payoo: {} };
    changed = true;
  }
  if (store.vnpay_khmoi_lock_date === undefined) {
    store.vnpay_khmoi_lock_date = "";
    changed = true;
  }
  if (migrateVnpayKhMoiInvoices(store)) changed = true;
  if (seedVnpayKhMoiManualMatchDefaults(store)) changed = true;
  return changed;
}

// Chi Nhan, 2026-07-30: "cái này lệch 40k của tàu á cộng vô doanh thu hôm đó
// ln nhá tiền có về rồi á là khớp còn mốt lệch cứ để lệch đi tôi sẽ cấn trừ
// hay nhập lên sao" -- Khoan ve 28/7 (doanh thu 27/7) kenh offline, gian KVC
// TIMES: file du lieu tai len thieu 40.000d doanh thu cua "Tàu Time"/"TUTU
// TRAIN VC TIMES" (co trong hoa don 10999 nhung khong co trong file gross),
// lam ca ngay lech dung 39.648d so ngan hang (39.648 = 40.000 sau khi tru % phi
// cung ty le voi phan con lai cua KVC TIMES ngay do -- xac nhan khop chinh
// xac). Day la 1 lan CHI NHAN CHU DONG xac nhan tien da ve/hop le (KHONG PHAI
// quy tac tu dong ap dung cho tuong lai -- Chi Nhan noi ro "mot lech cu de
// lech di toi se can" nen KHONG tao co che tu dong nao khac ngoai ban ghi cu
// the nay). Dung netAdjustment (xem ghi chu utils/zvpReconcile.js) de cong
// dung ty le phi thay vi lam trung net = gross (mat % phi cua phan gross cu).
// Seed KHONG DIEU KIEN moi lan load() (giong MANUAL_MATCH_DEFAULTS ben VietQR)
// de khong bi mat khi server dang chay cua Chi Nhan tu ghi de store.json.
const VNPAY_KHMOI_MANUAL_MATCH_DEFAULTS = {
  offline: {
    "2026-07-28|KVC TIMES": {
      invoiceNumbers: [10999],
      grossAdjustment: 40000,
      netAdjustment: 39648,
      note: "Bo sung 40.000d doanh thu Tau Time (TUTU TRAIN VC TIMES) bi thieu trong file du lieu tai len ngay 27/7 (co trong hoa don 10999) - tien da ve ngan hang, Chi Nhan xac nhan 30/7.",
    },
  },
};

function seedVnpayKhMoiManualMatchDefaults(store) {
  if (!store.vnpay_khmoi_manual_matches) store.vnpay_khmoi_manual_matches = { offline: {}, payoo: {} };
  let changed = false;
  Object.keys(VNPAY_KHMOI_MANUAL_MATCH_DEFAULTS).forEach((ch) => {
    if (!store.vnpay_khmoi_manual_matches[ch]) store.vnpay_khmoi_manual_matches[ch] = {};
    Object.keys(VNPAY_KHMOI_MANUAL_MATCH_DEFAULTS[ch]).forEach((key) => {
      const desired = VNPAY_KHMOI_MANUAL_MATCH_DEFAULTS[ch][key];
      const cur = store.vnpay_khmoi_manual_matches[ch][key];
      if (!cur || JSON.stringify(cur) !== JSON.stringify(Object.assign({ created_at: cur && cur.created_at }, desired))) {
        store.vnpay_khmoi_manual_matches[ch][key] = Object.assign(
          { created_at: (cur && cur.created_at) || new Date().toISOString() },
          desired
        );
        changed = true;
      }
    });
  });
  return changed;
}

// Same overlap-safe merge as ZVP/Momo: newer upload wins per (ngay, ma) key,
// so re-uploading a corrected/extended export supersedes older data instead
// of double-counting -- dung chung cho ca VNPay (mergeResolvedGross(store.vnpay_khmoi_uploads))
// va Payoo (mergeResolvedGross(store.vnpay_khmoi_payoo_uploads)).
function mergeResolvedGross(uploads) {
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  const sorted = [...uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  for (const u of sorted) {
    (u.codes || []).forEach((c) => codes.add(c));
    for (const [k, v] of Object.entries(u.grossByCode || {})) grossByCode[k] = v;
    for (const [k, v] of Object.entries(u.netByCode || {})) netByCode[k] = v;
  }
  return { codes: Array.from(codes), grossByCode, netByCode };
}

// Chi Nhan, 2026-07-29: "nó cũng check các cột như thế này nè ... thêm số
// tiền sau khi trừ phí ... cái phí thì bạn cộng trừ ra giống như của vn pay
// offline kh cũ" -- muon cot "TIỀN VỀ NH (NET)" giong het trang VNPay Offline
// KH Cu (ACB31268), luon khop dung Ngan hang (Chenh lech = 0). Ben KH Cu co
// file "phi" rieng (muc 1b) cho net THAT tung giao dich; file "Danh sach
// giao dich" (muc 1) hien tai KHONG co cot phi (da kiem tra: "Số tiền trước
// KM" == "Số tiền sau KM" tren toan bo du lieu mau, khong co gia tri chenh
// lech nao de tach), nen KHONG THE tinh phi that tung giao dich khi chua co
// file phi that (muc 1b). Trong luc cho file do, "cong tru ra" phi = phan bo
// LAI dung so tien NGAN HANG THAT (da xac nhan dung) theo ty le doanh thu gop
// cua tung gian trong CHINH dot ve tien do -- dam bao Tong NET luon = dung
// Ngan hang (Chenh lech = 0), giu gross de doi soat hoa don rieng. Neu 1
// dong DA co net that (vd Payoo, hoac VNPay sau khi tai file phi muc 1b --
// nhan biet qua net !== gross), GIU NGUYEN, khong ghi de.
function applyProportionalNetFallback(results) {
  results.forEach((r) => {
    if (r.pendingBank || !r.bankAmount) return;
    const hasRealNet = r.lines.some((l) => l.net !== l.gross);
    if (hasRealNet) return; // da co net that (vd Payoo, hoac VNPay da tai file phi) -- khong dong den.
    const includable = r.lines.filter((l) => l.tkCo !== "SKIP");
    const totalGross = includable.reduce((s, l) => s + l.gross, 0);
    if (totalGross <= 0) return;
    let allocated = 0;
    includable.forEach((l, i) => {
      if (i === includable.length - 1) {
        l.net = r.bankAmount - allocated;
      } else {
        l.net = Math.round((l.gross * r.bankAmount) / totalGross);
        allocated += l.net;
      }
    });
    r.totalNetComputed = includable.reduce((s, l) => s + l.net, 0);
    r.diffVsBank = r.totalNetComputed - r.bankAmount;
  });
}

// Chi Nhan, 2026-07-29: "chỗ này cấn trừ qua lại á bỏ cho tôi luôn nhá" --
// KVC TIMES ngay 07-27 (doanh thu 24-26) Lech -40.000d (invoiceTotal it hon
// gross), ngay 07-28 (doanh thu 27) Lech +40.000d (invoiceTotal nhieu hon
// gross) -- 2 ngay LIEN TIEP cung 1 gian bu tru vua khop (-40.000 + 40.000 =
// 0), dung mau hinh "hoa don xuat sang ngay ke tiep/truoc do" da xu ly cho
// VietQR (xem applyMultiDayGroupConsolidation trong utils/vietqrReconcile.js).
// Ap dung y het nguyen tac do cho ket qua ZVP o day: gom cac dong CUNG ma
// cong trinh theo THU TU ngay, tim moi day LIEN TIEP cac dong CHUA khop (bi
// chan boi dong DA khop/da khoa/da co ghi de thu cong o 2 dau), neu tong
// diff (co dau) cua ca day = 0 (trong sai so lam tron) thi coi CA DAY la da
// khop, khong con hien "Lệch" nua.
function applyConsecutiveRunNetting(results) {
  const byCode = new Map();
  results.forEach((r) => {
    r.lines.forEach((l) => {
      if (l.tkCo === "SKIP") return;
      if (!byCode.has(l.code)) byCode.set(l.code, []);
      byCode.get(l.code).push(l);
    });
  });
  byCode.forEach((lines) => {
    // `results` da duoc goi voi settlements sap theo thu tu, nhung de chac
    // chan (phong khi ham nay duoc goi truoc buoc sort ben ngoai), khong sap
    // lai o day -- lines da theo dung thu tu duyet `results` (calling code
    // phai dam bao results da sap theo settlementDate truoc khi goi ham nay).
    let i = 0;
    while (i < lines.length) {
      const isFree = (l) => !l.matched && !l.manualOverride && !l.locked && l.invoiceNumbers.length > 0;
      if (!isFree(lines[i])) {
        i++;
        continue;
      }
      let j = i;
      let sumDiff = 0;
      while (j < lines.length && isFree(lines[j])) {
        sumDiff += lines[j].diff;
        j++;
      }
      if (Math.abs(sumDiff) < 1 && j > i) {
        for (let k = i; k < j; k++) {
          lines[k].matched = true;
          lines[k].diff = 0;
          lines[k].manualOverride = true;
          lines[k].manualNote =
            "Tự động cấn trừ: lệch của kỳ này bù trừ vừa khớp với kỳ liền kề (tổng lệch gộp = 0).";
        }
      }
      i = j;
    }
  });
}

// Chi Nhan, 2026-07-30: "ngày 8 có cái 12.120 trên dữ liệu payoo này là text
// thôi bạn xóa ra ln đi cho tôi nhá" -- file "PY-GiaoDichBanHangPayoo-
// 01072026-29072026.xlsx" co 1 dong bi doc nham thanh "2026-07-08|KVC ROYAL"
// = 12.120d (thuc ra la text, khong phai giao dich that -- ngay 08/7 KVC
// ROYAL khong co doanh thu Payoo nao ca, xac nhan: bo dong nay di thi ca
// "Ngan hang" (3.757.098d) lan hoa don (8987,8988 = 3.780.000d) deu khop
// TUYET DOI voi rieng ngay 09/7, giai thich dung 2 cho lech "Chenh lech
// 11.892d" (= net cua dong rac) VA "Lech -12.120d" (= gross cua dong rac)
// cung luc). Loc bo SAU KHI merge (khong sua truc tiep tung file upload) de
// ap dung du la ban ghi nam trong file nao, KHONG BI MAT khi server dang
// chay cua Chi Nhan tu ghi de store.json.
const PAYOO_BAD_GROSS_KEYS = new Set(["2026-07-08|KVC ROYAL"]);

function stripPayooBadGrossKeys(merged) {
  PAYOO_BAD_GROSS_KEYS.forEach((key) => {
    delete merged.grossByCode[key];
    delete merged.netByCode[key];
  });
  return merged;
}

function buildReconciliation(store) {
  if (ensureVnpayKhMoiShape(store)) save(store);
  const bank = store.banks.find((b) => b.name === VTB_BANK_NAME);
  if (!bank) {
    return { error: `Chua co ngan hang "${VTB_BANK_NAME}" trong he thong.` };
  }
  const txs = store.transactions.filter((t) => t.bank_id === bank.id);
  const settlements = extractZvpSettlements(txs);

  const grossMerged = mergeResolvedGross(store.vnpay_khmoi_uploads);
  const payooGrossMerged = stripPayooBadGrossKeys(mergeResolvedGross(store.vnpay_khmoi_payoo_uploads));
  // Chi Nhan, 2026-07-29: "KVC ROYAL có lệch đâu đây nó chỉ có 1 hóa đơn mà
  // cộng chi mà nhiều vậy" -- store.viet_qr_invoices.vnpayKhMoi la 1 pool
  // hoa don DUNG CHUNG cho ca 2 kenh (gop tu ca 2 pool zvp_invoices.vnpay VA
  // .payoo qua migrateVnpayKhMoiInvoices), nen truoc day dua NGUYEN pool nay
  // vao CA 2 lan goi reconcileZvpChannel -- 1 hoa don Payoo (vd so 9274/9275,
  // raw "Payoo QR - .../Payoo thẻ - ...") bi khop NHAM vao ca dong VNPay
  // (offline) cua CUNG gian/ky, cong don sai invoiceTotal (vd 9133 that cua
  // Vnpay + 9274,9275 cua Payoo -> tao "Lech" gia trong khi 9133 mot minh da
  // khop dung roi). Tach lai theo dung tag "raw" da luu san tren tung hoa don
  // (giu nguyen qua spread {...inv} luc migrate) -- "Vnpay CS MB..." ve kenh
  // offline, "Payoo ... - ..." ve kenh payoo, giong cach ZALO_MAI_TAG_PATTERN
  // da dung de tach nham Zalo truoc do.
  const allInvoices = store.viet_qr_invoices.vnpayKhMoi || [];
  const vnpayInvoiceData = { invoices: allInvoices.filter((inv) => /vnpay/i.test(inv.raw || "")) };
  const payooInvoiceData = { invoices: allInvoices.filter((inv) => /payoo/i.test(inv.raw || "")) };
  const manualMatches = store.vnpay_khmoi_manual_matches || { offline: {}, payoo: {} };

  const reconciled = {
    // Doanh thu VNPay (Offline-style QR tai co so) tu file "DanhSachGiaoDich"
    // tai o muc 1 duoi day.
    offline: reconcileZvpChannel(settlements.offline, grossMerged, vnpayInvoiceData, store.gian_mapping, manualMatches.offline, store.invoice_diem_alias, new Set()),
    // Payoo: doanh thu tu file "Payoo thu hộ..."/"PY-GiaoDichBanHangPayoo..."
    // (muc 2 ben duoi), gan het ve 1 gian "KVC ROYAL" duy nhat.
    payoo: reconcileZvpChannel(settlements.payoo, payooGrossMerged, payooInvoiceData, store.gian_mapping, manualMatches.payoo, store.invoice_diem_alias, new Set()),
  };
  applyProportionalNetFallback(reconciled.offline);
  applyProportionalNetFallback(reconciled.payoo);

  // Chi Nhan, 2026-07-29: "sắp theo thời gian cho tôi sao lại loạn xạ cá
  // ngày như thế" -- reconcileZvpChannel tra ve theo thu tu duyet settlements
  // (goc + cac dong "cho ngan hang" them vao SAU CUNG), khong tu sap xep theo
  // ngay, nen cac dong "Chua co ngan hang" (settlementDate = ngay ke tiep du
  // kien) luon bi day xuong CUOI danh sach thay vi nam dung vi tri thoi gian
  // cua no. Sap lai theo settlementDate TANG DAN o day cho ca 2 kenh truoc
  // khi tra ve, dam bao hien dung thu tu ngay ca khi loc theo thang.
  ["offline", "payoo"].forEach((ch) => {
    (reconciled[ch] || []).sort((a, b) => (a.settlementDate < b.settlementDate ? -1 : a.settlementDate > b.settlementDate ? 1 : 0));
  });

  // Phai chay SAU khi da sap theo settlementDate (dua vao thu tu `results`
  // de gom dung theo thoi gian cho tung ma cong trinh) -- xem ghi chu day du
  // tai applyConsecutiveRunNetting o tren.
  applyConsecutiveRunNetting(reconciled.offline);
  applyConsecutiveRunNetting(reconciled.payoo);

  const lockDate = store.vnpay_khmoi_lock_date || "";
  if (lockDate) {
    ["offline", "payoo"].forEach((ch) => {
      (reconciled[ch] || []).forEach((r) => {
        if (r.settlementDate <= lockDate) {
          r.locked = true;
          r.lines.forEach((l) => {
            l.locked = true;
            l.diff = 0;
          });
        }
      });
    });
  }

  return { reconciled, lockDate };
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 80 * 1024 * 1024 } });

router.get("/doi-soat/vnpay-khmoi", (req, res) => {
  const store = load();
  const built = buildReconciliation(store);
  const reconciledAll = built.reconciled || { offline: [], payoo: [] };

  const monthSet = new Set();
  ["offline", "payoo"].forEach((ch) => (reconciledAll[ch] || []).forEach((r) => monthSet.add(r.settlementDate.slice(0, 7))));
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";

  const reconciled = { offline: [], payoo: [] };
  ["offline", "payoo"].forEach((ch) => {
    reconciled[ch] = selectedMonth
      ? (reconciledAll[ch] || []).filter((r) => r.settlementDate.slice(0, 7) === selectedMonth)
      : reconciledAll[ch] || [];
  });

  res.render("doisoat-vnpay-khmoi", {
    userName: req.session.userName,
    uploads: store.vnpay_khmoi_uploads || [],
    payooUploads: store.vnpay_khmoi_payoo_uploads || [],
    reconciled,
    months,
    selectedMonth,
    lockDate: built.lockDate || "",
    error: built.error || req.query.error || null,
    success: req.query.success || null,
  });
});

// ---------- Upload: file "Danh sách giao dịch" (cổng VNPay Merchant) ----------
router.post("/doi-soat/vnpay-khmoi/upload", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const parsed = parseVnpayPortalWorkbook(req.file.buffer);
    const resolved = resolveGianGross(parsed.rows, parsed.storeMap, [], {}, null, TEN_DIEM_TO_MA_CONG_TRINH);
    if (resolved.codes.length === 0) {
      throw new Error('Khong doc duoc dong giao dich nao khop voi diem thu da biet (AE HUE KVCN/KVC TIMES/KVC ROYAL/SAVICO KVCN) trong file nay.');
    }
    const dates = Array.from(new Set(parsed.rows.map((r) => r.date))).sort();

    if (!store.vnpay_khmoi_uploads) store.vnpay_khmoi_uploads = [];
    store.vnpay_khmoi_uploads.push({
      id: nextId(store, "vnpay_khmoi_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      dates,
      codes: resolved.codes,
      grossByCode: resolved.grossByCode,
      unmapped: resolved.unmapped,
    });
    save(store);

    let successMsg = `Da nap "${parsed.sheetName}" (${dates[0]} - ${dates[dates.length - 1]}), ${resolved.codes.length} ma cong trinh.`;
    if (resolved.unmapped.length > 0) {
      successMsg += ` CANH BAO: ${resolved.unmapped.length} diem thu chua khop mapping (${resolved.unmapped.join(", ")}).`;
    }
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: file "Dữ liệu báo cáo phí theo GD thanh toán" (cổng
// VNPay Merchant) -- CÓ tách phí, giống hệt file "phí" chị dùng cho VNPay
// Offline bên KH Cũ (ACB31268, routes/doisoat-zvp.js "upload-offline-raw").
// Chi Nhan, 2026-07-29: "phần check ngân hàng bạn phải check số sau khi trừ
// phí ... giống bên kh cũ 1268" -- file "Danh sách giao dịch" (mục 1 ở trên)
// CHỈ có "Số tiền sau KM" (KM = khuyến mãi, KHÔNG PHẢI phí VNPay thu), không
// có cột phí nào -- nên "Ngân hàng" (đã trừ phí thật) luôn lệch nhẹ so với
// dữ liệu tải lên ở mục 1. File này (xuất từ CÙNG cổng VNPay, nhưng chọn báo
// cáo "phí theo GD thanh toán" thay vì "Danh sách giao dịch") có thêm cột
// "Số tiền phí thu hộ"/"Số tiền sau khi trừ phí" -- gross (số trước trừ, đối
// chiếu hóa đơn) VÀ net (số sau trừ phí, đối chiếu ngân hàng) tách riêng,
// đúng như bên KH Cũ. Đẩy vào CHUNG store.vnpay_khmoi_uploads với mục 1 (key
// "ngày|mã", tải mới nhất thắng) nên ngày nào có file phí sẽ tự thay số gross-
// only cũ bằng số gross+net chính xác hơn, không cần xoá tay. ----------
router.post("/doi-soat/vnpay-khmoi/upload-phi", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const parsed = parseVnpayOfflineFeeReport(req.file.buffer);
    if (parsed.transactions.length === 0) {
      throw new Error('Khong doc duoc giao dich nao trong file "Du lieu bao cao phi theo GD thanh toan" nay.');
    }
    const grossByCode = {};
    const netByCode = {};
    const codes = new Set();
    const dates = new Set();
    const unmappedAgg = new Map();
    for (const tx of parsed.transactions) {
      const key = normText(tx.chiNhanh);
      const code = TEN_DIEM_TO_MA_CONG_TRINH[key];
      if (!code) {
        unmappedAgg.set(tx.chiNhanh, (unmappedAgg.get(tx.chiNhanh) || 0) + 1);
        continue;
      }
      codes.add(code);
      dates.add(tx.date);
      const k = `${tx.date}|${code}`;
      grossByCode[k] = (grossByCode[k] || 0) + tx.gross;
      netByCode[k] = (netByCode[k] || 0) + tx.net;
    }
    const unmapped = Array.from(unmappedAgg.entries()).map(([name, count]) => `${name}: ${count} giao dich`);
    const sortedDates = Array.from(dates).sort();
    if (sortedDates.length === 0) {
      throw new Error('File nay khong co diem thu nao khop voi 4 ma da biet (AE HUE KVCN/KVC TIMES/KVC ROYAL/SAVICO KVCN).');
    }

    if (!store.vnpay_khmoi_uploads) store.vnpay_khmoi_uploads = [];
    store.vnpay_khmoi_uploads.push({
      id: nextId(store, "vnpay_khmoi_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: `[File phí] ${req.file.originalname}`,
      sheetName: parsed.sheetName,
      dates: sortedDates,
      codes: Array.from(codes),
      grossByCode,
      netByCode,
      unmapped,
    });
    save(store);

    let successMsg = `Da nap file phi "${req.file.originalname}" (${sortedDates[0]} - ${sortedDates[sortedDates.length - 1]}) -- da tach rieng gross/phi/net, "Ngan hang" se doi soat dung so sau khi tru phi.`;
    if (unmapped.length > 0) {
      successMsg += ` CANH BAO: ${unmapped.length} diem thu chua khop mapping (${unmapped.join(", ")}).`;
    }
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vnpay-khmoi/upload/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  try {
    const id = Number(req.params.id);
    store.vnpay_khmoi_uploads = (store.vnpay_khmoi_uploads || []).filter((u) => u.id !== id);
    save(store);
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent("Da xoa bang doanh thu nay."));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload Payoo: file tong hop ("Payoo thu hộ KH ...xlsx", sheet
// "Dữ liệu Payoo cơ sở ..."). Chi Nhan xac nhan Payoo KH Moi chi co 1 gian
// duy nhat ("KVC ROYAL") nen KHONG can bang tra "Danh muc ten diem" -- gan
// thang ca file ve 1 ma nay. ----------
router.post("/doi-soat/vnpay-khmoi/upload-payoo", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const parsed = parsePayooWorkbook(req.file.buffer);
    if (parsed.rowsMatched === 0) throw new Error("Khong doc duoc dong giao dich Payoo nao trong file nay.");
    const grossByCode = {};
    const netByCode = {};
    for (const [key, v] of Object.entries(parsed.grossByCode)) {
      const date = key.split("|")[0];
      const newKey = `${date}|${PAYOO_FIXED_CODE}`;
      grossByCode[newKey] = (grossByCode[newKey] || 0) + v;
    }
    for (const [key, v] of Object.entries(parsed.netByCode)) {
      const date = key.split("|")[0];
      const newKey = `${date}|${PAYOO_FIXED_CODE}`;
      netByCode[newKey] = (netByCode[newKey] || 0) + v;
    }

    if (!store.vnpay_khmoi_payoo_uploads) store.vnpay_khmoi_payoo_uploads = [];
    store.vnpay_khmoi_payoo_uploads.push({
      id: nextId(store, "vnpay_khmoi_payoo_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      dates: parsed.dates,
      codes: [PAYOO_FIXED_CODE],
      grossByCode,
      netByCode,
    });
    save(store);

    res.redirect(
      "/doi-soat/vnpay-khmoi?success=" +
        encodeURIComponent(`Da nap "${parsed.sheetName}" (${parsed.dates[0]} - ${parsed.dates[parsed.dates.length - 1]}), ${parsed.rowsMatched} giao dich Payoo -> gian "${PAYOO_FIXED_CODE}".`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload Payoo: file THO xuat truc tiep tu cong Payoo ("PY-
// GiaoDichBanHangPayoo-...xlsx", "BÁO CÁO GIAO DỊCH BÁN HÀNG HỢP TÁC VỚI
// PAYOO") -- Chi Nhan: "hàng ngày tôi sẽ tải lên gian dữ liệu như thế này
// nhá" -- file re-export MOI NGAY thuong COVER LAI tu dau thang (khong chi 1
// ngay moi), nen KHONG dedup theo tung giao dich (khong can, vi grossByCode
// duoc TINH LAI TU DAU tu toan bo file, ghi de theo tung upload giong het
// mergeResolvedGross ben tren) -- lan tai MOI NHAT cho 1 ngay se tu dong thay
// the (khong cong don) lan tai cu cho CHINH ngay do. ----------
router.post("/doi-soat/vnpay-khmoi/upload-payoo-raw", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const parsed = parsePayooRawReport(req.file.buffer);
    const grossByCode = {};
    const netByCode = {};
    for (const tx of parsed.transactions) {
      const key = `${tx.date}|${PAYOO_FIXED_CODE}`;
      grossByCode[key] = (grossByCode[key] || 0) + tx.gross;
      netByCode[key] = (netByCode[key] || 0) + tx.net;
    }

    if (!store.vnpay_khmoi_payoo_uploads) store.vnpay_khmoi_payoo_uploads = [];
    store.vnpay_khmoi_payoo_uploads.push({
      id: nextId(store, "vnpay_khmoi_payoo_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: `[File thô] ${req.file.originalname}`,
      sheetName: parsed.sheetName,
      dates: parsed.dates,
      codes: [PAYOO_FIXED_CODE],
      grossByCode,
      netByCode,
    });
    save(store);

    let successMsg = `Da nap "${req.file.originalname}" (${parsed.dates[0]} - ${parsed.dates[parsed.dates.length - 1]}), ${parsed.rowsMatched} giao dich -> gian "${PAYOO_FIXED_CODE}".`;
    if (parsed.rowsSkippedInvalid > 0) successMsg += ` (bo qua ${parsed.rowsSkippedInvalid} giao dich khong hop le/khong thanh cong)`;
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent(successMsg));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vnpay-khmoi/upload-payoo/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  try {
    const id = Number(req.params.id);
    store.vnpay_khmoi_payoo_uploads = (store.vnpay_khmoi_payoo_uploads || []).filter((u) => u.id !== id);
    save(store);
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent("Da xoa bang doanh thu Payoo nay."));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Khoa so ----------
router.post("/doi-soat/vnpay-khmoi/khoa-so", requireAdmin, (req, res) => {
  const store = load();
  try {
    const lockDate = (req.body.lockDate || "").trim();
    if (lockDate && !/^\d{4}-\d{2}-\d{2}$/.test(lockDate)) throw new Error("Ngay khoa khong hop le (dang YYYY-MM-DD).");
    store.vnpay_khmoi_lock_date = lockDate;
    save(store);
    const msg = lockDate ? `Da khoa so den het ngay ${lockDate}.` : "Da mo khoa so.";
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Danh dau thu cong ----------
router.post("/doi-soat/vnpay-khmoi/manual-match", requireDataEntry, (req, res) => {
  const store = load();
  try {
    const { channel, settlementDate, code, invoiceNumbers, amount, grossAdjustment, note } = req.body;
    if (!["offline", "payoo"].includes(channel)) throw new Error("Kenh khong hop le.");
    if (!settlementDate || !code) throw new Error("Thieu thong tin dong can danh dau.");
    if (!store.vnpay_khmoi_manual_matches) store.vnpay_khmoi_manual_matches = { offline: {}, payoo: {} };
    if (!store.vnpay_khmoi_manual_matches[channel]) store.vnpay_khmoi_manual_matches[channel] = {};
    const key = `${settlementDate}|${code}`;
    const invoiceList = (invoiceNumbers || "").split(/[,\s]+/).map((s) => s.trim()).filter(Boolean);
    const amt = amount ? Number(String(amount).replace(/[^\d-]/g, "")) : null;
    const grossAdj = grossAdjustment ? Number(String(grossAdjustment).replace(/[^\d-]/g, "")) : 0;
    store.vnpay_khmoi_manual_matches[channel][key] = {
      invoiceNumbers: invoiceList,
      amount: amt,
      grossAdjustment: grossAdj,
      note: note || "",
      created_at: new Date().toISOString(),
    };
    save(store);
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent(`Da danh dau thu cong dong "${code}" ngay ${settlementDate}.`));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vnpay-khmoi/manual-match/delete", requireAdmin, (req, res) => {
  const store = load();
  try {
    const { channel, settlementDate, code } = req.body;
    if (store.vnpay_khmoi_manual_matches && store.vnpay_khmoi_manual_matches[channel]) {
      delete store.vnpay_khmoi_manual_matches[channel][`${settlementDate}|${code}`];
    }
    save(store);
    res.redirect("/doi-soat/vnpay-khmoi?success=" + encodeURIComponent("Da xoa danh dau thu cong."));
  } catch (e) {
    res.redirect("/doi-soat/vnpay-khmoi?error=" + encodeURIComponent(e.message));
  }
});

// Exposed for reuse/testing, cung kieu voi router.buildChannelReconciliation
// trong routes/doisoat-vietqr.js.
router.buildReconciliation = buildReconciliation;

module.exports = router;
