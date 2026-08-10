const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const {
  parseChiPhiRawWorkbook,
  resolveVendor,
  extractSiteFragment,
  matchGianForFragment,
  buildNccIndex,
  matchNccForVendor,
  findNccByCode,
  parseNccWorkbook,
  parseHddvInvoiceWorkbook,
  buildInvoiceIndex,
  matchInvoiceForPayment,
  parseUncWorkbook,
  buildUncIndex,
  matchUncForPayment,
  extractHdNumberFromText,
  isoToDmy,
  DEFAULT_NCC_LIST,
  normText,
} = require("../utils/chiphiReconcile");
const { getCompany } = require("../utils/companies");
const overviewAggregate = require("../utils/overviewAggregate");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

// 2 tai khoan chi phi rieng (tien CHI RA cho NCC/mat bang/luong...), doc lap
// voi cac kenh doanh thu Momo/ZVP/Viet QR. Khong co danh sach "hoa don" rieng
// de doi chieu nhu ben doanh thu -- ban than file sao ke NGAN HANG da la du
// lieu chi tiet tung dong roi, nen trang nay la "phan loai + xuat Misa" chu
// khong phai "doi soat 2 nguon". Ma kenh dat theo ten sheet Luyen dang dung
// (NH8651/NH9997) de de doi chieu khi upload lai file.
const CHANNELS = {
  nh8651: { bankName: "BIDV8651", label: "BIDV 8651", sheetMatch: /8651/, company: "kh_cu" },
  vpbank9997: { bankName: "VPBANK9997", label: "VPBank 9997", sheetMatch: /9997/, company: "kh_cu" },
  // Cong ty "KH Moi" (TNHH GIAI TRI K&H) -- tai khoan chi phi VPBank 58888.
  // sheetMatch dung "58888" (khong dung rieng "8888" vi co the trung voi so
  // khac trong ten sheet) de tu nhan dien khi upload file co sheet ten
  // "VP58888" giong file sao ke thuc te Luyen da dung.
  vp58888: { bankName: "VP58888", label: "VP 58888 (KH Mới)", sheetMatch: /58888/, company: "kh_moi" },
  // Chi Nhan, 2026-07-29: "thêm 1 tk chi bên kh mới cho tôi đi ... với gắn vô
  // mục tk chi cho tôi nhá" -- them tai khoan chi phi moi BIDV8681 (KH Moi).
  // bankName PHAI khop CHINH XAC "Tên hiển thị" cua tai khoan trong Quan ly
  // ngan hang (xem buildChannelChiPhi: tim theo store.banks.find(b => b.name
  // === cfg.bankName)) -- Chi Nhan can tao tai khoan nay voi Ten hien thi
  // dung la "BIDV8681" thi kenh moi nay moi nhan dien duoc.
  bidv8681: { bankName: "BIDV8681", label: "BIDV 8681 (KH Mới)", sheetMatch: /8681/, company: "kh_moi" },
};
const CHANNEL_KEYS = Object.keys(CHANNELS);
const UPDATED_NOTE = " Ket qua ben duoi da tu cap nhat theo du lieu moi.";

// TK theo Hinh Thuc Hop Tac: "chia se" -> TK Co 1388 (khoan phai thu),
// con lai ("tien thue"...) -> TK No 331 (phai tra nguoi ban).
// Dung store.hoa_don_dau_vao_gian_list de xay lookup dong (user da upload
// "Danh sach cac gian" Google Sheet). Neu chua co list -> tra chuoi rong.
function buildChiaSeTKSet(gianList) {
  const chiaSe = new Set();
  const tienThue = new Set();
  (gianList || []).forEach((g) => {
    const isChiaSe = normText(g.hinhThucHopTac || "").includes("chia");
    const words = normText((g.maDiemThue || "") + " " + (g.gianHang || ""))
      .split(/\s+/).filter((w) => w.length >= 3);
    words.forEach((w) => (isChiaSe ? chiaSe : tienThue).add(w));
  });
  // Xoa cac tu xuat hien ca 2 nhom (khong phan biet duoc)
  for (const w of chiaSe) if (tienThue.has(w)) chiaSe.delete(w);
  return chiaSe;
}

function getGianTK(gian, chiaSeTKSet) {
  if (!gian) return "331"; // mac dinh: tien thue -> 331
  const words = normText(gian).split(/\s+/).filter((w) => w.length >= 3);
  return words.some((w) => chiaSeTKSet.has(w)) ? "1388" : "331";
}

function ensureShape(store) {
  if (!store.chi_phi_raw_uploads) store.chi_phi_raw_uploads = {};
  if (!store.chi_phi_vendor_tk_map) store.chi_phi_vendor_tk_map = {};
  if (!store.chi_phi_gian_override) store.chi_phi_gian_override = {};
  if (!store.chi_phi_vendor_ncc_map) store.chi_phi_vendor_ncc_map = {}; // NCC go tay khi khong tu khop duoc
  if (!store.chi_phi_ncc_list) store.chi_phi_ncc_list = null; // null = dung danh sach goc 884 NCC di kem app
  if (!store.chi_phi_ncc_meta) store.chi_phi_ncc_meta = null; // { uploaded_at, file_name, count }
  if (!store.chi_phi_invoice_list) store.chi_phi_invoice_list = []; // danh sach hoa don NCC (sheet HDDV KH CU) -- doi chieu Ten NCC+So tien
  if (!store.chi_phi_invoice_meta) store.chi_phi_invoice_meta = null; // { uploaded_at, file_name, count, sheetName }
  if (!store.chi_phi_unc_list) store.chi_phi_unc_list = []; // bang lenh chi UNC -- doi chieu Ten NCC+So tien de lay Dien giai sach
  if (!store.chi_phi_unc_meta) store.chi_phi_unc_meta = null; // { uploaded_at, file_name, count, sheetName }
  CHANNEL_KEYS.forEach((ch) => {
    if (!store.chi_phi_raw_uploads[ch]) store.chi_phi_raw_uploads[ch] = [];
  });
}

// Danh sach NCC dang dung: uu tien ban Luyen tu upload lai, khong co thi
// dung ban goc 884 NCC "KH cu" di kem app.
function activeNccList(store) {
  return store.chi_phi_ncc_list && store.chi_phi_ncc_list.length ? store.chi_phi_ncc_list : DEFAULT_NCC_LIST;
}

// Newest upload wins per (ngay, so chung tu, chi ra, thu vao) key -- cung 1
// quy uoc voi cach cac kenh doanh thu gop nhieu lan tai len (Luyen thuong tai
// lai ban "tu dau den nay" moi ky, khong chi phan moi).
function mergeChiPhiUploads(uploads) {
  const sorted = [...uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  const byKey = {};
  for (const u of sorted) {
    for (const row of u.rows || []) {
      const key = `${row.date}|${row.docNo}|${row.debit}|${row.credit}`;
      byKey[key] = row;
    }
  }
  return Object.values(byKey);
}

// Luyen yeu cau 2026-07-17: moi dong Chi phi (gan voi 1 gian) can biet gian
// DO da "hach toan doanh thu" (co hoa don khop -- Khop) hay "chua hach toan
// doanh thu" (Chua co HD/Lech) trong cung thang, dua theo dung ket qua doi
// soat doanh thu (Momo/ZVP/Viet QR) da co san qua utils/overviewAggregate --
// KHONG tinh rieng, tranh lech so voi trang Cong no/Tong quan. Tinh 1 lan
// cho ca request (goi lai o day thay vi trong buildChannelChiPhi) vi ham
// nay duoc goi nhieu lan (moi kenh chi phi 1 lan) trong 1 request GET.
function buildRevenueStatusMap(store) {
  const flat = overviewAggregate.buildAllFlatLines(store);
  const map = {};
  flat.forEach((l) => {
    if (!l.gian) return;
    const key = `${l.gian}|${l.month}`;
    if (!map[key]) map[key] = { matched: 0, total: 0 };
    map[key].total++;
    if (overviewAggregate.isResolved(l)) map[key].matched++;
  });
  return map;
}

function revenueStatusFor(gian, month, revenueStatusMap) {
  if (!gian) return "Không xác định gian";
  const info = revenueStatusMap[`${gian}|${month}`];
  if (!info) return "Không có doanh thu ghi nhận";
  return info.matched === info.total ? "Đã hạch toán doanh thu" : "Chưa hạch toán doanh thu";
}

function buildChannelChiPhi(store, channelKey, revenueStatusMap) {
  ensureShape(store);
  const cfg = CHANNELS[channelKey];
  const bank = store.banks.find((b) => b.name === cfg.bankName);
  if (!bank) {
    return { error: `Chua co ngan hang "${cfg.bankName}" trong he thong.`, lines: [] };
  }
  const revMap = revenueStatusMap || buildRevenueStatusMap(store);

  const merged = mergeChiPhiUploads(store.chi_phi_raw_uploads[channelKey]);
  const vendorTkMap = store.chi_phi_vendor_tk_map || {};
  const gianOverride = store.chi_phi_gian_override || {};
  const gianList = store.zvp_gian_list || [];
  const nccList = activeNccList(store);
  const nccIndex = buildNccIndex(nccList);
  const vendorNccMap = store.chi_phi_vendor_ncc_map || {};
  const invoiceIndex = buildInvoiceIndex(store.chi_phi_invoice_list);
  const uncIndex = buildUncIndex(store.chi_phi_unc_list);

  const lines = merged
    .filter((r) => r.debit > 0) // chi ra -- credit-side rows (tien VAO tk chi phi, vd chuyen noi bo) khong thuoc pham vi trang nay
    .map((r) => {
      const { vendor, vendorKnown } = resolveVendor(r);
      const key = `${channelKey}|${r.date}|${r.docNo}|${r.debit}`;
      const siteFrag = extractSiteFragment(r.description);
      const { maCongTrinh: autoGian, candidates } = matchGianForFragment(siteFrag, gianList);
      const hasOverride = Object.prototype.hasOwnProperty.call(gianOverride, key);
      const finalGian = hasOverride ? gianOverride[key] : autoGian || "";

      // Ma NCC: uu tien ban Luyen go tay (theo NCC), khong co thi tu tra theo
      // danh sach dang dung. Go tay van tra Ten NCC/MST tu danh sach neu ma do
      // co trong danh sach, khong thi chi hien dung ma go tay.
      const manualNccCode = vendorNccMap[vendor];
      let maNCC = "", tenNCC = "", mstNCC = "", nccAmbiguous = false;
      if (manualNccCode) {
        maNCC = manualNccCode;
        const rec = findNccByCode(manualNccCode, nccList);
        tenNCC = rec ? rec.tenNCC : "";
        mstNCC = rec ? rec.mst : "";
      } else {
        const nccMatch = matchNccForVendor(vendor, nccIndex);
        maNCC = nccMatch.maNCC || "";
        tenNCC = nccMatch.tenNCC || "";
        mstNCC = nccMatch.mstNCC || "";
        nccAmbiguous = nccMatch.candidates.length > 1;
      }

      // Khop voi bang lenh chi UNC (Ten NCC + So tien) de lay Dien giai sach
      // hon dong sao ke ngan hang thuc te, roi khop voi danh sach hoa don
      // HDDV (MST/Ten NCC + So tien) de tu dien So hoa don/Ngay hoa don cho
      // file xuat Misa. Khong khop duoc hoa don thi thu rut so HD ngay tu
      // dien giai (UNC hoac sao ke) -- khong khop gi ca thi de trong, van
      // xuat dong nay nhu binh thuong (khong bo qua).
      const searchName = tenNCC || vendor;
      const uncMatch = matchUncForPayment(searchName, r.debit, uncIndex);
      const invoiceMatch = matchInvoiceForPayment(mstNCC, searchName, r.debit, invoiceIndex);
      const hdNumberFallback = !invoiceMatch ? extractHdNumberFromText((uncMatch && uncMatch.noiDungUnc) || r.description) : null;

      return {
        key,
        channelKey,
        month: r.date.slice(0, 7),
        date: r.date,
        debit: r.debit,
        description: r.description,
        docNo: r.docNo,
        taiKhoanDoiUng: r.taiKhoanDoiUng,
        nganHangDoiUng: r.nganHangDoiUng,
        vendor,
        vendorKnown,
        tkNo: vendorTkMap[vendor] || "",
        siteFrag,
        autoGian: autoGian || "",
        gianCandidates: candidates,
        finalGian,
        gianOverridden: hasOverride,
        maNCC,
        tenNCC,
        mstNCC,
        nccAmbiguous,
        nccOverridden: !!manualNccCode,
        uncDienGiai: uncMatch ? uncMatch.noiDungUnc : "",
        invoiceMatch,
        hdNumberFallback,
        // Luyen yeu cau 2026-07-17: 2 trang thai them de loc/xem nhanh --
        // "coHoaDon" la hoa don CHI PHI (NCC) cua chinh dong nay (da co san
        // qua invoiceMatch/hdNumberFallback, chi lam ro thanh 1 nhan don);
        // "revenueStatus" la trang thai DOANH THU cua GIAN nay trong THANG
        // nay (doc lap, lay tu ket qua doi soat doanh thu Momo/ZVP/VietQR).
        coHoaDon: invoiceMatch ? "Có" : hdNumberFallback ? "Có (theo diễn giải)" : "Không",
        revenueStatus: revenueStatusFor(finalGian, r.date.slice(0, 7), revMap),
      };
    })
    .sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : 0));

  // "Khoa so" -- Luyen, 2026-07-21: "TẤT CẢ CÁC TRANG ĐIỀU CÓ KHÓA SỔ CHO TÔI
  // NHÁ". Trang nay KHAC 3 trang doanh thu (Momo/ZVP/VietQR): khong co tong
  // "lech" theo ngay, chi gan nhan CANH BAO tung dong rieng le ("Can xac
  // dinh NCC", "chua khop hoa don", "Chua hach toan doanh thu"...). Luyen xac
  // nhan (2026-07-21, AskUserQuestion): khoa o day nghia la AN CANH BAO cho
  // TUNG DONG co ngay <= ngay khoa (giu nguyen du lieu, chi khong con hien
  // badge "can xu ly" nua) -- luu theo TUNG CONG TY (giong Momo), vi trang
  // nay cung chuyen theo cong ty dang xem (topbar Cu/Moi).
  const companyKey = cfg.company;
  const lockDate = (store.chi_phi_lock_date && store.chi_phi_lock_date[companyKey]) || "";
  if (lockDate) {
    lines.forEach((l) => {
      if (l.date <= lockDate) l.locked = true;
    });
  }

  return { lines, error: null, lockDate };
}

// Chi Nhan, 2026-07-29: "check theo đối tượng NCC ... đối tượng đó nó xuất
// cho mình các hóa đơn nào mình thanh toán chưa ngày nào từ tài khoản nào số
// tiền bao nhiêu để cấn trừ công nợ ... thêm cái công nợ ncc theo tên cho
// tôi" -- gop hoa don Hóa Đơn Đầu Vào (store.hoa_don_dau_vao, DUNG cong ty
// dang xem) theo Ten NCC: Tong hoa don / Da chi / Con no, kem drill-down
// TUNG HOA DON (KHONG phai tung dong hang hoa -- xem ghi chu ben trong ham).
// Tach thanh ham rieng (khong inline trong route) de dung LAI duoc cho ca
// trang GET (hien thi) va route export.xlsx (xuat file, Chi Nhan yeu cau
// 2026-07-29 "xuất ra chi tiết chỗ này á cho tôi").
function computeNccDebtRows(store, activeCompany) {
  const activeKeys = CHANNEL_KEYS.filter((ch) => CHANNELS[ch].company === activeCompany);
  const revenueStatusMap = buildRevenueStatusMap(store);
  const allLines = activeKeys.flatMap((ch) => buildChannelChiPhi(store, ch, revenueStatusMap).lines || []);

  const bankIdsForCompany = new Set(
    (store.banks || []).filter((b) => (b.company || "kh_cu") === activeCompany).map((b) => b.id)
  );
  const generalChiTx = (store.transactions || []).filter(
    (t) => t.type === "chi" && bankIdsForCompany.has(t.bank_id) && (t.tenDoiUng || "").trim()
  );
  const bankNameById = new Map((store.banks || []).map((b) => [b.id, b.name]));

  // Chi Nhan, 2026-07-29: "sao dò dc chi rồi chưa mà chưa dò được nguồn chi
  // vậy" -- BUG: 1 hóa đơn co the co NHIEU DONG hang hoa (vd hoa don
  // 00008443 co 23 dong Kem Celano/Merino rieng le), moi dong chi la 1 PHAN
  // nho cua tong tien hoa don -- daChiTien duoc gan THEO CA NHOM (so sanh
  // TONG TIEN CA HOA DON voi 1 giao dich/dong chi phi, xem cap-nhat-da-chi
  // ben Hoa Don Dau Vao) roi to CA nhom la Da chi, chu KHONG PHAI tung dong
  // rieng le co giao dich rieng. Truoc do code nay do tim giao dich khop
  // DUNG SO TIEN TUNG DONG (vd 309.000d) nen khong bao gio khop dc, luon ra
  // "chua do duoc nguon chi". Sua: GOP cac dong CUNG 1 So hoa don (+Ky hieu)
  // lai THANH 1 HOA DON (cong don Tien) truoc khi do nguon chi, dung LAI
  // dung 1 cach key voi groupRowsByInvoice ben Hoa Don Dau Vao de ra cung
  // ket qua. LUU Y (Chi Nhan, 2026-07-29 "có thể chi theo đợt hay công nợ
  // cuối tháng mới xuất"): neu 1 giao dich tra GOP nhieu hoa don lai lam 1
  // lan (hoac tra tung dot/1 phan hoa don), so tien giao dich do se KHONG
  // khop dung tong 1 hoa don rieng le -- nhung truong hop nay se KHONG do ra
  // duoc nguon chi (van hien "Đã chi" dung theo du lieu da danh dau, chi la
  // khong biet chinh xac giao dich nao), Chi Nhan tu doi chieu them neu can.
  function findPaymentForGroup(soTien, tenNCC, mstNCC) {
    const nccNormShort = normText(tenNCC).slice(0, 12);
    const viaChiPhi = allLines.find((l) => {
      if (Math.abs(l.debit - soTien) > 1000) return false;
      const mstMatch = mstNCC && l.mstNCC && l.mstNCC === mstNCC;
      const tenMatch = (l.tenNCC || l.vendor) && normText(l.tenNCC || l.vendor).includes(nccNormShort);
      return mstMatch || tenMatch;
    });
    if (viaChiPhi) return { date: viaChiPhi.date, source: CHANNELS[viaChiPhi.channelKey].label };
    const viaTx = generalChiTx.find(
      (t) => Math.abs(t.amount - soTien) <= 1000 && normText(t.tenDoiUng).includes(nccNormShort)
    );
    if (viaTx) return { date: viaTx.date, source: bankNameById.get(viaTx.bank_id) || "" };
    return null;
  }

  const hddvRows = (store.hoa_don_dau_vao || []).filter((r) => r.congTy === activeCompany);
  const nccDebtMap = new Map();
  hddvRows.forEach((r) => {
    const nccKey = (r.tenNCC || "(chưa rõ NCC)").trim();
    if (!nccDebtMap.has(nccKey)) {
      nccDebtMap.set(nccKey, { tenNCC: nccKey, mstNCC: r.mstNCC || "", invoiceGroups: new Map(), invoiceOrder: [] });
    }
    const g = nccDebtMap.get(nccKey);
    // Gop dung 1 kieu voi groupRowsByInvoice (Hoa Don Dau Vao): cung So hoa
    // don + Ky hieu la 1 hoa don, khong co So hoa don thi coi moi dong 1 hoa don rieng.
    const invKey = r.soHoaDon ? `${r.soHoaDon}|${r.kyHieuHD || ""}` : `__single__${r.id}`;
    if (!g.invoiceGroups.has(invKey)) {
      g.invoiceGroups.set(invKey, {
        ngayHD: r.ngayHD,
        soHoaDon: r.soHoaDon,
        dienGiaiList: [],
        soTien: 0,
        daChiTienCount: 0,
        soDong: 0,
      });
      g.invoiceOrder.push(invKey);
    }
    const inv = g.invoiceGroups.get(invKey);
    inv.soTien += r.soTien || 0;
    inv.soDong++;
    if (r.dienGiai) inv.dienGiaiList.push(r.dienGiai);
    if (r.daChiTien) inv.daChiTienCount++;
  });
  const nccDebtRows = Array.from(nccDebtMap.values())
    .map((g) => {
      const invoices = g.invoiceOrder.map((k) => {
        const inv = g.invoiceGroups.get(k);
        const daChiTien = inv.daChiTienCount > 0; // dong nao trong nhom cung da danh dau Da chi thi ca hoa don la Da chi
        const payment = daChiTien ? findPaymentForGroup(inv.soTien, g.tenNCC, g.mstNCC) : null;
        return {
          ngayHD: inv.ngayHD,
          soHoaDon: inv.soHoaDon,
          dienGiai: inv.soDong > 1 ? `${inv.dienGiaiList[0] || ""} (+${inv.soDong - 1} dòng khác)` : inv.dienGiaiList[0] || "",
          soTien: inv.soTien,
          daChiTien,
          paymentDate: payment ? payment.date : "",
          paymentSource: payment ? payment.source : "",
        };
      });
      const tongHoaDon = invoices.reduce((s, i) => s + i.soTien, 0);
      const daChi = invoices.reduce((s, i) => s + (i.daChiTien ? i.soTien : 0), 0);
      return { tenNCC: g.tenNCC, mstNCC: g.mstNCC, tongHoaDon, daChi, invoices };
    })
    .map((g) => ({ ...g, conNo: g.tongHoaDon - g.daChi, soHoaDonCount: g.invoices.length }))
    .filter((g) => g.tongHoaDon !== 0)
    .sort((a, b) => Math.abs(b.conNo) - Math.abs(a.conNo));
  const nccDebtGrandTotal = {
    tongHoaDon: nccDebtRows.reduce((s, g) => s + g.tongHoaDon, 0),
    daChi: nccDebtRows.reduce((s, g) => s + g.daChi, 0),
    conNo: nccDebtRows.reduce((s, g) => s + g.conNo, 0),
  };
  return { nccDebtRows, nccDebtGrandTotal };
}

router.get("/doi-soat/chi-phi", (req, res) => {
  const store = load();
  ensureShape(store);

  // Chi hien cac kenh thuoc cong ty dang chon (nut chuyen cong ty tren
  // topbar) -- KH Cu thay BIDV 8651 + VPBank 9997 nhu truoc gio, KH Moi thay
  // VP 58888. Cac route upload/xoa/xuat file van nhan moi channelKey hop le
  // (khong gioi han theo cong ty) vi Luyen la nguoi dung duy nhat.
  const activeCompany = getCompany(req);
  const activeKeys = CHANNEL_KEYS.filter((ch) => CHANNELS[ch].company === activeCompany);

  const revenueStatusMap = buildRevenueStatusMap(store);
  const built = {};
  activeKeys.forEach((ch) => {
    built[ch] = buildChannelChiPhi(store, ch, revenueStatusMap);
  });

  const allLines = activeKeys.flatMap((ch) => built[ch].lines || []);
  const monthSet = new Set(allLines.map((l) => l.month));
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";
  const selectedChannel = req.query.channel !== undefined ? req.query.channel : "";
  // 2 bo loc them (Luyen yeu cau 2026-07-17): "coHoaDon" (Co/Khong -- hoa don
  // NCC cua chinh dong chi phi) va "revenueStatus" (trang thai doanh thu cua
  // GIAN trong thang -- doc lap voi hoa don NCC ben tren).
  const selectedInvoiceStatus = req.query.invoiceStatus || "";
  const selectedRevenueStatus = req.query.revenueStatus || "";

  let lines = allLines;
  if (selectedMonth) lines = lines.filter((l) => l.month === selectedMonth);
  if (selectedChannel) lines = lines.filter((l) => l.channelKey === selectedChannel);
  if (selectedInvoiceStatus) {
    lines = lines.filter((l) =>
      selectedInvoiceStatus === "co" ? l.coHoaDon.startsWith("Có") : l.coHoaDon === "Không"
    );
  }
  if (selectedRevenueStatus) lines = lines.filter((l) => l.revenueStatus === selectedRevenueStatus);

  const vendorSet = new Map();
  allLines.forEach((l) => {
    if (l.vendor) vendorSet.set(l.vendor, l.tkNo);
  });
  const vendors = Array.from(vendorSet.entries())
    .map(([vendor, tkNo]) => ({ vendor, tkNo }))
    .sort((a, b) => a.vendor.localeCompare(b.vendor));

  const totalAmount = lines.reduce((s, l) => s + l.debit, 0);

  const { nccDebtRows, nccDebtGrandTotal } = computeNccDebtRows(store, activeCompany);

  res.render("doisoat-chiphi", {
    userName: req.session.userName,
    channels: activeKeys.map((ch) => ({ key: ch, label: CHANNELS[ch].label })),
    uploadsByChannel: activeKeys.reduce((acc, ch) => {
      acc[ch] = store.chi_phi_raw_uploads[ch];
      return acc;
    }, {}),
    errorsByChannel: activeKeys.reduce((acc, ch) => {
      acc[ch] = built[ch].error;
      return acc;
    }, {}),
    lines,
    vendors,
    months,
    selectedMonth,
    selectedChannel,
    selectedInvoiceStatus,
    selectedRevenueStatus,
    totalAmount,
    nccCount: activeNccList(store).length,
    nccIsCustom: !!(store.chi_phi_ncc_list && store.chi_phi_ncc_list.length),
    nccMeta: store.chi_phi_ncc_meta,
    invoiceMeta: store.chi_phi_invoice_meta,
    invoiceCount: (store.chi_phi_invoice_list || []).length,
    uncMeta: store.chi_phi_unc_meta,
    uncCount: (store.chi_phi_unc_list || []).length,
    lockDate: (store.chi_phi_lock_date && store.chi_phi_lock_date[activeCompany]) || "",
    nccDebtRows,
    nccDebtGrandTotal,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Chi Nhan, 2026-07-29: "xuất ra chi tiết chỗ này á cho tôi cái cộng từng cái
// đi" -- xuat Excel muc "Cong no NCC theo ten": 1 dong / 1 hoa don (da GOP
// theo So hoa don, cong don so tien -- xem computeNccDebtRows), kem 1 dong
// tong ket dau moi NCC de de cong doi chieu ngoai Excel.
router.get("/doi-soat/chi-phi/cong-no-ncc/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const { nccDebtRows } = computeNccDebtRows(store, activeCompany);

  const exportRows = [];
  nccDebtRows.forEach((g) => {
    exportRows.push({
      "Tên NCC": g.tenNCC,
      "MST": g.mstNCC,
      "Số HĐ": "",
      "Ngày HĐ": "",
      "Diễn giải": `--- TỔNG ${g.soHoaDonCount} hóa đơn ---`,
      "Số tiền": g.tongHoaDon,
      "Đã chi": g.daChi,
      "Còn nợ": g.conNo,
      "Đã chi?": "",
      "Ngày thanh toán (dò được)": "",
      "Nguồn thanh toán (dò được)": "",
    });
    g.invoices.forEach((inv) => {
      exportRows.push({
        "Tên NCC": g.tenNCC,
        "MST": g.mstNCC,
        "Số HĐ": inv.soHoaDon,
        "Ngày HĐ": inv.ngayHD,
        "Diễn giải": inv.dienGiai,
        "Số tiền": inv.soTien,
        "Đã chi": inv.daChiTien ? inv.soTien : 0,
        "Còn nợ": inv.daChiTien ? 0 : inv.soTien,
        "Đã chi?": inv.daChiTien ? "Có" : "Không",
        "Ngày thanh toán (dò được)": inv.paymentDate,
        "Nguồn thanh toán (dò được)": inv.paymentSource,
      });
    });
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Cong no NCC");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=cong-no-ncc-${activeCompany}.xlsx`);
  res.send(buf);
});

// ---------- Khoa so (giong Momo/ZVP/VietQR, nhung an canh bao TUNG DONG theo
// ngay thay vi 1 tong "lech" theo ngay -- xem ghi chu tai buildChannelChiPhi).
// Luu theo TUNG CONG TY dang xem (topbar Cu/Moi). Gui lockDate rong de mo
// khoa lai. ----------
router.post("/doi-soat/chi-phi/khoa-so", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const lockDate = (req.body.lockDate || "").trim();
    if (lockDate && !/^\d{4}-\d{2}-\d{2}$/.test(lockDate)) throw new Error("Ngay khoa khong hop le (dang YYYY-MM-DD).");
    if (!store.chi_phi_lock_date) store.chi_phi_lock_date = {};
    store.chi_phi_lock_date[activeCompany] = lockDate;
    save(store);
    const msg = lockDate
      ? `Da khoa so den het ngay ${lockDate}. Cac dong tu do tro ve truoc se khong con hien canh bao nua.`
      : "Da mo khoa so.";
    res.redirect("/doi-soat/chi-phi?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

// 1 file duy nhat co the chua CA 2 sheet (1 cho moi tai khoan, dung nhu file
// Luyen thuong tai xuong tu 2 ngan hang) -- moi sheet duoc tu dong gan vao
// dung kenh theo TEN SHEET (chua "8651" hoac "9997"), khong can tach file/tai
// rieng tung tai khoan. Sheet nao khong khop ten kenh nao thi bao loi ro rang
// thay vi am tham bo qua hoac gan nham tai khoan.
router.post("/doi-soat/chi-phi/upload", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const sheets = parseChiPhiRawWorkbook(req.file.buffer);

    const assigned = {};
    const unassignedSheets = [];
    for (const sheet of sheets) {
      const chKey = CHANNEL_KEYS.find((k) => CHANNELS[k].sheetMatch.test(sheet.sheetName));
      if (!chKey) {
        unassignedSheets.push(sheet.sheetName);
        continue;
      }
      if (!assigned[chKey]) assigned[chKey] = [];
      assigned[chKey].push(...sheet.rows);
    }

    const summaryParts = [];
    for (const chKey of Object.keys(assigned)) {
      store.chi_phi_raw_uploads[chKey].push({
        id: nextId(store, "chi_phi_raw_uploads_seq") || Date.now(),
        uploaded_at: new Date().toISOString(),
        file_name: req.file.originalname,
        rows: assigned[chKey],
      });
      summaryParts.push(`${CHANNELS[chKey].label}: ${assigned[chKey].length} dong`);
    }
    if (summaryParts.length === 0) {
      // Chi Nhan, 2026-07-29: "tôi nạp sao kê chi tk này không dc" -- thong
      // bao loi CU hardcode chi nhac "8651"/"9997" (2 kenh dau tien), da CU
      // tu khi them VP58888/BIDV8681 -- gio liet ke DUNG toan bo kenh + tu
      // khoa can co trong ten sheet, lay THANG tu CHANNELS de khong bao gio
      // lech nua khi them kenh moi ve sau.
      const channelHints = CHANNEL_KEYS.map(
        (k) => `${CHANNELS[k].label} (tên sheet cần chứa "${String(CHANNELS[k].sheetMatch).replace(/^\/|\/[a-z]*$/g, "")}")`
      ).join(", ");
      throw new Error(
        `File co ${sheets.length} sheet du lieu (${sheets.map((s) => s.sheetName).join(", ")}) nhung khong sheet nao khop ten voi cac tai khoan chi phi dang co: ${channelHints}. Doi ten sheet trong file Excel cho khop roi tai lai.`
      );
    }
    save(store);

    let msg = `Da nap "${req.file.originalname}": ${summaryParts.join(", ")}.${UPDATED_NOTE}`;
    if (unassignedSheets.length > 0) {
      msg += ` CANH BAO: bo qua sheet "${unassignedSheets.join(", ")}" -- ten sheet khong khop tai khoan nao.`;
    }
    res.redirect("/doi-soat/chi-phi?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/chi-phi/upload/:channel/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const channelKey = req.params.channel;
  if (!CHANNELS[channelKey]) return res.redirect("/doi-soat/chi-phi");
  store.chi_phi_raw_uploads[channelKey] = store.chi_phi_raw_uploads[channelKey].filter(
    (u) => String(u.id) !== req.params.id
  );
  save(store);
  res.redirect("/doi-soat/chi-phi");
});

// ---------- Uploat lai danh sach NCC (Ma NCC/Ten NCC/MST) -- thay the ban 884
// NCC "KH cu" di kem app khi Luyen co ban moi hon (them NCC, sua ten, doi ma...) ----------
router.post("/doi-soat/chi-phi/upload-ncc", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const list = parseNccWorkbook(req.file.buffer);
    store.chi_phi_ncc_list = list;
    store.chi_phi_ncc_meta = {
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      count: list.length,
    };
    save(store);
    res.redirect(
      "/doi-soat/chi-phi?success=" +
        encodeURIComponent(`Da cap nhat danh sach NCC: ${list.length} nha cung cap.${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Tai danh sach hoa don NCC (sheet "HDDV KH CU") -- de doi chieu
// Ten NCC + So tien va tu dien So hoa don/Ngay hoa don vao file xuat Misa.
// Thay the toan bo ban cu moi lan tai len (giong cach lam voi danh sach NCC),
// vi day la "ban day du hien tai" Luyen xuat lai tu he thong hoa don dien tu,
// khong phai tang du lieu can gop theo tung lan. ----------
router.post("/doi-soat/chi-phi/upload-invoice", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const { sheetName, list } = parseHddvInvoiceWorkbook(req.file.buffer);
    store.chi_phi_invoice_list = list;
    store.chi_phi_invoice_meta = {
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      count: list.length,
      sheetName,
    };
    save(store);
    res.redirect(
      "/doi-soat/chi-phi?success=" +
        encodeURIComponent(`Da nap danh sach hoa don NCC (sheet "${sheetName}"): ${list.length} hoa don.${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Tai bang lenh chi UNC -- doi chieu Ten NCC + So tien de lay Dien
// giai sach hon dong sao ke ngan hang thuc te. Thay the toan bo ban cu moi
// lan tai len, cung ly do nhu tren. ----------
router.post("/doi-soat/chi-phi/upload-unc", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const { sheetName, list } = parseUncWorkbook(req.file.buffer);
    store.chi_phi_unc_list = list;
    store.chi_phi_unc_meta = {
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      count: list.length,
      sheetName,
    };
    save(store);
    res.redirect(
      "/doi-soat/chi-phi?success=" +
        encodeURIComponent(`Da nap bang lenh chi UNC (sheet "${sheetName}"): ${list.length} dong.${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- TK No theo tung NCC (Ten doi ung) -- Luyen tu go, khong co gia tri mac dinh ep buoc ----------
router.post("/doi-soat/chi-phi/vendor-tk", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const body = req.body || {};
  for (const [key, val] of Object.entries(body)) {
    if (key.startsWith("tkno_")) {
      const vendor = decodeURIComponent(key.slice("tkno_".length));
      if (val && val.trim()) store.chi_phi_vendor_tk_map[vendor] = val.trim();
      else delete store.chi_phi_vendor_tk_map[vendor];
    }
  }
  save(store);
  res.redirect("/doi-soat/chi-phi?success=" + encodeURIComponent("Da luu TK No theo NCC."));
});

// ---------- Gan/sua Ma cong trinh cho 1 dong cu the (khi tu dong khong xac dinh duoc hoac xac dinh sai) ----------
router.post("/doi-soat/chi-phi/gian-override", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const { key, maCongTrinh } = req.body;
    if (!key) throw new Error("Thieu dong can gan.");
    if (maCongTrinh && maCongTrinh.trim()) {
      store.chi_phi_gian_override[key] = maCongTrinh.trim();
    } else {
      store.chi_phi_gian_override[key] = ""; // luu tuong minh "khong co gian" (khac voi chua xet)
    }
    save(store);
    res.redirect("/doi-soat/chi-phi?success=" + encodeURIComponent("Da gan Ma cong trinh cho dong nay."));
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Gan tay Ma NCC theo tung NCC (Ten doi ung) khi khong tu khop duoc
// (khong co trong danh sach, hoac trung ten voi nhieu NCC) ----------
router.post("/doi-soat/chi-phi/vendor-ncc", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const { vendor, maNCC } = req.body;
    if (!vendor) throw new Error("Thieu NCC can gan.");
    if (maNCC && maNCC.trim()) {
      store.chi_phi_vendor_ncc_map[vendor] = maNCC.trim();
    } else {
      delete store.chi_phi_vendor_ncc_map[vendor];
    }
    save(store);
    res.redirect("/doi-soat/chi-phi?success=" + encodeURIComponent("Da luu Ma NCC."));
  } catch (e) {
    res.redirect("/doi-soat/chi-phi?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Export Misa: dung DUNG 57 cot cua mau that "MISAKHCU" (Phieu chi
// tien gui de nhap AMIS Accounting), lay tu file "Chi phi KH NCC T7 2026.xlsx"
// (2026-07-16). TK Co co dinh la 1121 (tien gui ngan hang VND) theo mau that,
// khac voi 112 dung tam truoc do. Cot Ma NCC/Ten NCC/Ma so thue NCC + Ma doi
// tuong duoc dien tu bang tra chi khi khop DUNG 1 NCC (utils/chiphiReconcile.js
// matchNccForVendor) -- khong khop thi de trong cho Luyen dien tay, KHONG doan.
// Doi soat hoa don (2026-07-16): danh sach hoa don HDDV + bang lenh chi UNC
// (muc 1b/1c) duoc doi chieu theo Ten/MST NCC + So tien -- khop duoc thi dien
// day du cot hoa don + lay Dien giai sach tu UNC; khong khop hoa don nao thi
// thu rut so HD ngay tu dien giai; khong co gi ca thi de trong toan bo cot
// hoa don nhung VAN xuat dong do (khong bo qua) theo yeu cau cua Luyen.
function buildExportRows(lines, startNo, bankAccount, bankFullName, chiaSeTKSet) {
  let seq = startNo;
  return lines
    .filter((l) => l.debit > 0)
    .sort((a, b) => (a.date > b.date ? 1 : -1))
    .map((l) => {
      const soCt = "UNC" + String(seq).padStart(4, "0");
      seq++;
      const ngayDmy = isoToDmy(l.date);
      // Uu tien Dien giai sach tu bang lenh chi UNC (khop Ten NCC + So tien)
      // hon dong sao ke ngan hang thuc te -- UNC thuong da co san "theo so HD
      // <n>" ro rang, con sao ke nhieu khi chi ghi ma but toan/FT... khong
      // doc duoc.
      const dienGiai = (l.uncDienGiai || l.description || "").replace(/\s+/g, " ").trim().slice(0, 250);
      const maDoiTuong = l.maNCC || "";
      const tenDoiTuong = l.tenNCC || l.vendor || "(chưa xác định NCC)";
      // Hoa don: khop duoc voi danh sach hoa don HDDV (Ten/MST NCC + So tien)
      // thi dien day du; khong khop duoc nhung dien giai co san "HD <so>" thi
      // lay rieng So hoa don (khong doan ngay/ky hieu); khong co gi ca thi de
      // trong toan bo cot hoa don -- van xuat dong nay nhu binh thuong.
      const inv = l.invoiceMatch;
      const soHoaDon = inv ? inv.soHoaDon : l.hdNumberFallback || "";
      const ngayHoaDon = inv && inv.ngayLap ? isoToDmy(inv.ngayLap) : "";
      const coHoaDon = inv ? "Có" : "Không";
      return {
        "Phương thức thanh toán": "Ủy nhiệm chi",
        "Ngày hạch toán (*)": ngayDmy,
        "Ngày chứng từ (*)": ngayDmy,
        "Số chứng từ (*)": soCt,
        "Lý do chi": "Chi trả nhà cung cấp / đối tác",
        "Là UNC chuyển tiền theo lô": "",
        "Nội dung thanh toán": dienGiai,
        "Số tài khoản chi": bankAccount,
        "Tên ngân hàng chi": bankFullName,
        "Mã đối tượng": maDoiTuong,
        "Tên đối tượng": tenDoiTuong,
        "Địa chỉ": "",
        "Số tài khoản nhận": l.taiKhoanDoiUng || "",
        "Tên ngân hàng nhận": l.nganHangDoiUng || "",
        "Người lĩnh tiền": "",
        "Số CMND": "",
        "Ngày cấp CMND": "",
        "Nơi cấp CMND": "",
        "Mã nhân viên": "",
        "Diễn giải (hạch toán)": dienGiai,
        // TK No: uu tien ban Luyen go tay (vendorTkMap); neu chua co thi tu
        // suy tu Hinh Thuc Hop Tac cua gian: "chia se" -> 1388, con lai -> 331.
        "TK Nợ (*)": l.tkNo || getGianTK(l.finalGian, chiaSeTKSet),
        "TK Có (*)": "1121",
        "Số tiền": l.debit,
        "Tên người hưởng": l.vendor || "",
        "TK hưởng": l.taiKhoanDoiUng || "",
        "Tên NH thụ hưởng": l.nganHangDoiUng || "",
        "Tên chi nhánh NH thụ hưởng": "",
        "Mã đối tượng (hạch toán)": maDoiTuong,
        "Số khế ước đi vay": "",
        "Số khế ước cho vay": "",
        "Mã khoản mục chi phí": "",
        "Nghiệp vụ": "",
        "Mã đơn vị": "",
        "Mã đối tượng THCP": "",
        "Mã công trình": l.finalGian || "",
        "Số đơn đặt hàng": "",
        "Số đơn mua hàng": "",
        "Số hợp đồng mua": "",
        "Số hợp đồng bán": "",
        "Mã thống kê": "",
        "CP không hợp lý": "",
        "Hạch toán gộp nhiều hóa đơn": "",
        "Diễn giải thuế": "",
        "Có hóa đơn": coHoaDon,
        "Giá trị HHDV chưa thuế": inv ? inv.tongTienChuaThue : "",
        "% thuế GTGT": inv && inv.tongTienChuaThue ? Math.round((inv.tongTienThue / inv.tongTienChuaThue) * 100) : "",
        "% thuế suất KHAC": "",
        "Tiền thuế GTGT": inv ? inv.tongTienThue : "",
        "TK thuế GTGT": "",
        "Ngày hóa đơn": ngayHoaDon,
        "Số hóa đơn": soHoaDon,
        "Mẫu số HĐ": inv ? inv.kyHieuMauSo : "",
        "Ký hiệu HĐ": inv ? inv.kyHieuHoaDon : "",
        "Nhóm HHDV mua vào": "",
        "Mã NCC": l.maNCC || "",
        "Tên NCC": l.tenNCC || "",
        "Mã số thuế NCC": l.mstNCC || "",
      };
    });
}

router.get("/doi-soat/chi-phi/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const channelKey = req.query.channel;
  if (!CHANNELS[channelKey]) return res.status(400).send("Kenh khong hop le.");

  const bank = store.banks.find((b) => b.name === CHANNELS[channelKey].bankName);
  if (!bank) return res.status(400).send(`Chua co ngan hang "${CHANNELS[channelKey].bankName}" trong he thong.`);

  const built = buildChannelChiPhi(store, channelKey);
  if (built.error) return res.status(400).send(built.error);

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;

  const chiaSeTKSetExport = buildChiaSeTKSet(store.hoa_don_dau_vao_gian_list);
  const rows = buildExportRows(built.lines, startNo, bank.account_number, `Ngân hàng ${bank.bank_name}`, chiaSeTKSetExport);

  const sheetLabel = CHANNELS[channelKey].company === "kh_moi" ? "MISAKHMOI" : "MISAKHCU";
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, sheetLabel);
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=${sheetLabel}-${channelKey}.xlsx`);
  res.send(buf);
});

// Luyen, 2026-08-09: "cho tôi thêm 1 cái nữa là đối chiếu tài khoản chi dưới
// cái tài khoản chi nhá trong đó sẽ chia ra làm 2 ngân hàng dựa vào sao kê á
// lấy ra 2 ngân hàng đó cho tôi và đối ứng chi của ngân hàng đó với lại dựa
// vào chi phí hay hóa đơn đầu vào hay diễn giải á để lấy ra cho tôi số hóa
// đơn với gian đó dựa vào chi phí á số tiền tên đối tác diễn giải"
// Trang nay: doc truc tiep tu store.transactions (sao ke song), lay tat ca GD
// type="chi" cua 2 tai khoan chi cua cong ty dang xem, cross-ref voi
// store.chi_phi (uu tien bankTxId chinh xac, fallback so tien +-1000 + ngay
// +-7d) de lay soHoaDon / gian / ncc / daHachToan. Khac trang "Tai khoan Chi"
// hien co (upload file raw, xuat MISA) -- trang nay de xem nhanh tung giao
// dich va doi chieu chi phi da nhap.
router.get("/doi-soat/chi-phi-saoke", (req, res) => {
  try {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);

  // Tai khoan chi cua cong ty dang xem
  const companyChannelKeys = CHANNEL_KEYS.filter((k) => CHANNELS[k].company === activeCompany);
  const chiBankNameSet = new Set(companyChannelKeys.map((k) => CHANNELS[k].bankName));

  const banksById = {};
  (store.banks || []).forEach((b) => { banksById[b.id] = b; });

  const chiBankIds = new Set(
    (store.banks || []).filter((b) => chiBankNameSet.has(b.name)).map((b) => b.id)
  );

  // Danh sach thang co GD chi tren 2 tai khoan nay
  const monthSet = new Set();
  (store.transactions || []).forEach((t) => {
    if (t.type !== "chi" || !chiBankIds.has(t.bank_id)) return;
    const m = (t.date || "").slice(0, 7);
    if (m) monthSet.add(m);
  });
  const months = [...monthSet].sort().reverse();

  const selectedMonth = req.query.thang || months[0] || "";
  const selectedBankName = req.query.nganHang || "";

  // Xay dung TK lookup tu danh sach gian (Hinh Thuc Hop Tac)
  const chiaSeTKSet = buildChiaSeTKSet(store.hoa_don_dau_vao_gian_list);

  // Chi phi index: exact by bankTxId
  const chiPhiByBankTxId = {};
  // Also fuzzy list for this company (so tien + ngay)
  const chiPhiCompany = [];
  (store.chi_phi || []).forEach((r) => {
    if ((r.congTy || "kh_cu") !== activeCompany) return;
    if (r.bankTxId) chiPhiByBankTxId[r.bankTxId] = r;
    chiPhiCompany.push(r);
  });

  const AMOUNT_TOL = 1000;
  const DATE_WIN_MS = 7 * 86400000;
  function dateMs(d) { return new Date(d + "T00:00:00").getTime(); }

  // Loc GD chi
  const chiTxs = (store.transactions || []).filter((t) => {
    if (t.type !== "chi" || !chiBankIds.has(t.bank_id)) return false;
    if (selectedMonth && (t.date || "").slice(0, 7) !== selectedMonth) return false;
    const bankName = (banksById[t.bank_id] || {}).name || "";
    if (selectedBankName && bankName !== selectedBankName) return false;
    return true;
  });

  const rows = chiTxs.map((t) => {
    const bank = banksById[t.bank_id] || {};
    const bankName = bank.name || "";
    const channelKey = companyChannelKeys.find((k) => CHANNELS[k].bankName === bankName);
    const bankLabel = channelKey ? CHANNELS[channelKey].label : bankName;

    // Exact match by bankTxId
    let cp = chiPhiByBankTxId[t.id] || null;
    let matchType = cp ? "exact" : "";

    // Fuzzy match: so tien +-1000 + ngay +-7d + NCC phai co tu nao khop tenDoiUng
    // Them dieu kien NCC de tranh ghep nham GD trung so tien nhung khac doi
    // tuong hoan toan (vd tien nop kho bac trung so tien voi hoa don do dau xe).
    if (!cp) {
      const tMs = dateMs(t.date);
      const tSearchText = normText((t.tenDoiUng || "") + " " + (t.description || ""));
      cp = chiPhiCompany.find((r) => {
        if (!r.ngay || !r.soTien) return false;
        if (Math.abs((r.soTien || 0) - t.amount) > AMOUNT_TOL) return false;
        if (Math.abs(dateMs(r.ngay) - tMs) > DATE_WIN_MS) return false;
        // Neu chi phi co NCC va GD co noi dung, yeu cau NCC phai co it nhat
        // 1 tu >= 5 ky tu (bo cac tu qua ngan/pho bien) khop voi tenDoiUng
        // hoac dien giai cua GD ngan hang. Neu NCC khong co tu nao du dai,
        // tha loi dieu kien nay (fallback ve so tien + ngay nhu cu).
        if (r.ncc && (t.tenDoiUng || t.description)) {
          const nccWords = normText(r.ncc).split(/\s+/).filter((w) => w.length >= 5);
          if (nccWords.length > 0 && !nccWords.some((w) => tSearchText.includes(w))) return false;
        }
        return true;
      }) || null;
      if (cp) matchType = "fuzzy";
    }

    return {
      id: t.id,
      date: t.date,
      bankName,
      bankLabel,
      tenDoiUng: t.tenDoiUng || "",
      description: t.description || "",
      amount: t.amount,
      soHoaDon: (cp && cp.soHoaDon) || "",
      gian: (cp && cp.gian) || "",
      ncc: (cp && cp.ncc) || "",
      daHachToan: cp ? !!cp.daHachToan : false,
      chiPhiId: (cp && cp.id) || "",
      matchType,
      tk: cp ? getGianTK(cp.gian, chiaSeTKSet) : "",
    };
  }).sort((a, b) => a.date.localeCompare(b.date));

  const { COMPANIES } = require("../utils/companies");
  const viewData = {
    COMPANIES,
    activeCompany,
    userName: req.session.userName,
    isAdmin: req.session.isAdmin,
    months,
    selectedMonth,
    selectedBankName,
    chiBankNames: [...chiBankNameSet],
    rows,
    totalRows: rows.length,
    totalAmount: rows.reduce((s, r) => s + r.amount, 0),
    matchedCount: rows.filter((r) => r.matchType).length,
    successMsg: req.query.success || "",
    errorMsg: req.query.error || "",
  };
  res.render("doisoat-chiphi-saoke", viewData);
  } catch (e) {
    console.error("[doi-soat/chi-phi-saoke] ERROR:", e.message);
    res.status(500).send("Lỗi: " + e.message);
  }
});

// Huy lien ket bankTxId sai (GD ngan hang bi ghep nham voi dong chi phi do
// fuzzy match cu chua co kiem tra NCC). Xoa bankTxId + daHachToan tren dong
// chi_phi de cho phep ghep lai dung sau khi sua.
router.post("/doi-soat/chi-phi-saoke/unlink-banktxid", requireDataEntry, (req, res) => {
  try {
    const store = load();
    const { chiPhiId } = req.body;
    if (!chiPhiId) throw new Error("Thieu chiPhiId.");
    const r = (store.chi_phi || []).find((x) => String(x.id) === String(chiPhiId));
    if (!r) throw new Error("Khong tim thay dong chi phi ID " + chiPhiId);
    const oldTxId = r.bankTxId;
    r.bankTxId = "";
    r.daHachToan = false;
    save(store);
    const back = req.headers.referer || "/doi-soat/chi-phi-saoke";
    res.redirect(back + (back.includes("?") ? "&" : "?") + "success=" + encodeURIComponent("Da huy lien ket GD " + (oldTxId || "") + " khoi dong chi phi nay."));
  } catch (e) {
    const back = req.headers.referer || "/doi-soat/chi-phi-saoke";
    res.redirect(back + (back.includes("?") ? "&" : "?") + "error=" + encodeURIComponent(e.message));
  }
});

module.exports = router;
// Luyen, 2026-07-19: xuat them vai ham/hang so noi bo (khong doi router
// chinh) de trang moi "Cong No NCC" (routes/congno-ncc.js) tai su dung DUNG
// logic khop NCC/hoa don/UNC da co o day, khong viet lai/tinh lech so.
module.exports.CHANNELS = CHANNELS;
module.exports.CHANNEL_KEYS = CHANNEL_KEYS;
module.exports.ensureShape = ensureShape;
module.exports.activeNccList = activeNccList;
module.exports.buildChannelChiPhi = buildChannelChiPhi;
module.exports.buildRevenueStatusMap = buildRevenueStatusMap;
