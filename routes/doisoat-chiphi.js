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
    error: req.query.error || null,
    success: req.query.success || null,
  });
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
function buildExportRows(lines, startNo, bankAccount, bankFullName) {
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
        "TK Nợ (*)": l.tkNo || "",
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

  const rows = buildExportRows(built.lines, startNo, bank.account_number, `Ngân hàng ${bank.bank_name}`);

  const sheetLabel = CHANNELS[channelKey].company === "kh_moi" ? "MISAKHMOI" : "MISAKHCU";
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, sheetLabel);
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=${sheetLabel}-${channelKey}.xlsx`);
  res.send(buf);
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
