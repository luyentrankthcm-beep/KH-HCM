// Luyen, 2026-08-24: "thêm cho tôi 1 mục nữa là hồ sơ" -- trang Hồ Sơ Hóa
// Đơn NCC: tổng hợp hóa đơn đầu vào theo NCC, link thẳng về file PDF trên
// Google Drive. Mỗi hóa đơn có thêm: tên đầy đủ NCC, tổng tiền, nội dung
// tóm tắt, hồ sơ liên quan.
const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { getCompany, COMPANIES } = require("../utils/companies");
const driveApi = require("../utils/driveApi");
const { Packer } = require("docx");
const AdmZip = require("adm-zip");
const { buildBienBanGiaoNhan, buildBienBanNghiemThu, buildBangKeHoaDon } = require("../utils/chungTuNccDoc");
const { extractHoaDonPdfInfo } = require("../utils/hoaDonPdf");

// Nhan, 2026-09-17: tai file PDF hoa don tren Drive va trich xuat dia chi/
// dien thoai Ben A + bang hang hoa chi tiet (xem utils/hoaDonPdf.js). Loi bat
// ky buoc nao (Drive tra loi, khong parse duoc PDF, bang hang hoa khong
// khop...) deu tra ve null, KHONG throw -- de cho goi noi lam viec fallback
// ve 1 dong noi dung rut gon nhu cu, khong bao gio lam sap export.
async function fetchInvoicePdfExtract(store, driveId) {
  if (!driveId) return null;
  try {
    const gmailApi = require("../utils/gmailApi");
    const accessToken = await gmailApi.getValidAccessToken(store);
    const dlResp = await fetch(`https://www.googleapis.com/drive/v3/files/${driveId}?alt=media`, {
      headers: { Authorization: "Bearer " + accessToken },
    });
    if (!dlResp.ok) return null;
    const buf = Buffer.from(await dlResp.arrayBuffer());
    const pdfParse = require("pdf-parse");
    const data = await pdfParse(buf);
    return extractHoaDonPdfInfo(data.text || "");
  } catch (e) {
    return null;
  }
}

const uploadMem = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

const router = express.Router();

// Luyen 2026-08-31: bulk-upsert tu Google Drive -- phai dang ky TRUOC router.use(requireLogin)
// de X-Internal-Key header bypass duoc auth ma khong can session.
const INTERNAL_SYNC_KEY = process.env.INTERNAL_SYNC_KEY || "";
// API JSON cho v2 service doc data ma khong can session (dung X-Internal-Key)
router.get("/api/ho-so/hoa-don-ncc", express.json(), (req, res) => {
  const keyOk = INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY;
  if (!keyOk) return res.status(401).json({ error: "Unauthorized" });
  const store = load();
  ensureHoSo(store);
  res.json({
    ho_so_hoa_don: store.ho_so_hoa_don || [],
    hoa_don_dau_vao: (store.hoa_don_dau_vao || []).map((r) => ({
      soHoaDon: r.soHoaDon, tenNCC: r.tenNCC, soTien: r.soTien,
      dienGiai: r.dienGiai || r.tenHangHoaMisa || "",
    })),
  });
});

// Bulk-upsert Phi Cang -- check theo driveId de tranh duplicate
router.post("/api/ho-so/phi-cang/bulk-upsert", express.json(), (req, res) => {
  const keyOk = INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY;
  if (!keyOk) return res.status(401).json({ error: "Unauthorized" });
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  const records = Array.isArray(req.body.records) ? req.body.records : [];
  const existingIds = new Set(store.ho_so_phi_cang.map(r => r.driveId).filter(Boolean));
  let added = 0, skipped = 0;
  records.forEach(r => {
    if (r.driveId && existingIds.has(r.driveId)) { skipped++; return; }
    store.ho_so_phi_cang.push({
      id: nextId(store),
      ngay: (r.ngay||"").trim(), soHoaDon: (r.soHoaDon||"").trim(),
      loaiPhi: (r.loaiPhi||"").trim(), soTien: (r.soTien||"").trim(),
      linkFile: r.driveId ? `https://drive.google.com/file/d/${r.driveId}/view` : (r.linkFile||"").trim(),
      driveId: (r.driveId||"").trim(), ghiChu: (r.ghiChu||"").trim(),
      createdAt: new Date().toISOString(),
    });
    if (r.driveId) existingIds.add(r.driveId);
    added++;
  });
  save(store);
  res.json({ success: true, added, skipped });
});

// API tong hop cho v2 -- tra ve du lieu ca 5 trang Ho So trong 1 request
router.get("/api/ho-so/all", express.json(), (req, res) => {
  const keyOk = INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY;
  if (!keyOk) return res.status(401).json({ error: "Unauthorized" });
  const store = load();
  ensureHoSo(store);
  res.json({
    ho_so_hoa_don: store.ho_so_hoa_don || [],
    hoa_don_dau_vao: (store.hoa_don_dau_vao || []).map((r) => ({
      soHoaDon: r.soHoaDon, tenNCC: r.tenNCC, soTien: r.soTien,
      dienGiai: r.dienGiai || r.tenHangHoaMisa || "",
    })),
    ho_so_phi_cang: store.ho_so_phi_cang || [],
    dtcs_chiase_diem: store.dtcs_chiase_diem || [],
    dtcs_chiase_thang: store.dtcs_chiase_thang || [],
    ho_so_tien_thue: store.ho_so_tien_thue || [],
    ho_so_chung_tu: store.ho_so_chung_tu || [],
  });
});

router.post("/ho-so/hoa-don-ncc/bulk-upsert", express.json(), (req, res) => {
  const keyOk = INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY;
  if (!keyOk && !(req.session && req.session.role === "admin")) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  const store = load();
  ensureHoSo(store);
  const files = req.body && Array.isArray(req.body.files) ? req.body.files : [];
  const company = req.body && req.body.company ? String(req.body.company).trim() : "kh_cu";
  const existingIds = new Set(store.ho_so_hoa_don.map((r) => r.driveId));
  let added = 0, skipped = 0;
  files.forEach((f) => {
    const driveId  = String(f.driveId  || "").trim();
    const fileName = String(f.fileName || "").trim();
    if (!driveId || !fileName) { skipped++; return; }
    if (existingIds.has(driveId)) { skipped++; return; }
    store.ho_so_hoa_don.push({ driveId, fileName, company, addedAt: new Date().toISOString() });
    existingIds.add(driveId);
    added++;
  });
  save(store);
  res.json({ success: true, added, skipped });
});

// Internal endpoint: add months to a diem by maCongTrinh (no session required)
router.post("/ho-so/dtcs-chiase/diem/bulk-add-thang", express.json(), (req, res) => {
  const keyOk = INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY;
  if (!keyOk) return res.status(401).json({ error: "Unauthorized" });
  const { maCongTrinh, thangList } = req.body;
  if (!maCongTrinh || !Array.isArray(thangList)) return res.status(400).json({ error: "maCongTrinh and thangList required" });
  const store = load();
  if (!store.dtcs_chiase_diem) store.dtcs_chiase_diem = [];
  if (!store.dtcs_chiase_thang) store.dtcs_chiase_thang = [];
  const diem = store.dtcs_chiase_diem.find(d => d.maCongTrinh === maCongTrinh);
  if (!diem) return res.status(404).json({ error: `Diem not found: ${maCongTrinh}` });
  let added = 0, skipped = 0;
  for (const t of thangList) {
    const existing = store.dtcs_chiase_thang.find(x => String(x.diemId) === String(diem.id) && x.thang === t.thang);
    if (existing) { skipped++; continue; }
    store.dtcs_chiase_thang.push({ id: nextId(store), diemId: diem.id, ...t });
    added++;
  }
  save(store);
  res.json({ success: true, diemId: diem.id, added, skipped });
});

router.use(requireLogin);

function parseInvoiceFileName(fileName) {
  const base = fileName.replace(/\.pdf$/i, "");
  const parts = base.split("_");
  const ncc = parts[0] || "";
  const soHoaDon = parts[1] || "";
  const dateStr = parts[2] || ""; // DD-MM-YYYY
  let ngay = "";
  if (dateStr && dateStr.match(/^\d{2}-\d{2}-\d{4}$/)) {
    const [d, m, y] = dateStr.split("-");
    ngay = `${y}-${m}-${d}`;
  }
  const thang = ngay ? ngay.slice(0, 7) : "";
  return { ncc, soHoaDon, ngay, thang };
}

function ensureHoSo(store) {
  if (!store.ho_so_hoa_don) store.ho_so_hoa_don = [];
  if (!store.ho_so_chung_tu) store.ho_so_chung_tu = [];
}

// Bo so 0 dau de doi chieu so HD khong phu thuoc format (2072 = 00002072)
function normSoHD(s) {
  const str = String(s || "").trim();
  return str.replace(/^0+/, "") || str;
}

// Luyen, 2026-08-25: tu dong dien Ten NCC va Tong tien tu bang hoa_don_dau_vao
// (doi chieu theo so hoa don tu ten file). Luu ca key goc lan key bo so 0 dau.
//
// Nhan, 2026-09-17: BUG -- so hoa don (vd "107") la so NHO, rat de TRUNG giua
// nhieu NCC khac nhau khong lien quan gi (vd hoa don 107 cua "Ngoc Ha Foods",
// "Khanh Thao" VA "CONG TY TNHH CNM" trung nhau). Ban cu gop CA 3 lai thanh 1
// (chi lay tenNCC cua ban ghi dau tien gap, CONG DON ca 3 tongTien lai) ->
// hien sai ten NCC + sai so tien cho hoa don that. Sua: nhom theo CA (so hoa
// don, ten NCC da chuan hoa) de khong gop nham; lookup[soHoaDon] tra ve MANG
// cac NCC trung so, de enrichRow tu chon dung theo ten NCC cua chinh file do
// (xem pickHddvCandidate).
function buildHddvLookup(store, activeCompany) {
  const lookup = {}; // soHoaDon (raw + normalized) -> [{ tenNCC, tongTien, dienGiai }, ...]
  const grouped = {}; // "soHoaDon||tenNCCChuanHoa" -> candidate (de gop nhieu dong CUNG 1 hoa don that)
  (store.hoa_don_dau_vao || []).forEach((r) => {
    if (activeCompany && r.congTy && r.congTy !== activeCompany) return;
    const no = String(r.soHoaDon || "").trim();
    if (!no) return;
    const noNorm = normSoHD(no);
    const nccNorm = normForMatch(r.tenNCC || "");
    // Nhan, 2026-09-17: bug cu -- khi so HD KHONG co so 0 dau (vd "107", rat
    // pho bien), no===noNorm nen forEach lap 2 LAN TREN CUNG 1 KEY, cong
    // tongTien HAI LAN cho cung 1 hoa don (sai gap doi so tien). Dung Set de
    // chi giu key duy nhat.
    [...new Set([no, noNorm])].forEach((k) => {
      const gkey = k + "||" + nccNorm;
      if (!grouped[gkey]) {
        grouped[gkey] = { tenNCC: r.tenNCC || "", tongTien: 0, dienGiai: r.dienGiai || r.tenHangHoaMisa || "" };
        if (!lookup[k]) lookup[k] = [];
        lookup[k].push(grouped[gkey]);
      }
      grouped[gkey].tongTien += r.soTien || 0;
      if (!grouped[gkey].dienGiai && (r.dienGiai || r.tenHangHoaMisa))
        grouped[gkey].dienGiai = r.dienGiai || r.tenHangHoaMisa || "";
    });
  });
  return lookup;
}

// Trong so cac ban ghi hoa_don_dau_vao trung SO HOA DON, chon dung ban ghi
// theo ten NCC lay tu chinh ten file (vd "KhanhThao"). Chi 1 ung vien thi lay
// luon (khong can khop ten, giu hanh vi cu cho truong hop khong trung). Nhieu
// ung vien ma KHONG chac chan cai nao dung thi tra null (KHONG tu dien bua --
// tha de trong con hon dien SAI ten/so tien cua NCC khac).
function pickHddvCandidate(candidates, nccShort) {
  if (!candidates || !candidates.length) return null;
  const shortNorm = normForMatch(nccShort);
  if (shortNorm.length >= 3) {
    const match = candidates.find((c) => {
      const cNorm = normForMatch(c.tenNCC);
      return cNorm.length > 0 && (cNorm.includes(shortNorm) || shortNorm.includes(cNorm));
    });
    if (match) return match;
  }
  // Nhan, 2026-09-17: phat hien them qua case CGV -- CHI 1 ung vien trung so
  // HD KHONG co nghia la dung ("CGV" trung so HD voi hoa don cua NCC hoan
  // toan khac ten "Kubo"/"HANG GIA DUNG TONG HOP" khong lien quan gi rap
  // chieu phim). Truoc day cho 1 ung vien la nhan luon -- gio BAT BUOC ten
  // phai lien quan (dung nhu truong hop nhieu ung vien), du chi co 1 cai,
  // de tranh dien nham hoan toan mot NCC khac chi vi trung so hoa don.
  return null;
}

// Fallback: dò tên đầy đủ NCC từ danh sách NCC (chi_phi_ncc_list) khi số HĐ
// không khớp trong hoa_don_dau_vao. Chuẩn hoá bằng cách bỏ ký tự đặc biệt,
// lowercase rồi kiểm tra nccShort có nằm trong tên NCC không.
function normForMatch(s) {
  // Nhan, 2026-09-17: bug cu -- chi lowercase roi xoa ky tu khong phai a-z0-9
  // se XOA LUON ca chu co dau (vd "ả" khong thuoc [a-z0-9] nen bi xoa thang,
  // KHONG chuyen ve "a"), lam 2 ten khac nhau bi trung lam do trung cac phan
  // con lai sau khi xoa dau ngau nhien (vd "Khanh Thảo" -> "khanhtho" thay vi
  // "khanhthao"), gay khop NHAM hop dong. Chuan hoa dung cach: NFD + bo dau
  // (combining marks) + doi "đ" -> "d" TRUOC khi xoa ky tu con lai.
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9]/g, "");
}
function lookupNccNameFallback(nccShort, store) {
  const shortNorm = normForMatch(nccShort);
  if (shortNorm.length < 3) return "";
  const allLists = [
    ...(store.chi_phi_ncc_list_cu || []),
    ...(store.chi_phi_ncc_list_moi || []),
    ...(store.chi_phi_ncc_list || []),
    ...(store.danh_muc_ma_nha_cung_cap_cu || []).map((r) => ({ tenNCC: r.ten })),
    ...(store.danh_muc_ma_nha_cung_cap_moi || []).map((r) => ({ tenNCC: r.ten })),
    ...(store.danh_muc_ma_nha_cung_cap || []).map((r) => ({ tenNCC: r.ten })),
  ];
  for (const rec of allLists) {
    const name = rec.tenNCC || "";
    if (!name) continue;
    if (normForMatch(name).includes(shortNorm)) return name;
  }
  return "";
}

function enrichRow(r, hddvLookup, store) {
  const parsed = parseInvoiceFileName(r.fileName || "");
  const candidates = hddvLookup && parsed.soHoaDon
    ? (hddvLookup[parsed.soHoaDon] || hddvLookup[normSoHD(parsed.soHoaDon)])
    : null;
  const inv = candidates ? pickHddvCandidate(candidates, parsed.ncc) : null;
  const autoTenFromHD = inv ? inv.tenNCC : "";
  const autoTenFromNCC = !autoTenFromHD ? lookupNccNameFallback(parsed.ncc, store) : "";
  const autoTen = autoTenFromHD || autoTenFromNCC;
  const autoTien = inv && inv.tongTien ? inv.tongTien.toLocaleString("vi-VN") + "đ" : "";
  return {
    ...r,
    nccShort: parsed.ncc,
    soHoaDon: parsed.soHoaDon,
    ngay: parsed.ngay,
    thang: parsed.thang,
    driveLink: r.driveId ? `https://drive.google.com/file/d/${r.driveId}/view` : "",
    tenDayDuNCC: r.tenDayDuNCC || autoTen,
    tongTien: r.tongTien || autoTien,
    noiDung: r.noiDung || (inv ? inv.dienGiai : "") || "",
    hoSoLienQuan: r.hoSoLienQuan || "",
    autoFilled: !r.tenDayDuNCC && !!autoTen,
  };
}

router.get("/ho-so/hoa-don-ncc", (req, res) => {
  const store = load();
  ensureHoSo(store);
  const activeCompany = getCompany(req);
  const thangFilter = req.query.thang || "";
  const hddvLookup = buildHddvLookup(store, activeCompany);

  // Loc theo cong ty: ban cu khong co truong company -> thuoc kh_cu
  const allForCompany = store.ho_so_hoa_don.filter(
    (r) => !r.company || r.company === activeCompany
  );

  let rows = allForCompany.map((r) => enrichRow(r, hddvLookup, store));
  if (thangFilter) rows = rows.filter((r) => r.thang === thangFilter);
  rows.sort((a, b) => {
    const nccCmp = a.nccShort.toLowerCase().localeCompare(b.nccShort.toLowerCase());
    return nccCmp !== 0 ? nccCmp : a.ngay.localeCompare(b.ngay);
  });

  const allThang = [...new Set(allForCompany.map((r) => {
    return parseInvoiceFileName(r.fileName || "").thang;
  }).filter(Boolean))].sort().reverse();

  // Nhom theo nccShort (ten ngan trong ten file) -- moi NCC 1 nhom du co
  // chi nhanh hay ten khac nhau. Ten day du NCC lay tu hoa don dau vao / ncc list.
  const groupMap = new Map();
  rows.forEach((r) => {
    const key = r.nccShort.toLowerCase();
    if (!groupMap.has(key)) {
      groupMap.set(key, { nccShort: r.nccShort, tenDayDuNCC: "", invoices: [], hopDong: [], chungTu: "" });
    }
    const g = groupMap.get(key);
    if (!g.tenDayDuNCC && r.tenDayDuNCC) g.tenDayDuNCC = r.tenDayDuNCC;
    g.invoices.push(r);
  });

  // Auto-match hop dong NCC tu phap_danh_hop_dong_ncc
  const hdNccList = store.phap_danh_hop_dong_ncc || [];
  const nccChungTuMap = store.ho_so_ncc_chungtu || {};
  groupMap.forEach((g, key) => {
    const shortNorm = normForMatch(g.nccShort);
    const fullNorm = normForMatch(g.tenDayDuNCC);
    g.hopDong = hdNccList.filter((hd) => {
      const hdShort = normForMatch(hd.tenNCC || "");
      const hdFull = normForMatch(hd.tenDayDuNCC || "");
      // Nhan, 2026-09-17: bug cu -- "".includes(x) luon false nhung x.includes("")
      // luon TRUE, nen hop dong thieu tenDayDuNCC/tenNCC (vd hop dong vua them
      // thu cong qua "+ Thêm hợp đồng NCC") se KHOP NHAM voi MOI NCC khac co
      // fullNorm/shortNorm >= do dai toi thieu. Them dieu kien do dai > 0 cho
      // ca 2 ve de chi khop khi CA HAI ben deu co ten thuc su.
      return (shortNorm.length >= 3 && hdShort.length > 0 && (hdShort.includes(shortNorm) || shortNorm.includes(hdShort))) ||
             (fullNorm.length >= 5 && hdFull.length > 0 && (hdFull.includes(fullNorm) || fullNorm.includes(hdFull)));
    });
    g.chungTu = (nccChungTuMap[key] || {}).chungTu || "";
    g.ghiChuNCC = (nccChungTuMap[key] || {}).ghiChu || "";
  });

  const groups = Array.from(groupMap.values()).sort((a, b) =>
    a.nccShort.toLowerCase().localeCompare(b.nccShort.toLowerCase())
  );

  res.render("ho-so-hoa-don", {
    userName: req.session.userName,
    rows,
    groups,
    thangFilter,
    allThang,
    total: rows.length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Luyen, 2026-09-05: trang v2 -- giao dien moi, sach hon, nhanh hon.
// Cung logic loc theo cong ty nhu route chinh, chi khac template render.
router.get("/v2/ho-so/hoa-don-ncc", (req, res) => {
  const store = load();
  ensureHoSo(store);
  const activeCompany = getCompany(req);
  const thangFilter = req.query.thang || "";
  const hddvLookup = buildHddvLookup(store, activeCompany);

  const allForCompany = store.ho_so_hoa_don.filter(
    (r) => !r.company || r.company === activeCompany
  );

  let rows = allForCompany.map((r) => enrichRow(r, hddvLookup, store));
  if (thangFilter) rows = rows.filter((r) => r.thang === thangFilter);
  rows.sort((a, b) => {
    const nccCmp = a.nccShort.toLowerCase().localeCompare(b.nccShort.toLowerCase());
    return nccCmp !== 0 ? nccCmp : a.ngay.localeCompare(b.ngay);
  });

  const allThang = [...new Set(allForCompany.map((r) =>
    parseInvoiceFileName(r.fileName || "").thang
  ).filter(Boolean))].sort().reverse();

  const groupMap = new Map();
  rows.forEach((r) => {
    const key = r.nccShort.toLowerCase();
    if (!groupMap.has(key)) {
      groupMap.set(key, { nccShort: r.nccShort, tenDayDuNCC: "", invoices: [] });
    }
    const g = groupMap.get(key);
    if (!g.tenDayDuNCC && r.tenDayDuNCC) g.tenDayDuNCC = r.tenDayDuNCC;
    g.invoices.push(r);
  });

  const groups = Array.from(groupMap.values()).sort((a, b) =>
    a.nccShort.toLowerCase().localeCompare(b.nccShort.toLowerCase())
  );

  res.render("v2-ho-so-hd", {
    userName: req.session.userName,
    rows,
    groups,
    thangFilter,
    allThang,
    total: rows.length,
    activeCompany,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// V2 placeholder tabs -- chuyen ve trang tuong ung ben v1 cho den khi co v2 rieng
router.get("/v2/ho-so/phi-cang", (req, res) => res.redirect("/ho-so/phi-cang-phu-quoc"));
router.get("/v2/ho-so/doanh-thu-chia-se", (req, res) => res.redirect("/ho-so/doanh-thu-chia-se"));
router.get("/v2/ho-so/tien-thue", (req, res) => res.redirect("/ho-so/tien-thue"));
router.get("/v2/ho-so/chung-tu", (req, res) => res.redirect("/ho-so/chung-tu"));

// Luu cac truong bo sung cho 1 hoa don
router.post("/ho-so/hoa-don-ncc/:driveId/sua", requireAdmin, (req, res) => {
  const store = load();
  ensureHoSo(store);
  const r = store.ho_so_hoa_don.find((x) => x.driveId === req.params.driveId);
  if (!r) return res.redirect("/ho-so/hoa-don-ncc?error=Không+tìm+thấy");
  const { tenDayDuNCC, tongTien, noiDung, hoSoLienQuan, thang } = req.body;
  r.tenDayDuNCC = (tenDayDuNCC || "").trim();
  r.tongTien = (tongTien || "").trim();
  r.noiDung = (noiDung || "").trim();
  r.hoSoLienQuan = (hoSoLienQuan || "").trim();
  save(store);
  const qs = thang ? "?thang=" + encodeURIComponent(thang) : "";
  res.redirect("/ho-so/hoa-don-ncc" + qs + "&success=Đã+lưu");
});

// Luu chung tu / ghi chu cap NCC (bien ban, bao gia, link, v.v.)
router.post("/ho-so/hoa-don-ncc/ncc/:nccKey/luu-chungtu", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_ncc_chungtu) store.ho_so_ncc_chungtu = {};
  const key = req.params.nccKey;
  store.ho_so_ncc_chungtu[key] = {
    chungTu: (req.body.chungTu || "").trim(),
    ghiChu: (req.body.ghiChu || "").trim(),
    updatedAt: new Date().toISOString(),
  };
  save(store);
  const qs = req.body.thang ? "?thang=" + encodeURIComponent(req.body.thang) : "";
  res.json({ ok: true });
});

// Nhan, 2026-09-17: "nếu đúng tên ncc á thì map còn không thì cho tôi nút
// thêm rồi tôi sẽ hiện ra 1 cái khung á nó sẽ có thông tin hợp đồng á thì đó
// tôi gắn link gg drive xong bạn lấy tên người đại diện ra từ hợp đồng" --
// khi 1 NCC o trang Ho So > Hoa Don NCC CHUA map duoc hop dong nao (khong
// trung ten voi phap_danh_hop_dong_ncc), cho phep them thang 1 ban ghi hop
// dong NCC (Dai dien/Chuc vu + link Google Drive) ngay tai modal, khong can
// qua trang Phap Danh > Hop Dong NCC rieng. Tra JSON de client cap nhat
// ngay khong can reload trang.
router.post("/ho-so/hoa-don-ncc/ncc/:nccKey/them-hop-dong", requireAdmin, (req, res) => {
  try {
    const store = load();
    if (!store.phap_danh_hop_dong_ncc) store.phap_danh_hop_dong_ncc = [];
    const activeCompany = getCompany(req);
    const { tenNCC, tenDayDuNCC, daiDien, chucVu, linkHopDong, soHopDong, ngayKy } = req.body;
    const ten = (tenNCC || req.params.nccKey || "").trim();
    if (!ten) return res.status(400).json({ error: "Thiếu tên NCC." });
    const rec = {
      id: nextId(store, "phap_danh_hop_dong_ncc_seq") || Date.now(),
      // Nhan, 2026-09-17: luu CA tenNCC (ten ngan, dung khop voi nccShort) LAN
      // tenDayDuNCC (ten day du neu co) -- logic khop hop dong (normForMatch)
      // dung ca 2 truong nay, thieu 1 trong 2 co the lam khop nham/khong khop
      // duoc voi chinh NCC vua them.
      tenNCC: ten,
      tenDayDuNCC: (tenDayDuNCC || "").trim(),
      congTy: activeCompany,
      daiDien: (daiDien || "").trim(),
      chucVu: (chucVu || "").trim(),
      linkHopDong: (linkHopDong || "").trim(),
      soHopDong: (soHopDong || "").trim(),
      ngayKy: ngayKy || "",
      chiTiet: [],
      ghiChu: "",
      createdAt: new Date().toISOString(),
      source: "them tu Ho So Hoa Don NCC",
    };
    store.phap_danh_hop_dong_ncc.push(rec);
    save(store);
    res.json({ ok: true, hopDong: rec });
  } catch (e) {
    res.status(500).json({ error: "Lỗi lưu hợp đồng: " + e.message });
  }
});

// Nhan, 2026-09-16: "cho tôi thêm 1 chỗ xuất chứng từ mẫu ... biên bản giao
// nhận hay biên bản nghiệm thu hay bảng kê hóa đơn" -- sinh file Word cho 1
// hoặc nhiều hóa đơn đã CHỌN của 1 NCC. Quyết định theo AskUserQuestion:
//  - Giao nhận/Nghiệm thu: chọn NHIỀU hóa đơn -> mỗi hóa đơn 1 file .docx
//    riêng (vì mỗi biên bản chỉ có đúng 1 ngày = ngày hóa đơn), nén ZIP nếu
//    >1 file; chỉ 1 hóa đơn thì trả thẳng .docx.
//  - Bảng kê hóa đơn: luôn 1 file .docx dù chọn bao nhiêu hóa đơn.
//  - Tên hàng hóa lấy nguyên "Nội dung" đã có (không tách ĐVT/SL).
//  - Info Bên A (NCC): lấy tenDayDuNCC; địa chỉ/điện thoại/đại diện/chức vụ
//    CHỈ điền nếu tìm thấy hợp đồng NCC khớp có sẵn field đó, không thì để
//    trống (gạch chấm) cho Nhan viết tay.
//  - Info Bên B (K&H): lấy từ utils/companies.js (đang để trống chờ Nhan
//    cung cấp, xem ghi chú trong file đó).
router.post("/ho-so/hoa-don-ncc/xuat-chung-tu", requireLogin, express.json(), async (req, res) => {
  try {
    const store = load();
    ensureHoSo(store);
    const activeCompany = getCompany(req);
    const nccKey = String(req.body.nccKey || "").toLowerCase().trim();
    const loai = String(req.body.loai || "").trim(); // 'giao-nhan' | 'nghiem-thu' | 'bang-ke'
    const driveIds = Array.isArray(req.body.driveIds) ? req.body.driveIds.map(String) : [];
    if (!nccKey) return res.status(400).json({ error: "Thiếu nccKey." });
    if (!driveIds.length) return res.status(400).json({ error: "Chưa chọn hóa đơn nào." });
    if (!["giao-nhan", "nghiem-thu", "bang-ke"].includes(loai)) return res.status(400).json({ error: "Loại chứng từ không hợp lệ." });

    const hddvLookup = buildHddvLookup(store, activeCompany);
    const allForCompany = store.ho_so_hoa_don.filter((r) => !r.company || r.company === activeCompany);
    const rows = allForCompany.map((r) => enrichRow(r, hddvLookup, store));
    const nccRows = rows.filter((r) => r.nccShort.toLowerCase() === nccKey);
    if (!nccRows.length) return res.status(404).json({ error: "Không tìm thấy NCC." });
    const invoices = nccRows.filter((r) => driveIds.includes(String(r.driveId)));
    if (!invoices.length) return res.status(404).json({ error: "Không tìm thấy hóa đơn đã chọn." });

    const nccShort = nccRows[0].nccShort;
    const tenDayDuNCC = nccRows.find((r) => r.tenDayDuNCC)?.tenDayDuNCC || "";

    // Do hop dong NCC khop de lay so HD (nghiem thu) + cac truong dia chi/dai
    // dien/chuc vu/dien thoai NEU CO san trong ho so hop dong.
    const hdNccList = store.phap_danh_hop_dong_ncc || [];
    const shortNorm = normForMatch(nccShort);
    const fullNorm = normForMatch(tenDayDuNCC);
    const matchedHopDong = hdNccList.find((hd) => {
      const hdShort = normForMatch(hd.tenNCC || "");
      const hdFull = normForMatch(hd.tenDayDuNCC || "");
      // Nhan, 2026-09-17: bug cu -- "".includes(x) luon false nhung x.includes("")
      // luon TRUE, nen hop dong thieu tenDayDuNCC/tenNCC (vd hop dong vua them
      // thu cong qua "+ Thêm hợp đồng NCC") se KHOP NHAM voi MOI NCC khac co
      // fullNorm/shortNorm >= do dai toi thieu. Them dieu kien do dai > 0 cho
      // ca 2 ve de chi khop khi CA HAI ben deu co ten thuc su.
      return (shortNorm.length >= 3 && hdShort.length > 0 && (hdShort.includes(shortNorm) || shortNorm.includes(hdShort))) ||
             (fullNorm.length >= 5 && hdFull.length > 0 && (hdFull.includes(fullNorm) || fullNorm.includes(hdFull)));
    });

    // Nhan, 2026-09-17: "Đại diện" (Ben A) CHI lay tu Hop dong NCC da map/them
    // (khong doan), con "Địa chỉ"/"Điện thoại" lay TRUC TIEP tu file PDF hoa
    // don (xem fetchInvoicePdfExtract ben duoi), khong lay tu Hop dong nua.
    const ncc = {
      nccShort,
      tenDayDuNCC,
      diaChi: "",
      dienThoai: "",
      daiDien: matchedHopDong && matchedHopDong.daiDien ? matchedHopDong.daiDien : "",
      chucVu: matchedHopDong && matchedHopDong.chucVu ? matchedHopDong.chucVu : "",
    };
    const company = COMPANIES[activeCompany] || COMPANIES.kh_cu;
    const safeName = (nccShort || "NCC").replace(/[^a-zA-Z0-9À-ỹ_-]+/g, "_");

    if (loai === "bang-ke") {
      const doc = buildBangKeHoaDon({ ncc, invoices, company });
      const buf = await Packer.toBuffer(doc);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="BangKeHoaDon_${safeName}.docx"`);
      return res.send(buf);
    }

    // Giao nhan / Nghiem thu: doc tung file PDF hoa don da chon de lay dia
    // chi/dien thoai NCC (dung chung cho ca nhom, lay tu hoa don dau tien
    // trich xuat thanh cong) + bang hang hoa chi tiet rieng cho tung hoa don.
    for (const inv of invoices) {
      const extract = await fetchInvoicePdfExtract(store, inv.driveId);
      if (extract) {
        if (extract.items && extract.items.length) inv.pdfItems = extract.items;
        if (extract.tongCong) inv.pdfTongCong = extract.tongCong;
        if (!ncc.diaChi && extract.diaChi) ncc.diaChi = extract.diaChi;
        if (!ncc.dienThoai && extract.dienThoai) ncc.dienThoai = extract.dienThoai;
      }
    }

    const buildOne = loai === "giao-nhan"
      ? (inv) => buildBienBanGiaoNhan({ ncc, invoice: inv, company })
      : (inv) => buildBienBanNghiemThu({ ncc, invoice: inv, company, hopDong: matchedHopDong });
    const labelPrefix = loai === "giao-nhan" ? "BienBanGiaoNhan" : "BienBanNghiemThu";

    if (invoices.length === 1) {
      const buf = await Packer.toBuffer(buildOne(invoices[0]));
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="${labelPrefix}_${safeName}_${invoices[0].soHoaDon || ""}.docx"`);
      return res.send(buf);
    }

    const zip = new AdmZip();
    for (const inv of invoices) {
      const buf = await Packer.toBuffer(buildOne(inv));
      const fname = `${labelPrefix}_${safeName}_${(inv.soHoaDon || "khong-so").replace(/[^a-zA-Z0-9_-]+/g, "_")}.docx`;
      zip.addFile(fname, buf);
    }
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${labelPrefix}_${safeName}.zip"`);
    return res.send(zip.toBuffer());
  } catch (e) {
    res.status(500).json({ error: "Lỗi tạo chứng từ: " + e.message });
  }
});

// ─── Phí Cảng Phú Quốc ────────────────────────────────────────────────────
router.get("/ho-so/phi-cang-phu-quoc", (req, res) => {
  const store = load();
  const rows = (store.ho_so_phi_cang || []).slice().sort((a,b) => (b.ngay||'').localeCompare(a.ngay||''));
  res.render("ho-so-phi-cang", {
    userName: req.session.userName,
    rows,
    error: req.query.error || null,
    success: req.query.success || null,
    warn: req.query.warn || null,
  });
});

router.post("/ho-so/phi-cang-phu-quoc/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  const { ngay, soHoaDon, loaiPhi, soTien, linkFile, ghiChu } = req.body;
  store.ho_so_phi_cang.push({
    id: nextId(store),
    ngay: (ngay||"").trim(),
    soHoaDon: (soHoaDon||"").trim(),
    loaiPhi: (loaiPhi||"").trim(),
    soTien: (soTien||"").trim(),
    linkFile: (linkFile||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+thêm");
});

router.post("/ho-so/phi-cang-phu-quoc/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  store.ho_so_phi_cang = (store.ho_so_phi_cang||[]).filter(r=>String(r.id)!==req.params.id);
  save(store);
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+xóa");
});

router.post("/ho-so/phi-cang-phu-quoc/:id/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_phi_cang||[]).find(x=>String(x.id)===req.params.id);
  if (r) { r.ngay=(req.body.ngay||'').trim(); r.soHoaDon=(req.body.soHoaDon||'').trim(); r.loaiPhi=(req.body.loaiPhi||'').trim(); r.soTien=(req.body.soTien||'').trim(); r.linkFile=(req.body.linkFile||'').trim(); r.ghiChu=(req.body.ghiChu||'').trim(); save(store); }
  res.redirect("/ho-so/phi-cang-phu-quoc?success=Đã+lưu");
});

// ─── Doanh Thu Chia Sẻ ────────────────────────────────────────────────────
const DTCS_TABS = [
  { key: 'mall-tru-chi-phi',   label: '🏦 Mall giữ tiền — Trừ chi phí' },
  { key: 'mall-phan-tram-dt',  label: '🏦 Mall giữ tiền — % Doanh thu' },
  { key: 'tien-thue-vuot',     label: '📈 Tiền thuê vượt' },
  { key: 'tien-thue-co-dinh',  label: '📊 Chia sẻ trên %DT (mình giữ tiền)' },
];
router.get("/ho-so/doanh-thu-chia-se", (req, res) => {
  const store = load();
  const activeTab = DTCS_TABS.some(t => t.key === req.query.tab) ? req.query.tab : DTCS_TABS[0].key;
  const allRows = (store.ho_so_doanhthu_chiase || []).slice().sort((a,b) => (b.ngay||'').localeCompare(a.ngay||''));
  const rows = allRows.filter(r => (r.loaiTab || DTCS_TABS[0].key) === activeTab);
  // Load danh sach Diem theo tab hiện tại
  const diemList = (store.dtcs_chiase_diem || [])
    .filter(d => (d.tabKey || 'tien-thue-co-dinh') === activeTab)
    .sort((a,b) => (a.tenDiem||'').localeCompare(b.tenDiem||''));
  res.render("ho-so-doanhthu-chiase", {
    userName: req.session.userName,
    rows,
    tabs: DTCS_TABS,
    activeTab,
    diemList,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// ─── DTCS Chia Sẻ — Điểm (location) CRUD ─────────────────────────────────────
router.post("/ho-so/dtcs-chiase/diem/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.dtcs_chiase_diem) store.dtcs_chiase_diem = [];
  const { tenDiem, maCongTrinh, ghiChu, tabKey, tuThang, denThang } = req.body;
  store.dtcs_chiase_diem.push({
    id: nextId(store),
    tabKey: (tabKey||"tien-thue-co-dinh").trim(),
    tenDiem: (tenDiem||"").trim(),
    maCongTrinh: (maCongTrinh||"").trim(),
    tuThang: (tuThang||"").trim(),
    denThang: (denThang||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.json({ success: true });
});

router.post("/ho-so/dtcs-chiase/diem/:id/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.dtcs_chiase_diem||[]).find(x=>String(x.id)===req.params.id);
  if (!r) return res.json({ success: false, error: "Không tìm thấy" });
  r.tenDiem = (req.body.tenDiem||"").trim();
  r.maCongTrinh = (req.body.maCongTrinh||"").trim();
  r.tuThang = (req.body.tuThang||"").trim();
  r.denThang = (req.body.denThang||"").trim();
  r.ghiChu = (req.body.ghiChu||"").trim();
  save(store);
  res.json({ success: true });
});

router.post("/ho-so/dtcs-chiase/diem/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  store.dtcs_chiase_diem = (store.dtcs_chiase_diem||[]).filter(x=>String(x.id)!==req.params.id);
  // Also delete all thang records for this diem
  store.dtcs_chiase_thang = (store.dtcs_chiase_thang||[]).filter(x=>String(x.diemId)!==req.params.id);
  save(store);
  res.json({ success: true });
});

// ─── DTCS Chia Sẻ — Tháng (month records per Điểm) ──────────────────────────
router.get("/ho-so/dtcs-chiase/diem/:id/thang", (req, res) => {
  const store = load();
  const diem = (store.dtcs_chiase_diem||[]).find(x=>String(x.id)===req.params.id);
  if (!diem) return res.json({ success: false, error: "Không tìm thấy điểm" });
  const thangList = (store.dtcs_chiase_thang||[])
    .filter(x=>String(x.diemId)===req.params.id)
    .sort((a,b)=>{
      // Sort by thang MM/YYYY descending
      const toNum = s => { const p=(s||'').split('/'); return (parseInt(p[1]||0)*100+parseInt(p[0]||0)); };
      return toNum(b.thang) - toNum(a.thang);
    });
  res.json({ success: true, diem, thangList });
});

router.post("/ho-so/dtcs-chiase/diem/:id/thang/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.dtcs_chiase_thang) store.dtcs_chiase_thang = [];
  const { thang, phanTramHo, phanTramMiNh, tongDT, soTienMinhNhan, soTienMinhTra,
          hoaDonHo, tkThanhToan, ngayThanhToan, linkDoiSoat, tienThue, khoauTru, ghiChu } = req.body;
  store.dtcs_chiase_thang.push({
    id: nextId(store),
    diemId: req.params.id,
    thang: (thang||"").trim(),
    phanTramHo: (phanTramHo||"").trim(),
    phanTramMiNh: (phanTramMiNh||"").trim(),
    tongDT: (tongDT||"").trim(),
    soTienMinhNhan: (soTienMinhNhan||"").trim(),
    soTienMinhTra: (soTienMinhTra||"").trim(),
    hoaDonHo: (hoaDonHo||"").trim(),
    tkThanhToan: (tkThanhToan||"").trim(),
    ngayThanhToan: (ngayThanhToan||"").trim(),
    linkDoiSoat: (linkDoiSoat||"").trim(),
    tienThue: (tienThue||"").trim(),
    khoauTru: (khoauTru||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.json({ success: true });
});

router.post("/ho-so/dtcs-chiase/diem/:id/thang/:thangId/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.dtcs_chiase_thang||[]).find(x=>String(x.id)===req.params.thangId && String(x.diemId)===req.params.id);
  if (!r) return res.json({ success: false, error: "Không tìm thấy" });
  ['thang','phanTramHo','phanTramMiNh','tongDT','soTienMinhNhan','soTienMinhTra',
   'hoaDonHo','tkThanhToan','ngayThanhToan','linkDoiSoat','tienThue','khoauTru','ghiChu'].forEach(k => {
    r[k] = (req.body[k]||"").trim();
  });
  save(store);
  res.json({ success: true });
});

router.post("/ho-so/dtcs-chiase/diem/:id/thang/:thangId/xoa", requireAdmin, (req, res) => {
  const store = load();
  store.dtcs_chiase_thang = (store.dtcs_chiase_thang||[]).filter(
    x=>!(String(x.id)===req.params.thangId && String(x.diemId)===req.params.id)
  );
  save(store);
  res.json({ success: true });
});

router.post("/ho-so/doanh-thu-chia-se/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_doanhthu_chiase) store.ho_so_doanhthu_chiase = [];
  const { ngay, gian, loai, soTien, linkFile, ghiChu, loaiTab } = req.body;
  const tab = DTCS_TABS.some(t => t.key === loaiTab) ? loaiTab : DTCS_TABS[0].key;
  store.ho_so_doanhthu_chiase.push({
    id: nextId(store),
    ngay: (ngay||"").trim(),
    gian: (gian||"").trim(),
    loai: (loai||"").trim(),
    soTien: (soTien||"").trim(),
    linkFile: (linkFile||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    loaiTab: tab,
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/doanh-thu-chia-se?tab=" + tab + "&success=Đã+thêm");
});

router.post("/ho-so/doanh-thu-chia-se/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_doanhthu_chiase||[]).find(x=>String(x.id)===req.params.id);
  const tab = r ? (r.loaiTab || DTCS_TABS[0].key) : DTCS_TABS[0].key;
  store.ho_so_doanhthu_chiase = (store.ho_so_doanhthu_chiase||[]).filter(r=>String(r.id)!==req.params.id);
  save(store);
  res.redirect("/ho-so/doanh-thu-chia-se?tab=" + tab + "&success=Đã+xóa");
});

router.post("/ho-so/doanh-thu-chia-se/:id/sua", requireAdmin, (req, res) => {
  const store = load();
  const r = (store.ho_so_doanhthu_chiase||[]).find(x=>String(x.id)===req.params.id);
  if (r) { r.ngay=(req.body.ngay||'').trim(); r.gian=(req.body.gian||'').trim(); r.loai=(req.body.loai||'').trim(); r.soTien=(req.body.soTien||'').trim(); r.linkFile=(req.body.linkFile||'').trim(); r.ghiChu=(req.body.ghiChu||'').trim(); save(store); }
  const tab = r ? (r.loaiTab || DTCS_TABS[0].key) : DTCS_TABS[0].key;
  res.redirect("/ho-so/doanh-thu-chia-se?tab=" + tab + "&success=Đã+lưu");
});

// ─── Đọc PDF biên lai → trả về fields để auto-fill form ─────────────────────
// Luyen, 2026-08-25: "tải file PDF, đọc nội dung, điền vào form" -- nhan PDF
// qua multer (memoryStorage), goi driveApi.extractPhiCangInfo, tra JSON.
// Khong can auth (chi doc, khong luu gi).
router.post("/ho-so/phi-cang-phu-quoc/doc-file", uploadMem.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Vui lòng chọn file PDF." });
    const mt = (req.file.mimetype || "").toLowerCase();
    if (!mt.includes("pdf")) return res.status(400).json({ error: "Chỉ hỗ trợ file PDF để đọc tự động." });
    const info = await driveApi.extractPhiCangInfo(req.file.buffer);
    return res.json({ success: true, ...info });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ─── Upload ảnh Phí Cảng → PDF → Google Drive ────────────────────────────────
// Luyen, 2026-08-25: "tải lên ảnh, lưu ảnh qua PDF đặt tên số HĐ + ngày, lưu
// vào Drive, điền thông tin cơ bản hiển thị trên web" -- nhan anh tu form,
// chuyen sang PDF bang pdf-lib, upload len folder "Phi Cang HK" tren Drive,
// luu record vao store.ho_so_phi_cang voi link Drive.
router.post("/ho-so/phi-cang-phu-quoc/upload-anh", requireDataEntry, uploadMem.array("anh", 10), async (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  if (!req.files || req.files.length === 0) {
    return res.redirect("/ho-so/phi-cang-phu-quoc?error=" + encodeURIComponent("Vui lòng chọn ít nhất 1 file."));
  }
  const { ngay, soHoaDon, loaiPhi, soTien, ghiChu } = req.body;
  const soHD = (soHoaDon || "").trim();
  // Dinh dang ten file: PhiCang_SoHD_DD-MM-YYYY.pdf
  const dateStr = ngay
    ? ngay.split("-").reverse().join("-") // YYYY-MM-DD → DD-MM-YYYY
    : new Date().toLocaleDateString("vi-VN").replace(/\//g, "-");

  const driveWarnings = [];
  let added = 0;

  for (let i = 0; i < req.files.length; i++) {
    const f = req.files[i];
    const suffix = req.files.length > 1 ? `_${i + 1}` : "";
    const fileName = `PhiCang_${soHD || "HoaDon"}${suffix}_${dateStr}.pdf`;
    let linkFile = "";
    let driveId = "";

    // Thu upload Drive; neu loi (chua ket noi OAuth) thi van luu record khong co link
    try {
      const isPdf = (f.mimetype || "").toLowerCase().includes("pdf");
      const pdfBuffer = isPdf ? f.buffer : await driveApi.imageToPdf(f.buffer, f.mimetype);
      const driveFile = await driveApi.uploadFileToDrive(
        store, driveApi.DRIVE_PHI_CANG_FOLDER_ID, fileName, pdfBuffer, "application/pdf"
      );
      linkFile = "https://drive.google.com/file/d/" + driveFile.id + "/view";
      driveId = driveFile.id;
    } catch (e) {
      driveWarnings.push("⚠️ Chưa kết nối Drive — đã lưu hồ sơ trên web nhưng chưa upload file lên Google Drive. Vào Chi Phí → Kết nối Gmail để cấp quyền.");
    }

    store.ho_so_phi_cang.push({
      id: nextId(store),
      ngay: ngay || "",
      soHoaDon: soHD,
      loaiPhi: (loaiPhi || "").trim(),
      soTien: (soTien || "").trim(),
      linkFile,
      driveId,
      ghiChu: (ghiChu || "").trim(),
      tenFile: fileName,
      createdAt: new Date().toISOString(),
    });
    added++;
  }

  save(store);
  let msg = "Đã lưu " + added + " hồ sơ" + (driveWarnings.length > 0 ? " (chưa có link Drive)." : " lên Drive.");
  if (driveWarnings.length > 0) {
    return res.redirect("/ho-so/phi-cang-phu-quoc?success=" + encodeURIComponent(msg) + "&warn=" + encodeURIComponent(driveWarnings[0]));
  }
  res.redirect("/ho-so/phi-cang-phu-quoc?success=" + encodeURIComponent(msg));
});

// Luyen 2026-08-31: route bulk-upsert duoc chuyen len truoc router.use(requireLogin)
// o dau file de X-Internal-Key co the bypass auth. Route cu ben duoi da xoa.

// Route cu - giu lai de khong bi 404 nhung khong dung OAuth nua
router.post("/ho-so/sync-drive-hoa-don", requireAdmin, (req, res) => {
  res.redirect("/ho-so/hoa-don-ncc?error=" + encodeURIComponent("Vui lòng dùng Claude Cowork → nhắn: sync hóa đơn drive"));
});

// ─── Sync Phí Cảng từ Google Drive ───────────────────────────────────────────
// Luyen, 2026-08-25: sync PDF tu folder "Phi Cang HK" Drive vao ho_so_phi_cang.
// Moi file tao 1 dong moi voi linkFile = link Drive, ten file lam loaiPhi tam,
// ngay/soTien de trong cho Luyen tu dien sau khi kiem tra.
router.post("/ho-so/phi-cang-phu-quoc/sync-drive", requireAdmin, async (req, res) => {
  const store = load();
  if (!store.ho_so_phi_cang) store.ho_so_phi_cang = [];
  try {
    const files = await driveApi.listFolderFiles(store, driveApi.DRIVE_PHI_CANG_FOLDER_ID, "application/pdf");
    const existingLinks = new Set(
      store.ho_so_phi_cang.map((r) => r.linkFile).filter(Boolean)
    );
    let added = 0;
    files.forEach((f) => {
      const link = "https://drive.google.com/file/d/" + f.id + "/view";
      if (!existingLinks.has(link)) {
        store.ho_so_phi_cang.push({
          id: nextId(store),
          ngay: "",
          loaiPhi: f.name.replace(/\.pdf$/i, ""),
          soTien: "",
          linkFile: link,
          ghiChu: "Sync từ Drive tự động",
          createdAt: new Date().toISOString(),
        });
        added++;
      }
    });
    save(store);
    res.redirect(
      "/ho-so/phi-cang-phu-quoc?success=" +
        encodeURIComponent(
          "Đã sync Drive: thêm " + added + " file mới (tổng " + files.length + " file trong folder)."
        )
    );
  } catch (e) {
    res.redirect("/ho-so/phi-cang-phu-quoc?error=" + encodeURIComponent(e.message));
  }
});

// ─── Tiền Thuê Bình Thường ────────────────────────────────────────────────────
// ─── Chứng Từ ────────────────────────────────────────────────────────────────
// Luyen 2026-08-31: lưu file PDF chứng từ (UNC, biên lai...) từ Drive,
// daHachToan = false mặc định, tick tay để đánh dấu đã hạch toán.
function nextChungTuId(store) {
  const ids = (store.ho_so_chung_tu || []).map((r) => Number(r.id) || 0);
  return (ids.length ? Math.max(...ids) : 0) + 1;
}

router.get("/ho-so/chung-tu", (req, res) => {
  const store = load();
  ensureHoSo(store);
  const q = (req.query.q || "").toLowerCase();
  const htFilter = req.query.hachToan || "";
  let rows = store.ho_so_chung_tu.slice().sort((a, b) =>
    (b.ngay || "").localeCompare(a.ngay || "")
  );
  if (q) rows = rows.filter((r) =>
    [r.loai, r.soChungTu, r.ncc, r.ghiChu, r.fileName].some(
      (f) => (f || "").toLowerCase().includes(q)
    )
  );
  if (htFilter === "1") rows = rows.filter((r) => r.daHachToan);
  else if (htFilter === "0") rows = rows.filter((r) => !r.daHachToan);
  res.render("ho-so-chung-tu", {
    userName: req.session ? req.session.userName : null,
    rows,
    q,
    htFilter,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/ho-so/chung-tu/them", requireDataEntry, (req, res) => {
  const store = load();
  ensureHoSo(store);
  const { loai, soChungTu, ngay, ncc, soTien, driveId, fileName, ghiChu } = req.body;
  store.ho_so_chung_tu.push({
    id: nextChungTuId(store),
    loai: (loai || "").trim(),
    soChungTu: (soChungTu || "").trim(),
    ngay: (ngay || "").trim(),
    ncc: (ncc || "").trim(),
    soTien: Number(String(soTien || "0").replace(/[^0-9]/g, "")) || 0,
    driveId: (driveId || "").trim(),
    fileName: (fileName || "").trim(),
    ghiChu: (ghiChu || "").trim(),
    daHachToan: false,
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/chung-tu?success=Đã+thêm+chứng+từ");
});

router.post("/ho-so/chung-tu/:id/hach-toan", requireDataEntry, (req, res) => {
  const store = load();
  ensureHoSo(store);
  const r = store.ho_so_chung_tu.find((x) => String(x.id) === req.params.id);
  if (r) r.daHachToan = req.body.value === "1";
  save(store);
  res.json({ success: true, daHachToan: r ? r.daHachToan : null });
});

router.post("/ho-so/chung-tu/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  ensureHoSo(store);
  store.ho_so_chung_tu = store.ho_so_chung_tu.filter((x) => String(x.id) !== req.params.id);
  save(store);
  res.redirect("/ho-so/chung-tu?success=Đã+xóa");
});

// Nhan, 2026-09-14: "cac gian Ma NCC se gop chung hien thi ngay thang ma
// cong trinh ngay nao hay dien giai va so hoa don... 1 ncc roi liet ke cac
// cai lien quan khi toi nhan vao roi hien bang giua cac thong tin lien
// quan" -- parse chuoi "Ma nha cung cap" dang TEN dinh lien MA_SO_THUE
// (vd "GIGAMALL0313114291", "Lotte Nha Trang0304741634-011") khong co dau
// phan cach, de tach ten rieng phuc vu nhom/hien thi.
function parseNccString(raw) {
  const s = (raw || "").trim();
  const m = s.match(/^(.*?)(\d{6,}[\d-]*)$/);
  if (m && m[1].trim()) return { tenNCC: m[1].trim(), maSoThueNCC: m[2].trim() };
  return { tenNCC: s, maSoThueNCC: "" };
}

// Doi serial ngay Excel (vd 46235) sang chuoi dd/mm/yyyy.
function excelSerialToDateStr(serial) {
  const n = Number(serial);
  if (!n || isNaN(n)) return "";
  try {
    const d = XLSX.SSF.parse_date_code(n);
    if (!d || !d.y) return "";
    const pad = (x) => String(x).padStart(2, "0");
    return `${pad(d.d)}/${pad(d.m)}/${d.y}`;
  } catch (e) {
    return "";
  }
}

router.get("/ho-so/tien-thue", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  // Loc theo cong ty: ban cu khong co truong company -> van hien cho ca 2
  // ben (giu nguyen hanh vi cu, tranh mat du lieu da nhap tay truoc day).
  const allRows = (store.ho_so_tien_thue || []).filter(
    (r) => !r.company || r.company === activeCompany
  );
  const rows = allRows.slice().sort((a, b) => (b.thang || "").localeCompare(a.thang || ""));

  // Nhom cac dong co Ma NCC (tu import Excel "So chi tiet mua hang") lai
  // theo tung NCC -- 1 dong tong hop / NCC, nhan vao xem chi tiet. Dong nhap
  // tay (khong co maNCC) van hien rieng le o bang ben duoi nhu truoc gio.
  const groupMap = new Map();
  rows.forEach((r) => {
    if (!r.maNCC) return;
    if (!groupMap.has(r.maNCC)) {
      groupMap.set(r.maNCC, {
        maNCC: r.maNCC,
        tenNCC: r.tenNCC || r.maNCC,
        maSoThueNCC: r.maSoThueNCC || "",
        rows: [],
        tongTien: 0,
      });
    }
    const g = groupMap.get(r.maNCC);
    g.rows.push(r);
    // Uu tien tong cong (da gom thue GTGT) neu co (du lieu import moi), fallback
    // ve so tien truoc thue cho du lieu cu chua co truong nay.
    const tongRow = r.soTienTongRaw !== undefined
      ? r.soTienTongRaw
      : (r.soTienRaw !== undefined ? r.soTienRaw : Number(String(r.soTien || "").replace(/[.,\s]/g, "")) || 0);
    g.tongTien += Number(tongRow) || 0;
  });
  const groupsNcc = Array.from(groupMap.values()).sort((a, b) =>
    a.tenNCC.localeCompare(b.tenNCC)
  );
  groupsNcc.forEach((g) => g.rows.sort((a, b) => (b.thang || "").localeCompare(a.thang || "")));

  const rowsManual = rows.filter((r) => !r.maNCC);

  res.render("ho-so-tien-thue", {
    userName: req.session.userName,
    rows: rowsManual,
    groupsNcc,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/ho-so/tien-thue/them", requireAdmin, (req, res) => {
  const store = load();
  if (!store.ho_so_tien_thue) store.ho_so_tien_thue = [];
  const { thang, gian, soHoaDon, soTien, linkFile, ghiChu } = req.body;
  store.ho_so_tien_thue.push({
    id: nextId(store),
    company: getCompany(req),
    thang: (thang||"").trim(),
    gian: (gian||"").trim(),
    soHoaDon: (soHoaDon||"").trim(),
    soTien: (soTien||"").trim(),
    linkFile: (linkFile||"").trim(),
    ghiChu: (ghiChu||"").trim(),
    createdAt: new Date().toISOString(),
  });
  save(store);
  res.redirect("/ho-so/tien-thue?success=Đã+thêm");
});

router.post("/ho-so/tien-thue/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  store.ho_so_tien_thue = (store.ho_so_tien_thue||[]).filter(r=>String(r.id)!==req.params.id);
  save(store);
  res.redirect("/ho-so/tien-thue?success=Đã+xóa");
});

// Nhan, 2026-09-14: "xay nut tai file len + nhap luon file nay" -- import
// file Excel "SO CHI TIET MUA HANG" (xuat tu Misa/phan mem ke toan), loc
// rieng cac dong Ma hang = "TIEN THUE", gan company = cong ty dang active
// luc bam nut import. Chong nhap trung: dedupe theo "So chung tu" (cot noi
// bo, on dinh hon so hoa don vi 1 so hoa don co the co nhieu dong).
router.post("/ho-so/tien-thue/nhap-excel", requireAdmin, uploadMem.single("file"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Vui lòng chọn file Excel." });
    const wb = XLSX.read(req.file.buffer, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    // Tim dong tieu de bang cach do cell "Mã hàng" (khong cung cung index 3,
    // phong khi file xuat co so dong tieu de/khoang trang khac nhau).
    let headerIdx = -1;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      if ((data[i] || []).some((c) => String(c).trim() === "Mã hàng")) { headerIdx = i; break; }
    }
    if (headerIdx === -1) {
      return res.status(400).json({ error: "Không tìm thấy dòng tiêu đề (cột 'Mã hàng') trong file. Vui lòng kiểm tra đúng file 'SỔ CHI TIẾT MUA HÀNG'." });
    }
    const header = data[headerIdx];
    const idx = {};
    header.forEach((h, i) => { idx[String(h).trim()] = i; });

    const need = ["Ngày hóa đơn", "Số hóa đơn", "Mã nhà cung cấp", "Mã hàng", "Tên hàng", "Giá trị mua", "Thuế GTGT", "Mã công trình", "Tên công trình", "Số chứng từ"];
    const missing = need.filter((k) => idx[k] === undefined);
    if (missing.length) {
      return res.status(400).json({ error: "File thiếu cột: " + missing.join(", ") });
    }

    const dataRows = data.slice(headerIdx + 1).filter((r) => (r || []).some((c) => c !== ""));
    const tienThueRows = dataRows.filter((r) => String(r[idx["Mã hàng"]] || "").trim() === "TIỀN THUÊ");

    const store = load();
    if (!store.ho_so_tien_thue) store.ho_so_tien_thue = [];
    // Nhan, 2026-09-14: 1 "So chung tu" co the co NHIEU dong TIEN THUE (vd
    // gop chung nhieu khoan cua cung 1 CT/1 hoa don, chi khac Gia tri mua)
    // -- dedupe CHI theo So chung tu se lam mat cac dong nay. Khoa dedupe
    // phai gom them So hoa don + Ma NCC + So tien (Gia tri mua) de phan
    // biet cac dong that su khac nhau, nhung van chan duoc import trung khi
    // tai lai CHINH XAC 1 file nhieu lan.
    function dedupeKey(soChungTu, soHoaDon, maNCC, giaTriMua) {
      return [soChungTu, soHoaDon, maNCC, Math.round(Number(giaTriMua) || 0)].join("|");
    }
    const existingKeys = new Set(
      store.ho_so_tien_thue
        .filter((r) => r.soChungTu)
        .map((r) => dedupeKey(
          r.soChungTu, r.soHoaDon, r.maNCC,
          r.soTienRaw !== undefined ? r.soTienRaw : String(r.soTien || "").replace(/[.,\s]/g, "")
        ))
    );
    const activeCompany = getCompany(req);

    let added = 0, skipped = 0;
    tienThueRows.forEach((r) => {
      const soChungTu = String(r[idx["Số chứng từ"]] || "").trim();
      const maNCCRaw = String(r[idx["Mã nhà cung cấp"]] || "").trim();
      const soHoaDon = String(r[idx["Số hóa đơn"]] || "").trim();
      const giaTriMua = Number(r[idx["Giá trị mua"]]) || 0;
      const key = dedupeKey(soChungTu, soHoaDon, maNCCRaw, giaTriMua);
      if (soChungTu && existingKeys.has(key)) { skipped++; return; }

      const { tenNCC, maSoThueNCC } = parseNccString(maNCCRaw);
      const maCongTrinh = String(r[idx["Mã công trình"]] || "").trim();
      const tenCongTrinh = String(r[idx["Tên công trình"]] || "").trim();
      const ngayHoaDon = excelSerialToDateStr(r[idx["Ngày hóa đơn"]]) ||
        (idx["Ngày chứng từ"] !== undefined ? excelSerialToDateStr(r[idx["Ngày chứng từ"]]) : "");
      const soTien = giaTriMua ? giaTriMua.toLocaleString("vi-VN") : "";
      // Nhan, 2026-09-14: "cho cai dien giai... so tien truoc thue, so tien
      // thue va so tien tong" -- "Ten hang" trong file thuong la "TIEN THUE"
      // (giu nguyen lam fallback), nhung doi khi co dien giai chi tiet hon
      // (vd "DOANH THU PHAN CHIA... KY 07.2026...") thi lay dung dien giai do.
      // Tong = Gia tri mua (truoc thue) + Thue GTGT.
      const dienGiaiRaw = String(r[idx["Tên hàng"]] || "").trim();
      const dienGiai = dienGiaiRaw || "Tiền thuê";
      const thueGtgt = Number(r[idx["Thuế GTGT"]]) || 0;
      const tongCong = giaTriMua + thueGtgt;
      let thang = "";
      const dm = ngayHoaDon.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (dm) thang = dm[2] + "/" + dm[3];

      store.ho_so_tien_thue.push({
        id: nextId(store),
        company: activeCompany,
        thang,
        gian: tenCongTrinh || maCongTrinh || "",
        maCongTrinh,
        tenCongTrinh,
        dienGiai,
        soHoaDon,
        soChungTu,
        soTien,
        soTienRaw: giaTriMua,
        thueGtgt: thueGtgt ? thueGtgt.toLocaleString("vi-VN") : "",
        thueGtgtRaw: thueGtgt,
        soTienTong: tongCong ? tongCong.toLocaleString("vi-VN") : "",
        soTienTongRaw: tongCong,
        maNCC: maNCCRaw,
        tenNCC,
        maSoThueNCC,
        ngayHoaDon,
        linkFile: "",
        ghiChu: "",
        createdAt: new Date().toISOString(),
      });
      if (soChungTu) existingKeys.add(key);
      added++;
    });

    save(store);
    res.json({ success: true, added, skipped, total: tienThueRows.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Nhan, 2026-09-14: cong cu don import loi (vd sau khi sua bug dedupe) -- xoa
// het cac dong da nhap tu Excel (co maNCC) cua cong ty dang active, de import
// lai sach tu dau. Khong co nut tren UI (chi goi thu cong khi can), tranh
// bam nham xoa nham nhap tay.
router.post("/ho-so/tien-thue/xoa-import", requireAdmin, (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const before = (store.ho_so_tien_thue || []).length;
  store.ho_so_tien_thue = (store.ho_so_tien_thue || []).filter(
    (r) => !(r.maNCC && (!r.company || r.company === activeCompany))
  );
  const deleted = before - store.ho_so_tien_thue.length;
  save(store);
  res.json({ success: true, deleted });
});

// Nhan, 2026-09-12: "cho toi chỗ tải link gg drive của hợp đồng á xong rồi
// bạn đọc và lấy ra thông tin thêm điểm cho tôi nhá rồi lưu lại 1 mã công
// trình và các thông tin cơ bản rồi thêm cái link đó lên nhá" -- o form them
// dong "Tien Thue" da co san o "Link tai lieu (Drive)". Them nut "Doc hop
// dong" doc PDF hop dong thue tu Drive link do (dung LAI dung 1 co che da
// chung minh hoat dong trong routes/phap-danh.js, muc /hop-dong-thue-gian-tong/
// :id/doc-hop-dong: lay access token qua utils/gmailApi.js (chung voi Gmail
// OAuth), tai file qua Drive API alt=media, parse bang pdf-parse), rieng ham
// trich xuat duoi day mo rong hon (rieng cho hop dong thue gian: ben cho
// thue, MST, tien thue/thang, ngay ky, thoi han...) thay vi chi 4 truong
// chung chung nhu ham cu. Chi la GOI Y (best-effort tren van ban hop dong
// that su rat da dang ve cach trinh bay) -- luon hien cho Luyen xem lai/sua
// truoc khi luu, khong tu dong ghi thang.
function extractTienThueHopDongInfo(text) {
  const norm = (s) => (s || "").replace(/[ \t]+/g, " ").replace(/\n+/g, " ").trim();
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const joined = lines.join("\n");

  function grabNear(re, span) {
    const idx = lines.findIndex((l) => re.test(l));
    if (idx < 0) return "";
    return norm(lines.slice(idx, idx + span).join(" "));
  }

  // Ben cho thue -- doan "BEN CHO THUE (BEN A)" hoac "BEN A:", lay ten cong ty/
  // ca nhan ngay sau, cat bot o cac tu khoa ke tiep (Dia chi/MST) neu dinh lien.
  let benChoThue = "";
  const benAIdx = lines.findIndex((l) => /bên\s*(cho\s*thuê|a)\b/i.test(l));
  if (benAIdx >= 0) {
    const seg = norm(lines.slice(benAIdx, benAIdx + 6).join(" "));
    const m = /(?:tên\s*(?:công\s*ty|doanh\s*nghiệp)?\s*[:\-]?\s*)([^;]{3,80})/i.exec(seg);
    if (m) benChoThue = m[1].split(/địa chỉ|mã số thuế|\bmst\b/i)[0].trim();
  }

  let mstBenChoThue = "";
  const mstM = /(?:mã\s*số\s*thuế|mst)\s*[:\-]?\s*([0-9\-]{8,15})/i.exec(joined);
  if (mstM) mstBenChoThue = mstM[1];

  let tienThueThang = "";
  const rentM = /(?:tiền\s*thuê|giá\s*thuê)[^\d]{0,40}([\d.,]{5,})\s*(?:đồng|vnđ|vnd)/i.exec(joined);
  if (rentM) tienThueThang = rentM[1].replace(/[.,](?=\d{3}\b)/g, "").replace(/[^\d]/g, "");

  let ngayKyHD = "";
  const kyM = /ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})/i.exec(joined);
  if (kyM) ngayKyHD = `${kyM[1].padStart(2, "0")}/${kyM[2].padStart(2, "0")}/${kyM[3]}`;

  const thoiHanHopDong = grabNear(/thời\s*hạn\s*(thuê|hợp\s*đồng)/i, 3);
  const viTri = grabNear(/vị\s*trí|địa\s*điểm\s*thuê|mặt\s*bằng/i, 3);
  const dieuKhoanThanhToan = grabNear(/thanh\s*toán/i, 5);

  return { benChoThue, mstBenChoThue, tienThueThang, ngayKyHD, thoiHanHopDong, viTri, dieuKhoanThanhToan };
}

router.post("/ho-so/tien-thue/doc-hop-dong", requireAdmin, express.json(), async (req, res) => {
  const link = (req.body && req.body.link || "").trim();
  if (!link) return res.json({ success: false, error: "Thiếu link Drive." });
  const driveMatch = link.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  const openMatch = link.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  const fileId = (driveMatch && driveMatch[1]) || (openMatch && openMatch[1]);
  if (!fileId) return res.json({ success: false, error: "Không nhận dạng được file ID từ link Drive (link phải dạng .../d/<id>/... hoặc ...?id=<id>)." });
  try {
    const store = load();
    const gmailApi = require("../utils/gmailApi");
    const accessToken = await gmailApi.getValidAccessToken(store);
    const dlResp = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
      headers: { Authorization: "Bearer " + accessToken },
    });
    if (!dlResp.ok) throw new Error("Drive trả lỗi " + dlResp.status + ": " + (await dlResp.text()).slice(0, 300));
    const arrayBuffer = await dlResp.arrayBuffer();
    const pdfBuffer = Buffer.from(arrayBuffer);
    const pdfParse = require("pdf-parse");
    const data = await pdfParse(pdfBuffer);
    const result = extractTienThueHopDongInfo(data.text || "");
    res.json({ success: true, ...result });
  } catch (e) {
    res.json({ success: false, error: "Không đọc được file (kiểm tra file có phải PDF, có chia sẻ quyền xem cho tài khoản Google đã kết nối chưa): " + e.message });
  }
});

// Tao 1 "diem" moi tu thong tin da doc/sua tren hop dong: (1) them vao Danh
// Muc Ma Cong Trinh (dung chung toan he thong cho doi soat), (2) them 1 ho so
// day du vao Phap danh > Hop Dong Thue Gian Hang (co san cot ben cho thue/tien
// thue/thoi han/link...). Ca 2 cho deu luu cung 1 link Drive vua doc.
router.post("/ho-so/tien-thue/tao-diem-tu-hop-dong", requireAdmin, express.json(), (req, res) => {
  try {
    const store = load();
    const company = getCompany(req);
    const {
      maCongTrinh, tenDiem, benChoThue, mstBenChoThue, tienThueThang,
      ngayKyHD, ngayBatDauHD, ngayHetHanHD, thoiHanHopDong, ghiChu, link,
    } = req.body || {};
    const ma = (maCongTrinh || "").trim();
    if (!ma) throw new Error("Thiếu Mã công trình.");
    const ten = (tenDiem || "").trim();

    // 1) Danh Muc Ma Cong Trinh (danh_muc_ma_cong_trinh_cu / _moi) -- xem
    // routes/danh-muc.js, cung 1 quy uoc storeKey + kiem tra trung ma.
    const dmKey = "danh_muc_ma_cong_trinh_" + (company === "kh_moi" ? "moi" : "cu");
    if (!Array.isArray(store[dmKey])) store[dmKey] = [];
    const dup = store[dmKey].find((r) => (r.ma || "").toLowerCase() === ma.toLowerCase());
    if (dup) throw new Error(`Mã công trình "${ma}" đã có sẵn trong Danh mục.`);
    store[dmKey].push({
      id: nextId(store),
      ma,
      ten: ten || "",
      ghiChu: "Tạo từ đọc hợp đồng Drive (Hồ Sơ > Tiền Thuê)",
      createdAt: new Date().toISOString(),
    });

    // 2) Phap Danh > Hop Dong Thue Gian Hang -- dung lai dung shape voi form
    // "them" cua routes/phap-danh.js (POST /phap-danh/hop-dong-thue-gian-hang).
    if (!Array.isArray(store.phap_danh_hop_dong_thue)) store.phap_danh_hop_dong_thue = [];
    const amt = tienThueThang ? Number(String(tienThueThang).replace(/[^\d]/g, "")) : 0;
    store.phap_danh_hop_dong_thue.push({
      id: nextId(store, "phap_danh_hop_dong_thue_seq") || Date.now(),
      loaiHinh: "",
      congTy: company,
      maDiemMisa: "",
      tenDiemNoiBo: ten,
      khuVuc: "",
      gian: ten || ma,
      maCongTrinh: ma,
      maKH: "",
      benChoThue: (benChoThue || "").trim(),
      mstBenChoThue: (mstBenChoThue || "").trim(),
      hinhThucHopTac: "",
      trangThaiHoatDong: "",
      thoiHanHopDong: (thoiHanHopDong || "").trim(),
      tienThueThang: amt,
      ghiChu: (ghiChu || "").trim(),
      ngayKyHD: (ngayKyHD || "").trim(),
      ngayBatDauHD: (ngayBatDauHD || "").trim(),
      ngayHetHanHD: (ngayHetHanHD || "").trim(),
      linkHopDongChuaDuDau: (link || "").trim(),
      createdAt: new Date().toISOString(),
      source: "Đọc từ Drive (Hồ Sơ > Tiền Thuê)",
    });

    save(store);
    res.json({ success: true, maCongTrinh: ma, tenDiem: ten });
  } catch (e) {
    res.json({ success: false, error: e.message });
  }
});

// POST /ho-so/dtcs-chiase/mall/bulk-upsert
// Body JSON: { records: [{diemId, thang, tongDT, tienThue, khoauTru, soTienMinhNhan, linkDoiSoat, ghiChu}] }
// Skip duplicates by (diemId, thang). Used by Cowork Claude Drive-sync skill.
router.post("/ho-so/dtcs-chiase/mall/bulk-upsert", requireAdmin, express.json(), (req, res) => {
  const store = load();
  if (!store.dtcs_chiase_thang) store.dtcs_chiase_thang = [];
  const records = req.body && Array.isArray(req.body.records) ? req.body.records : [];
  let added = 0;
  let skipped = 0;
  records.forEach(r => {
    const diemId = String(r.diemId || "").trim();
    const thang  = String(r.thang  || "").trim();
    if (!diemId || !thang) { skipped++; return; }
    const exists = store.dtcs_chiase_thang.some(
      t => String(t.diemId) === diemId && String(t.thang) === thang
    );
    if (exists) { skipped++; return; }
    store.dtcs_chiase_thang.push({
      id: nextId(store),
      diemId,
      thang,
      tongDT:          String(r.tongDT          || "").trim(),
      tienThue:        String(r.tienThue        || "").trim(),
      khoauTru:        String(r.khoauTru        || "").trim(),
      soTienMinhNhan:  String(r.soTienMinhNhan  || "").trim(),
      linkDoiSoat:     String(r.linkDoiSoat     || "").trim(),
      ghiChu:          String(r.ghiChu          || "").trim(),
      createdAt: new Date().toISOString(),
    });
    added++;
  });
  save(store);
  res.json({ success: true, added, skipped });
});

module.exports = router;
