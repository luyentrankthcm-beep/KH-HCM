const XLSX = require("xlsx");

// "Chi phi" (money PAID OUT to vendors/landlords/staff) reconciliation --
// mirror of Momo/ZVP/Viet QR's revenue reconciliation, but for the OUTGOING
// side: 2 dedicated "chi phi" bank accounts (BIDV8651 + VPBank9997) whose raw
// statement exports get parsed here, each debit ("chi ra") row tagged with a
// vendor (NCC) and, where UNAMBIGUOUSLY possible, a Ma cong trinh (gian) for
// rent-style payments and a Ma NCC (tra theo danh sach NCC "KH cu"), then
// exported to the real 57-cot "MISAKHCU" (Phieu chi tien gui AMIS) layout.
//
// Unlike the revenue channels, there is no separate "hoa don" list to match
// against here -- Luyen's own decision (2026-07-16) was to identify vendors
// straight off the bank's own "Ten doi ung" column, and to ONLY auto-assign
// Ma cong trinh when the site name in the description matches EXACTLY ONE
// gian in the existing list -- anything ambiguous or unrecognized is left
// blank for her to fill in by hand rather than guessed at, since a wrong Ma
// cong trinh on a real accounting export is worse than a blank one. Ma NCC
// (2026-07-16, file "Chi phi KH NCC T7 2026.xlsx") follows the SAME
// safe-by-default rule: chi gan khi ten doi ung khop DUNG 1 NCC trong danh
// sach, con lai de trong cho gan tay.
// data/chi-phi-ncc-list.json chi la danh sach NCC MAC DINH ban dau (import tu
// file Excel "Chi phi KH NCC T7 2026.xlsx") -- danh sach NCC THAT, cap nhat
// lien tuc, nam trong store.chi_phi_ncc_list (file store.json, xem
// routes/doisoat-chiphi.js). Thu muc data/ khong duoc dua len Git (chua du
// lieu tai chinh thuc), nen tren moi may/moi ban deploy moi (chua tung chay
// qua) file nay co the chua ton tai -- doc an toan bang try/catch, fallback
// ve mang rong, de app khong bi crash luc khoi dong. Sau khi phuc hoi du lieu
// thuc qua trang "Sao luu" (store.chi_phi_ncc_list), danh sach nay khong con
// duoc dung nua (chi la fallback khi store rong).
let NCC_LIST = [];
try {
  NCC_LIST = require("../data/chi-phi-ncc-list.json");
} catch (e) {
  console.error(
    "[chiphiReconcile] Khong tim thay data/chi-phi-ncc-list.json -- dung danh sach NCC mac dinh RONG. " +
      "Se duoc thay the boi du lieu thuc sau khi phuc hoi ban sao luu qua trang Sao luu."
  );
}

function removeDiacritics(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}

function normText(s) {
  return removeDiacritics(String(s || ""))
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function normHeader(s) {
  return normText(s);
}

// ---------- Ma NCC (vendor code) lookup, tra theo danh sach "NCC KH cu" ----------
// Mot vai chu viet tat pho bien tren sao ke ngan hang (CT/CTCP/CP/SX/TM/XNK/DV)
// duoc mo rong truoc khi so khop, vi danh sach NCC luon ghi ten day du.
const ABBREV_MAP = {
  ctcp: "cong ty co phan",
  ct: "cong ty",
  cp: "co phan",
  sx: "san xuat",
  tm: "thuong mai",
  xnk: "xuat nhap khau",
  dv: "dich vu",
  kd: "kinh doanh",
  tp: "thanh pho",
  vh: "van hanh",
};
function expandAbbrev(normStr) {
  return normStr
    .split(" ")
    .map((tok) => ABBREV_MAP[tok] || tok)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

// Danh sach NCC gio co the duoc CHI upload lai qua UI (luu store.chi_phi_ncc_list)
// nen index phai duoc dung ra tu danh sach truyen vao, khong con co dinh 1 lan
// luc load module nua. buildNccIndex() goi 1 lan/luot xu ly (buildChannelChiPhi),
// matchNccForVendor() nhan index da dung san.
function buildNccIndex(nccList) {
  const index = new Map();
  for (const rec of nccList || []) {
    const key = normText(rec.tenKhongDau || rec.tenNCC);
    if (!key) continue;
    if (!index.has(key)) index.set(key, []);
    index.get(key).push(rec);
  }
  return index;
}

// Chi tra ve ket qua khi khop DUNG 1 NCC (thu ten nguyen ban truoc, roi thu
// ban da mo rong chu viet tat) -- neu trung ten voi NHIEU NCC (vd nhieu chi
// nhanh cua cung 1 cong ty) hoac khong tim thay, tra ve rong + candidates de
// hien thi "can gan tay" thay vi doan mo.
function matchNccForVendor(tenDoiUng, nccIndex) {
  if (!tenDoiUng) return { maNCC: null, tenNCC: null, mstNCC: null, candidates: [] };
  const raw = normText(tenDoiUng);
  if (!raw) return { maNCC: null, tenNCC: null, mstNCC: null, candidates: [] };
  const expanded = expandAbbrev(raw);

  let recs = nccIndex.get(raw) || [];
  if (recs.length === 0 && expanded !== raw) recs = nccIndex.get(expanded) || [];

  if (recs.length === 1) {
    return { maNCC: recs[0].maNCC, tenNCC: recs[0].tenNCC, mstNCC: recs[0].mst, candidates: [recs[0].maNCC] };
  }
  return { maNCC: null, tenNCC: null, mstNCC: null, candidates: recs.map((r) => r.maNCC) };
}

// Tra 1 NCC theo DUNG Ma NCC (dung khi Luyen go tay ma NCC thu cong -- lay lai
// Ten NCC/MST de hien thi/xuat cho dep, khong bat buoc phai co trong danh sach).
function findNccByCode(maNCC, nccList) {
  if (!maNCC) return null;
  const code = String(maNCC).trim();
  return (nccList || []).find((r) => r.maNCC === code) || null;
}

// Parse file danh sach NCC (cot: Ma nha cung cap / Ten nha cung cap / Ma so
// thue hoac CCCD / Ten khong dau) -- do header linh hoat (accent/case-
// insensitive substring) vi Luyen co the tai lai file voi ten cot hoi khac.
const NCC_CODE_PATTERNS = ["ma nha cung cap", "ma ncc"];
const NCC_NAME_PATTERNS = ["ten nha cung cap", "ten ncc"];
const NCC_MST_PATTERNS = ["ma so thue", "cccd", "mst"];
const NCC_NODIACRITIC_PATTERNS = ["ten khong dau"];

function findNccHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 20); r++) {
    const row = grid[r] || [];
    const cols = { code: -1, name: -1, mst: -1, nodiacritic: -1 };
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = normHeader(cell);
      if (cols.code === -1 && NCC_CODE_PATTERNS.some((p) => h.includes(p))) cols.code = c;
      if (cols.name === -1 && NCC_NAME_PATTERNS.some((p) => h.includes(p))) cols.name = c;
      if (cols.mst === -1 && NCC_MST_PATTERNS.some((p) => h.includes(p))) cols.mst = c;
      if (cols.nodiacritic === -1 && NCC_NODIACRITIC_PATTERNS.some((p) => h.includes(p))) cols.nodiacritic = c;
    });
    if (cols.code !== -1 && cols.name !== -1) {
      return { headerRowIdx: r, ...cols };
    }
  }
  return null;
}

function parseNccWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const found = findNccHeaderRow(grid);
    if (!found) continue;
    const { headerRowIdx, code, name, mst, nodiacritic } = found;
    const list = [];
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const maNCC = row[code] !== null && row[code] !== undefined ? String(row[code]).trim() : "";
      const tenNCC = row[name] !== null && row[name] !== undefined ? String(row[name]).trim() : "";
      if (!maNCC && !tenNCC) continue;
      const mstVal = mst >= 0 && row[mst] !== null && row[mst] !== undefined ? String(row[mst]).trim() : "";
      const tenKhongDau =
        nodiacritic >= 0 && row[nodiacritic] ? normText(row[nodiacritic]) : normText(tenNCC);
      list.push({ maNCC, tenNCC, mst: mstVal, tenKhongDau });
    }
    if (list.length > 0) return list;
  }
  throw new Error(
    'Khong nhan dien duoc danh sach NCC trong file nay: can co cot "Ma nha cung cap" va "Ten nha cung cap".'
  );
}

function excelSerialToIso(serial) {
  const days = Math.round(serial);
  const utcMillis = (days - 25569) * 86400 * 1000;
  return new Date(utcMillis).toISOString().slice(0, 10);
}

function parseDateCell(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return excelSerialToIso(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const [, y, mo, d] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return null;
}

function parseNumberCell(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// Header keyword sets -- accent/case-insensitive substring match against
// each normalized header cell. Two known real layouts so far (ACB-style
// "Số tiền ghi nợ/có" vs BIDV-style "Phát sinh nợ/có"), written generically
// so a similar export from a 3rd bank should also be picked up.
const DATE_PATTERNS = ["ngay gia tri", "ngay giao dich", "ngay hieu luc"];
const DEBIT_PATTERNS = ["so tien ghi no", "phat sinh no"];
const CREDIT_PATTERNS = ["so tien ghi co", "phat sinh co"];
const DESC_PATTERNS = ["noi dung giao dich", "dien giai", "noi dung"];
const VENDOR_PATTERNS = ["ten doi ung"];
const VENDOR_ACCOUNT_PATTERNS = ["tai khoan doi ung", "tai khoan doi u"];
const VENDOR_BANK_PATTERNS = ["ngan hang doi ung"];
const DOC_NO_PATTERNS = ["so but toan", "so chung tu", "ma giao dich"];

function findHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 20); r++) {
    const row = grid[r] || [];
    const cols = { date: -1, debit: -1, credit: -1, desc: -1, vendor: -1, vendorAccount: -1, vendorBank: -1, docNo: -1 };
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = normHeader(cell);
      if (cols.date === -1 && DATE_PATTERNS.some((p) => h.includes(p))) cols.date = c;
      if (cols.debit === -1 && DEBIT_PATTERNS.some((p) => h.includes(p))) cols.debit = c;
      if (cols.credit === -1 && CREDIT_PATTERNS.some((p) => h.includes(p))) cols.credit = c;
      if (cols.desc === -1 && DESC_PATTERNS.some((p) => h.includes(p))) cols.desc = c;
      if (cols.vendor === -1 && VENDOR_PATTERNS.some((p) => h.includes(p))) cols.vendor = c;
      if (cols.vendorAccount === -1 && VENDOR_ACCOUNT_PATTERNS.some((p) => h.includes(p))) cols.vendorAccount = c;
      if (cols.vendorBank === -1 && VENDOR_BANK_PATTERNS.some((p) => h.includes(p))) cols.vendorBank = c;
      if (cols.docNo === -1 && DOC_NO_PATTERNS.some((p) => h.includes(p))) cols.docNo = c;
    });
    if (cols.date !== -1 && cols.debit !== -1 && cols.credit !== -1) {
      return { headerRowIdx: r, ...cols };
    }
  }
  return null;
}

// Parse EVERY sheet in the workbook that looks like a raw bank statement
// (has a recognizable date + debit + credit column trio) -- Luyen's monthly
// export may contain either 1 or both accounts' sheets in the same file, so
// we don't hard-code sheet names/positions.
function parseChiPhiRawWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheets = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const found = findHeaderRow(grid);
    if (!found) continue; // sheet khong phai sao ke (vd sheet trong "Sheet2")

    const { headerRowIdx, date, debit, credit, desc, vendor, vendorAccount, vendorBank, docNo } = found;
    const rows = [];
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const dateIso = parseDateCell(row[date]);
      if (!dateIso) continue;
      const debitAmt = parseNumberCell(row[debit]);
      const creditAmt = parseNumberCell(row[credit]);
      if (debitAmt === 0 && creditAmt === 0) continue;
      rows.push({
        date: dateIso,
        debit: debitAmt,
        credit: creditAmt,
        description: desc >= 0 ? String(row[desc] || "").trim() : "",
        tenDoiUng: vendor >= 0 ? String(row[vendor] || "").trim() : "",
        taiKhoanDoiUng: vendorAccount >= 0 ? String(row[vendorAccount] || "").trim() : "",
        nganHangDoiUng: vendorBank >= 0 ? String(row[vendorBank] || "").trim() : "",
        docNo: docNo >= 0 ? String(row[docNo] || "").trim() : "",
      });
    }
    if (rows.length > 0) sheets.push({ sheetName, rows });
  }
  if (sheets.length === 0) {
    throw new Error(
      'Khong nhan dien duoc sheet sao ke nao trong file nay: can co cot ngay giao dich va 2 cot so tien (ghi no/ghi co hoac phat sinh no/co).'
    );
  }
  return sheets;
}

// Vendor identity: Luyen's explicit decision (2026-07-16) is to use the
// bank's own "Ten doi ung" column as-is, with NO separate Ma NCC mapping
// table for now. Rows that don't carry that column at all (some export
// layouts only have free-text "Noi dung giao dich", no structured
// counterparty field) are left with vendor = "" and flagged so they show up
// clearly for manual review instead of silently mis-labeled.
function resolveVendor(row) {
  if (row.tenDoiUng) return { vendor: row.tenDoiUng, vendorKnown: true };
  return { vendor: "", vendorKnown: false };
}

// Detect a "site name" fragment from a rent-style description -- these
// consistently follow a mall-brand keyword (VC/VINCOM/AEON/VINPEARL/...)
// immediately before an "HD <ma hop dong>" token, e.g.:
//   "... VC Royal HD VCRCP.2308.2025.POSH.KH ..."
//   "... VINCOM PHAM HUNG  HD VCRVH10012026.KvaH..."
// Returns null if the description doesn't look like a rent/site payment at
// all (no such keyword pair found) -- other expense types (goods, salary,
// tax, insurance, bank fees) are never even attempted for Ma cong trinh.
const SITE_BRAND_KEYWORDS = /\b(VC|VINCOM|VINPEARL|AEON|AEONMALL|LOTTE|BIG ?C|ESPACE)\b/i;
function extractSiteFragment(description) {
  const m = String(description || "").match(
    /\b(?:VC|VINCOM|VINPEARL|AEON(?:MALL)?|LOTTE|BIG ?C|ESPACE)\s*[:.\-]?\s*([A-Za-zÀ-ỹ0-9 ]{2,40}?)\s+HD\b/i
  );
  if (!m) return null;
  const frag = m[1].trim();
  if (frag.length < 2) return null;
  return frag;
}

// Match a site-name fragment against the known gian list (reusing the SAME
// list revenue reconciliation already maintains, store.zvp_gian_list) --
// ONLY returns a match when EXACTLY ONE gian's name contains the fragment
// (or the fragment contains the gian's own distinguishing keyword), in
// either direction, so short/generic fragments never silently pick a wrong
// site out of several plausible candidates.
function matchGianForFragment(fragment, gianList) {
  if (!fragment) return { maCongTrinh: null, candidates: [] };
  const fragNorm = normText(fragment);
  if (!fragNorm || fragNorm.length < 3) return { maCongTrinh: null, candidates: [] };

  const candidates = new Set();
  for (const g of gianList || []) {
    const codeNorm = normText(g.maCongTrinh);
    const nameNorm = normText(g.tenDiem);
    if (!codeNorm && !nameNorm) continue;
    if ((codeNorm && (codeNorm.includes(fragNorm) || fragNorm.includes(codeNorm))) ||
        (nameNorm && (nameNorm.includes(fragNorm) || fragNorm.includes(nameNorm)))) {
      candidates.add(g.maCongTrinh);
    }
  }
  const list = Array.from(candidates);
  return { maCongTrinh: list.length === 1 ? list[0] : null, candidates: list };
}

function isoToDmy(iso) {
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

// ---------- Danh sach hoa don NCC (sheet "HDDV KH CU"/"HDDV KH Moi") ----------
// Hoa don NCC XUAT cho K&H (mua hang/dich vu). Dung de doi chieu Ten nguoi
// ban (hoac MST) + So tien voi tung dong chi phi (bank/UNC), tu dien So hoa
// don/Ngay hoa don... vao file xuat Misa thay vi de trong toan bo cho Luyen
// tra tay tung dong -- quyet dinh 2026-07-16 cua Luyen (chi lam truoc cho
// sheet "HDDV KH CU", danh cho 2 kenh chi phi KH cu hien co: BIDV8651/VPBank9997).
const HDDV_SOHD_PATTERNS = ["so hoa don"];
const HDDV_NGAYLAP_PATTERNS = ["ngay lap"];
const HDDV_KYHIEUMAU_PATTERNS = ["ky hieu mau so"];
const HDDV_KYHIEUHD_PATTERNS = ["ky hieu hoa don"];
const HDDV_TENBAN_PATTERNS = ["ten nguoi ban"];
const HDDV_MSTBAN_PATTERNS = ["mst nguoi ban"];
const HDDV_TONGCHUATHUE_PATTERNS = ["tong tien chua thue"];
const HDDV_TONGTHUE_PATTERNS = ["tong tien thue"];
const HDDV_TONGTT_PATTERNS = ["tong tien thanh toan"];

function findHddvHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = grid[r] || [];
    const cols = {};
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = normHeader(cell);
      if (cols.soHd === undefined && HDDV_SOHD_PATTERNS.some((p) => h.includes(p))) cols.soHd = c;
      if (cols.ngayLap === undefined && HDDV_NGAYLAP_PATTERNS.some((p) => h.includes(p))) cols.ngayLap = c;
      if (cols.kyHieuMau === undefined && HDDV_KYHIEUMAU_PATTERNS.some((p) => h.includes(p))) cols.kyHieuMau = c;
      if (cols.kyHieuHd === undefined && HDDV_KYHIEUHD_PATTERNS.some((p) => h.includes(p))) cols.kyHieuHd = c;
      if (cols.tenBan === undefined && HDDV_TENBAN_PATTERNS.some((p) => h.includes(p))) cols.tenBan = c;
      if (cols.mstBan === undefined && HDDV_MSTBAN_PATTERNS.some((p) => h.includes(p))) cols.mstBan = c;
      if (cols.tongChuaThue === undefined && HDDV_TONGCHUATHUE_PATTERNS.some((p) => h.includes(p))) cols.tongChuaThue = c;
      if (cols.tongThue === undefined && HDDV_TONGTHUE_PATTERNS.some((p) => h.includes(p)) && !h.includes("chua")) cols.tongThue = c;
      if (cols.tongTt === undefined && HDDV_TONGTT_PATTERNS.some((p) => h.includes(p))) cols.tongTt = c;
    });
    if (cols.soHd !== undefined && cols.tenBan !== undefined && cols.tongTt !== undefined) {
      return { headerRowIdx: r, ...cols };
    }
  }
  return null;
}

// Uu tien sheet co ten chua "hddv" (danh sach hoa don dich vu NCC ban cho
// K&H) -- neu file co nhieu sheet (KH cu/KH moi/MISAKHCU/NCC...) thi chi lay
// dung sheet hoa don, khong lay nham sheet khac co cau truc tuong tu.
function parseHddvInvoiceWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const preferred = wb.SheetNames.filter((n) => normHeader(n).includes("hddv"));
  const order = [...preferred, ...wb.SheetNames.filter((n) => !preferred.includes(n))];
  for (const sheetName of order) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const found = findHddvHeaderRow(grid);
    if (!found) continue;
    const { headerRowIdx, soHd, ngayLap, kyHieuMau, kyHieuHd, tenBan, mstBan, tongChuaThue, tongThue, tongTt } = found;
    const list = [];
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const tenNguoiBan = tenBan !== undefined ? String(row[tenBan] || "").trim() : "";
      if (!tenNguoiBan) continue;
      const soHoaDonRaw = soHd !== undefined ? row[soHd] : null;
      if (soHoaDonRaw === null || soHoaDonRaw === undefined || soHoaDonRaw === "") continue;
      list.push({
        soHoaDon: String(soHoaDonRaw).trim(),
        kyHieuMauSo: kyHieuMau !== undefined ? String(row[kyHieuMau] || "").trim() : "",
        kyHieuHoaDon: kyHieuHd !== undefined ? String(row[kyHieuHd] || "").trim() : "",
        ngayLap: ngayLap !== undefined ? parseDateCell(row[ngayLap]) : null,
        tenNguoiBan,
        mstNguoiBan: mstBan !== undefined ? String(row[mstBan] || "").trim() : "",
        tongTienChuaThue: tongChuaThue !== undefined ? parseNumberCell(row[tongChuaThue]) : 0,
        tongTienThue: tongThue !== undefined ? parseNumberCell(row[tongThue]) : 0,
        tongTienThanhToan: tongTt !== undefined ? parseNumberCell(row[tongTt]) : 0,
      });
    }
    if (list.length > 0) return { sheetName, list };
  }
  throw new Error(
    'Khong nhan dien duoc danh sach hoa don trong file nay: can co cot "So hoa don", "Ten nguoi ban" va "Tong tien thanh toan".'
  );
}

// Ca ten nguoi ban tren hoa don HDDV VA ten doi ung tren sao ke/UNC deu co
// the viet tat kieu doanh nghiep khac nhau ("CONG TY CO PHAN..." vs "CONG TY
// CP..."), da xac nhan tren du lieu thuc te (BLUECOM: hoa don ghi "CO PHAN",
// UNC ghi "CP") -- danh index duoi CA 2 dang (nguyen ban + mo rong viet tat)
// cho moi ban ghi, roi khi tra cuu cung thu ca 2 dang cua ten tim kiem, giong
// HET nguyen tac matchNccForVendor da dung, de khop duoc bat ke ben nao viet
// tat hay viet day du.
function addToNameIndex(map, key, rec) {
  if (!key) return;
  if (!map.has(key)) map.set(key, []);
  const list = map.get(key);
  if (!list.includes(rec)) list.push(rec);
}

function buildInvoiceIndex(invoiceList) {
  const byMst = new Map();
  const byName = new Map();
  for (const inv of invoiceList || []) {
    if (inv.mstNguoiBan) {
      const k = normText(inv.mstNguoiBan);
      if (!byMst.has(k)) byMst.set(k, []);
      byMst.get(k).push(inv);
    }
    const nameKey = normText(inv.tenNguoiBan);
    addToNameIndex(byName, nameKey, inv);
    const expanded = expandAbbrev(nameKey);
    if (expanded !== nameKey) addToNameIndex(byName, expanded, inv);
  }
  return { byMst, byName };
}

// Chi tra ve hoa don khi khop DUNG 1 hoa don (theo MST hoac ten NCC, VA so
// tien phai khop CHINH XAC (sai so < 500d cho lam tron) voi Tong tien thanh
// toan) -- trung ten/MST voi nhieu hoa don CUNG so tien (hiem nhung co the,
// vd 2 don hang gia tri giong nhau) duoc coi la khong the phan biet, tra ve
// rong cho Luyen tu kiem tra thay vi doan sai hoa don.
function matchInvoiceForPayment(vendorMst, vendorName, amount, invoiceIndex) {
  let candidates = [];
  const mstKey = vendorMst ? normText(vendorMst) : "";
  if (mstKey && invoiceIndex.byMst.has(mstKey)) candidates = invoiceIndex.byMst.get(mstKey);
  if (candidates.length === 0 && vendorName) {
    const raw = normText(vendorName);
    candidates = invoiceIndex.byName.get(raw) || [];
    if (candidates.length === 0) {
      const expanded = expandAbbrev(raw);
      if (expanded !== raw) candidates = invoiceIndex.byName.get(expanded) || [];
    }
  }
  if (candidates.length === 0) return null;
  const hits = candidates.filter((inv) => Math.abs(inv.tongTienThanhToan - amount) < 500);
  return hits.length === 1 ? hits[0] : null;
}

// ---------- Bang lenh chi UNC (uy nhiem chi) ----------
// Luyen tu lap bang nay khi xin thanh toan tung khoan (Google Sheet rieng,
// tai xuong .xlsx roi tai len day). Cot "Noi dung tren UNC"/"Noi dung unc"
// thuong la mo ta SACH hon nhieu so voi dong sao ke ngan hang thuc te (vd co
// san "theo so HD <n>"), nen duoc uu tien dung lam Dien giai cho file xuat
// Misa khi khop duoc voi 1 dong chi phi. Header linh hoat vi Luyen dang dung
// nhieu file/tab hoi khac ten cot nhau (vd "Ten cong ty" vs "Ten don vi thu huong").
const UNC_DATE_PATTERNS = ["ngay di tien", "ngay xin payment", "ngay"];
const UNC_VENDOR_PATTERNS = ["ten cong ty", "ten don vi thu huong", "ten don vi nhan"];
const UNC_AMOUNT_PATTERNS = ["so tien"];
const UNC_CONTENT_PATTERNS = ["noi dung tren unc", "noi dung unc"];
const UNC_ACCOUNT_PATTERNS = ["tai khoan nguoi thu huong", "so tk"];
const UNC_BANK_PATTERNS = ["ngan hang nguoi thu huong", "ngan hang"];

// Yeu cau header ro rang cho vendor + content -- amount CHUA bat buoc o day,
// vi it nhat 1 layout thuc te cua Luyen ("KVC MB") de trong tieu de cot So
// tien (chi co so lieu o dong du lieu, khong co chu "So tien" nao ca) --
// parseUncWorkbook() se tu do vi tri cot so tien bang cach xem cot ngay ben
// trai cot Ten don vi thu huong khi khong tim duoc qua header.
function findUncHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = grid[r] || [];
    const cols = {};
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = normHeader(cell);
      if (cols.vendor === undefined && UNC_VENDOR_PATTERNS.some((p) => h.includes(p))) cols.vendor = c;
      if (cols.amount === undefined && UNC_AMOUNT_PATTERNS.some((p) => h.includes(p))) cols.amount = c;
      if (cols.content === undefined && UNC_CONTENT_PATTERNS.some((p) => h.includes(p))) cols.content = c;
      if (cols.date === undefined && UNC_DATE_PATTERNS.some((p) => h.includes(p))) cols.date = c;
      if (cols.account === undefined && UNC_ACCOUNT_PATTERNS.some((p) => h.includes(p))) cols.account = c;
      if (cols.bank === undefined && UNC_BANK_PATTERNS.some((p) => h.includes(p))) cols.bank = c;
    });
    if (cols.vendor !== undefined && cols.content !== undefined) {
      return { headerRowIdx: r, ...cols };
    }
  }
  return null;
}

// Doan vi tri cot So tien khi header khong ghi ro: thu cot ngay ben trai cot
// Ten don vi thu huong -- neu phan lon (>= 70%) cac dong du lieu o cot do doc
// duoc thanh 1 so duong, coi day la cot So tien. Chi doan 1 cot DUY NHAT ngay
// ben trai (khong do xa hon) de tranh nham voi cac cot khac.
function guessUncAmountCol(grid, headerRowIdx, vendorCol) {
  const candidate = vendorCol - 1;
  if (candidate < 0) return undefined;
  let checked = 0;
  let numeric = 0;
  for (let r = headerRowIdx + 1; r < grid.length && checked < 30; r++) {
    const v = (grid[r] || [])[candidate];
    if (v === null || v === undefined || v === "") continue;
    checked++;
    if (parseNumberCell(v) > 0) numeric++;
  }
  if (checked >= 5 && numeric / checked >= 0.7) return candidate;
  return undefined;
}

// Quet TAT CA sheet trong file (Luyen co the tai file co nhieu tab thang) --
// gop chung tat ca dong tim duoc thanh 1 danh sach.
function parseUncWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const list = [];
  let anySheet = null;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const found = findUncHeaderRow(grid);
    if (!found) continue;
    const { headerRowIdx, vendor, content, date, account, bank } = found;
    const amount = found.amount !== undefined ? found.amount : guessUncAmountCol(grid, headerRowIdx, vendor);
    if (amount === undefined) continue;
    anySheet = anySheet ? `${anySheet}, ${sheetName}` : sheetName;
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const tenCongTy = String(row[vendor] || "").trim();
      const soTien = parseNumberCell(row[amount]);
      if (!tenCongTy || !soTien) continue;
      list.push({
        ngayDiTien: date !== undefined ? parseDateCell(row[date]) : null,
        tenCongTy,
        soTien,
        noiDungUnc: content !== undefined ? String(row[content] || "").trim() : "",
        soTk: account !== undefined ? String(row[account] || "").trim() : "",
        nganHang: bank !== undefined ? String(row[bank] || "").trim() : "",
      });
    }
  }
  if (list.length === 0) {
    throw new Error(
      'Khong nhan dien duoc bang lenh chi UNC trong file nay.'
    );
  }
  return { sheetName: anySheet, list };
}

function buildUncIndex(uncList) {
  const byName = new Map();
  for (const u of uncList || []) {
    const key = normText(u.tenCongTy);
    if (!key) continue;
    addToNameIndex(byName, key, u);
    const expanded = expandAbbrev(key);
    if (expanded !== key) addToNameIndex(byName, expanded, u);
  }
  return byName;
}

// Cung nguyen tac "chi tra ve khi khop DUNG 1" nhu matchInvoiceForPayment, va
// cung thu ca dang viet tat/mo rong cua ten tim kiem (xem addToNameIndex).
function matchUncForPayment(vendorName, amount, uncIndex) {
  if (!vendorName) return null;
  const raw = normText(vendorName);
  let candidates = uncIndex.get(raw) || [];
  if (candidates.length === 0) {
    const expanded = expandAbbrev(raw);
    if (expanded !== raw) candidates = uncIndex.get(expanded) || [];
  }
  if (candidates.length === 0) return null;
  const hits = candidates.filter((u) => Math.abs(u.soTien - amount) < 500);
  return hits.length === 1 ? hits[0] : null;
}

// Rut so hoa don ngay tu 1 doan van ban khi KHONG khop duoc voi danh sach hoa
// don co cau truc (hoa don cu chua co trong danh sach, hoac day chi la 1
// dong chi khac khong xuat hoa don chinh thuc) -- chi bat cac cum "HD <so>"/
// "hoa don <so>" nhu Luyen tu ghi tren dien giai UNC/sao ke thuc te ("theo so
// HD 770", "theo HD so 4426", "theo hoa don 9342"...). Chi tra ve so hoa don,
// KHONG doan them ngay/ky hieu hoa don di kem -- nhung truong do de trong.
const HD_NUMBER_RE = /\b(?:HD|h[oó]a\s*[dđ]on)\b\s*(?:s[oố])?\s*[:.\-]?\s*(\d{3,7})\b/i;
function extractHdNumberFromText(text) {
  if (!text) return null;
  const m = String(text).match(HD_NUMBER_RE);
  return m ? m[1] : null;
}

module.exports = {
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
  normText,
  DEFAULT_NCC_LIST: NCC_LIST,
};
