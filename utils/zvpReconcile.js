// ---------- Doi soat Zalo App / VNPay (Online + Offline) / Payoo ----------
// Same overall philosophy as utils/momoReconcile.js: parse whatever real
// export files Luyen uploads, never guess a mapping silently -- anything
// that can't be confidently matched is surfaced as "CHUA MAP: ..." so she
// can fix it once, rather than misattributing revenue to the wrong Ma cong
// trinh (which would produce wrong tax/accounting entries).
//
// Bank account: ACB31268 (so TK 12131268, sheet "NH1268" in her files).
// On the bank statement this account receives 3 kinds of settlement, each
// identifiable purely from the transaction description:
//   - "...VNPAY ... DV CTT NGAY <ngay>"        -> Online (Zalo mini app checkout via VNPay)
//   - "...VNPAY ... DV QR OFFLINE NGAY <ngay>" -> Offline (walk-up QR at a physical diem)
//   - "...PAYOO ... TT TD NGAY <ngay>..."       -> Payoo (card + QR, may be 2 lines per ngay)
// <ngay> is either a single date (dd.mm.yy) or an inclusive range
// (dd-dd.mm.yy, dd.mm-dd.mm.yy, or dd.mm_dd.mm.yyyy for Payoo) -- weekends
// commonly get consolidated into one multi-day settlement the following
// business day, exactly like Momo.
//
// CSE handling: any gian tagged "CSE" in "Hinh thuc hop tac" (on the shared
// "gian hang xuat HD" list, and mirrored on the invoice list's own "Hinh
// thuc hop tac" column) is doanh thu chia se -> TK Co 1388, tracked as a
// SEPARATE Ma cong trinh internally by appending momoReconcile's "__FF"
// suffix -- exactly the same convention already used for FF SC VIVO under
// Momo, just generalized here to any gian instead of one hardcoded parent.

const XLSX = require("xlsx");
const m = require("./momoReconcile");
const toIsoDate = m.toIsoDate;
const isoToDmy = m.isoToDmy;
const normCode = m.normCode;
const normText = m.normText;
const dateRange = m.dateRange;
const displayCode = m.displayCode;
const FF_SUFFIX = m.FF_SUFFIX;
const buildPendingDaySettlements = m.buildPendingDaySettlements;

// ---------- Settlement extraction from bank statement (ACB31268) ----------

function pad2(s) {
  return String(s).padStart(2, "0");
}

function parseVnpayNgayExpr(expr) {
  let mm = expr.match(/^(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (mm) {
    const [, d1, mo1, d2, mo2, yy] = mm;
    const yyyy = yy.length === 2 ? "20" + yy : yy;
    return { fromIso: `${yyyy}-${pad2(mo1)}-${pad2(d1)}`, toIso: `${yyyy}-${pad2(mo2)}-${pad2(d2)}` };
  }
  mm = expr.match(/^(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (mm) {
    const [, d1, d2, mo, yy] = mm;
    const yyyy = yy.length === 2 ? "20" + yy : yy;
    return { fromIso: `${yyyy}-${pad2(mo)}-${pad2(d1)}`, toIso: `${yyyy}-${pad2(mo)}-${pad2(d2)}` };
  }
  mm = expr.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (mm) {
    const [, d, mo, yy] = mm;
    const yyyy = yy.length === 2 ? "20" + yy : yy;
    const iso = `${yyyy}-${pad2(mo)}-${pad2(d)}`;
    return { fromIso: iso, toIso: iso };
  }
  return null;
}

function parsePayooNgayExpr(expr) {
  let mm = expr.match(/^(\d{1,2})\.(\d{1,2})_(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (mm) {
    const [, d1, mo1, d2, mo2, yyyy] = mm;
    return { fromIso: `${yyyy}-${pad2(mo1)}-${pad2(d1)}`, toIso: `${yyyy}-${pad2(mo2)}-${pad2(d2)}` };
  }
  mm = expr.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (mm) {
    const [, d, mo, yyyy] = mm;
    const iso = `${yyyy}-${pad2(mo)}-${pad2(d)}`;
    return { fromIso: iso, toIso: iso };
  }
  return null;
}

function extractZvpSettlements(transactions) {
  const onlineByRange = {};
  const offlineByRange = {};
  const payooByRange = {};

  // Luyen, 2026-07-21: "sao zalo app offline bị double lên 2 lần vậy" -- VNPay
  // thinh thoang tra tien cua CUNG 1 ngay doanh thu qua 2 giao dich ngan hang
  // rieng (vd 1 khoan chinh + 1 khoan dieu chinh/bo sung, cung ghi "NGAY
  // 15.07.26" nhung so tien khac nhau). Truoc day moi giao dich VNPay (CTT/QR
  // OFFLINE) duoc coi la 1 "khoan ve" DOC LAP, nen 2 giao dich cung ngay ->
  // ca 2 deu tu keo doanh thu-theo-gian CHO NGAY DO, xuat ra file bi trung
  // lap y het (cung so HD, cung so tien, lap lai 2 lan). Fix bang cach GOM
  // theo (fromIso, toIso) VA CONG don so tien ngan hang lai -- giong dung
  // cach Payoo da lam ben duoi tu truoc -- de moi khoang ngay doanh thu chi
  // con DUY NHAT 1 "khoan ve" (voi tong tien ngan hang la tong ca 2 giao
  // dich), tranh xuat trung dong.
  function addToRange(map, t, fromIso, toIso) {
    const key = `${fromIso}|${toIso}`;
    if (!map[key]) {
      map[key] = { date: t.date, amount: 0, fromIso, toIso, txIds: [] };
    }
    map[key].amount += t.amount;
    map[key].txIds.push(t.id);
    if (t.date > map[key].date) map[key].date = t.date;
  }

  for (const t of transactions) {
    if (t.type !== "thu") continue;
    const desc = t.description || "";
    if (/VNPAY/i.test(desc)) {
      const mCtt = desc.match(/DV\s+CTT\s+NGAY\s+([0-9.\-_]+)/i);
      const mOff = desc.match(/DV\s+QR\s+OFFLINE\s+NGAY\s+([0-9.\-_]+)/i);
      if (mCtt) {
        const r = parseVnpayNgayExpr(mCtt[1]);
        if (r) addToRange(onlineByRange, t, r.fromIso, r.toIso);
        continue;
      }
      if (mOff) {
        const r = parseVnpayNgayExpr(mOff[1]);
        if (r) addToRange(offlineByRange, t, r.fromIso, r.toIso);
        continue;
      }
      continue;
    }
    if (/PAYOO/i.test(desc)) {
      const mPayoo = desc.match(/NGAY\s+([0-9._]+)/i);
      if (mPayoo) {
        const cleaned = mPayoo[1].replace(/\.+$/, "");
        const r = parsePayooNgayExpr(cleaned);
        if (r) addToRange(payooByRange, t, r.fromIso, r.toIso);
      }
    }
  }

  // "id" field tren cac ket qua nay chi con y nghia hien thi (khong dung de
  // join nua, xem txIds cho danh sach day du) -- giu lai id CUOI CUNG gop
  // vao de khong pha vo cho nao con doc s.id truc tiep.
  const finalize = (byRange) =>
    Object.values(byRange).map((s) => ({ ...s, id: s.txIds[s.txIds.length - 1] }));

  return { online: finalize(onlineByRange), offline: finalize(offlineByRange), payoo: Object.values(payooByRange) };
}

// Extract the day-list from a "Dich vu thu ho" cell. Unlike Momo's raw
// format ("momo 11,12", digits immediately after the tag), the real Zalo/
// VNPay/Payoo tags often have extra descriptive words in between, e.g.
// "ZALO MINI APP 19", "VNPAY CƠ SỞ 22", "Payoo QR 11,12" -- so instead of
// requiring digits right after the tag word, this scans the WHOLE string
// for every run of digits (joined by "," "+" or "/") and takes the LAST
// one, since the day list consistently appears at the end. Cells with no
// digits at all ("Zalo", "Vnpay cs") have no day info and are left
// unmatched (empty days) rather than guessed.
function extractDayList(text) {
  const groups = text.match(/\d+(?:\s*[,+\/]\s*\d+)*/g);
  if (!groups || groups.length === 0) return [];
  const lastGroup = groups[groups.length - 1];
  if (lastGroup.includes("/")) {
    const d = parseInt(lastGroup.split("/")[0], 10);
    return isNaN(d) ? [] : [d];
  }
  return lastGroup
    .split(/[,+]/)
    .map((x) => parseInt(x.trim(), 10))
    .filter((x) => !isNaN(x));
}

function parseInvoiceWorkbookByTag(buffer, tag) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const kdsCandidates = wbLite.SheetNames.filter((n) => /k.\s*ds\s*xu.t/i.test(n));
  const sheetName =
    kdsCandidates.find((n) => /989/.test(n)) ||
    kdsCandidates[0] ||
    wbLite.SheetNames.find((n) => {
      const t = normText(n);
      return t.includes("hoa don") || (t.includes("danh sach") && t.includes("don"));
    });
  if (!sheetName) {
    throw new Error("Khong tim thay sheet danh sach hoa don trong file.");
  }
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.soHd === undefined && s.includes("so hd")) idx.soHd = c;
      if (idx.ngayHd === undefined && s.includes("ngay hd")) idx.ngayHd = c;
      if (idx.thang === undefined && s.includes("thang")) idx.thang = c;
      if (idx.maDiem === undefined && s.includes("ma diem")) idx.maDiem = c;
      if (s.includes("tong tt hd") || (s.includes("tong") && s.includes("hd"))) idx.tongTt = c;
      if (idx.dichVuThuHo === undefined && s.includes("thu ho")) idx.dichVuThuHo = c;
      if (idx.hinhThuc === undefined && s.includes("hinh thuc hop tac")) idx.hinhThuc = c;
      if (idx.tenDiem === undefined && s.includes("ten diem xuat hoa don")) idx.tenDiem = c;
    });
    if (idx.soHd !== undefined && idx.maDiem !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de (can cot "So HD" va "Ma diem") trong sheet hoa don.');
  }

  const invoices = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const dvth = cols.dichVuThuHo !== undefined ? row[cols.dichVuThuHo] : null;
    if (!dvth || !new RegExp(tag, "i").test(String(dvth))) continue;
    const days = extractDayList(String(dvth).trim());
    const ngayHdRaw = cols.ngayHd !== undefined ? row[cols.ngayHd] : null;
    let maDiem = normCode(cols.maDiem !== undefined ? row[cols.maDiem] : null);
    const hinhThuc = cols.hinhThuc !== undefined ? String(row[cols.hinhThuc] || "").toUpperCase() : "";
    if (hinhThuc.includes("CSE") && !maDiem.endsWith(FF_SUFFIX)) {
      maDiem = maDiem + FF_SUFFIX;
    }

    invoices.push({
      soHd: cols.soHd !== undefined ? row[cols.soHd] : null,
      ngayHd: toIsoDate(ngayHdRaw),
      thang: cols.thang !== undefined ? row[cols.thang] : null,
      maDiem,
      // Ten diem xuat hoa don (raw site name on the invoice itself) -- kept
      // alongside maDiem so the caller can redirect an invoice's maDiem
      // through the SAME zvp_gian_list mapping used for Online revenue
      // (see applyGianRedirectToInvoices in routes/doisoat-zvp.js). Needed
      // because "Ma diem tren misa thue" on the invoice sheet is sometimes
      // just the site name typed in directly (not yet a real accounting
      // code) even when zvp_gian_list already has the correct redirect for
      // that exact site name -- e.g. an invoice for "GHOST BRIDE AE HUE"
      // carries maDiem "GHOST BRIDE AE HUE" even though zvp_gian_list
      // already maps that tenDiem to "AE HUE KVCN".
      tenDiem: cols.tenDiem !== undefined ? String(row[cols.tenDiem] || "").trim() : "",
      tongTt: cols.tongTt !== undefined ? Number(row[cols.tongTt]) || 0 : 0,
      days,
      raw: String(dvth).trim(),
    });
  }

  return { sheetName, invoices };
}

function parseOfflineVnpayWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => /du\s*lieu.*vnpay/i.test(normText(n))) ||
    wbLite.SheetNames.find((n) => normText(n).includes("du lieu"));
  if (!sheetName) {
    throw new Error('Khong tim thay sheet "du lieu VNpay co so ..." trong file.');
  }
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 6); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.grossCol === undefined && s.includes("so tien hach toan thu ho")) idx.grossCol = c;
      if (idx.netCol === undefined && s.includes("so tien sau khi tru phi")) idx.netCol = c;
      if (idx.dateCol === undefined && s.includes("ngay hach toan thu ho")) idx.dateCol = c;
      if (s === "trang thai") idx.statusCol = c;
      if (idx.diemCol === undefined && s.includes("ten diem xuat hoa don")) idx.diemCol = c;
    });
    if (idx.grossCol !== undefined && idx.dateCol !== undefined && idx.diemCol !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de trong sheet du lieu VNpay offline.');
  }

  const dates = new Set();
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  let matched = 0;
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const status = cols.statusCol !== undefined ? row[cols.statusCol] : null;
    if (status !== null && !/thanh cong|th.nh c.ng/i.test(String(status))) continue;
    const diem = cols.diemCol !== undefined ? String(row[cols.diemCol] || "").trim() : "";
    if (!diem) continue;
    const gross = cols.grossCol !== undefined ? Number(row[cols.grossCol]) || 0 : 0;
    if (!gross) continue;
    const net = cols.netCol !== undefined ? Number(row[cols.netCol]) || 0 : gross;
    const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
    const date = toIsoDate(dateRaw ? String(dateRaw).split(" ")[0] : null);
    if (!date) continue;
    matched++;
    dates.add(date);
    codes.add(diem);
    const key = `${date}|${diem}`;
    grossByCode[key] = (grossByCode[key] || 0) + gross;
    netByCode[key] = (netByCode[key] || 0) + net;
  }

  return { sheetName, dates: Array.from(dates).sort(), codes: Array.from(codes), grossByCode, netByCode, rowsMatched: matched };
}

function parseDiemMappingSheet(buffer, sheetNameHint) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => normText(n).includes(normText(sheetNameHint))) ||
    wbLite.SheetNames.find((n) => /gian\s*hang/i.test(n));
  if (!sheetName) return {};
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 6); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.tenDiem === undefined && s.includes("ten diem xuat hoa don")) idx.tenDiem = c;
      if (s.includes("ma diem tren misa") || s.includes("ma cong trinh")) idx.maCongTrinh = c;
      if (idx.hinhThuc === undefined && s.includes("hinh thuc hop tac")) idx.hinhThuc = c;
      if (s === "chi nhanh") idx.chiNhanh = c;
    });
    if (idx.tenDiem !== undefined && idx.maCongTrinh !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) return {};

  const map = {};
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const joinKey =
      cols.chiNhanh !== undefined
        ? String(row[cols.chiNhanh] || "").trim()
        : cols.tenDiem !== undefined
        ? String(row[cols.tenDiem] || "").trim()
        : "";
    if (!joinKey) continue;
    const maCongTrinh = normCode(cols.maCongTrinh !== undefined ? row[cols.maCongTrinh] : null);
    if (!maCongTrinh) continue;
    const hinhThuc = cols.hinhThuc !== undefined ? String(row[cols.hinhThuc] || "").toUpperCase() : "";
    map[joinKey] = { maCongTrinh, isCse: hinhThuc.includes("CSE") };
  }
  return map;
}

function parsePayooWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => /du\s*lieu.*payoo/i.test(normText(n))) ||
    wbLite.SheetNames.find((n) => normText(n).includes("payoo"));
  if (!sheetName) {
    throw new Error('Khong tim thay sheet "Du lieu Payoo ..." trong file.');
  }
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 6); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.grossCol === undefined && s.includes("so tien thanh toan")) idx.grossCol = c;
      if (idx.netCol === undefined && s.includes("thanh tien")) idx.netCol = c;
      if (s === "ngay giao dich") idx.dateCol = c;
      if (idx.invalidCol === undefined && s.includes("ly do gd khong hop le")) idx.invalidCol = c;
      if (s === "gian hang") idx.gianCol = c;
      if (s === "cua hang" && idx.gianColFallback === undefined) idx.gianColFallback = c;
    });
    if (idx.grossCol !== undefined && idx.dateCol !== undefined && (idx.gianCol !== undefined || idx.gianColFallback !== undefined)) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de trong sheet du lieu Payoo.');
  }
  const gianCol = cols.gianCol !== undefined ? cols.gianCol : cols.gianColFallback;

  const dates = new Set();
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  let matched = 0;
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const invalidReason = cols.invalidCol !== undefined ? row[cols.invalidCol] : null;
    if (invalidReason) continue;
    const gian = gianCol !== undefined ? String(row[gianCol] || "").trim() : "";
    if (!gian) continue;
    const gross = cols.grossCol !== undefined ? Number(row[cols.grossCol]) || 0 : 0;
    if (!gross) continue;
    const net = cols.netCol !== undefined ? Number(row[cols.netCol]) || 0 : gross;
    const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
    const date = toIsoDate(dateRaw);
    if (!date) continue;
    matched++;
    dates.add(date);
    codes.add(gian);
    const key = `${date}|${gian}`;
    grossByCode[key] = (grossByCode[key] || 0) + gross;
    netByCode[key] = (netByCode[key] || 0) + net;
  }

  return { sheetName, dates: Array.from(dates).sort(), codes: Array.from(codes), grossByCode, netByCode, rowsMatched: matched };
}

// ---------- Payoo: tai file THO truc tiep tu cong Payoo/VNPay, khong can
// gop tay vao "Danh muc ten diem" moi lan (Luyen, 2026-07-24: "cho phép úp
// này lên chỗ Payoo đi để có gì tôi úp lên đó cập nhật luôn tất cả điều bỏ
// qua cái trùng nha") ----------
// Khac voi parsePayooWorkbook (doc file TONG HOP cua Luyen, sheet ten co
// "Payoo", da co san cot phu "Gian hang"/"Loc ngay"), ham nay doc THANG file
// export goc tu cong ("BÁO CÁO GIAO DỊCH BÁN HÀNG HỢP TÁC VỚI PAYOO" -- sheet
// thuong ten "Báo cáo", tieu de/tong o vai dong dau) -- khong loc theo TEN
// SHEET (co the la "Báo cáo" bat ky), ma quet header THUC SU (co the nam o
// dong 2-15 tuy file) co du cot "Số tiền thanh toán", "Ngày giao dịch", va
// "Cửa hàng"/"Mã cửa hàng". Tra ve TUNG giao dich rieng le (kem 1 khoa duy
// nhat moi dong -- uu tien "Số tham chiếu"/"Mã giao dịch ĐVTT"/"Mã QR", neu ca
// 3 deu trong thi ghep ngay+gio+so tien+STT lam khoa du phong) thay vi gop
// san theo ngay|gian nhu parsePayooWorkbook, de caller tu khu trung theo TUNG
// GIAO DICH (xem store.zvp_payoo_raw_tx trong routes/doisoat-zvp.js) -- an
// toan ngay ca khi Luyen tai chong 2 file khac dinh dang cung ngay (vd bao
// cao "Giao dich ban hang" co ca the+QR, va bao cao "Giao dich QR" rieng chi
// co QR -- CUNG 1 giao dich QR se trung khoa, chi tinh 1 lan).
function parsePayooRawReport(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  for (const sheetName of wbLite.SheetNames) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    let headerRowIdx = -1;
    let cols = {};
    for (let r = 0; r < Math.min(grid.length, 15); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (idx.grossCol === undefined && s.includes("so tien thanh toan")) idx.grossCol = c;
        if (idx.grossColAlt === undefined && s.includes("so tien giao dich")) idx.grossColAlt = c;
        if (idx.feeCol === undefined && s.includes("phi xu ly giao dich")) idx.feeCol = c;
        if (idx.netCol === undefined && s.includes("thanh tien")) idx.netCol = c;
        if (idx.dateCol === undefined && (s === "ngay giao dich" || s === "ngay thanh toan")) idx.dateCol = c;
        if (idx.invalidCol === undefined && s.includes("ly do gd khong hop le")) idx.invalidCol = c;
        if (idx.trangThaiCol === undefined && s === "trang thai") idx.trangThaiCol = c;
        if (idx.gianCol === undefined && s === "cua hang") idx.gianCol = c;
        if (idx.gianCol === undefined && s === "ten cua hang") idx.gianCol = c;
        if (idx.thamChieuCol === undefined && s.includes("so tham chieu")) idx.thamChieuCol = c;
        if (idx.dvttCol === undefined && s.includes("ma giao dich dvtt")) idx.dvttCol = c;
        if (idx.qrCol === undefined && s === "ma qr") idx.qrCol = c;
      });
      const grossIdx = idx.grossCol !== undefined ? idx.grossCol : idx.grossColAlt;
      if (grossIdx !== undefined && idx.dateCol !== undefined && idx.gianCol !== undefined) {
        headerRowIdx = r;
        cols = idx;
        break;
      }
    }
    if (headerRowIdx < 0) continue; // sheet nay khong phai sheet du lieu -- thu sheet tiep theo

    const grossCol = cols.grossCol !== undefined ? cols.grossCol : cols.grossColAlt;
    const dates = new Set();
    const codes = new Set();
    const transactions = [];
    let matched = 0;
    let skippedInvalid = 0;
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const invalidReason = cols.invalidCol !== undefined ? row[cols.invalidCol] : null;
      if (invalidReason) {
        skippedInvalid++;
        continue;
      }
      if (cols.trangThaiCol !== undefined) {
        const trangThai = String(row[cols.trangThaiCol] || "");
        if (trangThai && !/thanh cong/i.test(normText(trangThai))) {
          skippedInvalid++;
          continue;
        }
      }
      const gian = String(row[cols.gianCol] || "").trim();
      if (!gian) continue;
      const gross = grossCol !== undefined ? Number(row[grossCol]) || 0 : 0;
      if (!gross) continue;
      const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
      const date = toIsoDate(dateRaw);
      if (!date) continue;
      const fee = cols.feeCol !== undefined ? Number(row[cols.feeCol]) || 0 : 0;
      const net = cols.netCol !== undefined ? Number(row[cols.netCol]) || 0 : gross - fee;

      const thamChieu = cols.thamChieuCol !== undefined ? String(row[cols.thamChieuCol] || "").trim() : "";
      const dvtt = cols.dvttCol !== undefined ? String(row[cols.dvttCol] || "").trim() : "";
      const qr = cols.qrCol !== undefined ? String(row[cols.qrCol] || "").trim() : "";
      const dateTimeRaw = dateRaw instanceof Date ? dateRaw.toISOString() : String(dateRaw);
      // Khoa duy nhat: uu tien ma tham chieu/DVTT/QR THAT (khong doi giua cac
      // bao cao khac dinh dang cho CUNG 1 giao dich, xem vi du doi chieu that
      // trong ghi chu tren ham nay) -- chi khi CA 3 deu trong (hiem, giao dich
      // loi/thieu du lieu) moi ghep ngay gio + so tien + so dong lam du phong.
      const txKey = thamChieu || dvtt || qr || `${dateTimeRaw}|${gross}|${r}`;

      matched++;
      dates.add(date);
      codes.add(gian);
      transactions.push({ txKey, date, gian, gross, fee, net });
    }
    if (matched === 0) continue;
    return {
      sheetName,
      dates: Array.from(dates).sort(),
      codes: Array.from(codes),
      transactions,
      rowsMatched: matched,
      rowsSkippedInvalid: skippedInvalid,
    };
  }
  throw new Error(
    'Khong tim thay bang du lieu giao dich Payoo hop le trong file nay (can cot "Số tiền thanh toán"/"Số tiền giao dịch", "Ngày giao dịch", "Cửa hàng").'
  );
}

function parsePayooDiemMapping(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => /danh\s*muc.*ten\s*diem/i.test(normText(n)));
  if (!sheetName) return {};
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 15); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.tenDiem === undefined && s.includes("ten diem xuat hoa don")) idx.tenDiem = c;
      if (idx.maCongTrinh === undefined && s.includes("ma diem tren misa")) idx.maCongTrinh = c;
      if (idx.hinhThuc === undefined && s.includes("hinh thuc hop tac")) idx.hinhThuc = c;
      if (s === "chi nhanh") idx.chiNhanh = c;
    });
    if (idx.chiNhanh !== undefined && idx.maCongTrinh !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) return {};

  const map = {};
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const chiNhanh = cols.chiNhanh !== undefined ? String(row[cols.chiNhanh] || "").trim() : "";
    if (!chiNhanh) continue;
    const maCongTrinh = normCode(cols.maCongTrinh !== undefined ? row[cols.maCongTrinh] : null);
    if (!maCongTrinh) continue;
    const hinhThuc = cols.hinhThuc !== undefined ? String(row[cols.hinhThuc] || "").toUpperCase() : "";
    map[chiNhanh] = { maCongTrinh, isCse: hinhThuc.includes("CSE") };
  }
  return map;
}

function parseOnlineVnpayWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  let sheetNames = wbLite.SheetNames.filter((n) => /doi\s*soat.*vnpay/i.test(normText(n)));
  if (sheetNames.length === 0) {
    sheetNames = wbLite.SheetNames.filter((n) => normText(n).includes("vnpay"));
  }
  if (sheetNames.length === 0) {
    throw new Error('Khong tim thay sheet "Doi soat Vnpay ..." trong file.');
  }
  const wb = XLSX.read(buffer, { type: "buffer", sheets: sheetNames });

  const dates = new Set();
  const products = new Set();
  const grossByProduct = {};
  const netByProduct = {};
  let matched = 0;

  for (const sheetName of sheetNames) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    let headerRowIdx = -1;
    let cols = {};
    for (let r = 0; r < Math.min(grid.length, 6); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (idx.grossCol === undefined && s.includes("so tien hach toan thu ho")) idx.grossCol = c;
        if (idx.netCol === undefined && s.includes("so tien sau khi tru phi")) idx.netCol = c;
        if (idx.dateCol === undefined && s.includes("ngay hach toan thu ho")) idx.dateCol = c;
        if (s === "trang thai") idx.statusCol = c;
        if (idx.productCol === undefined && s.includes("ten san pham vnpay")) idx.productCol = c;
      });
      if (idx.grossCol !== undefined && idx.dateCol !== undefined && idx.productCol !== undefined) {
        headerRowIdx = r;
        cols = idx;
        break;
      }
    }
    if (headerRowIdx < 0) continue;

    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const status = cols.statusCol !== undefined ? row[cols.statusCol] : null;
      if (status !== null && !/thanh cong|th.nh c.ng/i.test(String(status))) continue;
      const product = cols.productCol !== undefined ? String(row[cols.productCol] || "").trim() : "";
      if (!product) continue;
      const gross = cols.grossCol !== undefined ? Number(row[cols.grossCol]) || 0 : 0;
      if (!gross) continue;
      const net = cols.netCol !== undefined ? Number(row[cols.netCol]) || 0 : gross;
      const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
      const date = toIsoDate(dateRaw ? String(dateRaw).split(" ")[0] : null);
      if (!date) continue;
      matched++;
      dates.add(date);
      products.add(product);
      const key = `${date}|${product}`;
      grossByProduct[key] = (grossByProduct[key] || 0) + gross;
      netByProduct[key] = (netByProduct[key] || 0) + net;
    }
  }
  if (matched === 0) {
    throw new Error('Khong doc duoc dong tieu de trong sheet Doi soat Vnpay (online).');
  }

  return {
    sheetName: sheetNames.join(", "),
    dates: Array.from(dates).sort(),
    products: Array.from(products),
    grossByProduct,
    netByProduct,
    rowsMatched: matched,
  };
}

function buildOnlineProductMatcher(diemList) {
  const entries = diemList.map((d) => ({
    ...d,
    keywords: extractKeywords(d.tenDiem),
    // No-space, no-accent version of the CODE itself (Ma cong trinh), e.g.
    // "JPSBNB" -- used as a much stronger disambiguation signal than the
    // generic tenDiem keywords below. Verified necessary against real
    // BIDV77020 data: "Sân bay vinh" (keywords "san"/"bay") and "JP SB NOI
    // BAI" (keywords "noi"/"bai") both score an identical best-keyword-
    // length match against a raw QR row literally titled "...Sân Bay Nội
    // Bài..." (contains "san"/"bay" from "Sân Bay" AND "noi"/"bai" from
    // "Nội Bài") -- a coin-flip tie that silently misattributed ALL of
    // JPSBNB's gross to SB VINH PHN (JPSBNB showed zero gian on the
    // reconciliation page, SB VINH PHN showed inflated gross vs invoice --
    // both symptoms Luyen spotted independently). The store's own "Ma cua
    // hang" naming convention (e.g. "JP SBNB 44", "SB Vinh 06", "SBCR 15")
    // embeds the real code as a literal, near-unbroken substring once
    // spaces/accents are stripped, so checking for that first resolves the
    // tie unambiguously before ever falling back to the generic keyword
    // scoring loop.
    codeNoSpace: normText(d.maCongTrinh || "").replace(/[^a-z0-9]/g, ""),
  }));
  return function matchProduct(productTitle) {
    const t = normText(productTitle);
    const tNoSpace = t.replace(/[^a-z0-9]/g, "");
    // Pass 1: exact-code substring match (strongest signal) -- only trust
    // codes at least 4 chars long (short codes risk accidental substring
    // hits). Checked BOTH directions (candidate code inside input, OR input
    // inside candidate code) -- Luyen, 2026-07-19: after switching BIDV7702's
    // store names to come straight from her "mã điểm xuất hóa đơn" column
    // (no more "PHCM"/"BIDV" suffix baked in, e.g. matchText is now literally
    // "Sense CT PVĐ" instead of "SENSE CT PVĐ PHCM"), the ORIGINAL one-way
    // check (candidate code must appear INSIDE input) silently failed for
    // every one of these -- input "sensectpv" is never a substring of the
    // candidate's own longer code "sensectpvphcm" -- so real, otherwise-exact
    // matches kept falling through to Pass 2 and colliding on a shared
    // generic keyword instead (verified: "Sense CT PVĐ"'s revenue landed on
    // "Sense Bến Tre" via the shared "sense" keyword). Checking the reverse
    // direction too (input's code IS a substring of the candidate's code)
    // catches this "same name, candidate just has an extra location suffix"
    // case without needing a fuzzy keyword fallback at all. Prefer whichever
    // candidate has the LONGEST overlapping code (most specific) on a tie.
    let codeBest = null;
    let codeBestOverlap = 0;
    for (const e of entries) {
      if (e.codeNoSpace.length < 4) continue;
      let overlap = 0;
      if (tNoSpace.includes(e.codeNoSpace)) overlap = e.codeNoSpace.length;
      else if (tNoSpace.length >= 4 && e.codeNoSpace.includes(tNoSpace)) overlap = tNoSpace.length;
      if (overlap > codeBestOverlap) {
        codeBest = e;
        codeBestOverlap = overlap;
      }
    }
    if (codeBest) return codeBest;
    // Pass 2: fallback to the original generic keyword-overlap heuristic --
    // matched against WHOLE WORD TOKENS of productTitle (not a raw substring
    // check) since a short keyword like "tra" (from "Go Trà Vinh") used to
    // match via t.includes(kw) even when it only appeared as a fragment
    // INSIDE a longer, unrelated word (e.g. "tra" inside "Nha TRAng" == Nha
    // Trang) -- verified against real BIDV7702 data, 2026-07-18: a "Nha
    // Trang" QR row (no invoice yet for that gian, so it's genuinely
    // unmatched right now) was silently misattributed to "Go Tra Vinh"'s
    // gross, inflating it above its real invoice total. Tokenizing productTitle
    // the same way extractKeywords() tokenizes the candidate name, then
    // requiring an exact token match, keeps legitimate short keywords (e.g.
    // "vinh") working while refusing purely-coincidental substrings.
    // Luyen, 2026-07-18: raised from >=3 to >=4 after a SECOND real collision
    // found the same day -- "tho" (from "Go Mỹ Tho") silently matched "SB
    // CAN THO PHCM" (Sân Bay CẦN THƠ, a totally different airport) because
    // Vietnamese accent-stripping makes "Tho" and "Thơ" the same normalized
    // token. 3-letter Vietnamese syllables collide too easily after accents
    // are stripped (homophones like tho/thơ, tra/trà, la/là...) to trust as a
    // matching signal on their own.
    // Luyen, 2026-07-19: a candidate with 2+ qualifying (len>=4) keywords now
    // needs ALL of them present in the input, not just the best single one --
    // "AE Bình Dương Ghế" (keywords "binh","duong") used to match a raw "AE
    // Bình Tân ghế" row via "binh" alone (Bình Dương/Bình Tân/Bình Thạnh...
    // all share that syllable) even though "duong" was nowhere in the input.
    // A candidate with only 1 qualifying keyword still uses that lone one
    // (nothing stronger available), same risk as before.
    const tTokens = new Set(t.split(/[^a-z0-9]+/).filter(Boolean));
    let best = null;
    for (const e of entries) {
      const qualifying = e.keywords.filter((kw) => kw.length >= 4);
      if (qualifying.length === 0) continue;
      if (!qualifying.every((kw) => tTokens.has(kw))) continue;
      const score = qualifying.reduce((sum, kw) => sum + kw.length, 0);
      if (!best || score > best.score) best = { ...e, score };
    }
    return best;
  };
}

function extractKeywords(name) {
  const STOP = new Set([
    "funzone", "tau", "kvc", "am", "ae", "sc", "farm", "lotte", "mall", "aeon",
    "adventure", "combo", "ve", "khu", "vui", "choi", "giai", "tri", "the", "mn",
    // Luyen, 2026-07-19: "mart" qua chung chung -- nhieu chi nhanh "Lotte
    // mart X" khac nhau deu co tu nay, gay trung khi dung 1 minh lam keyword
    // (vd "Lotte mart Gò Vấp" bi gan nham vao "Lotte mart Nam Sài Gòn").
    "mart",
  ]);
  const t = normText(name);
  return t.split(/[^a-z0-9]+/).filter((w) => w.length >= 3 && !STOP.has(w));
}

function parseGianXuatHdSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => {
    const t = normText(n);
    return t.includes("gian hang") && (t.includes("xuat hd") || (t.includes("xuat") && t.includes("hd")));
  });
  if (!sheetName) return [];
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 6); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.tenDiem === undefined && s.includes("ten diem xuat hoa don")) idx.tenDiem = c;
      if (idx.maCongTrinh === undefined && s.includes("ma diem tren misa")) idx.maCongTrinh = c;
      if (idx.hinhThuc === undefined && s.includes("hinh thuc hop tac")) idx.hinhThuc = c;
    });
    if (idx.tenDiem !== undefined && idx.maCongTrinh !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) return [];

  const out = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const tenDiem = cols.tenDiem !== undefined ? String(row[cols.tenDiem] || "").trim() : "";
    if (!tenDiem) continue;
    const maCongTrinh = normCode(cols.maCongTrinh !== undefined ? row[cols.maCongTrinh] : null);
    if (!maCongTrinh) continue;
    const hinhThuc = cols.hinhThuc !== undefined ? String(row[cols.hinhThuc] || "").toUpperCase() : "";
    out.push({ tenDiem, maCongTrinh, isCse: hinhThuc.includes("CSE") });
  }
  return out;
}

// ---------- Product catalog sheet ("Danh muc ten san pham") ----------
// Authoritative, per-PRODUCT, per-DAY breakdown maintained by hand inside
// the SAME "Tong hop Zalo App" file Luyen already uploads for Online.
// Unlike parseOnlineVnpayWorkbook (the raw "Doi soat Vnpay" export -- only a
// product TITLE, no gian/Ma cong trinh column, so Online gross has to be
// resolved via fuzzy keyword matching through buildOnlineProductMatcher),
// this sheet already carries the EXACT "Ma cong trinh misa thue" for every
// product row, with day-by-day amounts broken out across 3 repeating
// 31-day blocks headed by real Excel date serials:
//   "Tong xuat hoa don"  -> gross (invoiced amount that day)
//   "Tong tien tra ve"   -> net (amount VNPay actually pays back, after fee)
//   "Tong phi thu ho"    -> fee (gross - net, not needed separately here)
// This is now the SOLE source of truth for Online/Zalo Mini App gross-by-
// code-by-day -- no fuzzy product-name matching needed. That fuzzy step was
// also the root cause of misattribution bugs (e.g. KVC ROYAL showing gross
// revenue on days it genuinely had zero sales, because some unrelated
// product's title happened to contain an overlapping keyword) -- an exact,
// per-row Ma cong trinh column can't misfire like that.
function parseProductCatalogSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => normText(n).includes("danh muc ten san pham"));
  if (!sheetName) return { sheetName: null, rows: [] };
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 15); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.tenSanPham === undefined && s.includes("ten san pham")) idx.tenSanPham = c;
      if (idx.tenGianXuatHd === undefined && s.includes("ten gian hang xuat hoa don")) idx.tenGianXuatHd = c;
      if (idx.maCongTrinh === undefined && s.includes("ma cong trinh misa thue")) idx.maCongTrinh = c;
    });
    if (idx.tenSanPham !== undefined && idx.maCongTrinh !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) return { sheetName, rows: [] };

  const headerRow = grid[headerRowIdx] || [];
  // The 3 date blocks are runs of consecutive Excel date-serial numbers
  // (roughly 40000-60000 covers any date from 2009 to 2064 -- comfortably
  // wide for this app's lifetime) in the header row; each run is followed
  // immediately by its "Tong ..." label cell, which tells us which block is
  // which without hardcoding column positions (the file's layout has
  // shifted before and will likely shift again).
  const dateCols = [];
  headerRow.forEach((v, c) => {
    if (typeof v === "number" && v > 40000 && v < 60000) dateCols.push(c);
  });
  const blocks = [];
  let cur = [];
  for (let i = 0; i < dateCols.length; i++) {
    if (cur.length === 0 || dateCols[i] === cur[cur.length - 1] + 1) {
      cur.push(dateCols[i]);
    } else {
      blocks.push(cur);
      cur = [dateCols[i]];
    }
  }
  if (cur.length) blocks.push(cur);

  function labelFor(block) {
    const afterCol = block[block.length - 1] + 1;
    return normText(String(headerRow[afterCol] || ""));
  }
  const grossBlock = blocks.find((b) => labelFor(b).includes("xuat hoa don"));
  const netBlock = blocks.find((b) => labelFor(b).includes("tien tra ve"));
  if (!grossBlock || !netBlock) return { sheetName, rows: [] };

  function isoFromSerial(serial) {
    const utcDays = Math.floor(serial - 25569);
    return new Date(utcDays * 86400 * 1000).toISOString().slice(0, 10);
  }

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const tenSanPham = cols.tenSanPham !== undefined ? String(row[cols.tenSanPham] || "").trim() : "";
    if (!tenSanPham) continue;
    const maCongTrinh = normCode(cols.maCongTrinh !== undefined ? row[cols.maCongTrinh] : null);
    if (!maCongTrinh) continue;
    const tenGianXuatHd = cols.tenGianXuatHd !== undefined ? String(row[cols.tenGianXuatHd] || "").trim() : "";
    const byDate = {};
    for (let i = 0; i < grossBlock.length; i++) {
      const iso = isoFromSerial(headerRow[grossBlock[i]]);
      const gross = Number(row[grossBlock[i]]) || 0;
      const net = Number(row[netBlock[i]]) || 0;
      if (gross || net) byDate[iso] = { gross, net };
    }
    if (Object.keys(byDate).length === 0) continue;
    rows.push({ tenSanPham, tenGianXuatHd, maCongTrinh, byDate });
  }
  return { sheetName, rows };
}

// Resolves parseProductCatalogSheet's rows into the same { dates, codes,
// grossByCode, netByCode, unmapped } shape resolveOnlineGross/resolveDiemGross
// already produce, so it plugs straight into reconcileZvpChannel unchanged.
// CSE is looked up by EXACT tenGianXuatHd against gianList (the same list
// built from "gian hang xuat HD" + any master-sheet merge, already used
// elsewhere) instead of fuzzy keyword matching, since each product row
// already tells us precisely which gian it belongs to.
// learnedGian: new {tenDiem, maCongTrinh, isCse:false} entries for any
// tenGianXuatHd not already on gianList -- the caller should merge these
// into store.zvp_gian_list so the gian is "remembered" going forward and
// stops showing up as unmapped on every future upload, without Luyen having
// to hand-edit the "gian hang xuat HD"/master sheet herself. Defaulting new
// gian to non-CSE (131) is deliberate: every genuinely-new product name seen
// in practice so far (Snow Fun / Nha Ma / Ghost variants) is a non-CSE
// product per Luyen's own rule, so this is the safe default even though it
// should still be spot-checked if a brand-new CSE-only gian ever appears.
function resolveProductCatalogGross(parsedRows, gianList) {
  // Keyed by EXACT tenGianXuatHd (normalized) -> the gian's own maCongTrinh
  // + isCse. This is what makes a "Sua ma gian" (1b) correction actually take
  // effect: when Luyen redirects a self-learned gian (e.g. "SNOWFUN TAN PHU")
  // to a real code (e.g. "AM TP KVCM") via /doi-soat/zvp/online-gian-fix,
  // gianList's entry for that tenDiem now carries the corrected maCongTrinh
  // -- and THAT is what determines the code below, not the raw "Ma cong
  // trinh misa thue" column read straight off this row (row.maCongTrinh),
  // which for a not-yet-reviewed product is often just the gian's own name
  // typed into that column rather than a real accounting code. Previously
  // this function used row.maCongTrinh unconditionally, so any 1b correction
  // silently had NO effect on the actual computed revenue -- gianList was
  // only consulted for the CSE flag, never for the code itself.
  const gianByTenDiem = new Map();
  for (const g of gianList || []) {
    gianByTenDiem.set(normText(g.tenDiem), { maCongTrinh: g.maCongTrinh, isCse: !!g.isCse });
  }
  const dates = new Set();
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  const unmapped = new Set();
  const learnedGian = [];
  const learnedKeys = new Set();
  for (const row of parsedRows) {
    const key = normText(row.tenGianXuatHd);
    const gianEntry = gianByTenDiem.get(key);
    let isCse, effectiveMaCongTrinh;
    if (gianEntry) {
      isCse = gianEntry.isCse;
      effectiveMaCongTrinh = gianEntry.maCongTrinh;
    } else {
      unmapped.add(row.tenGianXuatHd || row.tenSanPham);
      isCse = false; // default to non-CSE (131) rather than silently dropping the revenue
      effectiveMaCongTrinh = row.maCongTrinh;
      if (row.tenGianXuatHd && !learnedKeys.has(key)) {
        learnedKeys.add(key);
        learnedGian.push({ tenDiem: row.tenGianXuatHd, maCongTrinh: row.maCongTrinh, isCse: false });
      }
    }
    const code = isCse ? effectiveMaCongTrinh + FF_SUFFIX : effectiveMaCongTrinh;
    for (const [iso, v] of Object.entries(row.byDate)) {
      dates.add(iso);
      codes.add(code);
      const k = `${iso}|${code}`;
      grossByCode[k] = (grossByCode[k] || 0) + v.gross;
      netByCode[k] = (netByCode[k] || 0) + v.net;
    }
  }
  return {
    dates: Array.from(dates).sort(),
    codes: Array.from(codes),
    grossByCode,
    netByCode,
    unmapped: Array.from(unmapped),
    learnedGian,
  };
}

// Redirects an invoice's own "Ma diem tren misa thue" (maDiem) through the
// SAME zvp_gian_list mapping used for Online product-catalog revenue
// (resolveProductCatalogGross), keyed by the invoice's own "Ten diem xuat
// hoa don" (tenDiem). This is what lets a correction made via "Sua ma gian"
// (1b) -- or a mapping that already exists on the "gian hang xuat HD"/master
// sheet -- ALSO apply to invoice matching, not just to revenue: without
// this, an invoice whose own maDiem column is just the site's own name (not
// yet redirected to a real code) would never line up against the correctly
// redirected settlement line, showing a permanent, wrong "Chua co HD"/"Lech"
// even though a matching invoice clearly exists (e.g. an invoice for
// "GHOST BRIDE AE HUE" carries maDiem "GHOST BRIDE AE HUE" even when
// zvp_gian_list already redirects that exact tenDiem to "AE HUE KVCN").
// Only the BASE code is redirected -- each invoice's own CSE/non-CSE tag
// (from its "Hinh thuc hop tac" column, already baked into maDiem's __FF
// suffix by the caller) is left untouched, since a single gian can
// legitimately have a genuine mix of CSE and non-CSE invoices.
function applyGianRedirectToInvoices(invoices, gianList) {
  const map = new Map();
  for (const g of gianList || []) {
    map.set(normText(g.tenDiem), g.maCongTrinh);
  }
  return (invoices || []).map((inv) => {
    const hadFF = inv.maDiem && inv.maDiem.endsWith(FF_SUFFIX);
    const maDiemBase = hadFF ? inv.maDiem.slice(0, -FF_SUFFIX.length) : inv.maDiem;

    // Try the invoice's own "Ten diem xuat hoa don" first (the raw site name
    // as typed by the accountant, independent of whatever ended up in the
    // Ma diem column).
    let target = inv.tenDiem ? map.get(normText(inv.tenDiem)) : undefined;

    // Fallback for invoices uploaded BEFORE "Ten diem xuat hoa don" was
    // captured at all (so inv.tenDiem is simply missing) -- their Ma diem
    // column is often just the raw site name too (e.g. "GHOST BRIDE AE HUE"),
    // never redirected. If that exact string is itself a known tenDiem on
    // zvp_gian_list, redirect using it directly, same as above. Safe because
    // tenDiem values are site nicknames, never real Ma cong trinh codes, so
    // this can never accidentally override an invoice whose Ma diem is
    // already a correct code.
    if (!target && maDiemBase) {
      target = map.get(normText(maDiemBase));
    }

    if (!target) return inv;
    const newMaDiem = hadFF ? target + FF_SUFFIX : target;
    if (newMaDiem === inv.maDiem) return inv;
    return { ...inv, maDiem: newMaDiem };
  });
}

// Same idea as applyGianRedirectToInvoices, but for an already-resolved
// gross/net-by-code-by-date result (as produced by mergeResolvedGross over
// zvp_offline_uploads / zvp_payoo_uploads). The Offline/Payoo channels each
// keep their OWN "Chi nhanh -> Ma cong trinh" mapping table (store.
// zvp_offline_diem_map / zvp_payoo_diem_map), separate from the shared
// zvp_gian_list -- if that channel-specific table was built/last-updated
// BEFORE zvp_gian_list learned a further redirect for one of its own target
// codes (e.g. zvp_offline_diem_map has "GHOST BRIDE MEGA DA NANG" ->
// "GHOST BRIDE AE HUE", treating that as a final code, while zvp_gian_list
// separately already redirects "GHOST BRIDE AE HUE" -> "AE HUE KVCN" for
// invoices), the two channels permanently disagree on the final code even
// though it's the exact same revenue -- Offline keeps showing "Chua co HD"
// under the stale intermediate code while invoices (now correctly
// redirected) pile up under the final code instead, showing "Lech" there.
// Applying ONE more zvp_gian_list hop here, at merge/reconcile time (same
// timing as applyGianRedirectToInvoices), fixes this for ALL historical
// uploads without needing Luyen to re-upload the Offline/Payoo file, and
// keeps working automatically if zvp_gian_list ever grows another such
// chain. Same safe-by-default rule: a code is only redirected when it
// EXACTLY matches a known tenDiem, so a genuine final code is never
// touched; multiple keys colliding onto the same final code after redirect
// are summed rather than overwritten.
function applyGianRedirectToResolvedGross(resolved, gianList) {
  const map = new Map();
  for (const g of gianList || []) {
    map.set(normText(g.tenDiem), g.maCongTrinh);
  }
  const grossByCode = {};
  const netByCode = {};
  const codesSet = new Set();
  for (const key of Object.keys(resolved.grossByCode || {})) {
    const sep = key.indexOf("|");
    const date = key.slice(0, sep);
    const code = key.slice(sep + 1);
    const hadFF = code.endsWith(FF_SUFFIX);
    const codeBase = hadFF ? code.slice(0, -FF_SUFFIX.length) : code;
    const target = map.get(normText(codeBase));
    const finalCode = target ? (hadFF ? target + FF_SUFFIX : target) : code;
    codesSet.add(finalCode);
    const newKey = `${date}|${finalCode}`;
    grossByCode[newKey] = (grossByCode[newKey] || 0) + resolved.grossByCode[key];
    netByCode[newKey] = (netByCode[newKey] || 0) + (resolved.netByCode ? resolved.netByCode[key] || 0 : 0);
  }
  return {
    dates: resolved.dates,
    codes: Array.from(codesSet),
    grossByCode,
    netByCode,
    unmapped: resolved.unmapped,
  };
}

// Scans invoices for any "Ten diem xuat hoa don" not already present on
// gianList and returns new self-referential {tenDiem, maCongTrinh, isCse}
// entries for them (maCongTrinh defaults to the invoice's own maDiem, base
// code only, or the tenDiem itself if maDiem is blank) -- same auto-learn
// convention as resolveProductCatalogGross's learnedGian, so a brand-new
// invoice-only site name shows up in "1b. Ra soat gian moi" for a one-time
// review instead of silently staying unmatched forever. Caller is
// responsible for merging + persisting into store.zvp_gian_list (see
// /doi-soat/zvp/upload-hoadon).
function learnGianFromInvoices(invoices, gianList) {
  const known = new Set((gianList || []).map((g) => normText(g.tenDiem)));
  const learned = [];
  const seen = new Set();
  for (const inv of invoices || []) {
    if (!inv.tenDiem) continue;
    const key = normText(inv.tenDiem);
    if (known.has(key) || seen.has(key)) continue;
    seen.add(key);
    const baseMaDiem = inv.maDiem && inv.maDiem.endsWith(FF_SUFFIX) ? inv.maDiem.slice(0, -FF_SUFFIX.length) : inv.maDiem;
    learned.push({ tenDiem: inv.tenDiem, maCongTrinh: baseMaDiem || inv.tenDiem, isCse: false });
  }
  return learned;
}

// ---------- Master gian catalog ("gian " sheet, uploaded daily) ----------
// A single hand-maintained sheet covering ALL settlement channels (Momo,
// Viet QR x3, Zalo Mini App, VNPay Co so, Payoo QR, Payoo the) in one place:
// col "raw" (labelled "Cac gian tren file ke ds xuat HD MTT - 989" in the
// real file -- the exact text a revenue-side product/diem needs to match
// against) | "Ma cong trinh" (canonical accounting code) | "Thuoc" (which
// channel this row belongs to) | CSE flag column (freeform, "CSE"/"cse").
//
// IMPORTANT: the SAME "raw" text can legitimately appear more than once
// under the SAME channel mapping to the SAME Ma cong trinh but with
// DIFFERENT CSE flags (e.g. "AE HUE KVCN" appears both as CSE and non-CSE
// under "ZALO MINI APP 11,12") -- this isn't a data error. It reflects that
// one accounting code can receive revenue from multiple distinct underlying
// products (Funzone Hue = CSE, Ecokids Farm Hue / Snow Fun Hue = not CSE),
// and the fuzzy product matcher (buildOnlineProductMatcher) already resolves
// per PRODUCT TITLE against each candidate's own keywords, so as long as
// EVERY distinct product-title variant is present as its own row here, gross
// revenue naturally splits into 2 separate lines (code and code+"__FF") at
// resolve time -- no extra "split" logic is needed anywhere else.
function parseGianMasterSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => normText(n) === "gian" || normText(n).startsWith("gian"));
  if (!sheetName) return { sheetName: null, rows: [] };
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let maCongTrinhCol = -1;
  let thuocCol = -1;
  let cseCol = -1;
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = grid[r] || [];
    let mc = -1;
    let th = -1;
    let cse = -1;
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (mc === -1 && s.includes("ma cong trinh")) mc = c;
      if (th === -1 && s === "thuoc") th = c;
      if (cse === -1 && s.includes("cse")) cse = c;
    });
    if (mc !== -1 && th !== -1) {
      headerRowIdx = r;
      maCongTrinhCol = mc;
      thuocCol = th;
      cseCol = cse;
      break;
    }
  }
  if (headerRowIdx < 0) return { sheetName, rows: [] };
  // The raw/tenDiem column is, by convention on this sheet, immediately to
  // the left of the "Ma cong trinh" column.
  const rawCol = maCongTrinhCol - 1;
  if (rawCol < 0) return { sheetName, rows: [] };

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const raw = String(row[rawCol] || "").trim();
    if (!raw) continue;
    const maCongTrinh = normCode(row[maCongTrinhCol]);
    if (!maCongTrinh) continue;
    const thuoc = String(row[thuocCol] || "").trim();
    const cseVal = cseCol >= 0 ? String(row[cseCol] || "") : "";
    rows.push({ raw, maCongTrinh, thuoc, isCse: /cse/i.test(cseVal) });
  }
  return { sheetName, rows };
}

// Merge a base tenDiem list (as used by buildOnlineProductMatcher: [{tenDiem,
// maCongTrinh, isCse}]) with master rows for the SAME channel -- master rows
// win on a normText(tenDiem) collision (they're the daily-refreshed, more
// complete source of truth), base entries not touched by master are kept
// as-is so uploading the master sheet never erases data from the older
// "gian hang xuat HD" upload pipeline.
function mergeGianListWithMaster(baseList, masterRowsForChannel) {
  const byKey = new Map();
  for (const g of baseList || []) {
    byKey.set(normText(g.tenDiem), { tenDiem: g.tenDiem, maCongTrinh: g.maCongTrinh, isCse: g.isCse });
  }
  for (const r of masterRowsForChannel || []) {
    byKey.set(normText(r.raw), { tenDiem: r.raw, maCongTrinh: r.maCongTrinh, isCse: r.isCse });
  }
  return Array.from(byKey.values());
}

// Same idea for the exact-key (Chi nhanh / diem name) maps used by
// Offline/Payoo: { "<raw text>": { maCongTrinh, isCse } }.
function mergeDiemMapWithMaster(baseMap, masterRowsForChannel) {
  const merged = Object.assign({}, baseMap || {});
  for (const r of masterRowsForChannel || []) {
    merged[r.raw] = { maCongTrinh: r.maCongTrinh, isCse: r.isCse };
  }
  return merged;
}

function resolveDiemGross(parsed, diemMap) {
  const grossByCode = {};
  const netByCode = {};
  const codes = new Set();
  const unmapped = new Set();
  for (const key of Object.keys(parsed.grossByCode)) {
    const sep = key.indexOf("|");
    const date = key.slice(0, sep);
    const diemRaw = key.slice(sep + 1);
    const mapped = diemMap[diemRaw];
    if (!mapped) {
      unmapped.add(diemRaw);
      continue;
    }
    const code = mapped.isCse ? mapped.maCongTrinh + FF_SUFFIX : mapped.maCongTrinh;
    codes.add(code);
    const newKey = `${date}|${code}`;
    grossByCode[newKey] = (grossByCode[newKey] || 0) + parsed.grossByCode[key];
    netByCode[newKey] = (netByCode[newKey] || 0) + (parsed.netByCode[key] || 0);
  }
  return { dates: parsed.dates, codes: Array.from(codes), grossByCode, netByCode, unmapped: Array.from(unmapped) };
}

// Luyen, 2026-07-24: "bây giờ khó hơn nhá tôi thấy có chỗ tải rồi nhưng chưa
// chính sát lắm giờ tôi sẽ thiết lập lại chính sát hơn chi tiết ... đưa lên
// đối soát dựa trên tên sản phẩm của đơn đó mua trên file Order có chỗ cột
// Tên sản phẩm là của gian nào dựa vào file hehehehehe sheet nối rồi gắn mã
// công trình vô ... còn cái nào sau này có tên sản phẩm mới bạn cảnh báo tên
// sản phẩm đó cho tôi để biết mã công trình nhá" -- thay vi fuzzy keyword
// matching (buildOnlineProductMatcher, co the doan sai khi 2 san pham trung
// tu khoa), Luyen tu duy tri 1 bang tra CHINH XAC "Ten san pham" -> "Ma cong
// trinh" (sheet "noi" trong file rieng cua chi) -- xem parseOnlineProductMapSheet
// ben duoi. Ham nay dung bang do de tra cuu THAY VI fuzzy: khop CHINH XAC
// (chi trim + gop khoang trang, KHONG bo dau -- Luyen muon "chinh xac hon",
// bo dau se lam long chinh xac), san pham nao KHONG co trong bang thi KHONG
// doan (tra ve unmappedProducts de canh bao, giong het "unmapped" cua
// resolveDiemGross) thay vi im lang gan sai hoac bo qua.
function normalizeProductKey(s) {
  return String(s || "").trim().replace(/\s+/g, " ");
}

// Parses the "noi" ("nối") sheet: 2 cot "Ten san pham" / "Ma cong trinh"
// (co the co them cot tham khao rieng nhu "Ma cong trinh co san", bo qua --
// KHONG duoc nham voi cot "Ma cong trinh" chinh vi cung chua "ma cong
// trinh" nhu 1 substring, nen loai tru rieng cot nao co them "co san").
function parseOnlineProductMapSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => normText(n).includes("noi"));
  if (!sheetName) return { sheetName: null, rows: [] };
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 5); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.tenSanPham === undefined && s.includes("ten san pham")) idx.tenSanPham = c;
      if (idx.maCongTrinh === undefined && s.includes("ma cong trinh") && !s.includes("co san")) idx.maCongTrinh = c;
    });
    if (idx.tenSanPham !== undefined && idx.maCongTrinh !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) return { sheetName, rows: [] };

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const tenSanPham = row[cols.tenSanPham] ? String(row[cols.tenSanPham]).trim() : "";
    const maCongTrinh = row[cols.maCongTrinh] ? normCode(row[cols.maCongTrinh]) : "";
    if (!tenSanPham || !maCongTrinh) continue;
    rows.push({ tenSanPham, maCongTrinh });
  }
  return { sheetName, rows };
}

// mergedInto: store.zvp_online_product_map -- moi lan tai file "noi" moi
// GHI DE (khong xoa cac dong cu neu ten san pham do khong con trong file
// lan nay, vi Luyen chi them dong moi vao cuoi file, khong xoa dong cu).
// isCse: suy ra tu store.zvp_gian_list theo cung 1 quy tac "chi tin CSE khi
// TOAN BO cac dong hien co cua ma cong trinh do deu la CSE" da dung cho
// cseOverrideCodes (xem chu thich o reconcileZvpChannel) -- vi ban "noi"
// khong co cot rieng danh dau CSE, ma nhieu ma cong trinh (vd "AE HUE KVCN",
// "AM HP KVCN") tren thuc te la HON HOP (vua co hoa don TK 131 vua co 1388),
// nen chi tin CSE=true khi chac chan 100% (khong co dong nao TK 131 tung
// ghi nhan cho ma do), con lai mac dinh false (TK 131) -- an toan hon doan.
function inferIsCseForCode(maCongTrinh, gianList) {
  const matches = (gianList || []).filter((g) => normCode(g.maCongTrinh) === maCongTrinh);
  if (matches.length === 0) return false;
  return matches.every((g) => g.isCse === true);
}

function mergeOnlineProductMap(existingMap, rows, gianList) {
  const map = Object.assign({}, existingMap || {});
  let added = 0;
  let updated = 0;
  for (const r of rows) {
    const key = normalizeProductKey(r.tenSanPham);
    const isCse = inferIsCseForCode(r.maCongTrinh, gianList);
    const cur = map[key];
    if (!cur) {
      added++;
    } else if (cur.maCongTrinh !== r.maCongTrinh || cur.isCse !== isCse) {
      updated++;
    }
    map[key] = { maCongTrinh: r.maCongTrinh, isCse };
  }
  return { map, added, updated };
}

function resolveOnlineGrossByProductMap(parsed, productMap) {
  const grossByCode = {};
  const netByCode = {};
  const codes = new Set();
  const unmappedProducts = new Set();
  for (const key of Object.keys(parsed.grossByProduct)) {
    const sep = key.indexOf("|");
    const date = key.slice(0, sep);
    const product = key.slice(sep + 1);
    const mapped = productMap[normalizeProductKey(product)];
    if (!mapped) {
      unmappedProducts.add(product);
      continue;
    }
    const code = mapped.isCse ? mapped.maCongTrinh + FF_SUFFIX : mapped.maCongTrinh;
    codes.add(code);
    const newKey = `${date}|${code}`;
    grossByCode[newKey] = (grossByCode[newKey] || 0) + parsed.grossByProduct[key];
    netByCode[newKey] = (netByCode[newKey] || 0) + (parsed.netByProduct[key] || 0);
  }
  return { dates: parsed.dates, codes: Array.from(codes), grossByCode, netByCode, unmappedProducts: Array.from(unmappedProducts) };
}

function resolveOnlineGross(parsed, gianList) {
  const matcher = buildOnlineProductMatcher(gianList);
  const grossByCode = {};
  const netByCode = {};
  const codes = new Set();
  const unmapped = new Set();
  for (const key of Object.keys(parsed.grossByProduct)) {
    const sep = key.indexOf("|");
    const date = key.slice(0, sep);
    const product = key.slice(sep + 1);
    const match = matcher(product);
    if (!match) {
      unmapped.add(product);
      continue;
    }
    const code = match.isCse ? match.maCongTrinh + FF_SUFFIX : match.maCongTrinh;
    codes.add(code);
    const newKey = `${date}|${code}`;
    grossByCode[newKey] = (grossByCode[newKey] || 0) + parsed.grossByProduct[key];
    netByCode[newKey] = (netByCode[newKey] || 0) + (parsed.netByProduct[key] || 0);
  }
  return { dates: parsed.dates, codes: Array.from(codes), grossByCode, netByCode, unmapped: Array.from(unmapped) };
}

// diemAlias: { "TEN/MA DIEM TREN HOA DON": "MA CONG TRINH CHINH" } -- optional,
// same shared table as Momo's reconcileMomo (see store.js's invoice_diem_alias
// comment); lets an invoice issued under an alternate site name (e.g.
// "SNOWFUN TAN PHU") count towards the correct Ma Cong Trinh's line (e.g.
// "AM TP KVCM") without needing to re-upload the invoice file.
// cseOverrideCodes: optional Set of BASE Ma Cong Trinh codes (no __FF suffix)
// that the shared "gian hang xuat HD" list (store.zvp_gian_list) marks as CSE
// (doanh thu chia se). For a gian that is MONOLITHIC -- every single one of
// its invoices is genuinely CSE (e.g. FUNZONE IPH KVCN, AM LBIEN KVCN: 100%
// TK 1388 per the posted "So tien gui ngan hang" ledger) -- the separate
// Offline/Payoo mapping sheet ("gian hang VNpay co so") sometimes fails to
// also mark that gian as CSE (a data gap on Luyen's mapping sheet, not a
// real difference in how the gian is set up), so its Offline/Payoo revenue
// lands on the PLAIN code while its invoices are correctly tagged __FF, and
// the two never do with each other ("Chua co HD" even though a matching
// invoice clearly exists). For THOSE gian, this override normalizes gross
// (and, symmetrically, any of that gian's invoices) onto the __FF form
// before grouping, using the gian list as the single source of truth.
//
// BUT some gian on that same list are marked CSE at the gian level even
// though in reality they have a genuine MIX of CSE and non-CSE invoices
// from day to day -- e.g. AM HP KVCN (real ledger: 6 invoices TK 131, 8
// invoices TK 1388) and AE HUE KVCN (35 invoices TK 131, only 1 stray TK
// 1388). Forcing ALL of a mixed gian's revenue onto TK 1388 would silently
// misclassify what is mostly (or partly) real TK 131 revenue -- a much
// worse outcome than the "Chua co HD" display bug this override exists to
// fix. So the override is only ever trusted for a base code when NO
// invoice anywhere in the CURRENT invoice list is genuinely tagged plain
// (non-__FF) under that same base code -- i.e. only when the gian's own
// invoice data confirms it really is monolithic. Otherwise each invoice's
// own "Hinh thuc hop tac" tag (already read correctly by the parser) is
// left to decide plain vs __FF per invoice, same as any other gian.
function reconcileZvpChannel(settlements, grossData, invoiceData, gianMapping, manualMatches, diemAlias, cseOverrideCodes) {
  const alias = diemAlias || {};
  const basesWithNativePlainInvoice = new Set();
  for (const inv of invoiceData.invoices) {
    const effective = alias[inv.maDiem] || inv.maDiem;
    if (effective && !effective.endsWith(FF_SUFFIX)) basesWithNativePlainInvoice.add(effective);
  }
  const cseOverride = new Set(
    Array.from(cseOverrideCodes || []).filter((c) => !basesWithNativePlainInvoice.has(c))
  );
  const invoicesByDiemDay = {};
  for (const inv of invoiceData.invoices) {
    if (!inv.ngayHd || !inv.days || inv.days.length === 0) continue;
    const [invY, invMo, invD] = inv.ngayHd.split("-").map(Number);
    const baseMaDiem = alias[inv.maDiem] || inv.maDiem;
    const effectiveMaDiem =
      baseMaDiem && !baseMaDiem.endsWith(FF_SUFFIX) && cseOverride.has(baseMaDiem) ? baseMaDiem + FF_SUFFIX : baseMaDiem;
    for (const day of inv.days) {
      let y = invY;
      let mo = invMo;
      if (day > 20 && invD <= 3) {
        mo -= 1;
        if (mo === 0) {
          mo = 12;
          y -= 1;
        }
      }
      const iso = `${y}-${String(mo).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      const key = `${effectiveMaDiem}|${iso}`;
      if (!invoicesByDiemDay[key]) invoicesByDiemDay[key] = [];
      invoicesByDiemDay[key].push(inv);
    }
  }

  // Chi Nhan, 2026-07-24: Luyen xac nhan Zalo/VNPay/Payoo cung tra tien tre 1
  // ngay giong Momo (doi chieu du lieu that: "Khoản về ngày 22/07" luon di
  // kem "doanh thu 21/07" cho ca 3 kenh online/offline/payoo) -- truyen 1 de
  // dong "gia" (chua co ngan hang) cung hien dung nhan ngay du kien (23/07
  // cho doanh thu 22/07), khong con trung nhan voi dong THAT cua ngay truoc
  // do nua. Xem ghi chu day du tai buildPendingDaySettlements (utils/momoReconcile.js).
  const allSettlements = settlements.concat(buildPendingDaySettlements(settlements, grossData, 1));
  const results = [];
  for (const s of allSettlements) {
    const days = dateRange(s.fromIso, s.toIso);
    const gianLines = {};
    for (const day of days) {
      for (const rawCode of grossData.codes) {
        const key = `${day}|${rawCode}`;
        const gross = grossData.grossByCode[key];
        if (gross && gross > 0) {
          const code = !rawCode.endsWith(FF_SUFFIX) && cseOverride.has(rawCode) ? rawCode + FF_SUFFIX : rawCode;
          if (!gianLines[code]) {
            gianLines[code] = { code, gross: 0, net: 0, invoices: new Set(), days: new Set() };
          }
          gianLines[code].gross += gross;
          gianLines[code].net += grossData.netByCode[key] || gross;
          gianLines[code].days.add(day);
          const invs = invoicesByDiemDay[`${code}|${day}`] || [];
          for (const inv of invs) gianLines[code].invoices.add(inv.soHd);
        }
      }
    }
    const lines = Object.values(gianLines).map((g) => {
      const invoiceList = Array.from(g.invoices);
      const invoiceTotal = invoiceList.reduce((sum, soHd) => {
        const inv = invoiceData.invoices.find((i) => i.soHd === soHd);
        return sum + (inv ? inv.tongTt : 0);
      }, 0);
      const line = {
        code: g.code,
        maCongTrinh: displayCode(g.code),
        // Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- khong con
        // fallback ve 1388 cho gian FF/CSE nua, mac dinh 131 neu chua co trong
        // gian_mapping.
        tkCo: gianMapping[g.code] || "131",
        gross: g.gross,
        net: Math.round(g.net),
        invoiceNumbers: invoiceList,
        invoiceTotal,
        diff: invoiceTotal - g.gross,
        days: Array.from(g.days).sort(),
        matched: invoiceList.length > 0 && Math.abs(invoiceTotal - g.gross) < 1,
        manualOverride: false,
      };
      // Manual correction: some gian settle their revenue into the bank
      // account without ever going through the normal VNPay/Zalo invoice
      // matching for that settlement period (e.g. the invoice for that
      // revenue is only issued the NEXT day, outside this settlement's own
      // date window) -- Luyen identifies these by hand and confirms them
      // via the "danh dau da co HD bu" form instead of the automatic dò.
      // Also usable to CORRECT a "Lech" line (invoices matched but total
      // doesn't add up -- e.g. an invoice was missed by the automatic dò):
      // no longer gated on invoiceNumbers.length === 0, so a manual match
      // record always wins once she's saved one for that settlement+code,
      // regardless of whether the line started out "Chua co HD" or "Lech".
      if (manualMatches) {
        const mm = manualMatches[`${s.date}|${line.code}`];
        if (mm) {
          line.invoiceNumbers = mm.invoiceNumbers || [];
          // grossAdjustment (Luyen, 2026-07-20): "bo het cai FUNZONE IPH KVCN
          // nay cho lech ra het di" -- mot vai gian (vd FUNZONE IPH KVCN) bi
          // trung/gop nham doanh thu cua gian khac vao gross cua no (xac nhan:
          // phan du dung bang toan bo doanh thu cua 1 gian khac da tu khop
          // hoa don rieng roi), nen can TRU BOT thang vao gross cua chinh dong
          // nay cho 1 ngay/ky cu the, giong co che +/- Sua DT da co ben VietQR
          // (utils/vietqrReconcile.js) -- truoc day ZVP chua co, chi sua duoc
          // "amount" (so tien HD hien thi) ma khong dong den gross that su nen
          // "Chenh lech" dau ky khong bao gio het du dong da danh dau "Khop".
          if (mm.grossAdjustment) {
            line.gross += mm.grossAdjustment;
            line.net = Math.round(line.gross);
          }
          line.invoiceTotal = mm.amount != null ? mm.amount : line.gross;
          line.diff = line.invoiceTotal - line.gross;
          line.matched = Math.abs(line.diff) < 1;
          line.manualOverride = true;
          line.manualNote = mm.note || "";
        }
      }
      return line;
    });
    const totalNet = lines.filter((l) => l.tkCo !== "SKIP").reduce((sum, l) => sum + l.net, 0);
    results.push({
      settlementDate: s.date,
      from: s.fromIso,
      to: s.toIso,
      bankAmount: s.amount,
      totalNetComputed: totalNet,
      diffVsBank: s.pendingBank ? null : totalNet - s.amount,
      pendingBank: !!s.pendingBank,
      lines: lines.sort((a, b) => b.gross - a.gross),
    });
  }
  return results;
}

function reconcileZvp(settlementsByChannel, grossByChannel, invoicesByChannel, gianMapping, manualMatchesByChannel, diemAlias, cseOverrideCodes) {
  const mm = manualMatchesByChannel || {};
  return {
    online: reconcileZvpChannel(settlementsByChannel.online, grossByChannel.online, invoicesByChannel.online, gianMapping, mm.online, diemAlias, cseOverrideCodes),
    offline: reconcileZvpChannel(settlementsByChannel.offline, grossByChannel.offline, invoicesByChannel.offline, gianMapping, mm.offline, diemAlias, cseOverrideCodes),
    payoo: reconcileZvpChannel(settlementsByChannel.payoo, grossByChannel.payoo, invoicesByChannel.payoo, gianMapping, mm.payoo, diemAlias, cseOverrideCodes),
  };
}

// ---------- Raw VNPay portal exports (order details + fee-by-transaction report) ----------
// Luyen can now download 2 raw files directly from the VNPay/Zalo merchant
// portal instead of hand-consolidating them into the "Tong hop Zalo App"
// sheet every time:
//   1) "OrderDetails..." (.xls) -- one row per Zalo Mini App order, with the
//      real product title ("Ten san pham").
//   2) "DuLieuBaoCaoPhiTheoGDThanhToan..." (.xlsx) -- one row per PAID
//      transaction (both Online AND Offline mixed together), with gross/net
//      amounts + settlement date, but NO product title -- only an order
//      reference buried in "Thong tin dat hang" ("... don hang <so> ...").
// The fee-report's own "Diem thu" column tells Online vs Offline apart:
// "FUNZONE MINI APP" = Zalo Mini App checkout (Online); every other value is
// a physical location's QR terminal (Offline), joinable directly by its
// "Chi nhanh" column against the SAME diem-mapping sheet already used by
// parseDiemMappingSheet/resolveDiemGross for the manually-uploaded Offline
// file -- no new mapping table needed for Offline.
// For Online, the order reference is used to look up the product title from
// the OrderDetails file, then fed through the SAME buildOnlineProductMatcher
// pipeline as the manually-consolidated "Doi soat Vnpay" sheet.

// Returns { orderNo: tenSanPham }, keyed by the numeric order number (the
// "#" prefix in "Ma don hang" is stripped so it matches the plain digits
// found inside the fee report's "Thong tin dat hang" text).
function parseOrderDetailsWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => normText(n).includes("sheet1")) || wbLite.SheetNames[0];
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 5); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.orderNo === undefined && s.includes("ma don hang")) idx.orderNo = c;
      if (idx.product === undefined && s.includes("ten san pham")) idx.product = c;
      if (idx.channel === undefined && s.includes("kenh ban hang")) idx.channel = c;
    });
    if (idx.orderNo !== undefined && idx.product !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de (can cot "Ma don hang" va "Ten san pham") trong file OrderDetails.');
  }

  // Luyen, 2026-07-24: "à quên chỗ kênh bán á bỏ cái HUB nha lấy zalo thôi
  // á" -- file OrderDetails gom ca don "Zalo" (Zalo Mini App, kenh dang doi
  // soat) VA don "HUB" (kenh ban hang khac, khong lien quan) -- CHI lay dong
  // "Zalo" cho orderMap, bo qua HUB de khong bao gio lo nham san pham cua 1
  // kenh khac vao doi soat Online (dau chua thay trung so don hang giua 2
  // kenh tren du lieu thuc te, van loc cho chac).
  const orderMap = {};
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    if (cols.channel !== undefined) {
      const channel = String(row[cols.channel] || "").trim();
      if (channel && normText(channel) !== "zalo") continue;
    }
    const rawOrderNo = row[cols.orderNo];
    if (!rawOrderNo) continue;
    const orderNo = String(rawOrderNo).replace(/[^0-9]/g, "");
    if (!orderNo) continue;
    const product = row[cols.product] ? String(row[cols.product]).trim() : "";
    if (product) orderMap[orderNo] = product;
  }
  return orderMap;
}

// Parses "Du lieu bao cao phi theo GD thanh toan ..." and splits it into
// Online (Diem thu = "FUNZONE MINI APP", product resolved via orderMap) and
// Offline (every other Diem thu, joined by "Chi nhanh"). Output shapes
// match parseOnlineVnpayWorkbook / parseOfflineVnpayWorkbook exactly, so
// they plug straight into resolveOnlineGross / resolveDiemGross unchanged.
function parseFeeReportWorkbook(buffer, orderMap) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => normText(n).includes("sheet1")) ||
    wbLite.SheetNames.find((n) => !/config/i.test(n)) ||
    wbLite.SheetNames[0];
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 6); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.diemThu === undefined && s.includes("diem thu") && !s.includes("ma diem thu")) idx.diemThu = c;
      if (idx.chiNhanh === undefined && s.includes("chi nhanh")) idx.chiNhanh = c;
      if (idx.orderInfo === undefined && s.includes("thong tin dat hang")) idx.orderInfo = c;
      if (idx.grossCol === undefined && s.includes("so tien hach toan thu ho")) idx.grossCol = c;
      if (idx.netCol === undefined && s.includes("so tien sau khi tru phi")) idx.netCol = c;
      if (idx.feeCol === undefined && s.includes("so tien phi thu ho")) idx.feeCol = c;
      if (idx.dateCol === undefined && s.includes("ngay hach toan thu ho")) idx.dateCol = c;
    });
    if (idx.diemThu !== undefined && idx.grossCol !== undefined && idx.dateCol !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de trong file "Du lieu bao cao phi theo GD thanh toan".');
  }

  const onlineDates = new Set();
  const products = new Set();
  const grossByProduct = {};
  const netByProduct = {};
  let onlineMatched = 0;
  const unmatchedOrders = new Set();

  const offlineDates = new Set();
  const codes = new Set();
  const grossByCode = {};
  const netByCode = {};
  let offlineMatched = 0;

  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const diemThu = cols.diemThu !== undefined ? String(row[cols.diemThu] || "").trim() : "";
    if (!diemThu) continue;
    const gross = cols.grossCol !== undefined ? Number(row[cols.grossCol]) || 0 : 0;
    if (!gross) continue;
    const netRaw = cols.netCol !== undefined ? row[cols.netCol] : null;
    const fee = cols.feeCol !== undefined ? Number(row[cols.feeCol]) || 0 : 0;
    const net = (netRaw !== null && netRaw !== undefined && netRaw !== "") ? (Number(netRaw) || 0) : (gross - fee);
    const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
    const date = toIsoDate(dateRaw);
    if (!date) continue;

    if (/^FUNZONE MINI APP$/i.test(diemThu)) {
      const orderInfo = cols.orderInfo !== undefined ? String(row[cols.orderInfo] || "") : "";
      const mOrder = orderInfo.match(/don\s*hang\s+(\d+)/i);
      const product = mOrder ? orderMap[mOrder[1]] : null;
      if (!product) {
        unmatchedOrders.add(mOrder ? mOrder[1] : "(khong doc duoc so don hang)");
        continue;
      }
      onlineMatched++;
      onlineDates.add(date);
      products.add(product);
      const key = `${date}|${product}`;
      grossByProduct[key] = (grossByProduct[key] || 0) + gross;
      netByProduct[key] = (netByProduct[key] || 0) + net;
    } else {
      const chiNhanh = cols.chiNhanh !== undefined ? String(row[cols.chiNhanh] || "").trim() : diemThu;
      if (!chiNhanh) continue;
      offlineMatched++;
      offlineDates.add(date);
      codes.add(chiNhanh);
      const key = `${date}|${chiNhanh}`;
      grossByCode[key] = (grossByCode[key] || 0) + gross;
      netByCode[key] = (netByCode[key] || 0) + net;
    }
  }

  return {
    online: {
      sheetName,
      dates: Array.from(onlineDates).sort(),
      products: Array.from(products),
      grossByProduct,
      netByProduct,
      rowsMatched: onlineMatched,
      unmatchedOrders: Array.from(unmatchedOrders),
    },
    offline: {
      sheetName,
      dates: Array.from(offlineDates).sort(),
      codes: Array.from(codes),
      grossByCode,
      netByCode,
      rowsMatched: offlineMatched,
    },
  };
}

// Luyen, 2026-07-24: "chỗ offline á có thể tôi tải đối soát này bạn lên đối
// soát vnpay offline cho tôi nhá á dựa vào điểm thu để lấy ra gian và ngày
// giao dịch phí hay là tổng tiền á lấy các giao dịch thành công nhá và bỏ qua
// điểm thu FUNZONE MINI APP nhá và thêm trên wed úp cái dữ liệu này lên nha
// thêm dạng này á" -- upload THO rieng cho VNPay Offline, dung THANG file
// "Du lieu bao cao phi theo GD thanh toan" (cung dinh dang voi
// parseFeeReportWorkbook o tren, dung chung cho combo Online+Offline) nhung
// KHONG can file OrderDetails di kem (file do chi dung de tra ten san pham
// cho phan Online/Zalo Mini App -- khong lien quan Offline). Chi lay phan
// Offline (moi "Diem thu" KHAC "FUNZONE MINI APP", con FUNZONE MINI APP la
// Online nen bo qua theo dung yeu cau), tra ve TUNG GIAO DICH rieng (khong
// gop san theo ngay+ma) de khu trung qua store.zvp_offline_raw_tx (cung 1
// kieu voi store.zvp_payoo_raw_tx o tren) -- dung "Ma giao dich" lam khoa
// (xac nhan 100% duy nhat, khong blank, tren file thuc te 24/07/2026).
// "Trang thai": tren file thuc te cot nay 100% rong (chua thay VNPay dien gi
// ca) nhung van kiem tra PHONG THU cho file tuong lai -- CHI bo qua 1 dong
// khi cot nay CO GIA TRI ro rang va gia tri do KHONG chua "thanh cong" (vd
// "that bai", "huy", "loi"); con rong hoac chua "thanh cong" thi van lay
// (dung yeu cau "lấy các giao dịch thành công").
function parseVnpayOfflineFeeReport(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => normText(n).includes("sheet1")) ||
    wbLite.SheetNames.find((n) => !/config/i.test(n)) ||
    wbLite.SheetNames[0];
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 6); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.diemThu === undefined && s.includes("diem thu") && !s.includes("ma diem thu")) idx.diemThu = c;
      if (idx.chiNhanh === undefined && s.includes("chi nhanh")) idx.chiNhanh = c;
      if (idx.maGiaoDich === undefined && s.includes("ma giao dich")) idx.maGiaoDich = c;
      if (idx.grossCol === undefined && s.includes("so tien hach toan thu ho")) idx.grossCol = c;
      if (idx.netCol === undefined && s.includes("so tien sau khi tru phi")) idx.netCol = c;
      if (idx.feeCol === undefined && s.includes("so tien phi thu ho")) idx.feeCol = c;
      if (idx.dateCol === undefined && s.includes("ngay hach toan thu ho")) idx.dateCol = c;
      // "Trang thai" dung khop CHINH XAC (khong phai substring) de khong bi
      // nham voi "Trang thai tra gop"/cac cot "Trang thai ..." khac dung
      // truoc no trong file thuc te.
      if (idx.statusCol === undefined && s === "trang thai") idx.statusCol = c;
    });
    if (idx.diemThu !== undefined && idx.grossCol !== undefined && idx.dateCol !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de trong file "Du lieu bao cao phi theo GD thanh toan".');
  }

  const transactions = [];
  let excludedFunzone = 0;
  let excludedFailedStatus = 0;
  let noTxKeyFallbackUsed = 0;
  let fallbackSeq = 0;

  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const diemThu = cols.diemThu !== undefined ? String(row[cols.diemThu] || "").trim() : "";
    if (!diemThu) continue;
    if (/^FUNZONE MINI APP$/i.test(diemThu)) {
      excludedFunzone++;
      continue; // Online (Zalo Mini App) -- khong phai Offline, bo qua theo dung yeu cau.
    }
    const statusRaw = cols.statusCol !== undefined ? String(row[cols.statusCol] || "").trim() : "";
    if (statusRaw && !normText(statusRaw).includes("thanh cong")) {
      excludedFailedStatus++;
      continue;
    }
    const gross = cols.grossCol !== undefined ? Number(row[cols.grossCol]) || 0 : 0;
    if (!gross) continue;
    const netRaw = cols.netCol !== undefined ? row[cols.netCol] : null;
    const fee = cols.feeCol !== undefined ? Number(row[cols.feeCol]) || 0 : 0;
    const net = (netRaw !== null && netRaw !== undefined && netRaw !== "") ? (Number(netRaw) || 0) : (gross - fee);
    const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
    const date = toIsoDate(dateRaw);
    if (!date) continue;
    const chiNhanh = cols.chiNhanh !== undefined ? String(row[cols.chiNhanh] || "").trim() : diemThu;
    if (!chiNhanh) continue;

    let txKey = cols.maGiaoDich !== undefined ? String(row[cols.maGiaoDich] || "").trim() : "";
    if (!txKey) {
      // Du phong (file thuc te 24/07/2026 khong gap truong hop nay -- "Ma
      // giao dich" xac nhan 100% co gia tri) -- khoa gop theo ngay+chi
      // nhanh+so tien+dong, de van khu trung duoc trong 1 lan tai neu file
      // tuong lai co dong thieu Ma giao dich.
      fallbackSeq++;
      noTxKeyFallbackUsed++;
      txKey = `NOKEY|${date}|${chiNhanh}|${gross}|${fallbackSeq}`;
    }

    transactions.push({ txKey, date, chiNhanh, gross, fee, net });
  }

  return {
    sheetName,
    transactions,
    excludedFunzone,
    excludedFailedStatus,
    noTxKeyFallbackUsed,
  };
}

// ---------- Shared invoice upload: 1 file (MTT) -> ca 4 kenh cung luc ----------
// Luyen upload 1 file danh sach hoa don duy nhat tren TRANG NAO CUNG DUOC (Momo
// hoac Zalo/VNPay/Payoo) va no cap nhat luon ca 4 danh sach (momo, zalo, vnpay,
// payoo) trong 1 lan -- khong can upload lai file nay tren tung trang rieng.
// Dùng lai parseInvoiceWorkbook cua Momo (da on dinh, khong doi) cho tag momo,
// va parseInvoiceWorkbookByTag (generic) cho 3 tag con lai.
function parseSharedInvoiceWorkbook(buffer, companyKey) {
  const momoParsed = m.parseInvoiceWorkbook(buffer, companyKey);
  const zaloParsed = parseInvoiceWorkbookByTag(buffer, "zalo");
  const vnpayParsed = parseInvoiceWorkbookByTag(buffer, "vnpay");
  const payooParsed = parseInvoiceWorkbookByTag(buffer, "payoo");
  return {
    sheetName: momoParsed.sheetName || zaloParsed.sheetName || vnpayParsed.sheetName || payooParsed.sheetName,
    momo: momoParsed.invoices,
    zalo: zaloParsed.invoices,
    vnpay: vnpayParsed.invoices,
    payoo: payooParsed.invoices,
  };
}

module.exports = {
  extractZvpSettlements,
  parseInvoiceWorkbookByTag,
  parseSharedInvoiceWorkbook,
  parseOfflineVnpayWorkbook,
  parseDiemMappingSheet,
  parsePayooWorkbook,
  parsePayooRawReport,
  parsePayooDiemMapping,
  parseOnlineVnpayWorkbook,
  parseGianXuatHdSheet,
  parseProductCatalogSheet,
  resolveProductCatalogGross,
  applyGianRedirectToInvoices,
  applyGianRedirectToResolvedGross,
  learnGianFromInvoices,
  parseGianMasterSheet,
  mergeGianListWithMaster,
  mergeDiemMapWithMaster,
  buildOnlineProductMatcher,
  extractKeywords,
  resolveDiemGross,
  resolveOnlineGross,
  parseOnlineProductMapSheet,
  mergeOnlineProductMap,
  resolveOnlineGrossByProductMap,
  reconcileZvp,
  parseOrderDetailsWorkbook,
  parseFeeReportWorkbook,
  parseVnpayOfflineFeeReport,
  toIsoDate,
  isoToDmy,
  normCode,
  normText,
  dateRange,
  displayCode,
  FF_SUFFIX,
};
