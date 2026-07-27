// ---------- Doi soat Viet QR (BIDV7704 / BIDV77020 / MB11521268) ----------
// Different shape from Momo/Zalo-VNPay-Payoo: VietQR payments hit the bank
// account ONE TRANSACTION AT A TIME in near real time -- there is no daily
// or weekly batch settlement line to match against. Each bank "thu"
// transaction's description embeds a "VQR<code> PaymentForOrder" token,
// which is the SAME token found in the "Noi dung TT" column of the raw QR
// export ("Du Lieu" / "He thong" sheet) -- that raw row also carries "Ma
// cua hang" (an internal store code) for that one payment. "Ma cua hang"
// has no direct code join to "Ma cong trinh" (the accounting code); it
// only maps to a descriptive "Ten diem ban" / "Ten cua hang" (via the
// "Cua hang" sheet), which must be FUZZY matched against the invoice
// list's own "Ten diem xuat hoa don" -- same approach as ZVP's Online
// channel product matching, reused directly here.
//
// Because there's no bank-side settlement batch, this module treats each
// CALENDAR DAY's total "thu" transactions for a bank as its "settlement"
// (bankAmount = that day's sum), and reconciles gross (computed fresh from
// the raw QR export, resolved to gian via fuzzy matching) against invoices
// for that same day, exactly like Momo/ZVP do per settlement.

const XLSX = require("xlsx");
const m = require("./momoReconcile");
const toIsoDate = m.toIsoDate;
const isoToDmy = m.isoToDmy;
const normCode = m.normCode;
const normText = m.normText;
const displayCode = m.displayCode;
const FF_SUFFIX = m.FF_SUFFIX;
const buildPendingDaySettlements = m.buildPendingDaySettlements;
const zvp = require("./zvpReconcile");
const buildOnlineProductMatcher = zvp.buildOnlineProductMatcher;
const extractDayList = null; // (kept for clarity -- VietQR days come straight from bank tx dates, not a day-list tag)

// ---------- VQR code extraction (bank statement + raw export share this token) ----------

function extractVqrCode(text) {
  if (!text) return null;
  const m2 = String(text).match(/VQR[0-9A-Z]+/i);
  return m2 ? m2[0].toUpperCase() : null;
}

// A VietQR "settlement" is just "this bank's thu transactions on this
// calendar day" -- no batching to decode, unlike Momo/ZVP.
//
// Luyen, 2026-07-17: mot so dong "thu" tren sao ke KHONG phai tien ban ve QR
// (vd chuyen tien noi bo giua cac tai khoan, hoac 1 khoan tien lon bat
// thuong) -- nhung dong nay van la tien that vao tai khoan nen VAN giu
// nguyen trong Giao dich/Sao ke, chi loai KHOI tong "Ngan hang" cua doi soat
// VietQR (t.excludeFromVietQrRecon === true), tranh lam sai lech ca ngay chi
// vi 1 dong khong phai doanh thu ve QR.
//
// Luyen, 2026-07-18 (phat hien tu BIDV7702 ngay 10/07: ngan hang hien
// 66.539.919d nhung Luyen xac nhan thuc te chi 25.520.000d la tien QR that):
// tai khoan BIDV7702 dung CHUNG cho ca VietQR LAN mot so dong tien Momo (vd
// "KH705KVCMN0002/4" -- ma KH Momo, hoac "REM ... CONG TY CP DICH VU DI DONG
// TRUC TUYEN" -- settlement Momo tu M-Service/CONG TY CP DICH VU DI DONG
// TRUC TUYEN), khong phai doanh thu QR.
//
// LUU Y: KHONG dung "phai co token VQR..." de loc (thu qua roi, sai) -- mot
// so giao dich QR THAT (Luyen xac nhan, vd "Nguyen Hoang Long from Liobank"
// 50.000d ngay 10/07) khong he co token VQR trong dien giai (tuy ngan hang
// nguon: Liobank/MB/MSB... ghi ten nguoi chuyen thay vi ma VQR), neu doi hoi
// token se loai NHAM ca doanh thu QR that. Dung dung dac diem MOMO (ma
// KH###KVCMN#### hoac REM tu M-Service) de loai TRU, giu lai tat ca con lai.
const MOMO_TX_PATTERN = /KH\d+KVCMN\d+|DICH\s+VU\s+DI\s+DONG\s+TRUC\s+TUYEN/i;
// Tach rieng buoc loc (dung chung cho ca extractVietQrSettlements o duoi VA
// resolveGianGrossByBankRef, xem ghi chu tai do) khoi buoc gop theo ngay.
function extractVietQrThuTransactions(transactions) {
  return transactions.filter(
    (t) => t.type === "thu" && !t.excludeFromVietQrRecon && t.date && !MOMO_TX_PATTERN.test(t.description || "")
  );
}
function extractVietQrSettlements(transactions) {
  const byDate = {};
  for (const t of extractVietQrThuTransactions(transactions)) {
    if (!byDate[t.date]) {
      byDate[t.date] = { date: t.date, fromIso: t.date, toIso: t.date, amount: 0, txIds: [] };
    }
    byDate[t.date].amount += t.amount;
    byDate[t.date].txIds.push(t.id);
  }
  return Object.values(byDate).sort((a, b) => (a.date < b.date ? -1 : 1));
}

// The raw QR export's own "Thoi gian TT"/"Thoi gian tao" columns are
// formatted "DD-MM-YYYY HH:MM:SS" (DASH separated), unlike the invoice
// list's "Ngay HD" column (DD/MM/YYYY, SLASH separated) that the shared
// toIsoDate() helper above was written for. Silently returning null for
// every row here caused every transaction to be dropped during gross
// aggregation (resolveGianGross skips rows with no date) -- verified this
// against real "VIET QR ...xlsx" exports from all 3 banks, all of which
// use the dash format exclusively. This local parser handles that format
// specifically, only falling back to the shared helper for other shapes.
function parseVqrDate(raw) {
  if (!raw) return null;
  const datePart = String(raw).trim().split(" ")[0];
  const m = datePart.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return toIsoDate(datePart);
}

// ---------- Store export catalog ("store-export" -- standalone POS-name
// master list, independent of any one bank's raw QR transaction upload) ----------
// Luyen downloads this periodically straight from the POS backend to
// pre-load/refresh Ma cua hang -> Ten diem ban BEFORE a bank's transactions
// even exist for a code, so unlike parseCuaHangSheet (which only reads a
// "Cua hang" tab bundled INSIDE a raw QR file), this reads its own file with
// its own richer column set: STT | Ten cua hang | Ma cua hang | Ma diem ban
// | Ten diem ban | Doanh thu ngay | So luong GD ngay | Ngay tao. Only the
// first 4 data columns matter for reconciliation; the rest are ignored.
function parseStoreExportSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  for (const sheetName of wbLite.SheetNames) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (let r = 0; r < Math.min(grid.length, 6); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (idx.tenCuaHang === undefined && s.includes("ten cua hang")) idx.tenCuaHang = c;
        if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
        if (idx.maDiemBan === undefined && s.includes("ma diem ban")) idx.maDiemBan = c;
        if (idx.tenDiemBan === undefined && s.includes("ten diem ban")) idx.tenDiemBan = c;
      });
      if (idx.maCuaHang !== undefined && idx.tenDiemBan !== undefined) {
        const map = {};
        for (let rr = r + 1; rr < grid.length; rr++) {
          const dataRow = grid[rr] || [];
          const maCuaHang = String(dataRow[idx.maCuaHang] || "").trim();
          if (!maCuaHang) continue;
          const tenCuaHang = idx.tenCuaHang !== undefined ? String(dataRow[idx.tenCuaHang] || "").trim() : "";
          const maDiemBan = idx.maDiemBan !== undefined ? String(dataRow[idx.maDiemBan] || "").trim() : "";
          const tenDiemBan = String(dataRow[idx.tenDiemBan] || "").trim();
          map[maCuaHang] = { tenCuaHang, maDiemBan, tenDiemBan, matchText: `${tenCuaHang} ${tenDiemBan}`.trim() };
        }
        if (Object.keys(map).length > 0) return { sheetName, map };
      }
    }
  }
  throw new Error('Khong tim thay cot "Ma cua hang" va "Ten diem ban" trong file danh sach diem ban nay.');
}

// ---------- Raw QR export ("Du Lieu" / "He thong") ----------
// Columns observed (name varies slightly file to file, matched loosely):
// STT | Thoi gian TT | So tien den (VND) | [So tien di (VND)] | Loai |
// Trang thai | Ma tham chieu | Ma don hang | Ma diem ban | Ma cua hang |
// Tai khoan nhan | Thoi gian tao | Noi dung TT | Ghi chu | Loai giao dich
//
// IMPORTANT: which sheet actually holds this transaction log varies by
// bank/export -- e.g. the BIDV7704 file names its transaction log sheet
// "He thong" while its "Du Lieu" sheet (which the OLD name-based guess
// picked first) is really just a store-code/name lookup table with no
// header row at all. So this scans EVERY sheet for one whose header row
// actually contains the required columns, rather than trusting sheet
// names -- verified necessary against the real BIDV7704/BIDV77020/MB
// export files, which don't agree on sheet naming.
function findVietQrDataSheet(wbLite, buffer) {
  const nameHinted = wbLite.SheetNames.filter((n) => {
    const t = normText(n);
    return t.includes("du lieu") || t.includes("he thong");
  });
  const ordered = [...nameHinted, ...wbLite.SheetNames.filter((n) => !nameHinted.includes(n))];
  for (const sheetName of ordered) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (let r = 0; r < Math.min(grid.length, 10); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (idx.thoiGian === undefined && s.includes("thoi gian tt")) idx.thoiGian = c;
        if (idx.soTien === undefined && s.includes("so tien den")) idx.soTien = c;
        if (idx.loai === undefined && s === "loai") idx.loai = c;
        if (idx.trangThai === undefined && s.includes("trang thai")) idx.trangThai = c;
        if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
        if (idx.noiDung === undefined && s.includes("noi dung tt")) idx.noiDung = c;
        // Chi Nhan, 2026-07-24: them cot "Ma tham chieu" -- can rieng cho
        // tinh nang khop voi "So tham chieu" ben sao ke ngan hang (kenh
        // BIDV7702 tu 22/07 tro di, xem resolveGianGrossByBankRef ben duoi).
        if (idx.maThamChieu === undefined && s.includes("ma tham chieu")) idx.maThamChieu = c;
      });
      if (idx.soTien !== undefined && idx.maCuaHang !== undefined && idx.noiDung !== undefined) {
        return { sheetName, grid, headerRowIdx: r, cols: idx };
      }
    }
  }
  return null;
}

function parseVietQrRawWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const found = findVietQrDataSheet(wbLite, buffer);
  if (!found) {
    throw new Error('Khong doc duoc dong tieu de (can cot "So tien den", "Ma cua hang", "Noi dung TT") trong sheet du lieu QR.');
  }
  const { sheetName, grid, headerRowIdx, cols } = found;

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const trangThai = cols.trangThai !== undefined ? row[cols.trangThai] : null;
    if (trangThai !== null && trangThai !== undefined && !/thanh cong/i.test(normText(String(trangThai)))) continue;
    const loai = cols.loai !== undefined ? row[cols.loai] : null;
    if (loai !== null && loai !== undefined && !normText(String(loai)).includes("giao dich den")) continue;
    const amount = cols.soTien !== undefined ? Number(row[cols.soTien]) || 0 : 0;
    if (!amount) continue;
    const maCuaHang = cols.maCuaHang !== undefined ? String(row[cols.maCuaHang] || "").trim() : "";
    if (!maCuaHang) continue;
    const noiDung = cols.noiDung !== undefined ? String(row[cols.noiDung] || "") : "";
    const vqrCode = extractVqrCode(noiDung);
    const thoiGianRaw = cols.thoiGian !== undefined ? row[cols.thoiGian] : null;
    const date = parseVqrDate(thoiGianRaw);
    const refCode = cols.maThamChieu !== undefined ? String(row[cols.maThamChieu] || "").trim() : "";
    rows.push({ vqrCode, maCuaHang, amount, date, raw: noiDung, refCode });
  }

  return { sheetName, rows };
}

// ---------- Store catalog ("Cua hang" / "Cua Hang") ----------
// Ma cua hang -> a descriptive name, built from BOTH "Ten cua hang" and
// "Ten diem ban" so the fuzzy matcher against invoices has more text to
// work with (neither column alone reliably matches the invoice list's own
// "Ten diem xuat hoa don" wording).
function parseCuaHangSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => {
    const t = normText(n);
    return t.includes("cua hang");
  });
  if (sheetName) {
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
        if (idx.tenCuaHang === undefined && s.includes("ten cua hang")) idx.tenCuaHang = c;
        if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
        if (idx.tenDiemBan === undefined && s.includes("ten diem ban")) idx.tenDiemBan = c;
      });
      if (idx.maCuaHang !== undefined && (idx.tenCuaHang !== undefined || idx.tenDiemBan !== undefined)) {
        headerRowIdx = r;
        cols = idx;
        break;
      }
    }
    if (headerRowIdx >= 0) {
      const map = {};
      for (let r = headerRowIdx + 1; r < grid.length; r++) {
        const row = grid[r] || [];
        const maCuaHang = cols.maCuaHang !== undefined ? String(row[cols.maCuaHang] || "").trim() : "";
        if (!maCuaHang) continue;
        const tenCuaHang = cols.tenCuaHang !== undefined ? String(row[cols.tenCuaHang] || "").trim() : "";
        const tenDiemBan = cols.tenDiemBan !== undefined ? String(row[cols.tenDiemBan] || "").trim() : "";
        map[maCuaHang] = { tenCuaHang, tenDiemBan, matchText: `${tenCuaHang} ${tenDiemBan}`.trim() };
      }
      if (Object.keys(map).length > 0) return map;
    }
  }

  // Fallback: some exports (verified on the real BIDV7704 file) have NO
  // dedicated "Cua hang" catalog sheet at all -- the store code/name pairs
  // only appear as the first two columns of a headerless per-day pivot
  // summary sheet (e.g. "Du Lieu": code | name | day1 amount | day2 amount
  // | ...). Scan every sheet for rows shaped like that: col0 looks like a
  // store code (short alnum token, no spaces -- matches "Ma cua hang"
  // values seen on the transaction rows) and col1 is a real name (a
  // non-numeric, non-date string).
  const map = {};
  for (const sn of wbLite.SheetNames) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: [sn] });
    const ws = wb.Sheets[sn];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (const row of grid) {
      const code = row && row[0] != null ? String(row[0]).trim() : "";
      const name = row && row[1] != null ? String(row[1]).trim() : "";
      if (!/^[A-Z0-9]{6,15}$/i.test(code)) continue;
      if (!name || /^\d+([./-]\d+)*$/.test(name)) continue;
      if (!map[code]) map[code] = { tenCuaHang: name, tenDiemBan: "", matchText: name };
    }
  }
  return map;
}

// ---------- "Ten diem- Ma cong trinh" master table (BIDV7702, tu 22/07/2026) ----------
// Chi Nhan, 2026-07-24: bang tra cuu THU CONG rieng cua Luyen -- 2 cot that
// su ("ten diem ban", "ma cong trinh"), cac cot con lai (C/D/E) chi la du
// lieu rac/danh sach dropdown validation cua Excel, khong lien quan, phai bo
// qua. Dung de tra ma cong trinh CHINH XAC (khong fuzzy) tu ten diem ban da
// tra duoc qua "Ma cua hang" -> "Cua hang"/"store-export" o tren, danh rieng
// cho tinh nang khop theo "So tham chieu" ngan hang (xem resolveGianGrossByBankRef).
function parseTenDiemMaCongTrinhSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  for (const sheetName of wbLite.SheetNames) {
    const t = normText(sheetName);
    if (!(t.includes("ten diem") && t.includes("ma cong trinh"))) continue;
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
        if (idx.tenDiem === undefined && s.includes("ten diem")) idx.tenDiem = c;
        if (idx.maCongTrinh === undefined && s.includes("ma cong trinh")) idx.maCongTrinh = c;
      });
      if (idx.tenDiem !== undefined && idx.maCongTrinh !== undefined) {
        headerRowIdx = r;
        cols = idx;
        break;
      }
    }
    if (headerRowIdx < 0) continue;
    const map = {}; // normText(ten diem) -> ma cong trinh (nguyen ban, chua chuan hoa)
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const tenDiem = String(row[cols.tenDiem] || "").trim();
      const maCongTrinh = String(row[cols.maCongTrinh] || "").trim();
      if (!tenDiem || !maCongTrinh) continue;
      map[normText(tenDiem)] = maCongTrinh;
    }
    if (Object.keys(map).length > 0) return { sheetName, map };
  }
  throw new Error('Khong tim thay sheet "Ten diem- Ma cong trinh" hop le (can cot "Ten diem ban" va "Ma cong trinh") trong file nay.');
}

// ---------- Invoice list (shared MTT-style sheet, filtered by a bank-specific tag) ----------
// tagPattern: RegExp (or string used as RegExp source) tested against the
// "Dich vu thu ho" cell -- VietQR invoices are tagged per bank, e.g.
// "POSH+JP MB (7020) 4" for BIDV77020, rather than a single fixed keyword
// like Momo's "momo" -- so the caller supplies the pattern per channel.
// Also captures each invoice's own "Ten diem xuat hoa don" (tenDiem),
// needed here (unlike Momo/ZVP) to build the fuzzy-match candidate list
// directly from invoices, since VietQR has no separate "gian hang xuat HD"
// master sheet.
// Parses ONE candidate sheet (by name) into {sheetName, invoices} filtered
// by tagRe, or returns null if this sheet doesn't even have the required
// header columns (so the caller can just try the next candidate instead of
// throwing).
function parseInvoiceSheetByName(buffer, sheetName, tagRe) {
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
      // Ma diem / Ma cong trinh (cot MA ke toan thuc su): "Mã điểm trên misa
      // thuế" (sheet KH Cu, "-989") hoac "Mã điểm ghi chú HT Misa" (sheet KH
      // Moi, "-705"). Da kiem tra truc tiep tren file that ("MTT 16.07
      // (1).xlsx", sheet "-705", cac dong tag "MTD MN"): cot "Mã điểm ghi
      // chú HT Misa" chua CAC MA THAT (vd "ESTELLA PHN", "AM TP PHCM", "AE
      // BD PHCM"...) dung 1-1 cho tung gian, con cot "Mã NCC HT Misa" ben
      // canh no LUON LA HANG SO "KL" cho moi dong (khong phai ma rieng tung
      // gian) nen KHONG dung duoc lam maDiem -- nguoc lai voi phong doan ban
      // dau. Cot "Mã đối tượng nội bộ" (truoc do) moi la cot TEN mo ta (vd
      // "Posh Estella", "AE Tân Phú ghế").
      if (idx.maDiem === undefined && s.includes("ma diem tren")) idx.maDiem = c;
      if (idx.maDiem === undefined && s.includes("ma diem ghi chu")) idx.maDiem = c;
      if (s.includes("tong tt hd") || (s.includes("tong") && s.includes("hd"))) idx.tongTt = c;
      if (idx.dichVuThuHo === undefined && s.includes("thu ho")) idx.dichVuThuHo = c;
      if (idx.hinhThuc === undefined && s.includes("hinh thuc hop tac")) idx.hinhThuc = c;
      // Ten diem (cot TEN mo ta): "Tên điểm xuất hóa đơn" (KH Cu) hoac "Mã
      // đối tượng nội bộ" (KH Moi -- xem giai thich o tren).
      if (idx.tenDiem === undefined && s.includes("ten diem xuat hoa don")) idx.tenDiem = c;
      if (idx.tenDiem === undefined && s.includes("ma doi tuong noi bo")) idx.tenDiem = c;
    });
    if (idx.soHd !== undefined && idx.maDiem !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) return null;

  const invoices = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const dvth = cols.dichVuThuHo !== undefined ? row[cols.dichVuThuHo] : null;
    if (!dvth || !tagRe.test(String(dvth))) continue;
    const days = String(dvth)
      .trim()
      .match(/\d+(?:\s*[,+\/]\s*\d+)*\s*$/);
    const dayList = days
      ? days[0]
          .split(/[,+\/]/)
          .map((x) => parseInt(x.trim(), 10))
          .filter((x) => !isNaN(x))
      : [];
    const ngayHdRaw = cols.ngayHd !== undefined ? row[cols.ngayHd] : null;
    let maDiem = normCode(cols.maDiem !== undefined ? row[cols.maDiem] : null);
    const hinhThuc = cols.hinhThuc !== undefined ? String(row[cols.hinhThuc] || "").toUpperCase() : "";
    if (hinhThuc.includes("CSE") && !maDiem.endsWith(FF_SUFFIX)) {
      maDiem = maDiem + FF_SUFFIX;
    }
    const tenDiem = cols.tenDiem !== undefined ? String(row[cols.tenDiem] || "").trim() : "";

    invoices.push({
      soHd: cols.soHd !== undefined ? row[cols.soHd] : null,
      ngayHd: toIsoDate(ngayHdRaw),
      thang: cols.thang !== undefined ? row[cols.thang] : null,
      maDiem,
      tenDiem,
      tongTt: cols.tongTt !== undefined ? Number(row[cols.tongTt]) || 0 : 0,
      days: dayList,
      raw: String(dvth).trim(),
    });
  }

  return { sheetName, invoices };
}

// Nhieu cong ty (KH Cu / KH Moi) dung CHUNG 1 file hoa don MTT nhung MOI
// cong ty co sheet "ke ds xuat..." rieng theo ma khach hang (vd "-989" cho
// KH Cu, "-705" cho KH Moi) -- chon sheet CHI theo TEN (uu tien "989") tung
// lam BIDV7702/KH Moi (tag "MTD MN", chi co tren sheet "-705") luon ra 0 hoa
// don vi luon bi doc nham sang sheet "-989". Fix: thu doc TUNG sheet ung
// vien va DUNG sheet dau tien THUC SU co dong khop dung tagPattern cua kenh
// nay, thay vi doan theo ten sheet -- chi fallback ve cach doan ten CU (uu
// tien "989") neu KHONG sheet nao khop (giu nguyen hanh vi cu cho truong hop
// chua tung gap).
function parseInvoiceWorkbookByTag(buffer, tagPattern) {
  const tagRe = tagPattern instanceof RegExp ? tagPattern : new RegExp(tagPattern, "i");
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const kdsCandidates = wbLite.SheetNames.filter((n) => /k.\s*ds\s*xu.t/i.test(n));
  const candidates =
    kdsCandidates.length > 0
      ? kdsCandidates
      : wbLite.SheetNames.filter((n) => {
          const t = normText(n);
          return t.includes("hoa don") || (t.includes("danh sach") && t.includes("don")) || t.includes("so hoa don");
        });
  if (candidates.length === 0) {
    throw new Error("Khong tim thay sheet danh sach hoa don trong file.");
  }

  for (const sheetName of candidates) {
    const parsed = parseInvoiceSheetByName(buffer, sheetName, tagRe);
    if (parsed && parsed.invoices.length > 0) return parsed;
  }

  // Fallback: khong sheet nao co dong khop -- giu hanh vi cu (uu tien "989")
  // de khong lam doi ket qua cho truong hop chua gap (van co the tra ve 0
  // hoa don, giong nhu truoc day, thay vi bao loi).
  const legacyName = candidates.find((n) => /989/.test(n)) || candidates[0];
  const legacyParsed = parseInvoiceSheetByName(buffer, legacyName, tagRe);
  if (legacyParsed) return legacyParsed;
  throw new Error('Khong doc duoc dong tieu de (can cot "So HD" va "Ma diem") trong sheet hoa don.');
}

// Builds the fuzzy-match candidate list directly from the invoice data
// itself (deduped by maDiem), since VietQR has no separate "gian hang xuat
// HD" master sheet the way Momo/ZVP's Online channel does.
function buildGianCandidatesFromInvoices(invoices) {
  const byCode = {};
  for (const inv of invoices) {
    if (!inv.maDiem || !inv.tenDiem) continue;
    const baseCode = inv.maDiem.endsWith(FF_SUFFIX) ? inv.maDiem.slice(0, -FF_SUFFIX.length) : inv.maDiem;
    if (!byCode[baseCode]) {
      byCode[baseCode] = { tenDiem: inv.tenDiem, maCongTrinh: baseCode, isCse: inv.maDiem.endsWith(FF_SUFFIX) };
    }
  }
  return Object.values(byCode);
}

// Resolves each raw QR row's "Ma cua hang" to a Ma Cong Trinh gian code via
// fuzzy text matching (store's combined name vs invoice-derived candidate
// list), then aggregates gross revenue by (code, date). Rows whose store
// code can't be confidently matched are surfaced as "unmapped" so Luyen can
// review/alias them instead of revenue silently vanishing.
// Unmapped rows are AGGREGATED by their raw "Ma cua hang" label (count + tong
// tien) rather than listed one bullet per row -- a raw code that's just "-"
// (no store code on that transaction at all, seen on some genuine incoming
// QR rows) used to show up as a bare, uninformative "-" bullet with zero
// context, repeated every single reload; showing the count and total amount
// instead makes clear there IS real, uncounted revenue behind it and roughly
// how much, so Luyen can decide whether it's worth tracking down.
// defaultBlankCode: Luyen, 2026-07-18 (BIDV7702) -- giao dich KHONG co ma cua
// hang/vqrCode nao het (vd chuyen khoan thuong tu Liobank, khong qua QR) thi
// KHONG THE gan tung dong qua nocode-assign (can 1 vqrCode duy nhat lam key,
// dong nay khong co). Thay vi de tien "boc hoi" trong unmapped mai mai, neu
// kenh co cau hinh 1 ma cong trinh mac dinh (vd "AM TP PHCM" - BIDV TAN PHU)
// thi cong luon doanh thu cac dong nay vao do, ap dung tu dong cho ca du lieu
// tai sau nay (khong can lam lai tay). Chi ap dung cho dong THUC SU khong co
// ma cua hang (isBlankCode) va khong khop duoc gian nao qua fuzzy matcher.
function resolveGianGross(rawRows, storeNameMap, gianCandidates, nocodeAssignments, defaultBlankCode) {
  const matcher = buildOnlineProductMatcher(gianCandidates);
  const grossByCode = {};
  const codes = new Set();
  const assignments = nocodeAssignments || {};
  // Keyed by raw "Ma cua hang" (not the display label) so the caller can
  // build a per-code mapping form (Luyen, 2026-07-17: "map giữ tên điểm nội
  // bộ ... các mã nào chưa map được hiện ra cho tôi" -- she wants each
  // unresolved store code actionable, not just a read-only warning string).
  const unmappedAgg = new Map();
  // Individual "không có mã cửa hàng" rows, kept separately (in addition to
  // the aggregated bucket above) so the caller can offer a PER-TRANSACTION
  // fix button -- Luyen, 2026-07-17: "còn lỗi chưa map dc cái nào thì báo
  // cho tôi và có nút sửa nhá". Each such row only has a raw QR ref
  // (row.vqrCode), no store name/code at all, so it can't be grouped by
  // code like a normal unmapped store -- has to be assignable one by one.
  const blankRows = [];
  for (const row of rawRows) {
    if (!row.date) continue;
    const isBlankCode = !row.maCuaHang || row.maCuaHang === "-";
    // Luyen has manually assigned this SPECIFIC transaction (by its unique
    // vqrCode) to a Ma cong trinh -- route its amount straight there instead
    // of falling into the unmapped bucket, no fuzzy-matching needed.
    if (isBlankCode && row.vqrCode && assignments[row.vqrCode]) {
      const code = assignments[row.vqrCode].targetCode || assignments[row.vqrCode];
      codes.add(code);
      const key = `${row.date}|${code}`;
      grossByCode[key] = (grossByCode[key] || 0) + row.amount;
      continue;
    }
    const storeInfo = storeNameMap[row.maCuaHang];
    const matchText = storeInfo ? storeInfo.matchText : row.maCuaHang;
    const match = matcher(matchText);
    if (!match) {
      // Luyen, 2026-07-19: "ngày nào ngân hàng trả tiền dư so với hệ thống
      // thì cho vô gian Tân Phú" -- truoc day chi giao dich KHONG co ma cua
      // hang moi duoc gap vao defaultBlankCode, con giao dich CO ma cua hang
      // nhung khong khop duoc voi hoa don nao (vd "SB CAN THO PHCM", "FARM
      // LOTTE NHA TRANG" -- chua co hoa don) bi loai hoan toan khoi
      // grossByCode, lam tong ngay do "hut" so voi ngan hang that (Chenh
      // lech am) ma khong the note ra gian nao. Gio: BAT KY giao dich nao
      // khong khop (co ma hay khong ma) deu gap vao defaultBlankCode (Tan
      // Phu) khi kenh co cau hinh nay, de tien khong "boc hoi" khoi tong
      // ngay -- van giu ghi nhan trong unmappedAgg/unmappedDetails phia duoi
      // de Luyen biet khoan nao da duoc gap vao, phong khi muon bo sung hoa
      // don/mapping that cho no sau nay.
      if (defaultBlankCode) {
        codes.add(defaultBlankCode);
        const key = `${row.date}|${defaultBlankCode}`;
        grossByCode[key] = (grossByCode[key] || 0) + row.amount;
      }
      const key = isBlankCode ? "" : row.maCuaHang;
      const agg = unmappedAgg.get(key) || {
        maCuaHang: row.maCuaHang || "",
        matchText: storeInfo ? storeInfo.matchText : "",
        isBlankCode,
        count: 0,
        total: 0,
        foldedIntoDefault: !!defaultBlankCode,
      };
      agg.count += 1;
      agg.total += row.amount;
      unmappedAgg.set(key, agg);
      if (isBlankCode) {
        blankRows.push({ vqrCode: row.vqrCode || "", date: row.date, amount: row.amount, raw: row.raw || "" });
      }
      continue;
    }
    const code = match.isCse ? match.maCongTrinh + FF_SUFFIX : match.maCongTrinh;
    codes.add(code);
    const key = `${row.date}|${code}`;
    grossByCode[key] = (grossByCode[key] || 0) + row.amount;
  }
  const unmappedDetails = Array.from(unmappedAgg.values()).sort((a, b) => b.total - a.total);
  const unmapped = unmappedDetails.map((agg) => {
    const label = agg.isBlankCode
      ? "Giao dịch không có mã cửa hàng (ghi \"-\")"
      : `${agg.maCuaHang}${agg.matchText ? " (" + agg.matchText + ")" : ""}`;
    const note = agg.foldedIntoDefault ? " -- đã gộp vào Tân Phú" : "";
    return `${label}: ${agg.count} giao dịch, tổng ${agg.total.toLocaleString("vi-VN")}đ${note}`;
  });
  blankRows.sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  return { codes: Array.from(codes), grossByCode, unmapped, unmappedDetails, blankRows };
}

// ---------- BIDV7702: khop theo "So tham chieu" ngan hang (tu 22/07/2026) ----------
// Chi Nhan, 2026-07-24 (yeu cau Luyen): "check từ số tham chiếu qua ngân
// hàng lấy ngân hàng làm chuẩn". Khac voi resolveGianGross o tren (fuzzy
// text, van dung cho cac ngay TRUOC 22/07), tu 22/07 tro di NGAN HANG la
// chuan: moi giao dich "thu" tren sao ke duoc gan vao dung 1 dong QR co CUNG
// "Ma tham chieu" (ca 2 file deu co cot nay -- xem findVietQrDataSheet/
// parseVietQrRawWorkbook va utils/bankStatementParser.js), roi tra Ma cua
// hang -> Ten diem ban (qua storeNameMap, giong het truoc gio) -> Ma cong
// trinh (qua bang tra cuu rieng "Ten diem- Ma cong trinh", KHOP TUYET DOI
// sau khi chuan hoa, KHONG fuzzy). 3 truong hop canh bao theo dung yeu cau:
//  1) Giao dich ngan hang KHONG tim thay Ma tham chieu nao khop (o bat ky
//     ngay nao) -> unmatchedBankTx (canh bao do).
//  2) Tim thay nhung o NGAY KHAC voi ngay ngan hang bao tien ve (tien den QR
//     tu ngay truoc, ngan hang moi ghi nhan tien sau) -> van cong doanh thu
//     vao dung NGAY NGAN HANG (ngan hang la chuan), nhung them 1 canh bao
//     rieng -- lateMatches (hien dong xanh, chi mang tinh thong bao).
//  3) Ma cua hang MOI (chua co trong storeNameMap) hoac Ten diem ban chua co
//     trong bang tra cuu Ma cong trinh -> unmappedStoreCodes/unmappedTenDiem.
// Luyen, 2026-07-24: "mã cửa hàng mới này ... chỗ chọn mã công trình để gán
// vào nhá" -- truoc gio muc "Ma cua hang MOI, chua co trong danh sach diem
// ban" chi bao/liet ke, khong co cach gan truc tiep (phai cho toi khi co file
// "Danh sach diem ban" moi de bo sung Ten diem ban, roi file "Ten diem - Ma
// cong trinh" phai co dung ten do nua -- 2 buoc gian tiep). storeCodeOverride
// (store.viet_qr_store_code_override[channel], xem routes/doisoat-vietqr.js)
// cho phep gan THANG 1 Ma cua hang -> 1 Ma cong trinh, bo qua ca 2 buoc gian
// tiep tren, ap dung ngay khong can tai lai file nao. Kiem tra TRUOC storeNameMap
// nen luon uu tien neu da gan thu cong.
// refOverride: Luyen, 2026-07-27 -- gan THANG 1 So tham chieu ngan hang ->
// 1 Ma cong trinh, dung cho giao dich ngan hang THAT SU khong co dong QR
// nao khop (vd file "Danh sach diem ban"/rawRows khong co dong nao co Ma
// tham chieu do -- se roi vao unmatchedBankTx va bi tru khoi "Ngan hang"
// boi buildChannelReconciliation neu khong co override nay), nhung Luyen
// tu xac dinh duoc dung tien do thuoc gian nao. Kiem tra TRUOC ca buoc tim
// rawRow theo ref, nen boi qua hoan toan yeu cau phai co dong QR khop.
function resolveGianGrossByBankRef(bankTxs, rawRows, storeNameMap, tenDiemToProjectMap, storeCodeOverride, refOverride) {
  const override = storeCodeOverride || {};
  const refOv = refOverride || {};
  const refIndex = new Map(); // Ma tham chieu -> [rawRow, ...]
  rawRows.forEach((row) => {
    const ref = (row.refCode || "").trim();
    if (!ref) return;
    if (!refIndex.has(ref)) refIndex.set(ref, []);
    refIndex.get(ref).push(row);
  });

  const grossByCode = {};
  const codes = new Set();
  const unmatchedBankTx = [];
  const lateMatches = [];
  const unmappedStoreCodesAgg = new Map();
  const unmappedTenDiemAgg = new Map();

  for (const tx of bankTxs) {
    const ref = (tx.reference || "").trim();
    const refOverrideCode = ref ? refOv[ref] : undefined;
    if (refOverrideCode) {
      codes.add(refOverrideCode);
      const key = `${tx.date}|${refOverrideCode}`;
      grossByCode[key] = (grossByCode[key] || 0) + tx.amount;
      continue;
    }
    const candidates = ref ? refIndex.get(ref) || [] : [];
    // Uu tien dong QR CUNG NGAY voi giao dich ngan hang; neu khong co, lay
    // dong som nhat trong so cac dong trung ma (de bao "tien den tre" nhat quan).
    let raw = candidates.find((r) => r.date === tx.date);
    if (!raw && candidates.length > 0) {
      raw = candidates.slice().sort((a, b) => (a.date || "").localeCompare(b.date || ""))[0];
    }
    if (!raw) {
      unmatchedBankTx.push({
        date: tx.date,
        amount: tx.amount,
        reference: tx.reference || "",
        description: tx.description || "",
      });
      continue;
    }
    if (raw.date && raw.date !== tx.date) {
      lateMatches.push({
        bankDate: tx.date,
        rawDate: raw.date,
        reference: ref,
        amount: tx.amount,
        maCuaHang: raw.maCuaHang || "",
      });
    }
    const overrideCode = override[raw.maCuaHang];
    if (overrideCode) {
      codes.add(overrideCode);
      const key = `${tx.date}|${overrideCode}`;
      grossByCode[key] = (grossByCode[key] || 0) + tx.amount;
      continue;
    }
    const storeInfo = storeNameMap[raw.maCuaHang];
    if (!storeInfo) {
      const agg = unmappedStoreCodesAgg.get(raw.maCuaHang) || {
        maCuaHang: raw.maCuaHang || "",
        count: 0,
        total: 0,
        sampleRaw: raw.raw || "",
      };
      agg.count += 1;
      agg.total += tx.amount;
      unmappedStoreCodesAgg.set(raw.maCuaHang, agg);
      continue;
    }
    const tenDiem = (storeInfo.tenDiemBan || storeInfo.matchText || "").trim();
    const maCongTrinh = tenDiem ? tenDiemToProjectMap[normText(tenDiem)] : undefined;
    if (!maCongTrinh) {
      const key = tenDiem || raw.maCuaHang;
      const agg = unmappedTenDiemAgg.get(key) || {
        tenDiemBan: tenDiem,
        maCuaHang: raw.maCuaHang || "",
        count: 0,
        total: 0,
      };
      agg.count += 1;
      agg.total += tx.amount;
      unmappedTenDiemAgg.set(key, agg);
      continue;
    }
    codes.add(maCongTrinh);
    const key = `${tx.date}|${maCongTrinh}`;
    grossByCode[key] = (grossByCode[key] || 0) + tx.amount;
  }

  return {
    codes: Array.from(codes),
    grossByCode,
    unmatchedBankTx: unmatchedBankTx.sort((a, b) => (a.date || "").localeCompare(b.date || "")),
    lateMatches: lateMatches.sort((a, b) => (a.bankDate || "").localeCompare(b.bankDate || "")),
    unmappedStoreCodes: Array.from(unmappedStoreCodesAgg.values()).sort((a, b) => b.total - a.total),
    unmappedTenDiem: Array.from(unmappedTenDiemAgg.values()).sort((a, b) => b.total - a.total),
  };
}

// ---------- Viet QR MN (BIDV7702 / KH Moi's own "VIETQR MN 7702.xlsx"
// export -- structurally different from the 3 bank exports above: no "Noi
// dung TT"/VQR-token column at all, so there's no shared join token with a
// bank statement description. Uses its own "Ma tham chieu" column (unique
// per transaction) for de-dup instead, and its own "Ma Cua Hang APP" sheet
// for store names -- bundled in the SAME workbook, unlike the other 3
// channels' separate "Cua hang" tab. Store names here ("AMTP 01", "AMBD
// 05" ...) are literal PREFIX abbreviations of the real Ma cong trinh code
// ("AM TP KVCM", "AM BD KVCM"), not free-text site names, so matching uses
// a PREFIX rule (buildGianPrefixMatcher) instead of the fuzzy text matcher
// used above -- confirmed with Luyen 2026-07-16. ----------

function findVietQrMnDataSheet(wbLite, buffer) {
  const nameHinted = wbLite.SheetNames.filter((n) => normText(n).includes("viet qr"));
  const ordered = [...nameHinted, ...wbLite.SheetNames.filter((n) => !nameHinted.includes(n))];
  for (const sheetName of ordered) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (let r = 0; r < Math.min(grid.length, 10); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (idx.thoiGian === undefined && s.includes("thoi gian tt")) idx.thoiGian = c;
        if (idx.soTien === undefined && s.includes("so tien den")) idx.soTien = c;
        if (idx.trangThai === undefined && s.includes("trang thai")) idx.trangThai = c;
        if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
        if (idx.maThamChieu === undefined && s.includes("ma tham chieu")) idx.maThamChieu = c;
      });
      if (idx.soTien !== undefined && idx.maCuaHang !== undefined && idx.trangThai !== undefined) {
        return { sheetName, grid, headerRowIdx: r, cols: idx };
      }
    }
  }
  return null;
}

// vqrCode field is filled with "Ma tham chieu" here (unique per giao dich)
// instead of a real VQR token -- reuses the SAME field name on purpose so
// doisoat-vietqr.js's mergeRawRows (de-dup by row.vqrCode across uploads)
// works unchanged for this channel too.
function parseVietQrMnRawWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const found = findVietQrMnDataSheet(wbLite, buffer);
  if (!found) {
    throw new Error('Khong doc duoc dong tieu de (can cot "So tien den", "Ma cua hang", "Trang thai") trong sheet "VIET QR".');
  }
  const { sheetName, grid, headerRowIdx, cols } = found;

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const trangThai = cols.trangThai !== undefined ? row[cols.trangThai] : null;
    if (trangThai !== null && trangThai !== undefined && !/thanh cong/i.test(normText(String(trangThai)))) continue;
    const amount = cols.soTien !== undefined ? Number(row[cols.soTien]) || 0 : 0;
    if (!amount) continue;
    const maCuaHang = cols.maCuaHang !== undefined ? String(row[cols.maCuaHang] || "").trim() : "";
    if (!maCuaHang) continue;
    const thoiGianRaw = cols.thoiGian !== undefined ? row[cols.thoiGian] : null;
    const date = parseVqrDate(thoiGianRaw);
    const maThamChieu = cols.maThamChieu !== undefined ? String(row[cols.maThamChieu] || "").trim() : "";
    rows.push({ vqrCode: maThamChieu || null, maCuaHang, amount, date, raw: maThamChieu });
  }

  return { sheetName, rows };
}

// "Ma Cua Hang APP" sheet: STT | Ten cua hang | Ma cua hang | Ma diem ban |
// Ten diem ban | Doanh thu ngay | So luong GD ngay | Ngay tao -- only "Ten
// cua hang"/"Ma cua hang" matter here (matchText = Ten cua hang itself,
// used directly by the prefix matcher below).
function parseMaCuaHangAppSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => normText(n).includes("ma cua hang"));
  if (!sheetName) {
    throw new Error('Khong tim thay sheet "Ma Cua Hang APP" trong file nay.');
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
      if (idx.tenCuaHang === undefined && s.includes("ten cua hang")) idx.tenCuaHang = c;
      if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
    });
    if (idx.maCuaHang !== undefined && idx.tenCuaHang !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de (can cot "Ten cua hang", "Ma cua hang") trong sheet "Ma Cua Hang APP".');
  }

  const map = {};
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const maCuaHang = cols.maCuaHang !== undefined ? String(row[cols.maCuaHang] || "").trim() : "";
    if (!maCuaHang) continue;
    const tenCuaHang = cols.tenCuaHang !== undefined ? String(row[cols.tenCuaHang] || "").trim() : "";
    if (!tenCuaHang) continue;
    map[maCuaHang] = { tenCuaHang, tenDiemBan: "", matchText: tenCuaHang };
  }
  return map;
}

// "AMTP 01" -> "AMTP" (bo so chay + khoang trang/gach ngang cuoi cung).
function extractStorePrefix(tenCuaHang) {
  return String(tenCuaHang || "")
    .trim()
    .replace(/[\s-]*\d+\s*$/, "")
    .trim();
}

// Khop tien to cua ten cua hang VietQR MN (vd "AMTP" tu "AMTP 01") voi ma
// cong trinh cua TUNG gian ung vien (da chuan hoa, bo khoang trang -- vd "AM
// TP KVCM" -> "amtpkvcm") -- CHI khop khi tien to la .startsWith() cua DUNG
// 1 ung vien; mo ho (2+ ung vien) hoac khong ung vien nao tra ve null de
// Luyen tu gan thu cong, thay vi doan sai theo Luyen yeu cau ro.
function buildGianPrefixMatcher(gianCandidates) {
  const normalized = gianCandidates.map((c) => ({
    ...c,
    normCode: normText(c.maCongTrinh).replace(/\s+/g, ""),
  }));
  return function matchByPrefix(tenCuaHang) {
    const prefix = normText(extractStorePrefix(tenCuaHang)).replace(/\s+/g, "");
    if (!prefix) return null;
    const matches = normalized.filter((c) => c.normCode.startsWith(prefix));
    if (matches.length !== 1) return null;
    return matches[0];
  };
}

// Song song voi resolveGianGross o tren nhung dung buildGianPrefixMatcher
// thay vi buildOnlineProductMatcher (fuzzy text) -- dung cho kenh BIDV7702/
// VietQR MN.
function resolveGianGrossPrefix(rawRows, storeNameMap, gianCandidates) {
  const matcher = buildGianPrefixMatcher(gianCandidates);
  const grossByCode = {};
  const codes = new Set();
  const unmappedAgg = new Map();
  for (const row of rawRows) {
    if (!row.date) continue;
    const storeInfo = storeNameMap[row.maCuaHang];
    const matchText = storeInfo ? storeInfo.matchText : row.maCuaHang;
    const match = matcher(matchText);
    if (!match) {
      const isBlankCode = !row.maCuaHang || row.maCuaHang === "-";
      const label = isBlankCode
        ? "Giao dịch không có mã cửa hàng (ghi \"-\")"
        : `${row.maCuaHang}${storeInfo ? " (" + storeInfo.matchText + ")" : ""}`;
      const agg = unmappedAgg.get(label) || { count: 0, total: 0 };
      agg.count += 1;
      agg.total += row.amount;
      unmappedAgg.set(label, agg);
      continue;
    }
    const code = match.isCse ? match.maCongTrinh + FF_SUFFIX : match.maCongTrinh;
    codes.add(code);
    const key = `${row.date}|${code}`;
    grossByCode[key] = (grossByCode[key] || 0) + row.amount;
  }
  const unmapped = Array.from(unmappedAgg.entries()).map(
    ([label, agg]) => `${label}: ${agg.count} giao dịch, tổng ${agg.total.toLocaleString("vi-VN")}đ`
  );
  return { codes: Array.from(codes), grossByCode, unmapped };
}

// ---------- Viet QR MN (BIDV7702 / KH Moi's own "VIETQR MN 7702.xlsx"
// export -- structurally different from the 3 bank exports above: no "Noi
// dung TT"/VQR-token column at all, so there's no shared join token with a
// bank statement description. Uses its own "Ma tham chieu" column (unique
// per transaction) for de-dup instead, and its own "Ma Cua Hang APP" sheet
// for store names -- bundled in the SAME workbook, unlike the other 3
// channels' separate "Cua hang" tab. Store names here ("AMTP 01", "AMBD
// 05" ...) are literal PREFIX abbreviations of the real Ma cong trinh code
// ("AM TP KVCM", "AM BD KVCM"), not free-text site names, so matching uses
// a PREFIX rule (buildGianPrefixMatcher) instead of the fuzzy text matcher
// used above -- confirmed with Luyen 2026-07-16. ----------

function findVietQrMnDataSheet(wbLite, buffer) {
  const nameHinted = wbLite.SheetNames.filter((n) => normText(n).includes("viet qr"));
  const ordered = [...nameHinted, ...wbLite.SheetNames.filter((n) => !nameHinted.includes(n))];
  for (const sheetName of ordered) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (let r = 0; r < Math.min(grid.length, 10); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (idx.thoiGian === undefined && s.includes("thoi gian tt")) idx.thoiGian = c;
        if (idx.soTien === undefined && s.includes("so tien den")) idx.soTien = c;
        if (idx.trangThai === undefined && s.includes("trang thai")) idx.trangThai = c;
        if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
        if (idx.maThamChieu === undefined && s.includes("ma tham chieu")) idx.maThamChieu = c;
      });
      if (idx.soTien !== undefined && idx.maCuaHang !== undefined && idx.trangThai !== undefined) {
        return { sheetName, grid, headerRowIdx: r, cols: idx };
      }
    }
  }
  return null;
}

// vqrCode field is filled with "Ma tham chieu" here (unique per giao dich)
// instead of a real VQR token -- reuses the SAME field name on purpose so
// doisoat-vietqr.js's mergeRawRows (de-dup by row.vqrCode across uploads)
// works unchanged for this channel too.
function parseVietQrMnRawWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const found = findVietQrMnDataSheet(wbLite, buffer);
  if (!found) {
    throw new Error('Khong doc duoc dong tieu de (can cot "So tien den", "Ma cua hang", "Trang thai") trong sheet "VIET QR".');
  }
  const { sheetName, grid, headerRowIdx, cols } = found;

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const trangThai = cols.trangThai !== undefined ? row[cols.trangThai] : null;
    if (trangThai !== null && trangThai !== undefined && !/thanh cong/i.test(normText(String(trangThai)))) continue;
    const amount = cols.soTien !== undefined ? Number(row[cols.soTien]) || 0 : 0;
    if (!amount) continue;
    const maCuaHang = cols.maCuaHang !== undefined ? String(row[cols.maCuaHang] || "").trim() : "";
    if (!maCuaHang) continue;
    const thoiGianRaw = cols.thoiGian !== undefined ? row[cols.thoiGian] : null;
    const date = parseVqrDate(thoiGianRaw);
    const maThamChieu = cols.maThamChieu !== undefined ? String(row[cols.maThamChieu] || "").trim() : "";
    rows.push({ vqrCode: maThamChieu || null, maCuaHang, amount, date, raw: maThamChieu });
  }

  return { sheetName, rows };
}

// "Ma Cua Hang APP" sheet: STT | Ten cua hang | Ma cua hang | Ma diem ban |
// Ten diem ban | Doanh thu ngay | So luong GD ngay | Ngay tao -- only "Ten
// cua hang"/"Ma cua hang" matter here (matchText = Ten cua hang itself,
// used directly by the prefix matcher below).
function parseMaCuaHangAppSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames.find((n) => normText(n).includes("ma cua hang"));
  if (!sheetName) {
    throw new Error('Khong tim thay sheet "Ma Cua Hang APP" trong file nay.');
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
      if (idx.tenCuaHang === undefined && s.includes("ten cua hang")) idx.tenCuaHang = c;
      if (idx.maCuaHang === undefined && s.includes("ma cua hang")) idx.maCuaHang = c;
    });
    if (idx.maCuaHang !== undefined && idx.tenCuaHang !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de (can cot "Ten cua hang", "Ma cua hang") trong sheet "Ma Cua Hang APP".');
  }

  const map = {};
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const maCuaHang = cols.maCuaHang !== undefined ? String(row[cols.maCuaHang] || "").trim() : "";
    if (!maCuaHang) continue;
    const tenCuaHang = cols.tenCuaHang !== undefined ? String(row[cols.tenCuaHang] || "").trim() : "";
    if (!tenCuaHang) continue;
    map[maCuaHang] = { tenCuaHang, tenDiemBan: "", matchText: tenCuaHang };
  }
  return map;
}

// "AMTP 01" -> "AMTP" (bo so chay + khoang trang/gach ngang cuoi cung).
function extractStorePrefix(tenCuaHang) {
  return String(tenCuaHang || "")
    .trim()
    .replace(/[\s-]*\d+\s*$/, "")
    .trim();
}

// Khop tien to cua ten cua hang VietQR MN (vd "AMTP" tu "AMTP 01") voi ma
// cong trinh cua TUNG gian ung vien (da chuan hoa, bo khoang trang -- vd "AM
// TP KVCM" -> "amtpkvcm") -- CHI khop khi tien to la .startsWith() cua DUNG
// 1 ung vien; mo ho (2+ ung vien) hoac khong ung vien nao tra ve null de
// Luyen tu gan thu cong, thay vi doan sai theo Luyen yeu cau ro.
function buildGianPrefixMatcher(gianCandidates) {
  const normalized = gianCandidates.map((c) => ({
    ...c,
    normCode: normText(c.maCongTrinh).replace(/\s+/g, ""),
  }));
  return function matchByPrefix(tenCuaHang) {
    const prefix = normText(extractStorePrefix(tenCuaHang)).replace(/\s+/g, "");
    if (!prefix) return null;
    const matches = normalized.filter((c) => c.normCode.startsWith(prefix));
    if (matches.length !== 1) return null;
    return matches[0];
  };
}

// Song song voi resolveGianGross o tren nhung dung buildGianPrefixMatcher
// thay vi buildOnlineProductMatcher (fuzzy text) -- dung cho kenh BIDV7702/
// VietQR MN.
function resolveGianGrossPrefix(rawRows, storeNameMap, gianCandidates) {
  const matcher = buildGianPrefixMatcher(gianCandidates);
  const grossByCode = {};
  const codes = new Set();
  const unmappedAgg = new Map();
  for (const row of rawRows) {
    if (!row.date) continue;
    const storeInfo = storeNameMap[row.maCuaHang];
    const matchText = storeInfo ? storeInfo.matchText : row.maCuaHang;
    const match = matcher(matchText);
    if (!match) {
      const isBlankCode = !row.maCuaHang || row.maCuaHang === "-";
      const label = isBlankCode
        ? "Giao dịch không có mã cửa hàng (ghi \"-\")"
        : `${row.maCuaHang}${storeInfo ? " (" + storeInfo.matchText + ")" : ""}`;
      const agg = unmappedAgg.get(label) || { count: 0, total: 0 };
      agg.count += 1;
      agg.total += row.amount;
      unmappedAgg.set(label, agg);
      continue;
    }
    const code = match.isCse ? match.maCongTrinh + FF_SUFFIX : match.maCongTrinh;
    codes.add(code);
    const key = `${row.date}|${code}`;
    grossByCode[key] = (grossByCode[key] || 0) + row.amount;
  }
  const unmapped = Array.from(unmappedAgg.entries()).map(
    ([label, agg]) => `${label}: ${agg.count} giao dịch, tổng ${agg.total.toLocaleString("vi-VN")}đ`
  );
  return { codes: Array.from(codes), grossByCode, unmapped };
}

// diemAlias: same shared table as Momo/ZVP (store.invoice_diem_alias). Chi
// Nhan, 2026-07-24 (fix): bang nay dung CHUNG cho ca Momo/ZVP/VietQR, nhung 1
// alias (vd "LM NHA TRANG KVC" -> "FARM LOTTE NHA TRANG") co the la ma GOC
// dung cho kenh nay (VietQR/BIDV7702 -- gross van con ghi "LM NHA TRANG KVC",
// chua doi ten ben do) nhung lai la ma DA DOI TEN can thiet cho kenh khac
// (Momo -- gross ben do dung dung "FARM LOTTE NHA TRANG"). Neu CHI index theo
// effectiveMaDiem (da alias), hoa don se "bien mat" khoi kenh nao con dung ma
// GOC (verified: hoa don 10547/22-07 co maDiem "LM NHA TRANG KVC", trung KHOP
// voi gross cua chinh kenh nay, nhung bi day sang key "FARM LOTTE NHA TRANG"
// nen khong con khop nua). Fix: index hoa don duoi CA HAI key (ma goc VA ma da
// alias, neu khac nhau) -- kenh nao co gross dung ma nao se tu tim thay, an
// toan vi 2 ma nay khong bao gio CUNG co gross trong CUNG 1 kenh/ngay.
function reconcileVietQr(settlements, grossData, invoiceData, gianMapping, manualMatches, diemAlias) {
  const alias = diemAlias || {};
  const invoicesByDiemDay = {};
  for (const inv of invoiceData.invoices) {
    if (!inv.ngayHd || !inv.days || inv.days.length === 0) continue;
    const [invY, invMo, invD] = inv.ngayHd.split("-").map(Number);
    const effectiveMaDiem = alias[inv.maDiem] || inv.maDiem;
    const diemKeysToIndex = effectiveMaDiem === inv.maDiem ? [effectiveMaDiem] : [effectiveMaDiem, inv.maDiem];
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
      for (const diemKey of diemKeysToIndex) {
        const key = `${diemKey}|${iso}`;
        if (!invoicesByDiemDay[key]) invoicesByDiemDay[key] = [];
        invoicesByDiemDay[key].push(inv);
      }
    }
  }

  const allSettlements = settlements.concat(buildPendingDaySettlements(settlements, grossData));
  const results = [];
  for (const s of allSettlements) {
    const day = s.date;
    const gianLines = {};
    for (const rawCode of grossData.codes) {
      const key = `${day}|${rawCode}`;
      const gross = grossData.grossByCode[key];
      if (gross && gross > 0) {
        if (!gianLines[rawCode]) {
          gianLines[rawCode] = { code: rawCode, gross: 0, invoices: new Set(), effectiveCode: rawCode };
        }
        gianLines[rawCode].gross += gross;
        let invs = invoicesByDiemDay[`${rawCode}|${day}`] || [];
        // A gian's "Hinh thuc hop tac" (CSE hay khong) co the doi giua chung
        // (vd: mot diem tu 131 chuyen sang 1388/chia se tu 1 ngay nao do ve
        // sau) trong khi ma resolve ben doanh thu (grossData.codes) van dong
        // cung mot trang thai CSE co dinh (lay tu HD DAU TIEN tung thay).
        // Neu khong tim thay HD nao khop dung rawCode cho ngay nay, thu tim
        // o bien the CSE nguoc lai (co/khong co __FF) CHI cho ngay do -- neu
        // co, dung luon HD do va coi trang thai CSE cua NGAY DO theo dung HD
        // (khong doi lai cac ngay khac). Verified against BIDV7704 "Sân bay
        // Phú Quốc" tu ngay 07/07 tro di (xac nhan tu Luyen: CSE tu ngay 7).
        if (invs.length === 0) {
          const baseCode = rawCode.endsWith(FF_SUFFIX) ? rawCode.slice(0, -FF_SUFFIX.length) : rawCode;
          const altCode = rawCode === baseCode ? baseCode + FF_SUFFIX : baseCode;
          const altInvs = invoicesByDiemDay[`${altCode}|${day}`] || [];
          if (altInvs.length > 0) {
            invs = altInvs;
            gianLines[rawCode].effectiveCode = altCode;
          }
        }
        for (const inv of invs) gianLines[rawCode].invoices.add(inv.soHd);
      }
    }
    const lines = Object.values(gianLines).map((g) => {
      const invoiceList = Array.from(g.invoices);
      const effCode = g.effectiveCode || g.code;
      // Luyen, 2026-07-27: "bên momo hay các gian khác xuất hóa đơn cũng gộp
      // thứ 2 xuất cho thứ 7 CN ... 2 cái lệch này là xuất chung 1 hóa đơn 2
      // ngày T7 CN" -- khac voi Momo/ZVP (settlement CHINH no gop lai thanh 1
      // dong "gia" khi dang cho ngan hang), VietQR moi settlement ngay la THAT
      // (tien ve rieng tung ngay), nen KHONG gop dong lai duoc -- nhung 1 hoa
      // don co the vAn duoc xuat GOP cho nhieu ngay doanh thu (inv.days co >1
      // phan tu). Truoc day moi ngay trong so do deu cong NGUYEN so tien HD
      // vao invoiceTotal cua rieng minh (dem trung toan bo tren tung ngay),
      // gay "Lech" gia tren TAT CA cac ngay du tong hop don da khop dung. Chia
      // ty le so tien HD theo % doanh thu (gross) cua CHINH gian nay trong
      // tung ngay so voi tong doanh thu ca cac ngay hoa don do gop, dam bao
      // tong cong don lai vAn dung bang so tien HD that, khong con dem trung.
      const invoiceTotal = invoiceList.reduce((sum, soHd) => {
        const inv = invoiceData.invoices.find((i) => i.soHd === soHd);
        if (!inv) return sum;
        if (!inv.days || inv.days.length <= 1 || !inv.ngayHd) return sum + inv.tongTt;
        const [invY, invMo, invD] = inv.ngayHd.split("-").map(Number);
        let totalGrossAcrossDays = 0;
        let thisDayGross = 0;
        for (const d of inv.days) {
          let y = invY;
          let mo = invMo;
          if (d > 20 && invD <= 3) {
            mo -= 1;
            if (mo === 0) {
              mo = 12;
              y -= 1;
            }
          }
          const iso = `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
          const gr = grossData.grossByCode[`${iso}|${effCode}`] || 0;
          if (iso === day) thisDayGross = gr;
          totalGrossAcrossDays += gr;
        }
        // Khong co du lieu doanh thu o ngay nao trong so hoa don gop (hiem,
        // vd chua tai du lieu ngay do) -- giu nguyen hanh vi cu (cong nguyen)
        // de khong lam mat canh bao "Lech" that su.
        if (totalGrossAcrossDays <= 0) return sum + inv.tongTt;
        return sum + inv.tongTt * (thisDayGross / totalGrossAcrossDays);
      }, 0);
      const line = {
        code: g.code,
        maCongTrinh: displayCode(effCode),
        // Luyen, 2026-07-17: "doi xuat ra 1388 thanh 131 het" -- khong con
        // fallback ve 1388 cho gian FF/CSE nua.
        tkCo: gianMapping[effCode] || "131",
        gross: g.gross,
        net: g.gross,
        invoiceNumbers: invoiceList,
        invoiceTotal,
        diff: invoiceTotal - g.gross,
        matched: invoiceList.length > 0 && Math.abs(invoiceTotal - g.gross) < 1,
        manualOverride: false,
      };
      // Not gated on invoiceNumbers.length === 0 -- a saved manual match also
      // lets Luyen CORRECT a "Lech" line (invoices matched but total doesn't
      // add up), not just fill in a "Chua co HD" one.
      if (manualMatches) {
        const mm = manualMatches[`${day}|${line.code}`];
        if (mm) {
          line.invoiceNumbers = mm.invoiceNumbers || [];
          // grossAdjustment: cong them vao DOANH THU cua dong nay (khong chi
          // vao so tien HD hien thi) -- dung cho truong hop 1 GD QR "khong co
          // ma cua hang" (raw code "-") ma Luyen da xac dinh la thuoc ve gian
          // nay (vd qua hoa don hom sau). Khac voi chi sua "amount" (chi doi
          // so tien HD hien thi, khong doi doanh thu): neu khong cong vao
          // gross, dong nay van "Khop" (vi matched luon = true) nhung tong
          // "Tinh tu du lieu tai len" ca ngay van thieu dung so tien do, nen
          // phan "Chenh lech" tren dau ngay khong bao gio het du la moi gian
          // rieng le da "Khop". Verified: BIDV77020 ngay 02/07, GD "-" 100k
          // duoc gan vao SB CAM RANH PHN qua HD 1763 -- neu khong cong vao
          // gross, dau ngay van hien "Chenh lech -100.000d" mai mai.
          if (mm.grossAdjustment) {
            line.gross += mm.grossAdjustment;
            line.net = line.gross;
          }
          line.invoiceTotal = mm.amount != null ? mm.amount : line.gross;
          line.diff = line.invoiceTotal - line.gross;
          // Luyen, 2026-07-17: mot ban ghi bu thu cong co the CHI can trung
          // MOT PHAN (vd chuyen bot tien tu 1 gian du sang 1 gian thieu, ma
          // gian du van con du sau khi chuyen) -- trong truong hop do dong
          // nay KHONG duoc coi la "Khop", phai van hien "Lech" (kem so du
          // con lai) de Luyen biet con thieu bao nhieu, thay vi an di sau
          // nhan "Khop (bu thu cong)" gay hieu lam da xu ly xong.
          line.matched = Math.abs(line.diff) < 1;
          line.manualOverride = true;
          line.manualNote = mm.note || "";
        }
      }
      return line;
    });
    const totalGross = lines.filter((l) => l.tkCo !== "SKIP").reduce((sum, l) => sum + l.gross, 0);
    results.push({
      settlementDate: day,
      from: day,
      to: day,
      bankAmount: s.amount,
      totalNetComputed: totalGross,
      diffVsBank: s.pendingBank ? null : totalGross - s.amount,
      pendingBank: !!s.pendingBank,
      lines: lines.sort((a, b) => b.gross - a.gross),
    });
  }

  // Second pass: an invoice that covers SEVERAL days at once (Luyen's normal
  // weekly workflow -- Sat + Sun revenue gets combined into one invoice
  // issued the following Monday, via the "days" list parsed off the
  // "Dich vu thu ho" text) makes EACH individual day look like a false
  // "Lech", because the check above compares that ONE day's gross against
  // the invoice's FULL (multi-day) total. Fix: group every still-unmatched
  // line by (code, exact invoice-number set); if that same invoice set
  // shows up on 2+ DIFFERENT settlement days for the same gian, compare the
  // SUM of gross across just those days against the shared invoice total --
  // if that balances, mark all of them matched instead of leaving each one
  // flagged. Verified against real BIDV7704 / MB11521268 data (SAN BAY PHU
  // QUOC, SC VIVO KVCM, CHKQT CAM RANH multi-day invoices all balance
  // exactly when combined this way).
  const multiDayGroups = new Map();
  results.forEach((r, ri) => {
    r.lines.forEach((l, li) => {
      if (l.manualOverride || l.invoiceNumbers.length === 0 || l.matched) return;
      const key = l.code + "||" + l.invoiceNumbers.slice().sort().join(",");
      if (!multiDayGroups.has(key)) multiDayGroups.set(key, []);
      multiDayGroups.get(key).push({ ri, li });
    });
  });
  multiDayGroups.forEach((refs) => {
    if (refs.length < 2) return; // only relevant when the SAME invoice set spans 2+ different days
    const groupLines = refs.map(({ ri, li }) => results[ri].lines[li]);
    const sumGross = groupLines.reduce((sum, l) => sum + l.gross, 0);
    const invoiceTotal = groupLines[0].invoiceTotal; // identical for all (same invoice set)
    if (Math.abs(sumGross - invoiceTotal) < 1) {
      groupLines.forEach((l) => {
        l.matched = true;
        l.diff = 0;
      });
    }
  });

  return results;
}

module.exports = {
  extractVqrCode,
  extractVietQrSettlements,
  extractVietQrThuTransactions,
  parseVietQrRawWorkbook,
  parseCuaHangSheet,
  parseStoreExportSheet,
  parseTenDiemMaCongTrinhSheet,
  parseInvoiceWorkbookByTag,
  buildGianCandidatesFromInvoices,
  resolveGianGross,
  resolveGianGrossByBankRef,
  reconcileVietQr,
  parseVietQrMnRawWorkbook,
  parseMaCuaHangAppSheet,
  extractStorePrefix,
  buildGianPrefixMatcher,
  resolveGianGrossPrefix,
  toIsoDate,
  isoToDmy,
  normCode,
  normText,
  displayCode,
  FF_SUFFIX,
};
