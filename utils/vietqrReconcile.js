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
function extractVietQrSettlements(transactions) {
  const byDate = {};
  for (const t of transactions) {
    if (t.type !== "thu") continue;
    if (!t.date) continue;
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
    rows.push({ vqrCode, maCuaHang, amount, date, raw: noiDung });
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

// ---------- Invoice list (shared MTT-style sheet, filtered by a bank-specific tag) ----------
// tagPattern: RegExp (or string used as RegExp source) tested against the
// "Dich vu thu ho" cell -- VietQR invoices are tagged per bank, e.g.
// "POSH+JP MB (7020) 4" for BIDV77020, rather than a single fixed keyword
// like Momo's "momo" -- so the caller supplies the pattern per channel.
// Also captures each invoice's own "Ten diem xuat hoa don" (tenDiem),
// needed here (unlike Momo/ZVP) to build the fuzzy-match candidate list
// directly from invoices, since VietQR has no separate "gian hang xuat HD"
// master sheet.
function parseInvoiceWorkbookByTag(buffer, tagPattern) {
  const tagRe = tagPattern instanceof RegExp ? tagPattern : new RegExp(tagPattern, "i");
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const kdsCandidates = wbLite.SheetNames.filter((n) => /k.\s*ds\s*xu.t/i.test(n));
  const sheetName =
    kdsCandidates.find((n) => /989/.test(n)) ||
    kdsCandidates[0] ||
    wbLite.SheetNames.find((n) => {
      const t = normText(n);
      return t.includes("hoa don") || (t.includes("danh sach") && t.includes("don")) || t.includes("so hoa don");
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
function resolveGianGross(rawRows, storeNameMap, gianCandidates) {
  const matcher = buildOnlineProductMatcher(gianCandidates);
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

// diemAlias: same shared table as Momo/ZVP (store.invoice_diem_alias).
function reconcileVietQr(settlements, grossData, invoiceData, gianMapping, manualMatches, diemAlias) {
  const alias = diemAlias || {};
  const invoicesByDiemDay = {};
  for (const inv of invoiceData.invoices) {
    if (!inv.ngayHd || !inv.days || inv.days.length === 0) continue;
    const [invY, invMo, invD] = inv.ngayHd.split("-").map(Number);
    const effectiveMaDiem = alias[inv.maDiem] || inv.maDiem;
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

  const results = [];
  for (const s of settlements) {
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
      const invoiceTotal = invoiceList.reduce((sum, soHd) => {
        const inv = invoiceData.invoices.find((i) => i.soHd === soHd);
        return sum + (inv ? inv.tongTt : 0);
      }, 0);
      const effCode = g.effectiveCode || g.code;
      const line = {
        code: g.code,
        maCongTrinh: displayCode(effCode),
        tkCo: gianMapping[effCode] || (effCode.endsWith(FF_SUFFIX) ? "1388" : "131"),
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
          line.matched = true;
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
      diffVsBank: totalGross - s.amount,
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
  parseVietQrRawWorkbook,
  parseCuaHangSheet,
  parseStoreExportSheet,
  parseInvoiceWorkbookByTag,
  buildGianCandidatesFromInvoices,
  resolveGianGross,
  reconcileVietQr,
  toIsoDate,
  isoToDmy,
  normCode,
  normText,
  displayCode,
  FF_SUFFIX,
};
