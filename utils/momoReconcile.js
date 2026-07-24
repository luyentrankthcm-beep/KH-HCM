const XLSX = require("xlsx");

// ---------- Helpers ----------
//
// IMPORTANT: we never use XLSX's `cellDates: true` option. On some server
// timezones (e.g. UTC+7), converting an Excel date serial to a JS Date and
// then reading it back with toISOString() silently shifts the date by one
// day (Excel serials sometimes carry a tiny fractional/rounding artifact,
// and toISOString() is UTC-based while the Date XLSX builds is anchored to
// local midnight). To avoid that whole class of off-by-one bugs we always
// read raw values (numbers) and convert Excel serials to ISO date strings
// ourselves using plain integer arithmetic (no timezone involved at all).

function excelSerialToIso(serial) {
  // Momo's own exports use whole-day integer serials, so Math.floor and
  // Math.round give identical results there -- but Payoo's raw transaction
  // export (utils/zvpReconcile.js reuses this same helper) carries a real
  // time-of-day fraction per row (e.g. 46215.9 = ~21:30). Rounding that
  // pushes every afternoon/evening transaction into the NEXT calendar day,
  // which silently moved ~18tr/settlement of revenue onto the wrong date
  // and made the Payoo bank-vs-data reconciliation mismatch. Floor (=
  // truncate the time-of-day) always yields the correct calendar day.
  const days = Math.floor(serial);
  const utcMillis = (days - 25569) * 86400 * 1000; // 25569 = days between 1899-12-30 and 1970-01-01
  return new Date(utcMillis).toISOString().slice(0, 10);
}

function toIsoDate(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) {
    // Only reached if some other code path handed us a Date directly.
    return v.toISOString().slice(0, 10);
  }
  if (typeof v === "number") {
    return excelSerialToIso(v);
  }
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
  // "DD-MM-YYYY[ HH:MM:SS]" -- format used by the MoMo portal's "Transaction
  // report" export (cot "Thời gian", vd "21-07-2026 21:25:17"). Checked AFTER
  // the "YYYY-MM-DD" pattern above so a real ISO string is never misread.
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return null;
}

// "YYYY-MM-DD" -> "DD/MM/YYYY" (plain string formatting, no Date object
// involved at all -> zero risk of the timezone bug described above).
function isoToDmy(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function normCode(v) {
  if (v === null || v === undefined) return "";
  return String(v).toUpperCase().replace(/\s+/g, " ").trim();
}

// Strip Vietnamese diacritics for fuzzy, accent-insensitive matching of
// sheet names / headers (real-world exports rename things constantly --
// e.g. "kê ds xuất HĐ MTT - 989" one time, "DANH SACH HOA DON" another --
// diacritic + case + whitespace insensitive matching is far more robust
// than trying to hand-craft a regex with "." wildcards for every accented
// character).
function removeDiacritics(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}

function normText(s) {
  return removeDiacritics(s).toLowerCase().replace(/\s+/g, " ").trim();
}

function addDays(isoDate, n) {
  const d = new Date(isoDate + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function dateRange(fromIso, toIso) {
  const out = [];
  let cur = fromIso;
  let guard = 0;
  while (cur <= toIso && guard < 400) {
    out.push(cur);
    cur = addDays(cur, 1);
    guard++;
  }
  return out;
}

// A number that plausibly represents an Excel date serial (years ~2015-2045).
function looksLikeDateSerial(n) {
  return typeof n === "number" && n > 42000 && n < 53000;
}

// SC VIVO KVCM is the one Ma Cong trinh that Luyen splits into 2 separate
// TK Co lines: the "FF" / "FUNFEST" gian (doanh thu chia se -> TK Co 1388)
// vs the "ADV" gian (binh thuong -> TK Co 131). Both still report under the
// SAME Ma cong trinh "SC VIVO KVCM" on the MISA export -- internally we tag
// the FF gian with a "__FF" suffix so it flows through the rest of the
// pipeline (gross aggregation, invoice matching, TK Co mapping) as if it
// were its own code, then strip the suffix again at export/display time.
const SPLIT_PARENT_CODE = "SC VIVO KVCM";
const FF_SUFFIX = "__FF";

function isFfVivoGian(text) {
  const n = normCode(text);
  return n.includes("FF") || n.includes("FUNFEST");
}

function displayCode(code) {
  return String(code || "").replace(FF_SUFFIX, "");
}

function isFfCode(code) {
  return String(code || "").endsWith(FF_SUFFIX);
}

function prettyCodeLabel(code) {
  return isFfCode(code)
    ? `${displayCode(code)} (gian FF - chia se)`
    : String(code || "");
}

// ---------- Parse "Tong Momo Tx [Gop]" workbook(s) ----------
//
// Two formats are supported:
//
// 1) NEW "... Gop" format (per-Gian granularity): header row has text columns
//    "Ma Cong trinh", "Gian", "Ten diem xuat hoa don", "ly do noi bo" followed
//    by one date column per day. Several Ma Cong trinh (e.g. AM TP KVCM,
//    AM BD KVCM) legitimately have more than one Gian row rolling up under
//    them -- we sum those together. The one exception is SC VIVO KVCM, whose
//    2 Gian rows ("FZ FF CENTRAL" / "FUNFEST SCVIVO" vs "FUNZONE ADVENTURE SC
//    VIVO QUAN 7" / "ADV SCVIVO") get different TK Co and so are kept apart
//    (see isFfVivoGian).
//
// 2) OLD flat format (kept for backward compatibility with sheets uploaded
//    before the "Gop" template existed): header row has (Ma KH, Gian, Ma cua
//    hang, Ma Cong trinh, day1, day2, ...) with one row per Ma Cong trinh
//    already pre-summed (no separate Gian breakdown).
//
// A workbook may contain several matching sheets (e.g. both "Tong Momo T6
// Gop" and "Tong Momo T7 Gop") -- all are parsed and merged together.
function parseTongMomoWorkbook(buffer) {
  // The real workbooks Luyen uses (e.g. DULIEUMOMO.xlsm) can be 40+ MB with
  // 15 sheets, most of them irrelevant to this parse (raw momo master data,
  // Sale Online, hidden helper sheets, ...). Fully parsing every sheet with
  // XLSX.read() is slow and can time out on a large file. So we first do a
  // very cheap "sheet names only" read (bookSheets:true skips all cell
  // parsing) to decide which sheet(s) we actually need, then do a second
  // read that parses ONLY those sheets.
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });

  const gopSheets = wbLite.SheetNames.filter(
    (n) => /momo/i.test(n) && /(g.p|gop)/i.test(n)
  );

  if (gopSheets.length > 0) {
    const wb = XLSX.read(buffer, { type: "buffer", sheets: gopSheets });
    return mergeParsedGross(
      gopSheets.map((sheetName) => parseTongMomoGopSheet(wb.Sheets[sheetName], sheetName))
    );
  }

  // Fallback: old flat format.
  const sheetName = wbLite.SheetNames.find((n) => /t.?ng\s*momo/i.test(n) || /tong momo/i.test(n));
  if (!sheetName) {
    throw new Error(
      'Khong tim thay sheet "Tong Momo..." trong file. Vui long kiem tra lai file tai len.'
    );
  }
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  return parseTongMomoFlatSheet(wb.Sheets[sheetName], sheetName);
}

function mergeParsedGross(parsedList) {
  const dateSet = new Set();
  const codesSeen = new Set();
  const grossByCode = {};
  const sheetNames = [];
  const cuaHangMap = {};
  for (const p of parsedList) {
    sheetNames.push(p.sheetName);
    p.dates.forEach((d) => dateSet.add(d));
    p.codes.forEach((c) => codesSeen.add(c));
    for (const [k, v] of Object.entries(p.grossByCode)) {
      grossByCode[k] = (grossByCode[k] || 0) + v;
    }
    Object.assign(cuaHangMap, p.cuaHangMap || {});
  }
  return {
    sheetName: sheetNames.join(" + "),
    dates: Array.from(dateSet).sort(),
    codes: Array.from(codesSeen),
    grossByCode,
    cuaHangMap,
  };
}

// Detect the header row + date columns of a "Gop" or flat sheet: a run of
// >=5 consecutive numeric, date-serial-looking cells increasing by 1 each
// column (one column per day).
function findDateHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 15); r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (!looksLikeDateSerial(row[c])) continue;
      let count = 1;
      for (let k = c + 1; k < Math.min(c + 12, row.length); k++) {
        if (typeof row[k] === "number" && row[k] === row[k - 1] + 1) count++;
        else break;
      }
      if (count >= 5) return { headerRowIdx: r, dateStartCol: c };
    }
  }
  return null;
}

function parseTongMomoGopSheet(ws, sheetName) {
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const found = findDateHeaderRow(grid);
  if (!found) {
    throw new Error(`Khong doc duoc dong tieu de ngay trong sheet "${sheetName}".`);
  }
  const { headerRowIdx, dateStartCol } = found;
  const headerRow = grid[headerRowIdx];
  const dateCols = [];
  for (let c = dateStartCol; c < headerRow.length; c++) {
    if (looksLikeDateSerial(headerRow[c])) {
      dateCols.push({ col: c, date: excelSerialToIso(headerRow[c]) });
    } else break;
  }

  // Locate "Ma Cong trinh" / "Gian" / "Ten diem xuat hoa don" / "Ma cua hang"
  // text columns among the columns to the left of the date columns.
  let codeCol = -1;
  let gianCol = -1;
  let tenDiemCol = -1;
  let maCuaHangCol = -1;
  for (let c = 0; c < dateStartCol; c++) {
    const v = headerRow[c];
    if (!v || typeof v !== "string") continue;
    const h = normText(v);
    if (codeCol === -1 && h.includes("cong trinh")) codeCol = c;
    if (gianCol === -1 && h.startsWith("gian")) gianCol = c;
    if (tenDiemCol === -1 && h.includes("ten diem")) tenDiemCol = c;
    if (maCuaHangCol === -1 && h.includes("ma cua hang")) maCuaHangCol = c;
  }
  if (codeCol === -1) codeCol = 0; // observed layout fallback

  const grossByCode = {};
  const codesSeen = new Set();
  const cuaHangMap = {}; // "MA CUA HANG" -> { maCongTrinh, code, gian }

  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const codeVal = row[codeCol];
    if (codeVal === null || codeVal === undefined || String(codeVal).trim() === "") break; // end of block
    const maCongTrinh = normCode(codeVal);
    if (!maCongTrinh) continue;

    let code = maCongTrinh;
    const gianText = (gianCol !== -1 ? row[gianCol] : "") || "";
    const tenDiemText = (tenDiemCol !== -1 ? row[tenDiemCol] : "") || "";
    if (maCongTrinh === SPLIT_PARENT_CODE) {
      if (isFfVivoGian(gianText) || isFfVivoGian(tenDiemText)) {
        code = maCongTrinh + FF_SUFFIX;
      }
    }

    // Remember Ma cua hang -> gian/code, so future raw MoMo portal exports
    // (which only carry Ma cua hang, not the friendly Ma Cong trinh) can be
    // mapped back automatically.
    const maCuaHangVal = maCuaHangCol !== -1 ? row[maCuaHangCol] : null;
    if (maCuaHangVal !== null && maCuaHangVal !== undefined && String(maCuaHangVal).trim() !== "") {
      const key = normCode(maCuaHangVal);
      cuaHangMap[key] = { maCongTrinh, code, gian: String(gianText || tenDiemText || "").trim() };
    }

    codesSeen.add(code);
    for (const { col, date } of dateCols) {
      const val = row[col];
      const num = typeof val === "number" ? val : parseFloat(val) || 0;
      if (!num) continue;
      const key = `${date}|${code}`;
      grossByCode[key] = (grossByCode[key] || 0) + num;
    }
  }

  return {
    sheetName,
    dates: dateCols.map((d) => d.date),
    codes: Array.from(codesSeen),
    grossByCode,
    cuaHangMap,
  };
}

// Old flat format: (Ma KH, Gian, Ma cua hang, Ma Cong trinh, day1, day2, ...)
// one pre-summed row per Ma Cong trinh, only "KH Cu" rows kept.
function parseTongMomoFlatSheet(ws, sheetName) {
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const found = findDateHeaderRow(grid);
  if (!found) {
    throw new Error('Khong doc duoc dong tieu de ngay trong sheet "Tong Momo...".');
  }
  const { headerRowIdx, dateStartCol } = found;
  const headerRow = grid[headerRowIdx];
  const dateCols = [];
  for (let c = dateStartCol; c < headerRow.length; c++) {
    if (looksLikeDateSerial(headerRow[c])) {
      dateCols.push({ col: c, date: excelSerialToIso(headerRow[c]) });
    } else break;
  }

  let codeCol = dateStartCol - 1;
  let khCol = 0;
  for (let c = 0; c < dateStartCol; c++) {
    const v = headerRow[c];
    if (v && /c.ng\s*tr.nh|cong trinh/i.test(String(v))) codeCol = c;
    if (v && /m.\s*kh|ma kh/i.test(String(v))) khCol = c;
  }

  const grossByCode = {};
  const codesSeen = new Set();

  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const khVal = row[khCol];
    const codeVal = row[codeCol];
    if ((khVal === null || khVal === undefined) && (codeVal === null || codeVal === undefined)) {
      const nextRow = grid[r + 1] || [];
      const nextKh = nextRow[khCol];
      if (codesSeen.size > 0 && (nextKh === null || nextKh === undefined)) break;
      continue;
    }
    const khStr = khVal ? String(khVal).toLowerCase() : "";
    if (khStr.includes("mới") || khStr.includes("moi")) break;
    if (!khStr.includes("cũ") && !khStr.includes("cu")) continue;
    const code = normCode(codeVal);
    if (!code) continue;
    codesSeen.add(code);
    for (const { col, date } of dateCols) {
      const val = row[col];
      const num = typeof val === "number" ? val : parseFloat(val) || 0;
      const key = `${date}|${code}`;
      grossByCode[key] = (grossByCode[key] || 0) + num;
    }
  }

  return {
    sheetName,
    dates: dateCols.map((d) => d.date),
    codes: Array.from(codesSeen),
    grossByCode,
    cuaHangMap: {},
  };
}

// ---------- Parse the RAW MoMo merchant-portal export ----------
// This is the "daily_report" zip Luyen downloads straight from the MoMo
// business portal (unrelated to the hand-maintained "Tong Momo Gop" sheet):
// an xlsx with sheets like "data" / "bnpl" / "account_to_account", each a
// per-transaction detail report with columns "MS.Total Amount", "MS.Ngay
// hoan thanh", "MS.Ma cua hang", "MS.Trang Thai GD". We sum every sheet that
// has this column layout (they're different payment rails -- QR billpay,
// buy-now-pay-later, bank-transfer-QR -- all additive momo revenue), only
// counting "Thanh cong" (successful) rows, and dedupe by MS.TransID in case
// a transaction is ever listed on more than one sheet.
function parseRawMomoPortalWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const seenTransIds = new Set();
  const transactions = []; // {transId, date, maCuaHang, amount}
  let matchedAnySheet = false;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    let headerRowIdx = -1;
    let cols = {};
    for (let r = 0; r < Math.min(grid.length, 12); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (s.includes("total amount")) idx.amount = c;
        if (s.includes("ngay hoan thanh")) idx.date = c;
        if (s.includes("ma cua hang")) idx.maCuaHang = c;
        if (s.includes("trang thai")) idx.trangThai = c;
        if (s.includes("transid") && !s.includes("parent")) idx.transId = c;
      });
      if (idx.amount !== undefined && idx.date !== undefined && idx.maCuaHang !== undefined) {
        headerRowIdx = r;
        cols = idx;
        break;
      }
    }
    if (headerRowIdx < 0) continue; // not a matching sheet, skip silently

    matchedAnySheet = true;
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const trangThai = cols.trangThai !== undefined ? row[cols.trangThai] : null;
      if (trangThai !== null && !/th.nh c.ng|thanh cong/i.test(String(trangThai))) continue;
      const amount = cols.amount !== undefined ? Number(row[cols.amount]) || 0 : 0;
      if (!amount) continue;
      const maCuaHang = cols.maCuaHang !== undefined ? normCode(row[cols.maCuaHang]) : "";
      if (!maCuaHang) continue;
      const dateRaw = cols.date !== undefined ? row[cols.date] : null;
      const date = toIsoDate(dateRaw ? String(dateRaw).split(" ")[0] : null);
      if (!date) continue;
      const transId = cols.transId !== undefined ? String(row[cols.transId] || "") : `${sheetName}-${r}`;
      if (seenTransIds.has(transId)) continue;
      seenTransIds.add(transId);
      transactions.push({ transId, date, maCuaHang, amount });
    }
  }

  if (!matchedAnySheet) {
    throw new Error(
      'Khong nhan dien duoc file bao cao MoMo: can co cot "MS.Total Amount", "MS.Ngay hoan thanh", "MS.Ma cua hang".'
    );
  }

  return { transactions };
}

// ---------- Parse the NEWER "Transaction report" MoMo portal export ----------
// Chi Nhan, 2026-07-22: "tôi mới tải file bạn xem nó trả qua ví trả sao hay
// gì đó lấy ra tiền sao phí cho tôi được không" -- khac voi zip "daily_report"
// o tren (cot "MS.*"), day la 1 file .xlsx (khong nen zip) tai truc tiep tu
// cong MoMo voi ten cot KHONG co tien to "MS.": "Thời gian", "Trạng thái",
// "Số tiền", "Phương thức thanh toán", "Nguồn tiền", "Mã cửa hàng"... Diem
// quan trong: file nay co du thong tin de tinh PHI THAT theo tung giao dich
// (KH Cu dung phi co dinh 1,1% nen khong can, nhung KH Moi dung 3 muc phi
// khac nhau tuy theo NGUON TIEN khach hang dung de thanh toan -- xac nhan
// bang du lieu that trong file "BÁO CÁO DOANH THU momo.xlsm" (sheet "Dữ liệu
// Momo KH", cot AJ/AK/AP): cot "Nguồn tiền" = "Ví MoMo" -> phi 1% (
// 1,1% da gom VAT); = "Ví trả sau" -> phi 1,2% (1,32% da gom VAT); moi truong
// hop khac (vd "Ngân hàng hoặc Ví khác", "Thẻ Quốc Tế") -> phi 0,3% (0,33% da
// gom VAT) -- day la muc "con lai" trong ghi chu "Phí MoMo KH mới: Ví
// MoMo*1%, Ví trả sau*1,2%, còn lại*0,3% (giá chưa thuế)" cua Luyen.
const KH_MOI_FEE_RATE_BY_NGUON_TIEN = {
  "vi momo": 0.011, // 1% + 10% VAT
  "vi tra sau": 0.0132, // 1.2% + 10% VAT
};
const KH_MOI_FEE_RATE_DEFAULT = 0.0033; // "con lai" -- 0.3% + 10% VAT

function feeRateForNguonTien(nguonTien) {
  const n = normText(nguonTien || "");
  return KH_MOI_FEE_RATE_BY_NGUON_TIEN[n] !== undefined ? KH_MOI_FEE_RATE_BY_NGUON_TIEN[n] : KH_MOI_FEE_RATE_DEFAULT;
}

function parseKhMoiFeeTransactionWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const seenTransIds = new Set();
  const transactions = []; // {transId, date, maCuaHang, amount, fee}
  let matchedAnySheet = false;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    let headerRowIdx = -1;
    let cols = {};
    for (let r = 0; r < Math.min(grid.length, 12); r++) {
      const row = grid[r] || [];
      const idx = {};
      row.forEach((v, c) => {
        if (!v || typeof v !== "string") return;
        const s = normText(v);
        if (s === "so tien" || s.includes("so tien") && !s.includes("giam gia")) idx.amount = c;
        if (s.includes("thoi gian")) idx.date = c;
        if (s.includes("ma cua hang")) idx.maCuaHang = c;
        if (s.includes("trang thai")) idx.trangThai = c;
        if (s.includes("nguon tien")) idx.nguonTien = c;
        if (s.includes("ma giao dich")) idx.transId = c;
      });
      if (idx.amount !== undefined && idx.date !== undefined && idx.maCuaHang !== undefined && idx.nguonTien !== undefined) {
        headerRowIdx = r;
        cols = idx;
        break;
      }
    }
    if (headerRowIdx < 0) continue; // not a matching sheet, skip silently

    matchedAnySheet = true;
    for (let r = headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const trangThai = cols.trangThai !== undefined ? row[cols.trangThai] : null;
      if (trangThai !== null && !/th.nh c.ng|thanh cong/i.test(String(trangThai))) continue;
      const amount = cols.amount !== undefined ? Number(row[cols.amount]) || 0 : 0;
      if (!amount) continue;
      const maCuaHang = cols.maCuaHang !== undefined ? normCode(row[cols.maCuaHang]) : "";
      if (!maCuaHang) continue;
      const dateRaw = cols.date !== undefined ? row[cols.date] : null;
      const date = toIsoDate(dateRaw);
      if (!date) continue;
      const nguonTien = cols.nguonTien !== undefined ? row[cols.nguonTien] : "";
      const fee = Math.round(amount * feeRateForNguonTien(nguonTien));
      const transId = cols.transId !== undefined ? String(row[cols.transId] || "") : `${sheetName}-${r}`;
      if (seenTransIds.has(transId)) continue;
      seenTransIds.add(transId);
      transactions.push({ transId, date, maCuaHang, amount, fee });
    }
  }

  if (!matchedAnySheet) {
    throw new Error(
      'Khong nhan dien duoc file "Transaction report": can co cot "Thời gian", "Số tiền", "Mã cửa hàng", "Nguồn tiền".'
    );
  }

  return { transactions };
}

// Giong resolveRawPortalGross nhung tinh THEM feeByCode/netByCode (net =
// gross - phi) tu cac giao dich da co san phi (xem parseKhMoiFeeTransactionWorkbook).
function resolveKhMoiFeeTransactionGross(transactions, cuaHangMapping) {
  const dateSet = new Set();
  const codesSeen = new Set();
  const grossByCode = {};
  const netByCode = {};
  const unmapped = new Set();

  for (const t of transactions) {
    const known = cuaHangMapping[t.maCuaHang];
    const code = known ? known.code : `CHUA MAP: ${t.maCuaHang}`;
    if (!known) unmapped.add(t.maCuaHang);
    dateSet.add(t.date);
    codesSeen.add(code);
    const key = `${t.date}|${code}`;
    grossByCode[key] = (grossByCode[key] || 0) + t.amount;
    netByCode[key] = (netByCode[key] || 0) + (t.amount - t.fee);
  }

  return {
    dates: Array.from(dateSet).sort(),
    codes: Array.from(codesSeen),
    grossByCode,
    netByCode,
    unmapped: Array.from(unmapped),
  };
}

// Resolve a raw portal export's per-transaction (date, Ma cua hang, amount)
// rows into the same {dates, codes, grossByCode} shape parseTongMomoWorkbook
// produces, using a Ma cua hang -> {code} mapping table (built up over time
// from parseTongMomoGopSheet's cuaHangMap -- see store.cua_hang_mapping).
// Any Ma cua hang with no known mapping is kept under its own clearly-marked
// "CHUA MAP: <ma cua hang>" code rather than silently dropped or guessed, so
// Luyen can see it and map it once.
function resolveRawPortalGross(transactions, cuaHangMapping) {
  const dateSet = new Set();
  const codesSeen = new Set();
  const grossByCode = {};
  const unmapped = new Set();

  for (const t of transactions) {
    const known = cuaHangMapping[t.maCuaHang];
    const code = known ? known.code : `CHUA MAP: ${t.maCuaHang}`;
    if (!known) unmapped.add(t.maCuaHang);
    dateSet.add(t.date);
    codesSeen.add(code);
    const key = `${t.date}|${code}`;
    grossByCode[key] = (grossByCode[key] || 0) + t.amount;
  }

  return {
    dates: Array.from(dateSet).sort(),
    codes: Array.from(codesSeen),
    grossByCode,
    unmapped: Array.from(unmapped),
  };
}

// ---------- Parse invoice list workbook ("ke ds xuat HD MTT" style) ----------
// Looks for a sheet name containing "xuat HD" / "hoa don" / "danh sach ...
// don" (accent/case insensitive), then a header row containing "So HD" and
// "Ma diem", and extracts momo-tagged rows (column whose header mentions
// "thu ho" containing "momo"). Also reads "Hinh thuc hop tac" + "Ten diem
// xuat hoa don" to tell apart the 2 SC VIVO KVCM gian (see SPLIT_PARENT_CODE
// above).
function parseInvoiceWorkbook(buffer, companyKey) {
  // See the comment in parseTongMomoWorkbook: cheap "names only" pass first,
  // then a targeted read of just the one sheet we need -- large real-world
  // workbooks (40+ MB, 15 sheets) parse far faster this way.
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const kdsCandidates = wbLite.SheetNames.filter((n) => /k.\s*ds\s*xu.t/i.test(n));
  // Our bank data belongs to phap nhan "K&H Cu" (ma so thue ...4989) = code "989"
  // in these workbooks. KH Moi (CONG TY TNHH GIAI TRI K&H, phap nhan khac) =
  // code "705" -- Luyen, 2026-07-17: "KH moi la sheet 'ke ds xuat HD MTT -
  // 705' con KH cu la sheet '... - 989'". If several "ke ds xuat HD" sheets
  // exist (one per phap nhan), prefer the one tagged to match the company
  // whose page this upload happened on; fall back to "989" (the old default)
  // if that tag isn't found, so older files with only 1 sheet still work.
  const preferredTag = companyKey === "kh_moi" ? "705" : "989";
  const sheetName =
    kdsCandidates.find((n) => new RegExp(preferredTag).test(n)) ||
    kdsCandidates.find((n) => /989/.test(n)) ||
    kdsCandidates[0] ||
    wbLite.SheetNames.find((n) => {
      const t = normText(n);
      return t.includes("hoa don") || (t.includes("danh sach") && t.includes("don"));
    });
  if (!sheetName) {
    throw new Error('Khong tim thay sheet danh sach hoa don trong file.');
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
      if (s.includes("so hd")) idx.soHd = c;
      if (s.includes("ngay hd")) idx.ngayHd = c;
      if (s.includes("thang")) idx.thang = c;
      if (s.includes("ma diem")) idx.maDiem = c;
      if (s.includes("tong tt hd") || (s.includes("tong") && s.includes("hd"))) idx.tongTt = c;
      if (s.includes("thu ho")) idx.dichVuThuHo = c;
      if (s.includes("hinh thuc hop tac")) idx.hinhThuc = c;
      if (s.includes("ten diem xuat hoa don")) idx.tenDiem = c;
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
    if (!dvth || !/momo/i.test(String(dvth))) continue;
    // Day list after "momo" is usually comma-separated ("MOMO 11, 12") but
    // some rows use "+" instead ("momo 11+12"); accept both separators.
    const m = String(dvth).trim().match(/momo\s+([\d,+\s]+)/i);
    const days = m
      ? m[1]
          .split(/[,+]/)
          .map((x) => parseInt(x.trim(), 10))
          .filter((x) => !isNaN(x))
      : [];
    const ngayHdRaw = cols.ngayHd !== undefined ? row[cols.ngayHd] : null;

    let maDiem = normCode(cols.maDiem !== undefined ? row[cols.maDiem] : null);
    if (maDiem === SPLIT_PARENT_CODE) {
      const hinhThuc = cols.hinhThuc !== undefined ? row[cols.hinhThuc] : "";
      const tenDiem = cols.tenDiem !== undefined ? row[cols.tenDiem] : "";
      const hinhThucStr = String(hinhThuc || "").toUpperCase();
      if (hinhThucStr.includes("CSE") || isFfVivoGian(tenDiem)) {
        maDiem = SPLIT_PARENT_CODE + FF_SUFFIX;
      }
    }

    invoices.push({
      soHd: cols.soHd !== undefined ? row[cols.soHd] : null,
      ngayHd: toIsoDate(ngayHdRaw),
      thang: cols.thang !== undefined ? row[cols.thang] : null,
      maDiem,
      tongTt: cols.tongTt !== undefined ? Number(row[cols.tongTt]) || 0 : 0,
      days,
      raw: String(dvth).trim(),
    });
  }

  return { sheetName, invoices };
}

// ---------- Extract momo settlements from already-imported bank transactions ----------
// A momo settlement is a "thu" transaction whose description contains
// "DI DONG TRUC TUYEN" and a "tu DD/MM/YYYY den DD/MM/YYYY" date range.
function extractMomoSettlements(transactions) {
  const pattern = /tu (\d{2})\/(\d{2})\/(\d{4}) den (\d{2})\/(\d{2})\/(\d{4})/i;
  const out = [];
  for (const t of transactions) {
    if (t.type !== "thu") continue;
    if (!/DI DONG TRUC TUYEN/i.test(t.description || "")) continue;
    const m = pattern.exec(t.description);
    if (!m) continue;
    const [, d1, mo1, y1, d2, mo2, y2] = m;
    out.push({
      id: t.id,
      date: t.date,
      amount: t.amount,
      fromIso: `${y1}-${mo1}-${d1}`,
      toIso: `${y2}-${mo2}-${d2}`,
    });
  }
  return out;
}

// ---------- Core reconciliation ----------
// settlements: [{date, amount, fromIso, toIso}] extracted from bank transactions
// grossData: output of parseTongMomoWorkbook (possibly merged across months)
// invoiceData: output of parseInvoiceWorkbook (possibly merged across months)
// gianMapping: { CODE: '131' | '1388' | 'SKIP' } -- CODE may carry the internal "__FF" suffix
// diemAlias: { "TEN/MA DIEM TREN HOA DON": "MA CONG TRINH CHINH" } -- optional,
// see store.js's invoice_diem_alias comment. Applied here (at reconcile time,
// not parse time) so it also fixes invoices already uploaded under an
// alternate name, without needing to re-upload the invoice file.
function reconcileMomo(settlements, grossData, invoiceData, gianMapping, diemAlias) {
  const alias = diemAlias || {};

  // Resolve every invoice's day-of-month list into full ISO calendar dates
  // once (handles the "day near start of month actually belongs to end of
  // PREVIOUS month" case), tagged with its effective Ma Cong trinh (after
  // alias) -- reused below both to detect duplicate combined invoices and to
  // build invoicesByDiemDay.
  const resolvedInvoices = [];
  for (const inv of invoiceData.invoices) {
    if (!inv.ngayHd || !inv.days || inv.days.length === 0) continue;
    // The invoice date tells us the month/year; revenue day(s) are inv.days,
    // which normally fall in the same month as ngayHd, but a day like 30/31
    // on an invoice dated the 1st-2nd of the next month means the revenue
    // was from the LAST day of the previous month.
    const [invY, invMo, invD] = inv.ngayHd.split("-").map(Number);
    const effectiveMaDiem = alias[inv.maDiem] || inv.maDiem;
    const isos = inv.days.map((day) => {
      let y = invY;
      let mo = invMo;
      if (day > 20 && invD <= 3) {
        // likely previous month's tail end
        mo -= 1;
        if (mo === 0) {
          mo = 12;
          y -= 1;
        }
      }
      return `${y}-${String(mo).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    });
    resolvedInvoices.push({ inv, effectiveMaDiem, isos });
  }

  // Luyen, 2026-07-22: "do bị trùng á bỏ qua cho tôi nhá" -- chi xac nhan qua
  // truong hop thuc te (thang 7/2026, gian DIY ESTELLA/FARM LOTTE PHAN
  // THIET/NHA TRANG/AE TAN AN): moi gian da co san 1 hoa don RIENG cho ngay
  // 18 (vd so HD 9977-9981), roi NCC lai xuat THEM 1 hoa don GOP ghi ca 2
  // ngay ("momo 18,19" -> days=[18,19], vd so HD 10008-10012) -- day la hoa
  // don xuat TRUNG/XUAT LAI (trung lap voi ngay 18 da co hoa don rieng roi),
  // KHONG phai doanh thu moi -- cong ca 2 hoa don vao lam doi soat bi "Lech"
  // dung bang chinh so tien hoa don gop do.
  //
  // Quy tac chung: 1 hoa don gop NHIEU ngay (>1 ngay) ma CO IT NHAT 1 trong
  // cac ngay do (theo ngay lich ISO day du, khong chi so ngay-trong-thang --
  // tranh nham giua cac thang khac nhau) DA CO SAN 1 hoa don RIENG (chi 1
  // ngay) khac cho CUNG gian, thi coi hoa don gop nay la TRUNG LAP -- loai
  // HOAN TOAN khoi doi soat (khong cong vao bat ky ngay nao), thay vi cong
  // them lam du doanh thu.
  const singleDayIsoByCode = {}; // "CODE|ISO" -> true, tu cac hoa don CHI 1 ngay
  for (const { inv, effectiveMaDiem, isos } of resolvedInvoices) {
    if (inv.days.length === 1) {
      singleDayIsoByCode[`${effectiveMaDiem}|${isos[0]}`] = true;
    }
  }

  const invoicesByDiemDay = {}; // "CODE|YYYY-MM-DD" -> [invoice,...]
  for (const { inv, effectiveMaDiem, isos } of resolvedInvoices) {
    const isDuplicateCombinedInvoice =
      inv.days.length > 1 &&
      isos.some((iso) => singleDayIsoByCode[`${effectiveMaDiem}|${iso}`]);
    if (isDuplicateCombinedInvoice) continue; // hoa don gop trung -- bo qua hoan toan

    for (const iso of isos) {
      const key = `${effectiveMaDiem}|${iso}`;
      if (!invoicesByDiemDay[key]) invoicesByDiemDay[key] = [];
      invoicesByDiemDay[key].push(inv);
    }
  }

  const allSettlements = settlements.concat(buildPendingDaySettlements(settlements, grossData));
  const results = [];
  for (const s of allSettlements) {
    const days = dateRange(s.fromIso, s.toIso);
    const gianLines = {};
    for (const day of days) {
      for (const code of grossData.codes) {
        const key = `${day}|${code}`;
        const val = grossData.grossByCode[key];
        if (val && val > 0) {
          if (!gianLines[code]) {
            gianLines[code] = { code, gross: 0, netOverride: 0, hasNetOverride: false, invoices: new Set(), days: new Set() };
          }
          gianLines[code].gross += val;
          gianLines[code].days.add(day);
          // Luyen, 2026-07-21: "Phi MoMo KH moi: Vi MoMo*1%, Vi tra sau*1,2%,
          // con lai*0,3%" -- khac han muc phi co dinh 1,1% (he so 0.989) dang
          // dung cho KH Cu. Thay vi doan mot he so chung cho ca 2 cong ty,
          // file "Tong Momo KH Moi" chi da tu tinh san so tien NET thuc te ve
          // ngan hang moi ma/ngay (xem grossData.netByCode, doc tu block "momo
          // tra ve ngan hang" cua file) -- neu co, dung THANG so nay lam net,
          // KHONG nhan lai voi 0.989 (se sai/lech kep). Khong co (nhu du lieu
          // KH Cu cu) thi giu nguyen cach tinh cu.
          const netVal = grossData.netByCode && grossData.netByCode[key];
          if (netVal !== undefined && netVal !== null) {
            gianLines[code].netOverride += netVal;
            gianLines[code].hasNetOverride = true;
          }
          const invKey = `${code}|${day}`;
          const invs = invoicesByDiemDay[invKey] || [];
          for (const inv of invs) gianLines[code].invoices.add(inv.soHd);
        }
      }
    }
    const lines = Object.values(gianLines).map((g) => {
      const net = g.hasNetOverride ? Math.round(g.netOverride) : Math.round(g.gross * 0.989);
      let invoiceList = Array.from(g.invoices);
      // Luyen, 2026-07-17: gian "AE TAN AN KVC" nhan 2 loai HD cung ngay --
      // 1 HD alias tu "TUTU MN AEON MALL TÂN AN" (doanh thu Momo THAT), va 1
      // HD ghi thang ma "AE TAN AN KVC" (thuong la tien cua hang truong nop
      // vao, KHONG phai doanh thu Momo -- nhung thinh thoang ngay do lai la
      // doanh thu Momo that, vd 17/7/2026 HD 9701: TUTU (HD 9700) mot minh
      // KHONG du khop doanh thu ngay do, can ca 900k cua HD "AE TAN AN KVC"
      // moi khop dung). Thay vi lien co dinh tung so HD (de sai, phai sua tay
      // tung lan nhu 9701), tu dong quyet dinh: neu CHI rieng cac HD alias
      // (khong phai raw "AE TAN AN KVC") da du khop doanh thu ngay do roi,
      // thi cac HD raw "AE TAN AN KVC" la du/tien cua hang truong -- loai
      // khoi invoiceTotal. Neu KHONG du (thieu doanh thu), la HD that -- GIU
      // lai. Tu ap dung dung cho moi ky trong tuong lai, khong can sua tay.
      if (g.code === "AE TAN AN KVC") {
        const rawDirect = invoiceList.filter((soHd) => {
          const inv = invoiceData.invoices.find((i) => i.soHd === soHd);
          return inv && inv.maDiem === "AE TAN AN KVC";
        });
        if (rawDirect.length > 0) {
          const otherTotal = invoiceList
            .filter((soHd) => !rawDirect.includes(soHd))
            .reduce((sum, soHd) => {
              const inv = invoiceData.invoices.find((i) => i.soHd === soHd);
              return sum + (inv ? inv.tongTt : 0);
            }, 0);
          if (otherTotal >= g.gross - 1) {
            // Cac HD "khac" (alias tu TUTU) da du khop doanh thu -- HD raw
            // "AE TAN AN KVC" la du, coi la tien cua hang truong, loai ra.
            invoiceList = invoiceList.filter((soHd) => !rawDirect.includes(soHd));
          }
          // Nguoc lai (otherTotal < gross): HD raw can thiet de khop du
          // doanh thu -- giu nguyen invoiceList, khong loai.
        }
      }
      const invoiceTotal = invoiceList.reduce((sum, soHd) => {
        const inv = invoiceData.invoices.find((i) => i.soHd === soHd);
        return sum + (inv ? inv.tongTt : 0);
      }, 0);
      return {
        code: g.code,
        maCongTrinh: displayCode(g.code),
        label: prettyCodeLabel(g.code),
        tkCo: gianMapping[g.code] || "131",
        gross: g.gross,
        net,
        invoiceNumbers: invoiceList,
        invoiceTotal,
        diff: invoiceTotal - g.gross,
        days: Array.from(g.days).sort(),
        matched: invoiceList.length > 0 && Math.abs(invoiceTotal - g.gross) < 1,
      };
    });
    // Gian mapped to "SKIP" (KH moi) settle into a DIFFERENT bank account,
    // not this one -- so their revenue must be excluded from the total used
    // to check against this settlement's bank amount, otherwise diffVsBank
    // would wrongly show a mismatch. They stay in `lines` so the review page
    // can still display them, just excluded from this reconciliation total.
    const totalNet = lines
      .filter((l) => l.tkCo !== "SKIP")
      .reduce((sum, l) => sum + l.net, 0);
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

// ---------- Pending-bank rows (dung chung cho Momo/ZVP/VietQR) ----------
// Luyen, 2026-07-24: "tôi có tải dữ liệu nên rồi á check cho tôi lên bảng
// luôn đi không cần chờ ngân hàng đâu ... lên số với gian trước đi nào có
// ngân hàng đối chiếu sau" -- truoc gio moi dong doi soat chi hien ra khi co
// 1 "settlement" (khoan tien ve ngan hang) THAT tuong ung, nen 1 ngay da co
// doanh thu tai len (va co the da co hoa don khop) nhung ngan hang CHUA ve
// (hoac sao ke chua tai) se AN HOAN TOAN khoi trang, du du lieu gian/doanh
// thu/hoa don da san sang tu lau.
//
// Ham nay quet grossData.grossByCode (key "NGAY|MA") de tim moi NGAY co
// doanh thu > 0 nhung KHONG nam trong bat ky khoang ngay settlement THAT nao
// (dateRange(fromIso,toIso) cua tung settlement) -- moi ngay con thieu duoc
// tao thanh 1 "settlement gia" rieng (fromIso=toIso=ngay do), danh dau
// pendingBank:true va amount:null, de cac ham reconcile*Channel xu ly y het
// mot settlement that (van dò hoa don, tinh TK Co, doanh thu... binh thuong)
// -- chi khac la khong co so tien ngan hang de doi chieu/tinh Chenh lech,
// hien "Chua co ngan hang" thay vi mot con so. Khi sao ke ve sau va tao ra
// settlement THAT bao gom dung ngay do, ngay do se tu dong chuyen sang dong
// binh thuong (khong con la "gia" nua) tu lan render tiep theo, khong lo bi
// dung 2 lan.
function buildPendingDaySettlements(settlements, grossData) {
  const covered = new Set();
  for (const s of settlements || []) {
    for (const day of dateRange(s.fromIso, s.toIso)) covered.add(day);
  }
  const pendingDays = new Set();
  const grossByCode = (grossData && grossData.grossByCode) || {};
  for (const key of Object.keys(grossByCode)) {
    const val = grossByCode[key];
    if (!val || val <= 0) continue;
    const day = key.split("|")[0];
    if (!covered.has(day)) pendingDays.add(day);
  }
  return Array.from(pendingDays)
    .sort()
    .map((day) => ({
      date: day,
      amount: null,
      fromIso: day,
      toIso: day,
      pendingBank: true,
      id: "pending-" + day,
      txIds: [],
    }));
}

module.exports = {
  parseTongMomoWorkbook,
  parseInvoiceWorkbook,
  parseRawMomoPortalWorkbook,
  resolveRawPortalGross,
  parseKhMoiFeeTransactionWorkbook,
  resolveKhMoiFeeTransactionGross,
  reconcileMomo,
  extractMomoSettlements,
  buildPendingDaySettlements,
  toIsoDate,
  isoToDmy,
  normCode,
  normText,
  dateRange,
  displayCode,
  prettyCodeLabel,
  SPLIT_PARENT_CODE,
  FF_SUFFIX,
};
