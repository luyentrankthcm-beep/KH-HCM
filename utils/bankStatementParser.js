const XLSX = require("xlsx");

// Generic bank-statement parser.
//
// Vietnamese bank exports (BIDV, ACB, MB, Vietcombank, ...) all differ in
// exact column layout, but every one of them has (a) a transaction date
// column and (b) a running balance ("Số dư") column. Rather than hard-code
// per-bank column positions (which breaks the moment a bank tweaks its
// export template), we detect those two columns by header keyword, then
// derive Thu/Chi purely from the balance's day-to-day DELTA — the balance
// column is bank-verified ground truth, so this sidesteps any ambiguity or
// data-entry glitches in the Nợ/Có columns themselves (we hit exactly this
// kind of glitch once already: a debit that was mistakenly typed as text
// "DW" instead of a number — the delta-from-balance approach recovers the
// correct amount automatically because it never reads that column at all).

function removeDiacritics(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}

function normHeader(s) {
  return removeDiacritics(String(s || ""))
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
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
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

const DATE_HEADER_PATTERNS = [
  "ngay giao dich",
  "ngay hieu luc",
  "thoi gian giao dich",
];
const BALANCE_HEADER_PATTERNS = ["so du"];
const DESC_HEADER_PATTERNS = [
  "dien giai",
  "noi dung giao dich",
  "noi dung",
];
// A per-row unique bank-assigned ID, when the export exposes one. This is
// the ONLY reliable way to tell apart two genuinely distinct transactions
// that share the same date + amount + type (extremely common for VietQR
// fixed-price ticket sales: dozens/hundreds of same-day 20.000d/50.000d
// transactions from different customers). "Ma giao dich"/"Trans.Code" is
// deliberately excluded here -- on BIDV exports that column just holds a
// transaction-type code like "DD" repeated on every row, not a unique ID.
const REF_HEADER_PATTERNS = [
  "so tham chieu",
  "so chung tu",
  "so ct",
  "reference",
];

function findHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 20); r++) {
    const row = grid[r] || [];
    let dateCol = -1;
    let balCol = -1;
    let descCol = -1;
    let refCol = -1;
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = normHeader(cell);
      if (dateCol === -1 && DATE_HEADER_PATTERNS.some((p) => h.includes(p))) dateCol = c;
      if (balCol === -1 && BALANCE_HEADER_PATTERNS.some((p) => h.includes(p))) balCol = c;
      if (descCol === -1 && DESC_HEADER_PATTERNS.some((p) => h.includes(p))) descCol = c;
      if (refCol === -1 && REF_HEADER_PATTERNS.some((p) => h.includes(p))) refCol = c;
    });
    if (dateCol !== -1 && balCol !== -1) {
      return { headerRowIdx: r, dateCol, balCol, descCol, refCol };
    }
  }
  return null;
}

// Pick the widest (most likely descriptive text) column in the data rows,
// excluding the date/balance columns, as a fallback when no header matched
// "Diễn giải"/"Nội dung".
function guessDescCol(grid, headerRowIdx, dateCol, balCol, maxCol) {
  const scores = {};
  const sampleEnd = Math.min(grid.length, headerRowIdx + 60);
  for (let r = headerRowIdx + 1; r < sampleEnd; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < Math.min(row.length, maxCol + 1); c++) {
      if (c === dateCol || c === balCol) continue;
      const v = row[c];
      if (typeof v === "string" && v.length > 8) {
        scores[c] = (scores[c] || 0) + v.length;
      }
    }
  }
  let best = -1;
  let bestScore = 0;
  for (const [c, score] of Object.entries(scores)) {
    if (score > bestScore) {
      bestScore = score;
      best = Number(c);
    }
  }
  return best;
}

function parseBankStatement(buffer, sheetNameHint) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const sheetName = sheetNameHint && wb.Sheets[sheetNameHint] ? sheetNameHint : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const found = findHeaderRow(grid);
  if (!found) {
    throw new Error(
      'Khong nhan dien duoc file sao ke: can co cot "Ngay giao dich" (hoac "Ngay hieu luc") va cot "So du".'
    );
  }
  let { headerRowIdx, dateCol, balCol, descCol, refCol } = found;
  const maxCol = (grid[headerRowIdx] || []).length;
  if (descCol === -1) {
    descCol = guessDescCol(grid, headerRowIdx, dateCol, balCol, maxCol);
  }

  const rows = [];
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const date = parseDateCell(row[dateCol]);
    const bal = parseNumberCell(row[balCol]);
    if (!date || bal === null) continue;
    let desc = descCol >= 0 ? row[descCol] : null;
    if (desc === null || desc === undefined) {
      // Fall back to concatenating any other text cells on the row.
      desc = row
        .filter((v, c) => c !== dateCol && c !== balCol && typeof v === "string" && v.trim())
        .join(" ")
        .trim();
    }
    const reference = refCol >= 0 && row[refCol] !== null && row[refCol] !== undefined
      ? String(row[refCol]).trim()
      : "";
    rows.push({ date, balance: bal, description: String(desc || "").trim(), reference });
  }

  if (rows.length === 0) {
    throw new Error("Khong doc duoc dong giao dich nao trong file (kiem tra lai dinh dang).");
  }

  return { sheetName, rows };
}

// Turn balance-anchored rows into thu/chi transactions.
// priorBalance: the account balance immediately before the first row in `rows`.
function computeThuChi(rows, priorBalance) {
  const out = [];
  let running = priorBalance;
  for (const r of rows) {
    const delta = r.balance - running;
    let type, amount;
    if (delta > 0) {
      type = "thu";
      amount = delta;
    } else if (delta < 0) {
      type = "chi";
      amount = -delta;
    } else {
      running = r.balance;
      continue; // no actual movement (delta ~ 0) -> skip
    }
    out.push({
      date: r.date,
      description: r.description,
      amount: Math.round(amount * 100) / 100,
      type,
      reference: r.reference || "",
    });
    running = r.balance;
  }
  return out;
}

module.exports = { parseBankStatement, computeThuChi, normHeader };
