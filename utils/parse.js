// Helpers to parse pasted bank statement text and Vietnamese-formatted numbers.

/**
 * Parse a Vietnamese-formatted amount string into a float.
 * Handles: "1.250.000", "1,250,000", "1250000", "-500000", "1.250.000,50"
 */
function parseAmount(raw) {
  if (raw === undefined || raw === null) return NaN;
  let s = String(raw).trim();
  if (s === "") return NaN;
  const negative = /^-/.test(s) || /^\(.*\)$/.test(s);
  s = s.replace(/[()]/g, "");
  // Keep digits, comma, dot, minus
  s = s.replace(/[^0-9.,-]/g, "");
  // If both , and . appear, assume the last one is the decimal separator
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma > -1 && lastDot > -1) {
    if (lastComma > lastDot) {
      // comma is decimal sep
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // dot is decimal sep
      s = s.replace(/,/g, "");
    }
  } else if (lastComma > -1) {
    // Only commas present -> likely thousands separators unless exactly 2 digits follow last comma
    const decimals = s.length - lastComma - 1;
    if (decimals === 2) {
      s = s.replace(/,/g, (m, i) => (i === lastComma ? "." : ""));
      s = s.replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (lastDot > -1) {
    const decimals = s.length - lastDot - 1;
    if (decimals !== 2) {
      s = s.replace(/\./g, "");
    }
  }
  s = s.replace(/-/g, "");
  let val = parseFloat(s);
  if (isNaN(val)) return NaN;
  if (negative) val = -val;
  return val;
}

/**
 * Parse a date string in dd/mm/yyyy, d/m/yyyy, yyyy-mm-dd, or dd-mm-yyyy into ISO yyyy-mm-dd.
 */
function parseDate(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const [, y, mo, d] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return null;
}

/**
 * Parse bulk-pasted statement text into an array of transaction rows.
 * Expected columns per line (tab or 2+ spaces or comma separated):
 *   Ngay  Dien giai  So tien  [Loai: Thu/Chi]
 * If "Loai" column is missing, sign of amount decides (positive = Thu, negative = Chi).
 * Returns { rows: [...], errors: [...] }
 */
function parsePastedTransactions(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const rows = [];
  const errors = [];

  lines.forEach((line, idx) => {
    let cols;
    if (line.includes("\t")) {
      cols = line.split("\t").map((c) => c.trim());
    } else if (line.includes(",")) {
      cols = line.split(",").map((c) => c.trim());
    } else {
      cols = line.split(/\s{2,}/).map((c) => c.trim());
    }
    cols = cols.filter((c) => c.length > 0);

    if (cols.length < 3) {
      errors.push(`Dong ${idx + 1}: khong du cot (can it nhat Ngay, Dien giai, So tien) -> "${line}"`);
      return;
    }

    const date = parseDate(cols[0]);
    if (!date) {
      errors.push(`Dong ${idx + 1}: khong doc duoc ngay "${cols[0]}"`);
      return;
    }

    const description = cols[1];
    const amountRaw = cols[2];
    const amount = parseAmount(amountRaw);
    if (isNaN(amount)) {
      errors.push(`Dong ${idx + 1}: khong doc duoc so tien "${amountRaw}"`);
      return;
    }

    let type;
    const typeCol = (cols[3] || "").toLowerCase();
    if (typeCol.startsWith("thu") || typeCol === "c" || typeCol === "co") type = "thu";
    else if (typeCol.startsWith("chi") || typeCol === "n" || typeCol === "no") type = "chi";
    else type = amount < 0 ? "chi" : "thu";

    rows.push({
      date,
      description,
      amount: Math.abs(amount),
      type,
    });
  });

  return { rows, errors };
}

module.exports = { parseAmount, parseDate, parsePastedTransactions };
