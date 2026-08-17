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

// Real bank-statement date range for this app's lifetime is comfortably
// within [2000, 2100] -- rejects the footer/summary rows some exports leave
// at the bottom of the sheet (e.g. "So du kha dung"/"Available Balance")
// which happen to have SOME value in the date/balance columns that would
// otherwise misparse as an Excel-epoch date like 1899-12-30, and get
// mistaken for a real (and enormous, since balance=0 there) transaction.
function isPlausibleYear(y) {
  return y >= 2000 && y <= 2100;
}

function parseDateCell(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    const iso = excelSerialToIso(v);
    const y = Number(iso.slice(0, 4));
    return isPlausibleYear(y) ? iso : null;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return isPlausibleYear(Number(y))
      ? `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`
      : null;
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const [, y, mo, d] = m;
    return isPlausibleYear(Number(y))
      ? `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`
      : null;
  }
  // Chi Nhan, 2026-07-29: file VTB (Vietinbank eFast) ghi ngay kieu
  // "29-07-2026 12:25:38" (GACH NGANG, ngay-thang-nam, khac voi kieu "/" o
  // tren) -- them nhan dien rieng, dat SAU pattern "yyyy-mm-dd" phia tren de
  // khong nham (nam luon 4 chu so dung dau, ngay/thang o day chi 1-2 chu so).
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return isPlausibleYear(Number(y))
      ? `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`
      : null;
  }
  return null;
}

// Chi Nhan, 2026-07-29: mot so file sao ke (vd MB "Thong tin giao dich") ghi
// ca gio:phut:giay trong CUNG 1 o voi ngay ("29/07/2026 13:38:59"), va liet
// ke giao dich theo thu tu MOI NHAT TRUOC ke ca khi TAT CA giao dich cung 1
// NGAY (vd tai 1 ngay trong 1 phien lam viec). Logic dao nguoc cu chi so
// sanh rows[0].date vs rows[last].date (CHI theo ngay, khong theo gio) nen
// khong phat hien duoc truong hop nay -- ket qua la delta so du bi tinh
// NGUOC chieu cho hang loat dong cung ngay (vd 25/26 dong bi tinh nham thanh
// "chi" thay vi "thu" thuc te, va dong dau tien bi cong don ca canh so du
// that lon). Ham nay tra ve 1 sort-key day du (ca ngay + gio, dang so) neu
// doc duoc gio trong o ngay, de sap xep lai CHINH XAC theo thoi gian thuc,
// thay vi chi dua vao ngay nhu truoc.
function parseDateTimeSortKey(v) {
  if (v === null || v === undefined || typeof v !== "string") return null;
  const s = v.trim();
  // Chi Nhan, 2026-07-29: VTB dung gach ngang "29-07-2026 12:25:38" thay vi
  // "/" -- chap nhan ca 2 kieu.
  const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  const [, d, mo, y, h, mi, se] = m;
  if (!isPlausibleYear(Number(y))) return null;
  return (
    Number(y) * 100000000 +
    Number(mo) * 1000000 +
    Number(d) * 10000 +
    Number(h) * 100 +
    Number(mi) * 1 +
    Number(se) / 100
  );
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
  // Chi Nhan, 2026-07-29: file "VTB KH MỚI.xls" (Vietinbank eFast, TK
  // 116003053982) dung tieu de rieng "Ngày hạch toán/Accounting date" thay vi
  // 3 kieu tren -- khong co cot nay se khong nhan dien duoc file, bao loi
  // "Khong nhan dien duoc file sao ke" du du lieu hop le.
  "ngay hach toan",
];
const BALANCE_HEADER_PATTERNS = ["so du"];
// Chi Nhan, 2026-07-30: "sao ngân hàng trả... vẫn thiếu 20k" -- dong DAU TIEN
// cua 1 lan tai file phu thuoc so du luy ke he thong dang co san TRUOC ngay
// do; neu so du luy ke nay bi lech (vi du do 1 sai sot lich su nao khac chua
// phat hien) thi rieng dong dau tien se bi tinh SAI ca loai (thu/chi) lan so
// tien theo dung phan chenh lech do -- moi dong SAU van dung vi tu chuoi theo
// dung so du cua tung dong ghi san trong file (khong phu thuoc gia tri
// "running" ben ngoai nua). Them phat hien cot "Phat sinh No/Co" (khi file co
// san) de dung lam LUOI AN TOAN CHI CHO DONG DAU TIEN: neu gia tri No/Co that
// cua file mau thuan ro voi ket qua tinh tu delta so du, uu tien gia trị No/
// Co (dang tin hon vi khong phu thuoc so du luy ke ben ngoai) -- KHONG ap
// dung cho cac dong con lai (van giu nguyen triet ly cu, tranh lap lai bug
// "DW" text da tung gap voi cot No/Co).
const DEBIT_HEADER_PATTERNS = ["phat sinh no", "debit amount", "debit"];
const CREDIT_HEADER_PATTERNS = ["phat sinh co", "credit amount", "credit"];
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
//
// Chi Nhan, 2026-07-24: "Số tham chiếu" (dung de khop voi "Ma tham chieu"
// ben file QR, xem resolveGianGrossByBankRef trong utils/vietqrReconcile.js)
// PHAI duoc uu tien hon "Số chứng từ"/"So CT" -- file that BIDV7702 co CA HAI
// cot nay CUNG luc, va "Số chứng từ" (so phieu tuan tu vd "75918", "75919")
// luon nam TRUOC "Số tham chiếu" theo thu tu cot trong file thuc te, nen
// PATTERN_LIST.some() + "lay cot dau tien khop" (thu tu quet TRAI SANG PHAI)
// truoc day luon chon nham "Số chứng từ" lam .reference, khien no KHONG BAO
// GIO khop duoc voi "Ma tham chieu" ben file QR (2 gia tri hoan toan khac
// nhau ban chat: 1 ben la so phieu ke toan, 1 ben la ma giao dich QR that).
// Tach rieng nhom pattern "chinh xac" (so tham chieu/reference) uu tien hon
// nhom "du phong" (so chung tu/so ct) -- chi dung du phong khi KHONG co cot
// nao khop nhom chinh trong ca dong tieu de.
const REF_HEADER_PATTERNS_PRIMARY = ["so tham chieu", "reference"];
// Chi Nhan, 2026-07-29: "chỗ viet qr 115 giao dịch về nhiều tiền mà chèn sao
// số ngân hàng ít vậy" -- sao ke MB (vd tai khoan 11521268) khong co cot "So
// chung tu"/"So tham chieu" nhung co cot "BÚT TOÁN" (vd "FT26210189800001"),
// la ma but toan RIENG BIET cho tung dong, hoan toan du dieu kien lam
// reference de dedup. Truoc khi them pattern nay, file sao ke MB bi doc voi
// refCol=-1 (khong nhan dien duoc cot nao), khien dedup phai roi ve fallback
// date+amount+type -- nhieu giao dich VietQR gia co dinh (20k/50k/100k/200k)
// trung ngay TRUNG SO TIEN bi gop nham lam 1, lam mat that hang chuc giao
// dich that moi lan tai file len (~drop 20/26 dong trong 1 lan test thuc te).
// Day chinh la nguyen nhan cot "Ngan hang" cua MB11521268 bi thap bat thuong
// so voi "Tinh tu du lieu tai len".
const REF_HEADER_PATTERNS_FALLBACK = ["so chung tu", "so ct", "but toan"];
// Chi Nhan, 2026-07-29: "thêm cho tôi trên cái ngân hàng có hiển thị cái tên
// đối ứng trên sao kê luôn nha các tài khoản khác cũng vậy á" -- doc them cot
// "Tên đối ứng" (ten cua ben kia giao dich, vd NCC/khach hang), CUNG 1 pattern
// da dung o utils/chiphiReconcile.js (VENDOR_PATTERNS) cho sao ke Chi Phi,
// gio ap dung chung cho MOI ngan hang o trang Giao dich/Sao ke thuong.
// Chi Nhan, 2026-07-30: "tất cả các tài khoản điều lấy tên đối ứng cho tôi
// đi" -- phat hien file MB (Military Bank, vd MB02865168) dung tieu de cot
// KHAC HAN "Tên đối ứng" ("ĐƠN VỊ THỤ HƯỞNG/ĐƠN VỊ CHUYỂN"), khong khop pattern
// cu nen MOI giao dich cua MB import qua file (khong phai dan tay) van bi bo
// trong ten doi ung du file THUC SU co cot nay. Them cac bien the thuong gap
// (MB dung "don vi thu huong"/"don vi chuyen").
const VENDOR_HEADER_PATTERNS = ["ten doi ung", "don vi thu huong", "don vi chuyen"];

function findHeaderRow(grid) {
  // Chi Nhan, 2026-07-29: file VTB (Vietinbank eFast) co phan "Thông tin chi
  // tiết tài khoản" (ten cong ty, so TK, so du dau/cuoi ky...) chiem toi 24
  // dong truoc dong tieu de that su -- gioi han cu (20 dong) bo lo, bao loi
  // "Khong nhan dien duoc file sao ke" du file hop le. Noi rong len 40 dong
  // (van an toan, cac file khac chua tung can qua 6-10 dong).
  for (let r = 0; r < Math.min(grid.length, 40); r++) {
    const row = grid[r] || [];
    let dateCol = -1;
    let balCol = -1;
    let descCol = -1;
    let refColPrimary = -1;
    let refColFallback = -1;
    let vendorCol = -1;
    let debitCol = -1;
    let creditCol = -1;
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = normHeader(cell);
      if (dateCol === -1 && DATE_HEADER_PATTERNS.some((p) => h.includes(p))) dateCol = c;
      if (balCol === -1 && BALANCE_HEADER_PATTERNS.some((p) => h.includes(p))) balCol = c;
      if (descCol === -1 && DESC_HEADER_PATTERNS.some((p) => h.includes(p))) descCol = c;
      if (refColPrimary === -1 && REF_HEADER_PATTERNS_PRIMARY.some((p) => h.includes(p))) refColPrimary = c;
      if (refColFallback === -1 && REF_HEADER_PATTERNS_FALLBACK.some((p) => h.includes(p))) refColFallback = c;
      if (vendorCol === -1 && VENDOR_HEADER_PATTERNS.some((p) => h.includes(p))) vendorCol = c;
      if (debitCol === -1 && DEBIT_HEADER_PATTERNS.some((p) => h.includes(p))) debitCol = c;
      if (creditCol === -1 && CREDIT_HEADER_PATTERNS.some((p) => h.includes(p))) creditCol = c;
    });
    const refCol = refColPrimary !== -1 ? refColPrimary : refColFallback;
    if (dateCol !== -1 && balCol !== -1) {
      return {
        headerRowIdx: r,
        dateCol,
        balCol,
        descCol,
        refCol,
        refIsPrimary: refColPrimary !== -1,
        vendorCol,
        debitCol,
        creditCol,
      };
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
  let { headerRowIdx, dateCol, balCol, descCol, refCol, refIsPrimary, vendorCol, debitCol, creditCol } = found;
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
    const tenDoiUng = vendorCol >= 0 && row[vendorCol] !== null && row[vendorCol] !== undefined
      ? String(row[vendorCol]).trim()
      : "";
    const sortKey = parseDateTimeSortKey(row[dateCol]);
    const debitVal = debitCol >= 0 ? parseNumberCell(row[debitCol]) : null;
    const creditVal = creditCol >= 0 ? parseNumberCell(row[creditCol]) : null;
    rows.push({
      date,
      balance: bal,
      description: String(desc || "").trim(),
      reference,
      tenDoiUng,
      debit: debitVal,
      credit: creditVal,
      _sortKey: sortKey,
    });
  }

  if (rows.length === 0) {
    throw new Error("Khong doc duoc dong giao dich nao trong file (kiem tra lai dinh dang).");
  }

  // Luyen, 2026-07-25: "tại giao dịch kia bị lỗi á không tải về được bạn đọc
  // tạm file này nhá" -- file thay the tai truc tiep tu internet banking
  // ("THÔNG TIN LỊCH SỬ GIAO DỊCH...") xuat theo thu tu MOI NHAT TRUOC (dong
  // dau la giao dich gan nhat), NGUOC voi cac file sao ke Luyen hay tai
  // truoc gio (CU NHAT TRUOC). computeThuChi() gia dinh `rows` di theo thu tu
  // THOI GIAN TANG DAN (dong dau = giao dich cu nhat) de tinh dung delta so
  // du tung dong -- neu doc file "moi nhat truoc" ma khong dao lai, MOI delta
  // se bi tinh NGUOC (thu thanh chi, chi thanh thu, sai het so tien). Tu dong
  // phat hien + dao lai o day (thay vi doi hoi Luyen phai tu sap xep file
  // truoc khi tai) de ca 2 dinh dang deu ra ket qua dung.
  // Chi Nhan, 2026-07-29: neu MOI dong deu doc duoc gio:phut:giay day du
  // (_sortKey khac null), sap xep lai truc tiep theo _sortKey TANG DAN --
  // chinh xac hon cach cu (chi so sanh ngay dau/cuoi) va xu ly duoc ca
  // truong hop file "moi nhat truoc" ma TAT CA giao dich nam trong CUNG 1
  // ngay (vd sao ke MB xuat cho 1 ngay le), khi so sanh CHI theo ngay se
  // khong phat hien duoc thu tu bi dao nguoc. Array.prototype.sort trong
  // Node/V8 hien la stable sort nen khong lam xao tron thu tu cac dong co
  // cung sort-key (neu co).
  const allHaveSortKey = rows.every((r) => r._sortKey !== null);
  if (allHaveSortKey && rows.length > 1) {
    // Chi Nhan, 2026-07-29: khi 2 dong TRUNG sortKey (cung ngay-gio-phut-giay
    // -- vd file VTB "29-07-2026 12:25:38" xuat hien 2 lan cho 2 giao dich
    // Payoo khac nhau trong cung 1 giay), Array.sort (stable) se GIU NGUYEN
    // thu tu doc tu file cho cac dong trung nhau. Nhung neu file goc la "moi
    // nhat truoc" (dong dau = giao dich gan nhat, quy uoc tren-moi duoi-cu ap
    // dung cho CA FILE) thi trong 1 nhom trung sortKey cung phai ap dung DUNG
    // quy uoc do: dong doc duoc SAU (nam duoi trong file goc) la giao dich
    // CU HON. Kiem chung bang chuoi so du that: giao dich xep sau trong file
    // (VD26 QRCODE) khop chinh xac voi so du TRUOC giao dich xep truoc (VD25
    // GIAITRIKH_TONG, so du sau = dung bang "So du cuoi ky" cua file) -- neu
    // giu nguyen thu tu doc (khong dao) se tinh SAI dau thu/chi cho giao dich
    // dung sau (thu bi tinh thanh chi vi so du "giam" so voi dong truoc no).
    const firstKey = rows[0]._sortKey;
    const lastKey = rows[rows.length - 1]._sortKey;
    const newestFirst = firstKey > lastKey;
    rows.forEach((r, i) => { r._origIdx = i; });
    rows.sort((a, b) => {
      if (a._sortKey !== b._sortKey) return a._sortKey - b._sortKey;
      return newestFirst ? b._origIdx - a._origIdx : a._origIdx - b._origIdx;
    });
    rows.forEach((r) => delete r._origIdx);
  } else if (rows.length > 1 && rows[0].date > rows[rows.length - 1].date) {
    // Luyen, 2026-07-25: "tại giao dịch kia bị lỗi á không tải về được bạn
    // đọc tạm file này nhá" -- file thay the tai truc tiep tu internet
    // banking ("THÔNG TIN LỊCH SỬ GIAO DỊCH...") xuat theo thu tu MOI NHAT
    // TRUOC (dong dau la giao dich gan nhat), NGUOC voi cac file sao ke
    // Luyen hay tai truoc gio (CU NHAT TRUOC). computeThuChi() gia dinh
    // `rows` di theo thu tu THOI GIAN TANG DAN (dong dau = giao dich cu
    // nhat) de tinh dung delta so du tung dong -- neu doc file "moi nhat
    // truoc" ma khong dao lai, MOI delta se bi tinh NGUOC (thu thanh chi,
    // chi thanh thu, sai het so tien). Tu dong phat hien + dao lai o day
    // (khi khong co du gio de sap xep chinh xac hon) de ca 2 dinh dang deu
    // ra ket qua dung.
    rows.reverse();
  }
  rows.forEach((r) => delete r._sortKey);

  // Chi Nhan, 2026-07-24: co "refIsPrimary" de goi noi (routes/transactions.js)
  // biet cot reference vua doc duoc la "So tham chieu" that (dang alnum co
  // gach ngang, dung de khop VietQR) hay chi la "So chung tu" du phong (chuoi
  // so tuan tu ngan, vd "75918") -- can phan biet de tu sua nhung dong DA
  // LUU TRUOC DAY bang gia tri SAI (bug cu: "So chung tu" bi nham la
  // reference) ma khong lam sai nhung ngan hang khac von chi CO san "So
  // chung tu" (hop le voi ho, khong phai loi).
  return { sheetName, rows, refIsPrimary: !!refIsPrimary };
}

// Turn balance-anchored rows into thu/chi transactions.
// priorBalance: the account balance immediately before the first row in `rows`.
function computeThuChi(rows, priorBalance) {
  const out = [];
  let running = priorBalance;
  rows.forEach((r, idx) => {
    const delta = r.balance - running;
    let type, amount;
    if (delta > 0) {
      type = "thu";
      amount = delta;
    } else if (delta < 0) {
      type = "chi";
      amount = -delta;
    } else if (!isNaN(delta)) {
      running = r.balance;
      return; // delta === 0: no actual movement -> skip
      // NaN delta (priorBalance unknown): fall through to idx===0 explicit
      // credit/debit recovery below rather than silently dropping the row.
    }
    // Chi Nhan, 2026-07-30: chi ap dung luoi an toan nay cho DONG DAU TIEN cua
    // lan tai nay (idx===0) -- day la dong DUY NHAT phu thuoc "running"
    // (=priorBalance, tinh tu du lieu DA LUU truoc do, co the bi lech neu co
    // sai sot lich su chua phat hien) thay vi so du CHINH cua dong truoc DO
    // trong CUNG file nay; cac dong sau deu tu chuoi dung theo file, khong can
    // (va khong nen) can thiep. Neu file co san cot No/Co hop le va gia tri do
    // mau thuan ro voi (type, amount) tinh tu delta, uu tien No/Co.
    // Luyen 2026-08-17: cung xu ly NaN delta (priorBalance = undefined khi goi
    // khong truyen tham so) -- neu khong co explicit No/Co thi skip hang nay.
    if (idx === 0 && (r.debit !== null || r.credit !== null)) {
      const explicitCredit = r.credit || 0;
      const explicitDebit = r.debit || 0;
      let explicitType = null;
      let explicitAmount = null;
      if (explicitCredit > 0 && explicitDebit === 0) {
        explicitType = "thu";
        explicitAmount = explicitCredit;
      } else if (explicitDebit > 0 && explicitCredit === 0) {
        explicitType = "chi";
        explicitAmount = explicitDebit;
      }
      if (explicitType && (!type || explicitType !== type || Math.abs(explicitAmount - amount) > 1)) {
        type = explicitType;
        amount = explicitAmount;
      }
    }
    if (!type) {
      // Could not determine direction (e.g. NaN delta + no usable explicit cols)
      running = r.balance;
      return;
    }
    out.push({
      date: r.date,
      description: r.description,
      amount: Math.round(amount * 100) / 100,
      type,
      reference: r.reference || "",
      tenDoiUng: r.tenDoiUng || "",
    });
    running = r.balance;
  });
  return out;
}

module.exports = { parseBankStatement, computeThuChi, normHeader };
