const XLSX = require("xlsx");
const { parseAmount } = require("./parse");

// Parser cho file Google Sheet "ĐI ỦY NHIỆM CHI KVC + MTĐ MN" (nguồn dữ liệu
// cho trang /chi-phi). Luyen, 2026-07-21: "thêm nút cập nhật chi phí" -- truoc
// gio moi lan nhap du lieu (thang 7, roi thang 1-6) deu lam bang script rieng
// (mot lan, khong luu lai trong app). Ham nay dung LAI DUNG logic da kiem
// chung o script parse_chiphi.py (dot nhap thang 1-6, 2026-07-21), chuyen
// sang JS va TONG QUAT HOA ten sheet (do bang regex "Tháng N.YYYY" thay vi
// hardcode danh sach ten sheet cu the) de nhung lan upload sau (thang 8, 9...)
// khong can sua code nua.
//
// Cot du lieu co dinh trong tung sheet "Thang N.YYYY" (bat dau tu dong 3, dong
// 1-2 la tieu de -- xem anh Luyen gui 2026-07-20):
//   cot A (0): BP/gian (co the la merged-cell, forward-fill xuong cac dong con)
//   cot B (1): "Ngày cần đi tiền" -- thuc te la SO UNC (vd "UNC 01/7"), TEN COT SAI
//   cot C (2): "Ngày tạo lệnh" -- van ban dang "PAYMENT dd/mm", dung lam fallback ngay
//   cot D (3): "Nội dung trên đề xuất" -- dien giai, thuong co "...theo hđ <so>"
//   cot F (5): "Số tiền"
//   cot G (6): "Tên đơn vị thụ hưởng" -- NCC
//   cot J (9): "Ngân hàng" -- thuc te ghi "K VÀ H CŨ"/"K VÀ H MỚI" (cong ty)
//
// TT TIỀN MẶT la 1 tab RIÊNG, cau truc cot hoan toan khac (khong co ngay/NCC/
// cong ty ro rang -- lan truoc phai mo tung link Drive/tra cuu hoa don bang
// tay/OCR moi xac dinh duoc). KHONG tu dong parse tab nay o day (de tranh
// nhap trung/sai du lieu da xu ly thu cong truoc do) -- chi bao cho Luyen biet
// da bo qua tab nay.

function removeDiacritics(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}

function normCongTy(v) {
  if (!v) return null;
  const n = removeDiacritics(String(v)).toLowerCase();
  if (n.includes("moi")) return "kh_moi";
  if (n.includes("cu")) return "kh_cu";
  return null;
}

const MONTH_TAB_RE = /th[áa]ng\s*(\d{1,2})[\s./-]+(\d{4})/i;
const DATE_RE = /(\d{1,2})\s*[.\/]\s*(\d{1,2})(?:\s*[.\/]\s*(\d{2,4}))?/;
const HD_RE = /h[đd]\.?\s*(?:s[ôo]\s*)?:?\s*#?(\d{1,6}(?:[\/\-]\d{1,4})?)/i;

function pad2(n) {
  return String(n).padStart(2, "0");
}

// Cong thuc chuyen serial-date Excel -> ISO, GIONG HET utils/bankStatementParser.js
// (excelSerialToIso, da kiem chung khop voi sao ke ngan hang that) -- dung UTC
// getters, KHONG dung Date object (`cellDates:true` cua thu vien xlsx tung gay
// LECH 1 NGAY khi test upload thu, xem ghi chu o duoi -- co le do cach thu vien
// xu ly bug "1900 la nam nhuan" cua Excel khac voi ham tinh serial thu cong nay).
function excelSerialToIso(serial) {
  const days = Math.round(serial);
  const utcMillis = (days - 25569) * 86400 * 1000;
  const d = new Date(utcMillis);
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

// Excel co the tra ve gia tri ngay dang SO SERIAL (khi cell duoc format dang
// ngay that -- doc bang raw:true, KHONG dung cellDates) hoac chuoi tu do (vd
// "PAYMENT 30/6", "UNC 01/7") -- xu ly ca 2 truong hop, giong extract_date()
// trong parse_chiphi.py.
function extractDate(cellValue, sheetMonth, sheetYear) {
  if (cellValue === null || cellValue === undefined || cellValue === "") return null;
  if (typeof cellValue === "number") {
    return excelSerialToIso(cellValue);
  }
  const s = String(cellValue);
  const m = s.match(DATE_RE);
  if (!m) return null;
  let dd = Number(m[1]);
  let mm = Number(m[2]);
  let yy = m[3] ? Number(m[3]) : sheetYear;
  if (yy < 100) yy += 2000;
  if (dd < 1 || dd > 31 || mm < 1 || mm > 12) return null;
  // Fix ngay giap ranh nam: 1 ngay "31/12" ghi trong sheet "Thang 1/2" (dau
  // nam) hau nhu chac chan la cuoi nam TRUOC do, khong phai cuoi nam nay --
  // chi ap dung khi KHONG co nam ro rang trong chinh text do.
  if (!m[3] && sheetMonth <= 2 && mm >= 11) yy -= 1;
  return `${yy}-${pad2(mm)}-${pad2(dd)}`;
}

function extractSoHoaDon(text) {
  if (!text) return "";
  const m = String(text).match(HD_RE);
  return m ? m[1] : "";
}

// Luyen, 2026-07-21 (xac nhan qua AskUserQuestion): cot A (gian) trong 1 so
// sheet (vd THÁNG 7.2026) khong luon ghi ten gian cu the -- nhieu dong chi ghi
// nhan nhom chung chung ("posh"/"JP"/"POSH+JP", do merge-cell forward-fill 1
// khoi lon nhieu dong khac gian nhau), trong khi ten gian THAT nam trong cot
// dien giai tu do (vd "...TT Tien thue ghe Van Hanh Mall thang 7.2026"). Quyet
// dinh cua Luyen: chi tu dong tach ten gian tu dien giai khi do CUM "ghế "
// (anchor do tin cay cao, xuat hien o hau het cac dong thue mat bang thuc te),
// KHONG chac chan (khong co "ghế" trong dien giai) thi de trong Gian cho
// Luyen tu dien tay -- tuyet doi khong dung nhan nhom chung chung lam Gian vi
// se ghi SAI du lieu ke toan thay vi chi thieu du lieu.
const BUCKET_LABELS = new Set(["posh", "jp", "posh+jp", "jp+posh"]);
const GIAN_FROM_DIENGIAI_RE = /gh(?:ế|e)\s+([^\n]+?)(?:\s+(?:th[áa]ng|ng[àa]y|t\d{1,2}[.\/]\d{4})\b|\s*$)/i;

function isGenericBucketLabel(gianText) {
  const n = removeDiacritics(gianText).toLowerCase().replace(/\s+/g, "");
  return !n || n === "x" || n === "z" || BUCKET_LABELS.has(n);
}

function extractGianFromDienGiai(dienGiai) {
  if (!dienGiai) return "";
  const m = String(dienGiai).match(GIAN_FROM_DIENGIAI_RE);
  return m ? m[1].trim() : "";
}

// Luyen, 2026-07-21: BUG tim thay qua test truoc khi upload that -- thu vien
// xlsx (JS) tra ve xuong dong trong 1 o dang "\r\n", trong khi ban ghi CU (import
// bang script Python/openpyxl truoc do) luu "\n" -- 2 chuoi nay KHAC NHAU nen
// khoa upsert (chiPhiRowKey) khong nhan ra dong da co, suyt gay nhan doi hang
// tram dong khi test (da phat hien va rollback truoc khi luu that). Chuan hoa
// ve "\n" o day de khop dung voi du lieu cu.
function cellText(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return "";
  return String(v).replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
}

// Cot B ("soUNC") doi khi lai la 1 o NGAY THAT (serial) thay vi chuoi "UNC
// dd/m" -- hien thi dang ngay ISO cho de doc thay vi in ra so serial tho.
function socUncText(v) {
  if (typeof v === "number") return excelSerialToIso(v);
  return cellText(v);
}

function buildMergeForwardFill(ws, colIndex) {
  const map = {};
  const merges = ws["!merges"] || [];
  merges.forEach((rng) => {
    if (rng.s.c === colIndex && rng.e.c === colIndex && rng.s.r !== rng.e.r) {
      const topAddr = XLSX.utils.encode_cell({ r: rng.s.r, c: colIndex });
      const topVal = ws[topAddr] ? ws[topAddr].v : null;
      for (let r = rng.s.r; r <= rng.e.r; r++) map[r] = topVal;
    }
  });
  return map;
}

function parseMonthSheet(ws, month, year) {
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const gianMergeMap = buildMergeForwardFill(ws, 0);
  const rows = [];
  for (let r = 2; r < grid.length; r++) {
    const row = grid[r] || [];
    let gianRaw = Object.prototype.hasOwnProperty.call(gianMergeMap, r) ? gianMergeMap[r] : row[0];
    const colB = row[1]; // soUNC
    const colC = row[2]; // ngay (text "PAYMENT dd/mm")
    const colD = row[3]; // dien giai
    const colF = row[5]; // so tien
    const colG = row[6]; // ncc
    const colJ = row[9]; // cong ty

    // Bo qua dong khong co so tien hop le (dong trong/dong tieu de phu).
    if (typeof colF !== "number" || colF === 0) continue;

    const gianRawText = cellText(gianRaw);
    const dienGiaiText = cellText(colD);
    const gianEmpty = !gianRawText || ["x", "z"].includes(gianRawText.toLowerCase());
    const dienGiaiEmpty = !dienGiaiText;
    const nccEmpty = !cellText(colG);
    // Dong tong nhom (subtotal) chen giua cac nhom trong sheet goc: chi co so
    // tien, khong co gian/dien giai/NCC nao ca -- bo qua, khong phai giao dich thuc.
    if (gianEmpty && dienGiaiEmpty && nccEmpty) continue;

    // Cot A la nhan nhom chung chung (posh/JP/...) hoac rong -- thu tach ten
    // gian tu dien giai (anchor "ghế "), khong tach duoc thi de trong (KHONG
    // dung nhan nhom lam Gian).
    const gianText = isGenericBucketLabel(gianRawText) ? extractGianFromDienGiai(dienGiaiText) : gianRawText;

    const ngay = extractDate(colB, month, year) || extractDate(colC, month, year) || "";
    rows.push({
      gian: gianText,
      soUNC: socUncText(colB),
      ngay,
      ncc: cellText(colG),
      dienGiai: cellText(colD),
      soTien: colF,
      congTy: normCongTy(colJ) || "kh_cu",
      soHoaDon: extractSoHoaDon(colD),
    });
  }
  return rows;
}

// Doc toan bo workbook, tu do dò cac sheet "Thang N.YYYY" bang regex (khong
// hardcode ten sheet) -- tra ve { monthRows: [{sheetName, month, year, rows}],
// skippedSheets: [ten cac sheet khong khop pattern thang, vd "TT TIỀN MẶT"] }.
function parseChiPhiSheetWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const monthRows = [];
  const skippedSheets = [];
  for (const sheetName of wb.SheetNames) {
    const m = sheetName.trim().match(MONTH_TAB_RE);
    if (!m) {
      skippedSheets.push(sheetName);
      continue;
    }
    const month = Number(m[1]);
    const year = Number(m[2]);
    const ws = wb.Sheets[sheetName];
    const rows = parseMonthSheet(ws, month, year);
    // Luu ten sheet DA TRIM (co the file goc co khoang trang thua dau ten
    // sheet, vd " THÁNG 5.2026") -- khop dung voi quy uoc nguon da dung tu
    // truoc gio (khong khoang trang thua).
    monthRows.push({ sheetName: sheetName.trim(), month, year, rows });
  }
  return { monthRows, skippedSheets };
}

// ---------- Parser cho file KVC MIEN BAC ("Tạo lệnh UNC KVC MB.xlsx") ----------
// Luyen, 2026-07-23: nguon rieng cho Chi Phi MIEN BAC -- KHAC HAN file MTĐ MN
// (Mien Nam): CHI 1 tab lien tuc, KHONG tach theo ten sheet "Tháng N.YYYY".
// Luyen xac nhan qua AskUserQuestion: "1 tab thôi KVC HN C THANH á" -- nhung
// KHONG hardcode ten sheet do (de phong doi ten sau nay), thay vao do nhan
// dien qua CHU KY COT: co ca "gian hàng", "Số tiền", "Tên đơn vị thụ hưởng",
// "Ngày tạo lệnh" trong 8 dong dau. Cong ty (KH Cu/Moi) KHONG co cot rieng --
// nam LAN trong 1 cot ghi chu ngan hang (header "...note rõ", gia tri thuc te
// vd "VP KH MỚI", "BIDV KH CŨ", "KH cũ VP" -- nhieu bien the hoa/thuong, co
// khong tien to VP/BIDV) -- dung LAI normCongTy() da co (chi can chua "moi"
// hoac "cu" sau khi bo dau). ~450/1057 dong lich su (truoc thang 7.2026) dong
// nay BO TRONG -- nhung dong do bi LOAI (unclassifiedCount++), KHONG doan mo
// hinh, tranh gan sai cong ty. "gian hàng" cot rieng cung hau nhu luon TRONG
// (chi 36/1057 dong co, da so la TEN NGUOI duyet don nhu "Hường"/"Dương" chu
// khong phai ten gian that) -- van lay nguyen gia tri cot nay neu co (khong
// tu suy doan tu dien giai nhu file Mien Nam, vi khong co anchor dang tin cay
// "ghế " o day), phan lon se trong cho Luyen tu dien tay giong nhu file Mien
// Nam. So hoa don: KHONG co cot rieng, nhung thuong nhac trong dien giai kieu
// "...theo hoa don so 12345 cho..." -- tan dung lai extractSoHoaDon() de lay
// duoc mien phi khi co.
function mbNormHeader(s) {
  return removeDiacritics(String(s || "")).toLowerCase().replace(/\s+/g, " ").trim();
}

function findMbHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 8); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = mbNormHeader(cell);
      if (idx.ngayTaoLenh === undefined && h.includes("ngay tao lenh")) idx.ngayTaoLenh = c;
      if (idx.ngayXinPm === undefined && h.includes("ngay xin pm")) idx.ngayXinPm = c;
      if (idx.gianHang === undefined && h.includes("gian hang")) idx.gianHang = c;
      if (idx.deXuat === undefined && h.includes("de xuat")) idx.deXuat = c;
      if (idx.unc === undefined && h.includes("tren unc")) idx.unc = c;
      if (idx.soTien === undefined && h === "so tien") idx.soTien = c;
      if (idx.thuHuong === undefined && h.includes("thu huong") && h.includes("don vi")) idx.thuHuong = c;
      if (idx.ghiChuCongTy === undefined && h.includes("note ro")) idx.ghiChuCongTy = c;
    });
    if (idx.soTien !== undefined && idx.thuHuong !== undefined && idx.gianHang !== undefined && idx.ghiChuCongTy !== undefined) {
      return { headerRowIdx: r, ...idx };
    }
  }
  return null;
}

// Luyen, 2026-07-23: "cái này dựa vào diễn giải lấy ra cái gian được mà có
// ghi đó" -- tu dong tach ten gian tu dien giai khi cot "gian hàng" rieng
// (neu co) dang trong, dung chung cho ca 2 nguon Mien Bac (sheet UNC KVC MB
// thu cong + sheet MTĐ MB "máy tự động"). Da test bang du lieu that (589 +
// 1283 dong dien giai) truoc khi dua vao code -- ty le trich duoc: ~25% file
// UNC KVC MB (anchor "cho "/"tại "), ~65% file MTĐ MB (anchor sau token
// thang "T11.2025 Ten Gian" hoac sau "thuê "/"thue "). Nguyen tac AN TOAN
// giong het extractGianFromDienGiai() cua file Mien Nam o tren: CHI tach khi
// co anchor dang tin cay, khop nhieu gian cung dong (noi boi "va"/"+"/",")
// hoac nhan chung chung (posh/JP/K&H/...) deu BO QUA (de trong) thay vi doan
// sai.
const MB_GENERIC_GIAN_LABELS = new Set([
  "", "x", "z", "cty", "cong ty", "kho", "van phong", "vp",
  "co so", "cac co so", "posh", "jp", "posh+jp", "jp+posh",
  "k h", "k&h", "k va h", "kvah",
]);
function mbIsGenericOrInvalidGian(t) {
  const n = removeDiacritics(t || "").toLowerCase().replace(/\s+/g, " ").trim();
  if (!n || MB_GENERIC_GIAN_LABELS.has(n)) return true;
  if (/^t\d{1,2}$/.test(n)) return true; // con sot lai 1 token thang don le (vd "T11"), khong phai ten gian
  return false;
}
function mbIsMultiSiteGian(t) {
  const n = removeDiacritics(t || "").toLowerCase();
  return / va | \+ |,/.test(" " + n + " ");
}
function mbCleanGianCandidate(raw) {
  if (!raw) return "";
  let g = String(raw).trim();
  // "P " don le dung dau ket qua la tan du cua cum "tiền P DTCS"/"tiền P dịch
  // vụ" (Luyen, 2026-07-23: bam vao vi du "tt tien P DTCS PF Vinke Timecity
  // T6.2026") -- KHONG phai 1 phan cua ten gian, bo di truoc khi kiem tra tiep.
  g = g.replace(/^p\s+/i, "").trim();
  if (/^(th[áa]ng)\b/i.test(g)) return ""; // bat dau bang "tháng" -- chac chan khong phai ten gian
  if (mbIsGenericOrInvalidGian(g) || mbIsMultiSiteGian(g)) return "";
  return g;
}

// Doan chu (thang/ngay/so HD/...) hay xuat hien NGAY SAU ten gian trong dien
// giai -- dung lam DIEM DUNG khi cat ten gian ra (nam TRONG lookahead, KHONG
// nam trong capture group, tranh dinh kem "thang 11"/"T11.2025" vao cuoi ten).
const MB_TRAILING_STOP =
  "th[áa]ng\\s*\\d+|ng[àa]y\\s*\\d{1,2}|hết\\s+th[áa]ng|\\d{1,2}\\/\\d{1,2}(?:\\/\\d{2,4})?|t\\d{1,2}[.\\/]\\d{1,4}|t\\d{1,2}\\b|hd\\b|hđ\\b|lan\\s*\\d";

// Anchor "cho "/"tại " -- file UNC KVC MB (vd "Thanh toán đơn kem cho nhà
// tuyết royal 8/7"). CHI dung "tại" CO DAU (khong dung "tai" khong dau -- da
// phat hien qua test that: "tai" khong dau khop NHAM vao trong tu "tai nạn").
// Loai truong hop "cho de/để..." (thuc ra la "chỗ để xe...", khong phai gioi
// tu "cho" + dia diem).
const MB_GIAN_RE_CHO_TAI = new RegExp(
  "\\b(?:cho|tại)\\s+(?!(?:de|để)\\b)([^\\n,.;:]+?)(?:\\s+(?:" + MB_TRAILING_STOP + ")\\b|[.,:]|\\s*$)",
  "i"
);
// File MTĐ MB "máy tự động": uu tien anchor "T<thang>.<nam> " + ten gian
// NGAY SAU (vd "tt tien thue T11.2025 Lotte Da Nang" -- chinh xac hon vi ten
// gian luon la cum tu CUOI CUNG trong cau); fallback anchor "thuê "/"thue " +
// ten gian (vd "tt tien thue Go Ha Long 1CL1 thang 11.2025").
const MB_GIAN_RE_AFTER_MONTHTOKEN = new RegExp(
  "\\bt\\d{1,2}[.\\/]\\d{1,4}\\s+([^\\n,.;:]+?)(?:\\s+(?:hd\\b|hđ\\b|lan\\s*\\d)\\b|[.,:]|\\s*$)",
  "i"
);
const MB_GIAN_RE_AFTER_THUE = new RegExp(
  "\\bthu[eê]\\s+(?!\\d)([^\\n,.;:]+?)(?:\\s+(?:" + MB_TRAILING_STOP + ")\\b|[.,:]|\\s*$)",
  "i"
);
// Luyen, 2026-07-23: "bên nội dung có tiền thuê gian nào nè posh gì đó hay jp
// gì đó hay gian long biên... quăng vô gian cho tôi" -- nhieu dong "Nội dung
// unc" ghi loai chi phi (DTCS/thuê/điện/dịch vụ) NGAY TRUOC ten gian, THUONG
// kem theo 1 nhan hieu chung chung (POSH/PF/JP/K&H) chen giua, vd "tt tien
// DTCS PF Vinke Timecity T6.2026" hay "tt tien dien Posh VW Vu Yen T6.2026"
// hay "tt tien phi dich vu Posh Aeonmall Long Bien" (khong co thang o cuoi).
// Fallback tang 3 (sau anchor thang-token va anchor "thuê " don thuan o tren)
// -- bo qua nhan hieu chung chung ngay sau loai chi phi, lay phan con lai.
const MB_GIAN_RE_AFTER_LOAICHIPHI = new RegExp(
  "\\b(?:thu[eê]|dtcs|đi[eệ]n|dien|d[iị]ch\\s*vu)\\s+(?:posh|pf|jp|k\\s*h|k&h|k\\s*va\\s*h)?\\s*([^\\n,.;:]+?)(?:\\s+(?:" +
    MB_TRAILING_STOP +
    ")\\b|[.,:]|\\s*$)",
  "i"
);

function extractGianChoTai(dienGiai) {
  if (!dienGiai) return "";
  const m = String(dienGiai).match(MB_GIAN_RE_CHO_TAI);
  return mbCleanGianCandidate(m ? m[1] : "");
}
function extractGianThueMayTuDong(dienGiai) {
  if (!dienGiai) return "";
  const text = String(dienGiai);
  let m = text.match(MB_GIAN_RE_AFTER_MONTHTOKEN);
  if (m) {
    const g = mbCleanGianCandidate(m[1]);
    if (g) return g;
  }
  m = text.match(MB_GIAN_RE_AFTER_THUE);
  if (m) {
    const g = mbCleanGianCandidate(m[1]);
    if (g) return g;
  }
  m = text.match(MB_GIAN_RE_AFTER_LOAICHIPHI);
  return mbCleanGianCandidate(m ? m[1] : "");
}

function mbParseDateCell(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return excelSerialToIso(v);
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})/);
  if (!m) return null;
  let yr = Number(m[3]);
  if (yr < 100) yr += 2000;
  return `${yr}-${pad2(Number(m[2]))}-${pad2(Number(m[1]))}`;
}

// Tra ve { sheetName, rows, unclassifiedCount, skippedNoAmountOrDate } -- 1
// tab duy nhat khop chu ky cot (dung tab dau tien khop, bo qua cac tab khac
// nhu "JP HN"/"đối trừ goldtrans"/"site code" vi KHONG khop chu ky nay).
function parseKvcMienBacWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  let sheetName = null;
  let rows = [];
  let unclassifiedCount = 0;

  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const found = findMbHeaderRow(grid);
    if (!found) continue;
    sheetName = sn;
    for (let r = found.headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const soTienRaw = row[found.soTien];
      const soTien = typeof soTienRaw === "number" ? soTienRaw : NaN;
      if (!soTien || isNaN(soTien) || soTien <= 0) continue; // dong trong/tieu de phu/tong nhom

      const ngay =
        (found.ngayTaoLenh !== undefined ? mbParseDateCell(row[found.ngayTaoLenh]) : null) ||
        (found.ngayXinPm !== undefined ? mbParseDateCell(row[found.ngayXinPm]) : null);
      if (!ngay) continue;

      const deXuatText = found.deXuat !== undefined ? cellText(row[found.deXuat]) : "";
      const uncText = found.unc !== undefined ? cellText(row[found.unc]) : "";
      const dienGiai = deXuatText || uncText;
      const ncc = found.thuHuong !== undefined ? cellText(row[found.thuHuong]) : "";
      const gianRaw = found.gianHang !== undefined ? cellText(row[found.gianHang]) : "";
      const gian = gianRaw || extractGianChoTai(dienGiai);
      const congTyNote = found.ghiChuCongTy !== undefined ? row[found.ghiChuCongTy] : "";
      const congTy = normCongTy(congTyNote);
      if (!congTy) {
        unclassifiedCount++;
        continue; // khong doan mo hinh KH Cu/Moi -- bo qua, khong nhap sai cong ty
      }
      const soHoaDon = extractSoHoaDon(uncText) || extractSoHoaDon(deXuatText);

      rows.push({ congTy, ngay, gian, ncc, soHoaDon, soUNC: "", dienGiai, soTien });
    }
    break; // chi lay tab DAU TIEN khop chu ky (Luyen xac nhan chi co 1 tab lien quan)
  }

  return { sheetName, rows, unclassifiedCount };
}

// ---------- Parser cho file "MÁY TỰ ĐỘNG" (TẠO LỆNH UNC MTĐ MB 2025.xlsx) ----------
// Luyen, 2026-07-23: nguon THU 2 cho Chi Phi Mien Bac (bo sung file UNC KVC MB
// thu cong o tren) -- "đây là của máy tự động" -- tuong duong file MTĐ MN cua
// Mien Nam ve VAI TRO (nguon xuat tu dong), nhung cau truc cot HOAN TOAN khac
// (khong tach theo "Tháng N.YYYY", 1 sheet du lieu lien tuc ten "MTĐ" + 1 sheet
// tra cuu "site code" rieng -- sheet site code CHUA dung toi, chi de tham
// khao). Cot: "Ngày xin payment"/"Ngày đi tiền" (uu tien cot dau), "Nội dung
// unc" (dien giai + nguon trich gian), "Số tiền" (dang CHUOI co dau phay vd
// "3,465,000" -- dung parseAmount, KHONG phai so thuan nhu file UNC KVC MB),
// "Tên công ty" (NCC). Cong ty (KH Cu/Moi) CUNG nam trong 1 cot ghi chu KHONG
// CO TIEU DE (chi co khoang trang) ngay sau cot "Ngân hàng" -- vi la file XUAT
// TU DONG (khong phai Luyen tu go tay), vi tri cot on dinh hon nhieu so voi
// file thu cong, nen xac dinh QUA VI TRI (ngay sau cot "Ngân hàng" tim duoc)
// thay vi qua ten cot (khong the tim theo ten vi khong co tieu de).
function findMtdAutoHeaderRow(grid) {
  for (let r = 0; r < Math.min(grid.length, 8); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((cell, c) => {
      if (cell === null || cell === undefined || typeof cell !== "string") return;
      const h = mbNormHeader(cell);
      if (idx.ngayXinPayment === undefined && h.includes("ngay xin payment")) idx.ngayXinPayment = c;
      if (idx.ngayDiTien === undefined && h.includes("ngay di tien")) idx.ngayDiTien = c;
      if (idx.noiDungUnc === undefined && h.includes("noi dung unc")) idx.noiDungUnc = c;
      if (idx.noiDungPm === undefined && h.includes("noi dung pm")) idx.noiDungPm = c;
      if (idx.soTien === undefined && h === "so tien") idx.soTien = c;
      if (idx.tenCongTy === undefined && h.includes("ten cong ty")) idx.tenCongTy = c;
      if (idx.nganHang === undefined && h === "ngan hang") idx.nganHang = c;
    });
    if (idx.noiDungUnc !== undefined && idx.soTien !== undefined && idx.tenCongTy !== undefined && idx.nganHang !== undefined) {
      // Cot ghi chu KH Cu/Moi nam NGAY SAU cot "Ngân hàng" (da kiem chung tren
      // file that Luyen gui, 2026-07-23) -- khong co tieu de rieng nen phai
      // suy ra qua vi tri nay.
      idx.ghiChuCongTy = idx.nganHang + 1;
      return { headerRowIdx: r, ...idx };
    }
  }
  return null;
}

function parseKvcMienBacAutoWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  let sheetName = null;
  let rows = [];
  let unclassifiedCount = 0;

  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const found = findMtdAutoHeaderRow(grid);
    if (!found) continue;
    sheetName = sn;
    for (let r = found.headerRowIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const soTienRaw = row[found.soTien];
      const soTien = typeof soTienRaw === "number" ? soTienRaw : parseAmount(soTienRaw);
      if (!soTien || isNaN(soTien) || soTien <= 0) continue; // dong trong/tieu de phu/tong nhom

      const ngay =
        mbParseDateCell(row[found.ngayXinPayment]) || mbParseDateCell(row[found.ngayDiTien]);
      if (!ngay) continue;

      const uncText = cellText(row[found.noiDungUnc]);
      const pmText = found.noiDungPm !== undefined ? cellText(row[found.noiDungPm]) : "";
      const dienGiai = uncText || pmText;
      const ncc = cellText(row[found.tenCongTy]);
      const gian = extractGianThueMayTuDong(dienGiai);
      const congTyNote = row[found.ghiChuCongTy];
      const congTy = normCongTy(congTyNote);
      if (!congTy) {
        unclassifiedCount++;
        continue; // khong doan mo hinh KH Cu/Moi -- bo qua, khong nhap sai cong ty
      }
      const soHoaDon = extractSoHoaDon(uncText) || extractSoHoaDon(pmText);

      rows.push({ congTy, ngay, gian, ncc, soHoaDon, soUNC: "", dienGiai, soTien });
    }
    break; // chi lay sheet DAU TIEN khop chu ky (sheet "MTĐ" -- sheet "site code" khong khop nen tu bo qua)
  }

  return { sheetName, rows, unclassifiedCount };
}

module.exports = {
  parseChiPhiSheetWorkbook,
  parseKvcMienBacWorkbook,
  parseKvcMienBacAutoWorkbook,
  extractDate,
  extractSoHoaDon,
  normCongTy,
};
