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
// Chi Nhan (2026-07-28): "cac dien giai ngoai viet qr a bo ra khoi doi soat
// nha 11521268 a" -- vai giao dich KHONG phai tien khach tra qua VietQR van
// lot vao so du "thu" cua tai khoan (vd lai tien gui ngan hang tra hang
// thang) va bi tinh nham vao "Ngan hang" cua ngay do. Dung them 1 danh sach
// LOAI TRU theo dien giai (giong huong MOMO_TX_PATTERN o tren, KHONG dung
// "phai co token X" vi mot so giao dich QR THAT lai khong co token dac
// trung -- xem ghi chu ngay tren). Sau khi normText (bo dau, thuong hoa) de
// khop du dien giai co dau hay khong.
// Chi Nhan, 2026-07-29: "hải phòng làm gì cộng dữ liệu nhiều vậy check lại
// xem có bị trùng hk 77021" -- dieu tra ra KHONG phai trung du lieu, ma la
// giao dich "CTNB" (Chuyen Tien Noi Bo -- giua CHINH cac tai khoan cua cong
// ty, vd "CTNB BIDV701-681") lap lai NHIEU LAN (kiem tra 2026-07-29: >15 lan
// tren rieng tai khoan BIDV77021, ngay 23/7 va 27-28/7 la 3 vi du da phat
// hien/sua tay qua excludeFromVietQrRecon truoc khi tim ra day la 1 PATTERN
// LAP LAI chu khong phai 1-2 truong hop don le) -- tien NAY khong phai khach
// tra qua VietQR nen KHONG duoc tinh vao doanh thu QR, loai TU DONG qua day
// (giong huong "tra lai tien gui") thay vi phai tu tay tick "Loai khoi doi
// soat VietQR" cho tung dong 1 moi lan phat sinh ve sau.
const NON_VQR_TX_PATTERN = /tra lai tien gui|thanh toan lai|lai nhap von|lai tien gui|ctnb/i;
// Tach rieng buoc loc (dung chung cho ca extractVietQrSettlements o duoi VA
// resolveGianGrossByBankRef, xem ghi chu tai do) khoi buoc gop theo ngay.
function extractVietQrThuTransactions(transactions) {
  return transactions.filter(
    (t) =>
      t.type === "thu" &&
      !t.excludeFromVietQrRecon &&
      t.date &&
      !MOMO_TX_PATTERN.test(t.description || "") &&
      !NON_VQR_TX_PATTERN.test(normText(t.description || ""))
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
        // Chi Nhan, 2026-07-27: "lấy giao dịch viet qr đối soát thôi đừng
        // giao dịch nào cũng lấy" -- file xuat tu doitac.vietqr.vn co rieng
        // cot "Loai giao dich" (KHAC voi cot "Loai" = Giao dich den/di o
        // tren) voi gia tri "QR giao dich" cho giao dich VietQR that, nhung
        // cung co the la "Vang lai" (giao dich khac, khong phai QR -- vd
        // chuyen khoan ca nhan/noi dung khong lien quan don hang, khong co
        // Ma cua hang) -- truoc gio cot nay bi bo qua hoan toan nen "Vang
        // lai" van duoc tinh vao doanh thu, gay du "Tinh tu du lieu tai len"
        // so voi "Ngan hang" that. Tu gio chi giu dong "QR giao dich".
        if (idx.loaiGiaoDich === undefined && s.includes("loai giao dich")) idx.loaiGiaoDich = c;
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
    // Chi Nhan, 2026-07-27: chi giu giao dich THAT SU la VietQR ("QR giao
    // dich") -- loai "Vang lai" (chuyen khoan/noi dung khac, khong qua QR,
    // thuong khong co Ma cua hang that -- vd Noi dung TT "massage", Ma cua
    // hang "-") van la tien that vao tai khoan nhung KHONG phai doanh thu
    // ban hang qua QR, phai loai khoi doi soat VietQR.
    const loaiGiaoDich = cols.loaiGiaoDich !== undefined ? row[cols.loaiGiaoDich] : null;
    if (loaiGiaoDich !== null && loaiGiaoDich !== undefined && !normText(String(loaiGiaoDich)).includes("qr")) continue;
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

// ---------- VNPay portal transaction export ("DanhSachGiaoDich...xlsx") ----------
// Chi Nhan, 2026-07-29: "thêm cho tôi chỗ đối soát vn pay kh mới ... đối
// soát với ngân hàng 02865168 bên Kh mới á trả về tiền theo ngày" -- kenh
// MOI cho KH Moi (vd cac may quet QR gan nhan "TUTU TRAIN ...") KHONG dung
// dinh dang "Transactions" chuan (cot "Noi dung TT"/"Ma cua hang"/token VQR)
// nhu 4 kenh VietQR kia, ma la file XUAT TRUC TIEP tu portal VNPay ("BAO CAO
// CHI TIET GIAO DICH", sheet "Sheet1"): STT | Thoi gian GD | Ma giao dich |
// Chi nhanh | Ma diem thu | Diem thu | So hoa don | ... | So tien truoc KM |
// So tien sau KM | ... | Trang thai | ... -- moi giao dich la 1 khoan tien
// VE THANG ngan hang 02865168 (giong VietQR, khong phai settlement gop nhieu
// ngay nhu VNPay Offline/Payoo cua KH Cu), nen tai su dung THANG toan bo may
// doi soat VietQR (resolveGianGross/reconcileVietQr) cho kenh nay -- chi can
// 1 parser rieng doc dung dinh dang nay, tra ve CUNG 1 shape voi
// parseVietQrRawWorkbook (rows: {vqrCode, maCuaHang, amount, date, raw}) de
// resolveGianGross/mergeRawRows dung lai duoc khong sua gi them. "Diem thu"
// da la ten mo ta day du san (khong can sheet "Cua hang" rieng nhu cac kenh
// kia), nen tra luon storeMap cung 1 lan doc.
function parseVnpayPortalWorkbook(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName =
    wbLite.SheetNames.find((n) => normText(n) === "sheet1") ||
    wbLite.SheetNames.find((n) => !/config/i.test(n)) ||
    wbLite.SheetNames[0];
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  let headerRowIdx = -1;
  let cols = {};
  for (let r = 0; r < Math.min(grid.length, 8); r++) {
    const row = grid[r] || [];
    const idx = {};
    row.forEach((v, c) => {
      if (!v || typeof v !== "string") return;
      const s = normText(v);
      if (idx.dateCol === undefined && s.includes("thoi gian gd")) idx.dateCol = c;
      if (idx.maDiemCol === undefined && s.includes("ma diem thu")) idx.maDiemCol = c;
      if (idx.diemCol === undefined && s === "diem thu") idx.diemCol = c;
      if (idx.amountCol === undefined && s.includes("so tien sau km")) idx.amountCol = c;
      if (idx.refCol === undefined && s.includes("ma giao dich")) idx.refCol = c;
      // Khop tuyet doi "trang thai" -- tranh nham voi "Trang thai tra gop".
      if (idx.statusCol === undefined && s === "trang thai") idx.statusCol = c;
    });
    if (idx.dateCol !== undefined && idx.diemCol !== undefined && idx.amountCol !== undefined) {
      headerRowIdx = r;
      cols = idx;
      break;
    }
  }
  if (headerRowIdx < 0) {
    throw new Error('Khong doc duoc dong tieu de (can cot "Thoi gian GD", "Diem thu", "So tien sau KM") trong file bao cao giao dich VNPay.');
  }

  const rows = [];
  const storeMap = {};
  for (let r = headerRowIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    // Cot "Trang thai" luon trong tren du lieu thuc te (khong co GD that bai
    // nao trong file mau) -- chi loai khi CO gia tri VA gia tri do KHONG phai
    // thanh cong, giong quy uoc cua parseOfflineVnpayWorkbook/parseVietQrRawWorkbook.
    const status = cols.statusCol !== undefined ? row[cols.statusCol] : null;
    if (status !== null && status !== undefined && String(status).trim() !== "" && !/thanh cong/i.test(normText(String(status)))) {
      continue;
    }
    const diem = cols.diemCol !== undefined ? String(row[cols.diemCol] || "").trim() : "";
    if (!diem) continue;
    const maDiem = cols.maDiemCol !== undefined ? String(row[cols.maDiemCol] || "").trim() : diem;
    const amountRaw = cols.amountCol !== undefined ? row[cols.amountCol] : null;
    const amount = amountRaw === null || amountRaw === undefined ? 0 : Number(String(amountRaw).replace(/,/g, "")) || 0;
    if (!amount) continue;
    const dateRaw = cols.dateCol !== undefined ? row[cols.dateCol] : null;
    const date = parseVqrDate(dateRaw);
    if (!date) continue;
    const ref = cols.refCol !== undefined ? String(row[cols.refCol] || "").trim() : "";
    rows.push({ vqrCode: ref || null, maCuaHang: maDiem, amount, date, raw: diem });
    if (!storeMap[maDiem]) {
      storeMap[maDiem] = { tenCuaHang: diem, tenDiemBan: diem, matchText: diem };
    }
  }

  return { sheetName, rows, storeMap };
}

// ---------- VNPay KH Moi (02865168): di hoa don ra khoi pool ZVP KH Cu ----------
// Chi Nhan, 2026-07-29: "chuyển hẳn sang KH Mới" (xac nhan qua AskUserQuestion)
// -- danh sach CO DINH cac "Ma diem misa" tren sheet hoa don MTT dung chung
// thuc ra la doanh thu cua kenh VNPay KH Moi (tai khoan 02865168), khong
// phai KH Cu (tien that KHONG ve tai khoan ACB31268 cua ZVP) -- day chinh la
// nguyen nhan "KVC TIMES" bao lech hoai ben doi soat ZVP KH Cu truoc gio. Map
// thang qua dung "Ma cong trinh" chuan (co the KHAC voi ma tren hoa don, vd
// hoa don ghi "KVC AE HUE" nhung ma cong trinh chuan la "AE HUE KVCN", theo
// bang chi Nhan cung cap 2026-07-29).
const VNPAY_KHMOI_INVOICE_MADIEM_MAP = {
  "KVC AE HUE": "AE HUE KVCN",
  "KVC TIMES": "KVC TIMES",
  "KVC ROYAL": "KVC ROYAL",
  "SAVICO PHN": "SAVICO KVCN",
  // Nhan, 2026-08-06: "payoo kh mới từ ngày 1 không còn có 1 điểm nữa mà có
  // rất nhiều điểm" -- them "FARM LOTTE BAC GIANG" (gian Payoo moi tu 01/08,
  // xem PAYOO_STORE_TO_MA_CONG_TRINH trong routes/doisoat-vnpay-khmoi.js) vao
  // day de hoa don cua gian nay cung duoc chuyen ra khoi pool ZVP (KH Cu) va
  // gop vao store.viet_qr_invoices.vnpayKhMoi giong 3 gian cu, neu khong hoa
  // don se ket ket qua trong pool KH Cu, khong bao gio khop duoc voi doanh
  // thu Payoo KH Moi cua gian nay.
  "FARM LOTTE BAC GIANG": "FARM LOTTE BAC GIANG",
  // Nhan, 2026-08-06: chi nhanh "FARM SAVICO" (VNPay offline) thuc ra la
  // "Pinball và ghế LOTTE BAC GIANG" (xem TEN_DIEM_TO_MA_CONG_TRINH trong
  // routes/doisoat-vnpay-khmoi.js) -- them de hoa don cung ten duoc chuyen ra
  // khoi pool ZVP (KH Cu) va gop vao store.viet_qr_invoices.vnpayKhMoi.
  "PINBALL VÀ GHẾ LOTTE BAC GIANG": "PINBALL VÀ GHẾ LOTTE BAC GIANG",
  // Luyen, 2026-08-08: gian "KVC ROYAL" doi ten thanh "PINBALL DA NANG" tu
  // 31/07 -- hoa don moi se ghi maDiem "PINBALL DA NANG" (khac "KVC ROYAL" cu),
  // can them vao day de migrate sang vnpayKhMoi pool dung.
  "PINBALL DA NANG": "PINBALL DA NANG",
};

// Tu dong chuyen (KHONG chi loc-khi-doc) cac hoa don co maDiem nam trong
// VNPAY_KHMOI_INVOICE_MADIEM_MAP ra khoi 3 pool hoa don ZVP cua KH Cu
// (store.zvp_invoices.zalo/vnpay/payoo) va gop vao store.viet_qr_invoices.vnpayKhMoi
// (da doi maDiem sang dung ma cong trinh chuan). Goi lai MOI LAN load() (ca
// tu trang VietQR lan trang ZVP -- xem loi goi o ca 2 route) de tu "don" moi
// lan chi Nhan tai them hoa don MTT moi (van tiep tuc duoc gan tag "Vnpay CS
// MB"/"Payoo ... ngân hàng"/"ZALO MINI APP" nhu cu, chua co tag rieng cho
// kenh nay). Dat trong utils (khong phai routes/doisoat-vietqr.js) de ca 2
// route file dung chung duoc ham nay ma khong phai export qua router.
// Chi Nhan, 2026-07-29: "sao lại cộng hóa đơn ... vậy lấy ra đi kh phải của
// VNPAY offline á" (phat hien khi xay trang /doi-soat/vnpay-khmoi) -- "KVC
// ROYAL" bi TRUNG TEN giua 2 nguon doanh thu HOAN TOAN KHAC nhau: (1) doanh
// thu Zalo Mini App (Online) that su cua KH Cu (ve TK ACB31268, khong lien
// quan VNPay/Payoo/VTB982), va (2) doanh thu VNPay/Payoo cua "TUTU TRAIN VC
// ROYAL"/"NHA TUYET VC ROYAL" (KH Moi, ve TK VTB982) -- ca 2 deu duoc nhan
// vien xuat hoa don ghi CUNG 1 chu "KVC ROYAL" nen migrate truoc day (quet ca
// 3 pool zalo/vnpay/payoo theo maDiem) da gom NHAM 44 hoa don Zalo (tag "ZALO
// MINI APP...") vao chung voi 47 hoa don Vnpay/Payoo that, lam sai ca doi
// soat Online cua KH Cu (thieu hoa don) LAN doi soat VNPay KH Moi (thua hoa
// don, "Lệch" gia). Fix: CHI di chuyen tu 2 pool "vnpay"/"payoo" (dung dich
// vu thu ho MTT ghi "Vnpay CS MB"/"Payoo ... ngân hàng"), KHONG dong den pool
// "zalo" nua -- Online/Zalo Mini App luon la doanh thu ACB31268 that, du co
// trung ten voi 1 gian VNPay KH Moi khac.
const ZALO_MAI_TAG_PATTERN = /zalo/i;

// Nhan, 2026-08-06: "hóa đơn ngày 4 5 của kvc hue đây á" -- hoa don maDiem
// "KVC AE HUE" cua Payoo van hien "Chưa có HĐ" du gross Payoo da khop dung
// (FARM LOTTE BAC GIANG/KVC TIMES/KVC ROYAL deu khop). Nguyen nhan:
// VNPAY_KHMOI_INVOICE_MADIEM_MAP map "KVC AE HUE" -> "AE HUE KVCN" (dung cho
// kenh VNPay OFFLINE cu, ma gross that cua kenh do la "AE HUE KVCN" theo
// TEN_DIEM_TO_MA_CONG_TRINH ben routes/doisoat-vnpay-khmoi.js), nhung Payoo
// (gian moi tu 01/08) lai dung THANG ten "KVC AE HUE" ben gross (xem
// PAYOO_STORE_TO_MA_CONG_TRINH) -- CUNG 1 maDiem hoa don "KVC AE HUE" can ra
// 2 ma DICH KHAC NHAU tuy thuoc hoa don do la Payoo hay Vnpay offline (phan
// biet qua tag "raw"). Rieng cho pool "payoo", GIU NGUYEN "KVC AE HUE" (khong
// doi qua "AE HUE KVCN").
const VNPAY_KHMOI_INVOICE_MADIEM_MAP_PAYOO_OVERRIDE = {
  "KVC AE HUE": "KVC AE HUE",
};

function migrateVnpayKhMoiInvoices(store) {
  if (!store.zvp_invoices) return false;
  if (!store.viet_qr_invoices) store.viet_qr_invoices = {};
  if (!store.viet_qr_invoices.vnpayKhMoi) store.viet_qr_invoices.vnpayKhMoi = [];
  let changed = false;
  const target = store.viet_qr_invoices.vnpayKhMoi;
  const existingKeys = new Set(target.map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
  ["vnpay", "payoo"].forEach((key) => {
    if (!store.zvp_invoices[key]) return;
    const mapForKey =
      key === "payoo"
        ? Object.assign({}, VNPAY_KHMOI_INVOICE_MADIEM_MAP, VNPAY_KHMOI_INVOICE_MADIEM_MAP_PAYOO_OVERRIDE)
        : VNPAY_KHMOI_INVOICE_MADIEM_MAP;
    const kept = [];
    store.zvp_invoices[key].forEach((inv) => {
      const mapped = mapForKey[inv.maDiem];
      if (!mapped) {
        kept.push(inv);
        return;
      }
      changed = true;
      const moved = { ...inv, maDiem: mapped };
      const k = `${moved.soHd}|${moved.ngayHd}|${moved.maDiem}`;
      if (!existingKeys.has(k)) {
        existingKeys.add(k);
        target.push(moved);
      }
    });
    if (kept.length !== store.zvp_invoices[key].length) {
      store.zvp_invoices[key] = kept;
    }
  });
  // Sua nguoc: hoa don Zalo bi gom nham TU TRUOC (luc migration con quet ca
  // pool "zalo") van con nam trong target -- tra ve lai zvp_invoices.zalo,
  // giu nguyen maDiem GOC (truoc khi bi doi qua VNPAY_KHMOI_INVOICE_MADIEM_MAP,
  // vd "KVC ROYAL" giu nguyen vi map la identity, nhung an toan cho ca truong
  // hop map KHAC ten sau nay).
  const stillZaloTagged = target.filter((inv) => ZALO_MAI_TAG_PATTERN.test(inv.raw || ""));
  if (stillZaloTagged.length > 0) {
    if (!store.zvp_invoices.zalo) store.zvp_invoices.zalo = [];
    const zaloKeys = new Set(store.zvp_invoices.zalo.map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
    stillZaloTagged.forEach((inv) => {
      // maDiem tren hoa don Zalo goc luon la ten SITE THAT (vd "KVC ROYAL"),
      // khac voi ma cong trinh chuan cua nguon VNPay KH Moi trong truong hop
      // sau nay co map doi ten khac nhau -- vi hien tai ca 4 map deu ve dung
      // ten gian that (identity hoac ve dung ma chuan), doi nguoc dua tren
      // reverse-lookup cua VNPAY_KHMOI_INVOICE_MADIEM_MAP.
      const originalMaDiem =
        Object.keys(VNPAY_KHMOI_INVOICE_MADIEM_MAP).find((k) => VNPAY_KHMOI_INVOICE_MADIEM_MAP[k] === inv.maDiem) || inv.maDiem;
      const restored = { ...inv, maDiem: originalMaDiem };
      const k = `${restored.soHd}|${restored.ngayHd}|${restored.maDiem}`;
      if (!zaloKeys.has(k)) {
        zaloKeys.add(k);
        store.zvp_invoices.zalo.push(restored);
      }
      changed = true;
    });
    store.viet_qr_invoices.vnpayKhMoi = target.filter((inv) => !ZALO_MAI_TAG_PATTERN.test(inv.raw || ""));
  }
  return changed;
}

// ---------- Cac tai khoan VietQR MOI dung CHUNG may quet QR/Ma cua hang voi
// BIDV77021 (San bay Cam Ranh / Vinpearl Nha Trang...) ----------
// Chi Nhan, 2026-07-30: "thêm 1 ngân hàng 02865168 á các điểm bán với mã
// công trình map với 77021 á" -- xac nhan 96/97 "Ma cua hang" tren sao ke tai
// khoan MB02865168 TRUNG KHOP TUYET DOI voi Ma cua hang da co san trong
// store.viet_qr_store_names.bidv77021 (cung 1 he thong may POS/QR, chi khac
// tai khoan nhan tien) -- thay vi bat cho tai file "Danh sach diem ban" rieng
// cho tai khoan nay (Luyen se phai tim/tai lai tu cong QR, trong khi du lieu
// giong het da co san), tu dong dong bo (merge, uu tien ban ghi cua bidv77021
// khi trung ma) MOI LAN load() -- viet_qr_ten_diem_master cung dong bo tuong
// tu, vi cung 1 danh sach "Ten diem -> Ma cong trinh".
// Chi Nhan, 2026-07-30 (lan 3): "đây là viet qr kh mới thêm cho tôi vô ...
// thêm 1 ngân hàng 8613600999" -- tai khoan BIDV8613600999 xac nhan CUNG he
// thong nay tiep (45/46 Ma cua hang tren sao ke trung khop voi bidv77021),
// nen tong quat hoa ham nay de ap dung cho NHIEU kenh dung chung nguon (thay
// vi 1 ham rieng cho tung kenh moi them vao).
const BIDV77021_SEED_TARGETS = ["mb02865168", "bidv8613600999"];

function seedFromBidv77021(store, targetChannels) {
  if (!store.viet_qr_store_names || !store.viet_qr_store_names.bidv77021) return false;
  if (!store.viet_qr_ten_diem_master) store.viet_qr_ten_diem_master = {};
  const targets = targetChannels || BIDV77021_SEED_TARGETS;
  let changed = false;
  targets.forEach((ch) => {
    if (!store.viet_qr_store_names[ch]) store.viet_qr_store_names[ch] = {};
    if (!store.viet_qr_ten_diem_master[ch]) store.viet_qr_ten_diem_master[ch] = {};
    Object.entries(store.viet_qr_store_names.bidv77021).forEach(([code, info]) => {
      const cur = store.viet_qr_store_names[ch][code];
      if (!cur || JSON.stringify(cur) !== JSON.stringify(info)) {
        store.viet_qr_store_names[ch][code] = info;
        changed = true;
      }
    });
    if (store.viet_qr_ten_diem_master.bidv77021) {
      Object.entries(store.viet_qr_ten_diem_master.bidv77021).forEach(([k, v]) => {
        if (store.viet_qr_ten_diem_master[ch][k] !== v) {
          store.viet_qr_ten_diem_master[ch][k] = v;
          changed = true;
        }
      });
    }
  });
  return changed;
}

// Giu lai ten cu (goi qua ham chung, chi 1 kenh) de khong phai doi cho goi o
// routes/doisoat-vietqr.js.
function seedMb02865168FromBidv77021(store) {
  return seedFromBidv77021(store, BIDV77021_SEED_TARGETS);
}

// Chi Nhan, 2026-07-30: "KVC ROYAL có lệch đâu đây..." kieu loi tuong tu --
// truoc khi sua tagPattern cua CHANNELS.bidv77021 (them negative lookahead
// loai "tk 168"), hoa don cua MB02865168 ("VietQR POSH MB tk 168 X") da bi
// gom NHAM vao store.viet_qr_invoices.bidv77021 tu cac lan tai file MTT
// truoc do. Tu dong don (khong chi loc-khi-doc) cac hoa don "tk 168" con sot
// lai trong pool bidv77021 sang dung pool mb02865168 moi lan load(), giong
// het co che ZALO_MAI_TAG_PATTERN o tren.
const TK168_TAG_PATTERN = /tk\s*168/i;

function migrateTk168Invoices(store) {
  if (!store.viet_qr_invoices || !store.viet_qr_invoices.bidv77021) return false;
  if (!store.viet_qr_invoices.mb02865168) store.viet_qr_invoices.mb02865168 = [];
  const straggler = store.viet_qr_invoices.bidv77021.filter((inv) => TK168_TAG_PATTERN.test(inv.raw || ""));
  if (straggler.length === 0) return false;
  const target = store.viet_qr_invoices.mb02865168;
  const existingKeys = new Set(target.map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
  straggler.forEach((inv) => {
    const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
    if (!existingKeys.has(k)) {
      existingKeys.add(k);
      target.push(inv);
    }
  });
  store.viet_qr_invoices.bidv77021 = store.viet_qr_invoices.bidv77021.filter(
    (inv) => !TK168_TAG_PATTERN.test(inv.raw || "")
  );
  return true;
}

// Chi Nhan, 2026-07-30: "7702 tôi để nhầm tên á này của ngày 3 á ... trừ cái
// của ngày 4 dó dem qua ngày 3 và nó xuất chung á cấn trừ hóa đơn ngày 4.5
// nhá" -- 29 hoa don dan tag "MTD MN 4" (rieng 1 ngay, KHONG phai "MTD MN
// 4,5") thuc ra la doanh thu ngay 3 (xac nhan: tong 29 hoa don nay = dung
// 20.000.000d = so tien ngan hang chuyen ve ngay 3; hoa don "MTD MN 4,5" moi
// dung la ngay 4-5 gop, tong = dung 33.880.000+37.660.000d cua 2 ngay do).
// Sua 1 lan qua script truc tiep vao store.json bi MAT (server dang chay cua
// Chi Nhan luu de (save()) tu ban nho cu, ghi de lai ban chua sua) -- chuyen
// thanh migration tu dong chay lai MOI LAN load() de KHONG BAO GIO mat lai
// du server co ghi de bao nhieu lan nua.
// Luyen, 2026-08-03: doi chieu voi file hoa don goc "MỚi XHD.xlsx" xac nhan
// ham nay (blanket "bat ky hoa don nao raw === 'MTD MN 4' deu la ngay 3")
// dang CHUYEN NHAM 29 hoa don thang 6 + 20 hoa don thang 5 dan tag DUNG la
// ngay 4 (khong lien quan gi den lo hoa don CU ma Chi Nhan xac nhan
// 2026-07-30) -- chinh la nguyen nhan Luyen bao "hóa đơn ngày 3,4" (ngay 3
// du, ngay 4 thieu dung so tien bi chuyen nham, "Chưa có HĐ" het). File hoa
// don hien tai KHONG CON hoa don thang 7 nao dan tag "MTD MN 4" nua (lo cu da
// duoc sua truc tiep trong file goc), nen ham nay VO HIEU HOA hoan toan (tra
// ve false, khong chuyen doi gi nua) de khong tiep tuc lam hong du lieu hoa
// don MOI. Xem restoreBidv7702JuneDay4Invoices trong store.js de biet cach
// khoi phuc lai cac hoa don thang 6 da bi chuyen nham truoc do.
function fixBidv7702Day3TaggedAsDay4(store) {
  return false;
}

// Chi Nhan, 2026-07-30: 07-03/07-04 cua BIDV8613600999 "2 ngay nay cấn trừ
// nhau á cấn trừ cho tôi đi để khớp" -- dieu tra thay day KHONG PHAI cap
// tru chuoi ngay (nhu applySharePairChainNetting) ma la CUNG 1 loi tag hoa
// don nhu fixBidv7702Day3TaggedAsDay4 o tren: 6 hoa don so 8261-8266 dan tag
// rieng "QR JP 4" (raw dung 1 ngay, KHONG phai "QR JP 4,5") nhung tong tien
// tung hoa don khop CHINH XAC voi doanh thu ngay 3 cua ngan hang (vd hoa don
// 8265 "JP VC BA TRIEU" = 450.000d = dung so ngan hang chuyen ngay 3), trong
// khi hoa don "QR JP 4,5" cung nhom (8402-8408) da dung va da duoc chia ty
// le dung cho ngay 4/5 boi applyMultiDayGroupConsolidation. Vay chi can sua
// rieng 6 hoa don "QR JP 4" nay thanh ngay 3, KHONG dung toi hoa don "4,5".
const BIDV8613600999_DAY3_TAGGED_AS_DAY4 = [8261, 8262, 8263, 8264, 8265, 8266];
function fixBidv8613600999Day3TaggedAsDay4(store) {
  if (!store.viet_qr_invoices || !store.viet_qr_invoices.bidv8613600999) return false;
  let changed = false;
  store.viet_qr_invoices.bidv8613600999.forEach((inv) => {
    if (BIDV8613600999_DAY3_TAGGED_AS_DAY4.includes(inv.soHd) && inv.raw === "QR JP 4") {
      inv.days = [3];
      inv.raw = "QR JP 3";
      changed = true;
    }
  });
  return changed;
}

// Chi Nhan, 2026-07-30: BIDV77021 07-03/07-04 "chỗ này bị nhầm của ghi lộn
// ngày 3 thành ngày 4 chưa giải quyết cho tôi à" -- cung loi giong het 2 ham
// tren, nhung lan nay anh huong CA LOAT gian (85 hoa don) chu khong phai vai
// gian le: MOI hoa don co raw dung "VietQR POSH MB 4" (1 ngay, KHONG phai
// "VietQR POSH MB 4,5") thuc ra la doanh thu ngay 3 -- xac nhan tong 85 hoa
// don loai nay = dung 57.260.000d = tong tien ngan hang ngay 3 cua CA kenh;
// hoa don "VietQR POSH MB 4,5" (gop dung ngay 4-5) van dung, khong dung toi.
// Nhan, 2026-08-06: ham nay (blanket "bat ky hoa don nao raw === 'VietQR POSH
// MB 4' deu la ngay 3") dang CHUYEN NHAM 93 hoa don thang 8 (ngayHd 2026-08-05,
// tong 46.320.000d) dan tag DUNG la ngay 4 -- khong lien quan gi den lo hoa
// don CU (85 hoa don, 57.260.000d, Chi Nhan xac nhan 2026-07-30) ma ham nay
// duoc tao ra de sua. Kiem tra lai: du lieu hien tai KHONG CON hoa don nao dan
// tag "VietQR POSH MB 4" thuoc lo cu (da duoc sua permanent thanh ngay 3 tu
// truoc, xem lich su); ham nay gio chi con bat nham cac hoa don MOI. Giong het
// fixBidv7702Day3TaggedAsDay4 o tren -- VO HIEU HOA hoan toan de khong tiep
// tuc lam hong du lieu hoa don moi.
function fixBidv77021Day3TaggedAsDay4(store) {
  return false;
}

// ---------- Store catalog ("Cua hang" / "Cua Hang") ----------
// Ma cua hang -> a descriptive name, built from BOTH "Ten cua hang" and
// "Ten diem ban" so the fuzzy matcher against invoices has more text to
// work with (neither column alone reliably matches the invoice list's own
// "Ten diem xuat hoa don" wording).
function parseCuaHangSheet(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  // Chi Nhan (2026-07-28, kenh BIDV77021 moi): mot so file dat ten sheet la
  // "Diem ban" thay vi "Cua hang" nhung cot ben trong giong het (Ten cua
  // hang/Ma cua hang/Ten diem ban) -- chap nhan ca 2 kieu ten sheet, viec
  // khop cot header van dam bao khong nham voi sheet "Ten diem - Ma cong
  // trinh" (sheet do khong co cot "ma cua hang").
  const sheetName = wbLite.SheetNames.find((n) => {
    const t = normText(n);
    return t.includes("cua hang") || t.includes("diem ban");
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
      // Luyen, 2026-07-27: thieu "idx.tongTt === undefined" (khac voi tat ca
      // cot khac o day) khien cot SAU CUNG trong dong header co chua ca "tong"
      // va "hd" (kieu khop long leo) DE LEN cot dung "Tong TT HD" -- verified
      // qua HD 10966 (07-25, AM TP PHCM): so tien that 14.980.000d nhung bi
      // doc thanh 6.745.026.882d, dung voi 1 cot tong hop/cong don khac cung
      // dong header vo tinh cung chua "tong" + "hd". Fix: chi lay cot DAU
      // TIEN khop (giong tat ca cac cot khac ben duoi).
      if (idx.tongTt === undefined && (s.includes("tong tt hd") || (s.includes("tong") && s.includes("hd")))) idx.tongTt = c;
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
    // Chi Nhan (2026-07-27): "không chia theo chia sẻ hay không chia sẻ nữa"
    // -- ngung gan hau to FF_SUFFIX theo "Hinh thuc hop tac" (CSE), khong con
    // tach dong CSE/khong-CSE nua.
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
    // Chi Nhan (2026-07-27): maDiem khong con bao gio bi gan hau to FF_SUFFIX
    // nua (khong con tach CSE/khong-CSE), nen dung thang inv.maDiem.
    const baseCode = stripCseSuffix(inv.maDiem);
    if (!byCode[baseCode]) {
      byCode[baseCode] = { tenDiem: inv.tenDiem, maCongTrinh: baseCode, isCse: false };
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
// Chi Nhan (2026-07-27): "các hóa đơn này map sai hết rồi" -- phat hien 1
// gian THAT (vd San bay Cam Ranh, kenh MB11521268) co the co NHIEU ten diem
// ban khac nhau tren cac may/thiet bi khac nhau (vd "POSH Sân bay Cam ranh",
// "1 JP SB Cam Ranh.new", "POSH Sân bay Quốc Tế Cam Ranh") -- ten nao du
// giong voi ten tren hoa don ("Sân bay Cam Ranh - Nhà ga quốc tế") thi fuzzy
// match dung ("CHKQT CAM RANH"), ten nao khac qua (thieu tu "quoc te" v.v.)
// thi KHONG khop duoc, bi tu tao thanh 1 "gian" rieng theo dung ten no (co
// che self-fallback), lam doanh thu 1 gian THAT bi xe le ra nhieu dong khac
// nhau -- hoa don chi khop voi 1 trong so do, con lai bi bao "Lech"/"thieu
// HD" gia tao. tenDiemOverride (dung lai chinh bang "Ten diem - Ma cong
// trinh" cua BIDV7702, key = normText(tenDiemBan)) cho phep gan THANG ten
// diem ban -> Ma cong trinh dung, bo qua fuzzy match hoan toan, khong can
// doi ten tren may POS that (van con nhieu ten khac nhau ve sau).
function resolveGianGross(rawRows, storeNameMap, gianCandidates, nocodeAssignments, defaultBlankCode, tenDiemOverride) {
  const matcher = buildOnlineProductMatcher(gianCandidates);
  const override = tenDiemOverride || {};
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
    const tenDiemBanKey = storeInfo && storeInfo.tenDiemBan ? normText(storeInfo.tenDiemBan) : "";
    const overrideCode = tenDiemBanKey ? override[tenDiemBanKey] : undefined;
    if (overrideCode) {
      codes.add(overrideCode);
      const key = `${row.date}|${overrideCode}`;
      grossByCode[key] = (grossByCode[key] || 0) + row.amount;
      continue;
    }
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
    // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong rieng.
    const code = match.maCongTrinh;
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
// Chi Nhan (2026-07-28, Luyen yeu cau cho kenh BIDV77021): "lệch tiền á nếu
// nó có diễn giải giống của viet qr nhưng nó không có tên điểm thì đưa vô
// hải phòng nhá" -- giao dich ngan hang da qua duoc bo loc "thu, khong phai
// lai/Momo" (extractVietQrThuTransactions) nen chac chan GIONG tien VietQR
// that, nhung neu KHONG khop duoc So tham chieu nao (unmatchedBankTx) hoac
// khop duoc dong QR nhung Ma cua hang lai chua co ten/diem nao (khong co
// "ten diem", unmappedStoreCodesAgg) thi gan THANG vao defaultBlankCode
// (truyen tu cfg.defaultBlankCode cua kenh, vd "AE HP PHN") thay vi de rieng
// trong 2 bang canh bao roi bi tru khoi "Ngan hang" (xem buildChannelReconciliation).
// KHONG ap dung cho unmappedTenDiemAgg (co ten diem that, chi la CHUA co
// trong bang tra ma cong trinh) -- truong hop do van can bao rieng vi co the
// la 1 diem MOI thuc su, khong nen am tham gap vao gian mac dinh.
function resolveGianGrossByBankRef(bankTxs, rawRows, storeNameMap, tenDiemToProjectMap, storeCodeOverride, refOverride, defaultBlankCode) {
  const override = storeCodeOverride || {};
  const refOv = refOverride || {};
  const refIndex = new Map(); // Ma tham chieu -> [rawRow, ...]
  // Luyen, 2026-08-11: index phu theo ma VQR (vqrCode) de xu ly truong hop
  // ngan hang gui tien qua cong trung gian (Liobank, SHB, Techcombank...) ma
  // he thong VietQR ghi refCode KHAC voi "So tham chieu" tren sao ke BIDV.
  // Nhung giao dich nay van co ma VQR trong dien giai ngan hang (vd
  // "VQR263075109OZ1J PaymentForOrder") trung voi truong vqrCode trong du lieu
  // QR -- dung do lam fallback truoc khi ve defaultBlankCode.
  const vqrIndex = new Map(); // vqrCode (upper) -> rawRow
  rawRows.forEach((row) => {
    const ref = (row.refCode || "").trim();
    if (!ref) return;
    if (!refIndex.has(ref)) refIndex.set(ref, []);
    refIndex.get(ref).push(row);
    const vc = (row.vqrCode || "").trim().toUpperCase();
    if (vc && !vqrIndex.has(vc)) vqrIndex.set(vc, row);
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
    // Luyen, 2026-08-11: fallback -- neu khong khop duoc bang refCode, thu
    // trich xuat ma VQR tu dien giai ngan hang va tra cuu lai. Ap dung cho
    // cac giao dich ngan hang cong trung gian (Liobank, SHB...) ma he thong
    // doi tac VietQR luu vqrCode trong du lieu QR nhung refCode la khac nhau.
    if (!raw) {
      const vqrInDesc = extractVqrCode(tx.description || "");
      if (vqrInDesc) raw = vqrIndex.get(vqrInDesc.toUpperCase()) || null;
    }
    if (!raw) {
      if (defaultBlankCode) {
        codes.add(defaultBlankCode);
        const key = `${tx.date}|${defaultBlankCode}`;
        grossByCode[key] = (grossByCode[key] || 0) + tx.amount;
        continue;
      }
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
      if (defaultBlankCode) {
        codes.add(defaultBlankCode);
        const key = `${tx.date}|${defaultBlankCode}`;
        grossByCode[key] = (grossByCode[key] || 0) + tx.amount;
        continue;
      }
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
    // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong rieng.
    const code = match.maCongTrinh;
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
    // Chi Nhan (2026-07-27): khong con tach CSE/khong-CSE thanh 2 dong rieng.
    const code = match.maCongTrinh;
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
// Chi Nhan (2026-07-27): "ngân hàng đối soát tất cả điều là 131 và không
// chia theo chia sẻ hay không chia sẻ nữa" -- thay vi doi theo tung noi tao
// ra hau to __FF (gan tren hoa don theo "Hinh thuc hop tac", ep buoc qua
// gianMerge isCse, hay qua resolveGianGross), chuan hoa (bo hau to) NGAY TAI
// DAY -- diem hoi tu duy nhat truoc khi gom nhom/hien thi -- de dam bao gop
// dung CSE + khong-CSE lam 1 du du lieu di qua duong nao.
function stripCseSuffix(code) {
  return code && code.endsWith(FF_SUFFIX) ? code.slice(0, -FF_SUFFIX.length) : code;
}

function reconcileVietQr(settlements, grossData, invoiceData, gianMapping, manualMatches, diemAlias) {
  const alias = diemAlias || {};
  const invoicesByDiemDay = {};
  for (const inv of invoiceData.invoices) {
    if (!inv.ngayHd || !inv.days || inv.days.length === 0) continue;
    const [invY, invMo, invD] = inv.ngayHd.split("-").map(Number);
    const effectiveMaDiem = stripCseSuffix(alias[inv.maDiem] || inv.maDiem);
    const maDiemNormalized = stripCseSuffix(inv.maDiem);
    const diemKeysToIndex =
      effectiveMaDiem === maDiemNormalized ? [effectiveMaDiem] : [effectiveMaDiem, maDiemNormalized];
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
        // Chi Nhan (2026-07-27): gop dong CSE/khong-CSE lam 1 -- luon gom
        // nhom theo ma DA BO hau to __FF, bat ke rawCode goc co hau to hay
        // khong (giu nguyen grossData.grossByCode de tra cuu dung so tien,
        // chi doi ma dung de GOM NHOM/hien thi).
        const bucketCode = stripCseSuffix(rawCode);
        if (!gianLines[bucketCode]) {
          gianLines[bucketCode] = { code: bucketCode, gross: 0, invoices: new Set(), effectiveCode: bucketCode };
        }
        gianLines[bucketCode].gross += gross;
        const invs = invoicesByDiemDay[`${bucketCode}|${day}`] || [];
        for (const inv of invs) gianLines[bucketCode].invoices.add(inv.soHd);
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
        // Chi Nhan, 2026-07-29: "hóa đơn tuyên quang nào mà ngày 26 là 300 mấy
        // chục triệu đâu" -- SO KHONG PHAI 300 trieu, chi la 395.675,676 d (395
        // NGAN, phan thap phan la ",676") -- phep chia ty le nay ra so thap
        // phan dai (vd 395675.6756756757), hien thi qua toLocaleString('vi-VN')
        // MAC DINH giu ca phan thap phan (dau phay) nen nhin nham thanh 1 nhom
        // nghin nua ("395.675,676" bi doc nham la "395.675.676"). Lam tron ve
        // dong nguyen (VND khong co don vi nho hon) truoc khi cong don, tranh
        // hien thi gay hieu lam nhu vay ve sau.
        return sum + Math.round(inv.tongTt * (thisDayGross / totalGrossAcrossDays));
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
  applyMultiDayGroupConsolidation(results);

  return results;
}

// Chi Nhan, 2026-07-29: "mấy cái lệch lẻ lẻ này nè là gộp lại xuất 2 ngày
// khớp mà đừng có để lệch cho tôi chớ" -- tach rieng doan gop-nhieu-ngay
// (truoc nam trong reconcileVietQr) thanh 1 ham rieng, EXPORT ra ngoai, de co
// the goi LAI lan 2 sau khi routes/doisoat-vietqr.js dieu chinh gross qua co
// che "lay ngan hang lam chuan" (cong/tru phan du ngan hang<->du lieu, xem
// bankExcessDefaultCode/bankDeficitApplied). Lan goi DAU (trong ham nay) chi
// thay duoc cac cap ngay DA CAN BANG SAN tu dau (vd hoa don 1990 ngay 11-12,
// 2167 ngay 18-19); cap nao CHI can bang SAU KHI dieu chinh theo ngan hang
// (vd hoa don 1824 ngay 4-5, ngay 5 co du ngan hang<->du lieu duoc tru bot)
// can goi lai ham nay LAN 2 (voi gross moi) moi phat hien duoc la da khop.
function applyMultiDayGroupConsolidation(results) {
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
    // Truoc day dung "groupLines[0].invoiceTotal" voi comment "identical for
    // all", nhung do gia da SAI tu khi them tinh nang chia ty le hoa don
    // theo ngay (Luyen, 2026-07-27, xem ghi chu Math.round o tren): moi ngay
    // trong nhom nay gio co invoiceTotal la PHAN CHIA RIENG cua ngay do
    // (khac nhau), khong con "giong het nhau" nua -- so sanh sumGross voi
    // CHI 1 ngay trong so do luon lech ra dung phan cua NHUNG NGAY CON LAI,
    // tao "Lech" gia tren CA 2 ngay. Cong don invoiceTotal CA NHOM lai (luon
    // bang dung tong tien hoa don goc, vi la chia ty le tu chinh no) de so
    // sanh dung tong that.
    const invoiceTotal = groupLines.reduce((sum, l) => sum + l.invoiceTotal, 0);
    if (Math.abs(sumGross - invoiceTotal) < 1) {
      groupLines.forEach((l) => {
        l.matched = true;
        l.diff = 0;
      });
    }
  });

  // Chi Nhan, 2026-07-29: "cái phần này cấn trừ qua là khớp á cấn trừ cho
  // tôi đi cho khớp luôn á" -- khac voi block tren (cung 1 SO HOA DON xuat
  // gop nhieu ngay), day la truong hop 2 HOA DON KHAC NHAU cua CUNG 1 gian
  // o 2 ky lien tiep bi ke toan chia sai RANH GIOI ngay (vd hoa don ngay 24
  // ghi du 100.000d, hoa don gop ngay 25-26 ghi thieu dung 100.000d) -- moi
  // hoa don rieng le deu "Lech" nhung CONG DON ca chuoi ngay lien tiep chua
  // khop (CAM RANH 24+25+26, VINH 24+25+26 deu cong don ve dung 0) thi thuc
  // ra tong doanh thu ca giai doan da duoc xuat du hoa don, chi lam RANH
  // GIOI ngay lech thoi. Gom theo TUNG MA (khong doi hoi cung so hoa don nhu
  // tren), tim CHUOI NGAY LIEN TIEP CHUA KHOP (bi chan boi 2 dau la ngay da
  // khop hoac het du lieu) -- neu tong ca chuoi bang nhau, coi CA CHUOI da
  // khop, KHONG doi tung hoa don rieng (chi doi trang thai hien thi).
  const byCode = new Map();
  results
    .slice()
    .sort((a, b) => (a.settlementDate < b.settlementDate ? -1 : 1))
    .forEach((r) => {
      r.lines.forEach((l) => {
        if (l.manualOverride || l.invoiceNumbers.length === 0) return;
        if (!byCode.has(l.code)) byCode.set(l.code, []);
        byCode.get(l.code).push(l);
      });
    });
  // Luyen, 2026-08-03: "sao có 230k mà hóa đơn tới 280k mà cx để chữ khớp z
  // chèn" -- phat hien qua vi du that BV UNG BUOU PHCM: ngay 30-31/05 (hoa don
  // 4320, gop 2 ngay, TU NO da chia dung ty le nhung CON THIEU dung 50.000d so
  // voi doanh thu 2 ngay do) nam KE BEN ngay 01/06 (hoa don 4451, rieng 1
  // ngay, THUA dung 50.000d) -- 2 hoa don HOAN TOAN KHONG LIEN QUAN nhau,
  // NHUNG vi 3 dong lien tiep nay cong don TINH CO ve dung 0 nen ca 3 bi gan
  // nham "Khop". Hoi lai Luyen: xac nhan CHON "that chat dieu kien can tru",
  // chap nhan rui ro co the lam lo lai vai truong hop dang "Khop" khac tren he
  // thong can ra soat lai (vd CAM RANH/VINH da tung duoc Chi Nhan xac nhan
  // truoc day). Vi KHONG THE phan biet bang toan hoc thuan tuy 1 "loi ranh
  // gioi ngay that su giua 2 hoa don" voi 1 "trung hop cong don ve 0", cach
  // that chat AN TOAN VA RO RANG nhat la GIOI HAN CHUOI TOI DA 2 DONG (dung 2
  // hoa don ke nhau, giong dung vi du goc "hoa don ngay 24" + "hoa don gop
  // 25-26" Chi Nhan neu -- ban chat van la 2 hoa don, nhung o day gioi han o
  // muc 2 DONG dang xet, khong cho chuoi keo dai qua 2 hoa don khac nhau lien
  // tiep nhu truong hop BV UNG BUOU PHCM (3 dong, dung 2 hoa don nhung 1 hoa
  // don gom 2 ngay) -- doi voi truong hop 1 hoa don gop nhieu ngay (dai hon 1
  // dong) can tru voi 1 hoa don khac, se KHONG con tu dong "Khop" nua, hien
  // lai dung "Lech" de Luyen tu kiem tra tung truong hop.
  byCode.forEach((lines) => {
    let i = 0;
    while (i < lines.length) {
      if (lines[i].matched) {
        i++;
        continue;
      }
      let j = i;
      while (j + 1 < lines.length && !lines[j + 1].matched) j++;
      if (j === i + 1) {
        const run = lines.slice(i, j + 1);
        const sumGross = run.reduce((sum, l) => sum + l.gross, 0);
        const sumInvoiceTotal = run.reduce((sum, l) => sum + l.invoiceTotal, 0);
        if (Math.abs(sumGross - sumInvoiceTotal) < 1) {
          run.forEach((l) => {
            l.matched = true;
            l.diff = 0;
          });
        }
      }
      i = j + 1;
    }
  });
}

module.exports = {
  extractVqrCode,
  extractVietQrSettlements,
  extractVietQrThuTransactions,
  parseVietQrRawWorkbook,
  parseVnpayPortalWorkbook,
  migrateVnpayKhMoiInvoices,
  seedMb02865168FromBidv77021,
  migrateTk168Invoices,
  fixBidv7702Day3TaggedAsDay4,
  fixBidv8613600999Day3TaggedAsDay4,
  fixBidv77021Day3TaggedAsDay4,
  parseCuaHangSheet,
  parseStoreExportSheet,
  parseTenDiemMaCongTrinhSheet,
  parseInvoiceWorkbookByTag,
  buildGianCandidatesFromInvoices,
  resolveGianGross,
  resolveGianGrossByBankRef,
  reconcileVietQr,
  applyMultiDayGroupConsolidation,
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
