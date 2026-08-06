// Simple JSON-file based data store.
// Chosen instead of a native SQLite binding so that `npm install` never needs
// to compile native code -- it works on any hosting provider / Node version
// without prebuilt-binary or build-toolchain headaches.
//
// Not meant for huge datasets, but is more than enough for a small
// accounting team tracking a handful of bank accounts and their transactions.

const fs = require("fs");
const path = require("path");
const bcrypt = require("bcryptjs");

const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "store.json");
// Tach 2 mang lon ra file rieng de giam RAM startup: store.json chi chua config
// (~7MB), transactions.json (~95MB) va viet_qr_raw.json (~36MB) duoc load lazy.
const TRANSACTIONS_FILE = path.join(DATA_DIR, "transactions.json");
const VIET_QR_RAW_FILE = path.join(DATA_DIR, "viet_qr_raw.json");

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// Set to true only when load() had to fall back to an EMPTY store because
// the on-disk file existed but was unreadable/unrecoverable. This is used
// to stop the first-run seeding logic below from silently overwriting a
// corrupted-but-still-present data file with a brand new empty one -- that
// exact failure mode has happened before on this machine's mounted drive
// and destroyed real accounting data before this guard existed.
let loadedFromUnrecoverableCorruption = false;

// Chi Nhan, 2026-07-22: cache trong bo nho -- file store.json da lon (30-40MB+
// sau khi them cac tinh nang luu lich su tai len). Truoc day MOI request (ke
// ca chi xem trang, khong sua gi) deu goi load() va load() lai doc + JSON.parse
// LAI TU DAU toan bo file nay (~300-500ms, chan CUNG luong Node.js vi la ham
// dong bo/synchronous) -- khi 2-3 request den gan nhau (vd 1 nguoi mo 2 tab,
// hoac Railway tu kiem tra "/" trong luc dang co request khac dang xu ly) thi
// request den sau phai cho, co the vuot qua thoi gian cho toi da cua Railway
// va tra ve loi "502 Application failed to respond" -- day chinh la nguyen
// nhan gay loi 502 xay ra ngau nhien tren nhieu trang (ke ca trang chu) sau
// khi du lieu lon dan. Fix: chi thuc su doc file tu dia 1 LAN (luc server moi
// khoi dong, hoac neu vi ly do gi do chua co cache), sau do giu lai trong bo
// nho (cachedStore) va tra ve THANG cho cac lan goi load() tiep theo -- save()
// cap nhat lai cache ngay sau khi ghi thanh cong. An toan vi server nay chi
// chay 1 tien trinh Node.js duy nhat cho 1 file du lieu (khong co tien trinh
// nao khac cung sua file nay cung luc).
let cachedStore = null;
// Cache rieng cho 2 mang lon -- chi load khi co route can den (lazy):
// - cachedTransactions: mang giao dich (~95MB, 219K dong) doc tu transactions.json
// - cachedVietQrRaw: du lieu VietQR raw uploads (~36MB) doc tu viet_qr_raw.json
// Ket qua: baseline RAM khi khoi dong chi ~200MB thay vi ~600MB nhu truoc.
let cachedTransactions = null;
let cachedVietQrRaw = null;

// Doc 1 file JSON lon tu dia, tra ve defaultValue neu file chua ton tai hoac loi.
function loadLargeSection(filePath, defaultValue) {
  if (!fs.existsSync(filePath)) return defaultValue;
  try {
    console.log("[store] Dang load " + path.basename(filePath) + "...");
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (e) {
    console.error("[store] Loi doc " + path.basename(filePath) + ":", e.message);
    return defaultValue;
  }
}

// Gan getter lazy-load cho transactions va viet_qr_raw_uploads tren store object.
// Truy cap store.transactions / store.viet_qr_raw_uploads se tu dong doc file
// rieng khi can (lan dau), va ghi nho ket qua trong cachedTransactions/cachedVietQrRaw.
// Setter van hoat dong: store.transactions = newArray ghi de cachedTransactions.
function addSplitGetters(storeObj) {
  Object.defineProperty(storeObj, "transactions", {
    get() {
      if (cachedTransactions === null) {
        cachedTransactions = loadLargeSection(TRANSACTIONS_FILE, []);
      }
      return cachedTransactions;
    },
    set(val) { cachedTransactions = val; },
    configurable: true,
    enumerable: true,
  });
  Object.defineProperty(storeObj, "viet_qr_raw_uploads", {
    get() {
      if (cachedVietQrRaw === null) {
        cachedVietQrRaw = loadLargeSection(
          VIET_QR_RAW_FILE,
          { bidv7704: [], bidv77020: [], mb11521268: [] }
        );
      }
      return cachedVietQrRaw;
    },
    set(val) { cachedVietQrRaw = val; },
    configurable: true,
    enumerable: true,
  });
  return storeObj;
}

function emptyStore() {
  return {
    users: [],
    banks: [],
    transactions: [],
    momo_gross_uploads: [], // [{ id, uploaded_at, file_name, sheetName, dates:[], codes:[], grossByCode:{} }]
    momo_invoices: [], // merged, deduped invoice rows parsed from uploaded invoice lists
    // Cong ty "KH Moi" (TK BIDV7701, TNHH GIAI TRI K&H) -- cung dang Momo
    // nhung la 1 phap nhan/ngan hang KHAC KH Cu, nen tach rieng 2 mang nay
    // (khong dung chung momo_gross_uploads/momo_invoices ben tren) de doi
    // soat 2 cong ty khong lan vao nhau. Xem utils/companies.js + CHANNELS
    // trong routes/doisoat.js.
    momo_moi_gross_uploads: [],
    momo_moi_invoices: [],
    gian_mapping: {}, // { "MA CONG TRINH": "131" | "1388" | "SKIP" } -- Momo ONLY (xem zvp_gian_mapping ben duoi)
    // Chi Nhan, 2026-07-31: "cứ hiện thị cái ch xuất phải chuyển qua 131 rồi
    // lưu lất qua lại vẫn hiện của kh mưới là sao" -- gian_mapping o tren
    // TUNG la "shared across Momo AND Zalo/VNPay/Payoo" (xem comment cu), gay
    // bug: 4 gian duoc Momo ep ve SKIP (vi thuoc KH Moi ben kenh Momo rieng --
    // xem ensureGianHidden trong routes/doisoat.js, chay lai MOI LAN mo trang
    // Momo) lam LUON ca ben ZVP hien "Chưa xuất MISA (KH mới)" cho CUNG ten
    // gian, du ben ZVP la doanh thu KH Cu binh thuong (131) -- Chi Nhan doi
    // lai 131 tren trang ZVP xong quay lai trang Momo la bi de len lai. Tach
    // rieng bang nay CHI cho ZVP (routes/doisoat-zvp.js), KHONG con dung
    // chung voi Momo nua -- moi ben tu quan ly TK Co cua minh doc lap.
    zvp_gian_mapping: {}, // { "MA CONG TRINH": "131" | "1388" | "SKIP" } -- Zalo/VNPay/Payoo ONLY
    cua_hang_mapping: {}, // { "MA CUA HANG": { maCongTrinh, code, gian } } -- learned from "Tong Momo Gop" uploads
    // Some invoices get issued with a "Ma diem" text that names a specific
    // sub-brand/corner (e.g. "SNOWFUN TAN PHU", "FUNFEST SC VIVO") instead of
    // the Ma Cong Trinh code used on the revenue side (Tong Momo Gop / VNPay
    // gian list) -- e.g. "AM TP KVCM" or "SC VIVO KVCM__FF". Without this,
    // those invoices never match ANY gian during reconciliation (wrong "Lech"
    // shown even though the invoice money is really there, just filed under a
    // different name). This table lets that alternate invoice Ma Diem text be
    // mapped onto the correct canonical code; applied at RECONCILE time (not
    // parse time) in reconcileMomo/reconcileZvp, so it also retroactively
    // fixes invoices already uploaded, without needing to re-upload anything.
    // { "TEN/MA DIEM TREN HOA DON": "MA CONG TRINH CHINH (co the co hau to __FF)" }
    // Shared across Momo AND Zalo/VNPay/Payoo, same sharing convention as gian_mapping.
    invoice_diem_alias: {},
    // Zalo App / VNPay (Online+Offline) / Payoo reconciliation -- same shape
    // pattern as the Momo fields above, one upload list per data source.
    zvp_online_uploads: [], // [{ id, uploaded_at, file_name, sheetName, dates, codes, grossByCode, netByCode }]
    zvp_offline_uploads: [], // [{ id, uploaded_at, file_name, sheetName, dates, codes, grossByCode, netByCode }]
    zvp_payoo_uploads: [], // [{ id, uploaded_at, file_name, sheetName, dates, codes, grossByCode, netByCode }]
    zvp_invoices: { zalo: [], vnpay: [], payoo: [] }, // parsed from the same shared MTT invoice file, tagged by channel
    // Last-known reference tables learned from the manually-uploaded files,
    // persisted so later raw-portal-export uploads (order details + fee
    // report) can resolve gian/product without needing those files re-uploaded.
    zvp_gian_list: [], // [{ tenDiem, maCongTrinh, isCse }] from "gian hang xuat HD"
    zvp_offline_diem_map: {}, // { "Chi nhanh": { maCongTrinh, isCse } } from "gian hang VNpay co so"
    zvp_payoo_diem_map: {}, // { "Chi nhanh": { maCongTrinh, isCse } } from Payoo's own "Danh muc ten diem" sheet
    // Chi Nhan, 2026-07-24: "giờ tôi sẽ thiết lập lại chính sát hơn ... dựa
    // vào file hehehehehe sheet nối rồi gắn mã công trình vô ... còn cái nào
    // sau này có tên sản phẩm mới bạn cảnh báo tên sản phẩm đó cho tôi" --
    // bang tra CHINH XAC (khong fuzzy) "Ten san pham" (dung nguyen van, chi
    // trim khoang trang, KHONG bo dau) -> { maCongTrinh, isCse } cho kenh
    // Online/Zalo Mini App, thay the buildOnlineProductMatcher (fuzzy) trong
    // luong upload-combo -- san pham nao khong co trong bang nay se duoc bao
    // "san pham moi" thay vi doan. Xem parseOnlineProductMapSheet /
    // resolveOnlineGrossByProductMap trong utils/zvpReconcile.js.
    zvp_online_product_map: {},
    // Chi Nhan, 2026-07-24: Luyen muon tai THANG file export goc tu cong
    // Payoo/VNPay (khong can gop tay vao "Danh muc ten diem" moi lan nua) --
    // xem parsePayooRawReport trong utils/zvpReconcile.js va route
    // /doi-soat/zvp/upload-payoo-raw. Luu TUNG giao dich rieng le (khong gop
    // theo ngay|ma nhu zvp_payoo_uploads) de khu trung XUYEN SUOT MOI LAN
    // TAI, du la 2 file KHAC dinh dang cho CUNG 1 giao dich (vd bao cao
    // "Giao dich ban hang" gom ca the+QR, va bao cao "Giao dich QR" rieng chi
    // co QR -- cung giao dich se trung "txKey", chi tinh 1 lan, khong cong
    // don). { "<txKey>": { date, gian, gross, fee, net } }
    zvp_payoo_raw_tx: {},
    // Chi Nhan, 2026-07-24: cung 1 kieu voi zvp_payoo_raw_tx o tren, nhung
    // cho VNPay Offline -- Luyen tai THANG file "Du lieu bao cao phi theo GD
    // thanh toan" (khong can file OrderDetails di kem, file do chi dung cho
    // phan Online). Xem parseVnpayOfflineFeeReport trong utils/zvpReconcile.js
    // va route /doi-soat/zvp/upload-offline-raw. { "<Ma giao dich>": { date,
    // chiNhanh, gross, fee, net } }
    zvp_offline_raw_tx: {},
    // Master gian catalog, uploaded daily by Luyen as its own file (sheet
    // "gian "): one table covering ALL channels (Momo, Viet QR, Zalo Mini
    // App, VNPay Co so, Payoo QR/the) with columns raw-text -> Ma cong trinh
    // -> "Thuoc" (channel) -> CSE flag. On upload this is merged (master
    // wins on key collision) into zvp_gian_list / zvp_offline_diem_map /
    // zvp_payoo_diem_map for their respective channels, and into
    // invoice_diem_alias for any row whose raw text differs from its Ma
    // cong trinh -- see utils/zvpReconcile.js parseGianMasterSheet and the
    // merge helpers for why the raw text is not required to be unique (the
    // same displayed gian name can legitimately cover both a CSE and a
    // non-CSE product).
    zvp_gian_master: null, // { uploaded_at, file_name, sheetName, rows: [{raw, maCongTrinh, thuoc, isCse}] }
    // Manual correction for gian that legitimately don't reconcile through
    // the normal VNPay/Zalo invoice do for a given settlement (e.g. the
    // invoice for that revenue is only issued the NEXT day, outside the
    // settlement's own date window). Luyen confirms these by hand once she
    // has identified the gian; the reconciliation then treats that one line
    // as matched instead of showing "Chua co HD".
    // { online: { "<settlementDate>|<code>": { invoiceNumbers:[], amount, note, created_at } }, offline: {...}, payoo: {...} }
    zvp_manual_matches: { online: {}, offline: {}, payoo: {} },
    // Viet QR (BIDV7704 / BIDV77020 / MB11521268) reconciliation. Unlike
    // Momo/ZVP, VietQR payments post to the bank one transaction at a time
    // (no daily/weekly batch), so this channel reconciles per CALENDAR DAY
    // instead of per settlement batch. Store codes ("Ma cua hang") have no
    // direct join to Ma Cong Trinh -- they're fuzzy-matched (same technique
    // as ZVP's Online channel) against each invoice's own "Ten diem xuat
    // hoa don", since there's no separate "gian hang xuat HD" master sheet
    // for this channel. One sub-object per bank channel (bidv7704 /
    // bidv77020 / mb11521268), same sharing convention as zvp_* above.
    viet_qr_raw_uploads: { bidv7704: [], bidv77020: [], mb11521268: [] }, // [{id, uploaded_at, file_name, rows:[{vqrCode,maCuaHang,amount,date}]}]
    viet_qr_store_names: { bidv7704: {}, bidv77020: {}, mb11521268: {} }, // { "MA CUA HANG": { tenCuaHang, tenDiemBan, matchText } } from "Cua hang" sheet
    viet_qr_invoices: { bidv7704: [], bidv77020: [], mb11521268: [] }, // parsed invoice rows tagged per bank
    viet_qr_manual_matches: { bidv7704: {}, bidv77020: {}, mb11521268: {} },
    // Luyen, 2026-07-31: "ngoài các giao dịch đối soát còn có doanh thu khách
    // thu bằng tiền mặt cửa hàng trưởng sẽ thu về á rồi nộp sale bạn cộng vô
    // ... mỗi gian điều có 1 mã nộp tiền á" -- bang tra "Ma noi dung nop
    // tien" (ma cua hang truong ghi vao noi dung khi nop tien mat vao ngan
    // hang) -> Ma cong trinh, hoc tu file upload (xem parseChtNopTienMasterSheet
    // trong utils/zvpReconcile.js va route /bao-cao/xuat-hoa-don-ban-ra/
    // upload-cht-nop-tien). Dung de nhan dien cac giao dich "thu" da co san
    // trong store.transactions (sao ke chung, khong phai 1 kenh doi soat
    // rieng) la tien nop tay cua cua hang truong, roi cong vao dong "CHT nộp
    // tiền" trong bao cao Xuat Hoa Don Ban Ra. { "MA NOI DUNG NOP TIEN": {
    // maCongTrinh, phapNhan, sheetName } }.
    cht_nop_tien_map: {},
    cht_nop_tien_uploads: [], // [{id, uploaded_at, file_name, sheetsParsed, rowCount}]
    // Luyen, 2026-08-01: "cần để mốt tôi tải nhầm tôi có thể xóa á" -- muc
    // "Tai file sao ke tai truc tiep tu ngan hang" (routes/transactions.js,
    // POST /transactions/upload-statement) truoc gio nap giao dich THANG vao
    // store.transactions, KHONG luu lai lich su tung lan tai (khac voi cac
    // kenh Momo/ZVP/VietQR da co uploads rieng) -- neu tai nham file/nham
    // ngan hang thi phai do tay tung dong de xoa (xem vu MB02865168 bi nap
    // nham sao ke BIDV8681, 2026-08-01). Them so nay de moi lan tai luu lai
    // DUNG cac id giao dich vua tao ra, cho phep xoa nguyen 1 dot bang 1 nut
    // bam thay vi do tay. [{ id, bank_id, bank_name, file_name, sheetName,
    // uploaded_at, uploaded_by, transaction_ids: [...], rows_inserted,
    // rows_skipped }]
    bank_statement_uploads: [],
    seq: { users: 0, banks: 0, transactions: 0 },
  };
}

function tryHealTrailingGarbage(rawBuffer) {
  // Some writes to this data file have occasionally left stale trailing
  // bytes (NUL padding, or leftover bytes from a longer previous version of
  // the file) after otherwise-complete, valid JSON. Trim trailing NUL bytes
  // and re-parse before giving up.
  let end = rawBuffer.length;
  while (end > 0 && rawBuffer[end - 1] === 0) end--;
  const trimmed = rawBuffer.slice(0, end).toString("utf8");
  return JSON.parse(trimmed); // throws if still invalid
}

function backupCorruptedFile() {
  try {
    const backupPath = DATA_FILE + ".corrupted-" + Date.now() + ".bak";
    fs.copyFileSync(DATA_FILE, backupPath);
    console.error("[store] Da sao luu file loi sang: " + backupPath);
    return backupPath;
  } catch (e) {
    console.error("[store] Khong the sao luu file loi:", e.message);
    return null;
  }
}

function load() {
  if (cachedStore) return cachedStore;
  if (!fs.existsSync(DATA_FILE)) {
    // File chua ton tai (cai dat moi) -- khong ghi file trong, de save() tu seed() ghi sau.
    const fresh = emptyStore();
    delete fresh.transactions;
    delete fresh.viet_qr_raw_uploads;
    cachedStore = addSplitGetters(fresh);
    return cachedStore;
  }
  const raw = fs.readFileSync(DATA_FILE, "utf8");
  try {
    const parsed = JSON.parse(raw);
    const base = emptyStore();
    const merged = Object.assign({}, base, parsed, {
      seq: Object.assign({}, base.seq, parsed.seq || {}),
    });
    // Di tru tu format cu (1 file) sang format moi (3 file):
    // Neu store.json con nhung transactions/viet_qr_raw_uploads nhung (format cu),
    // luu vao cache va xoa khoi config object truoc khi gan getter.
    if (Array.isArray(merged.transactions) && merged.transactions.length > 0) {
      if (cachedTransactions === null) cachedTransactions = merged.transactions;
    }
    if (
      merged.viet_qr_raw_uploads &&
      typeof merged.viet_qr_raw_uploads === "object" &&
      !Array.isArray(merged.viet_qr_raw_uploads) &&
      Object.values(merged.viet_qr_raw_uploads).some((v) => Array.isArray(v) && v.length > 0)
    ) {
      if (cachedVietQrRaw === null) cachedVietQrRaw = merged.viet_qr_raw_uploads;
    }
    delete merged.transactions;
    delete merged.viet_qr_raw_uploads;
    cachedStore = addSplitGetters(merged);
    return cachedStore;
  } catch (e) {
    console.error("[store] Loi doc file du lieu, thu tu phuc hoi:", e.message);
    try {
      const rawBuffer = fs.readFileSync(DATA_FILE);
      const parsed = tryHealTrailingGarbage(rawBuffer);
      console.error("[store] Da tu phuc hoi du lieu thanh cong (cat bo byte rac cuoi file).");
      const base = emptyStore();
      const healed = Object.assign({}, base, parsed, {
        seq: Object.assign({}, base.seq, parsed.seq || {}),
      });
      if (Array.isArray(healed.transactions) && healed.transactions.length > 0) {
        if (cachedTransactions === null) cachedTransactions = healed.transactions;
      }
      if (
        healed.viet_qr_raw_uploads &&
        typeof healed.viet_qr_raw_uploads === "object" &&
        !Array.isArray(healed.viet_qr_raw_uploads) &&
        Object.values(healed.viet_qr_raw_uploads).some((v) => Array.isArray(v) && v.length > 0)
      ) {
        if (cachedVietQrRaw === null) cachedVietQrRaw = healed.viet_qr_raw_uploads;
      }
      delete healed.transactions;
      delete healed.viet_qr_raw_uploads;
      cachedStore = addSplitGetters(healed);
      save(cachedStore); // luu ngay phien ban da sua de on dinh
      return cachedStore;
    } catch (e2) {
      console.error("[store] KHONG THE tu phuc hoi du lieu:", e2.message);
      backupCorruptedFile();
      console.error(
        "[store] CANH BAO NGHIEM TRONG: file data/store.json bi loi va KHONG the tu phuc hoi. " +
          "App se chay tam voi du lieu RONG trong bo nho (KHONG ghi de len file that) de ban con co the yeu cau khoi phuc thu cong. " +
          "Kiem tra file .corrupted-*.bak trong thu muc data/."
      );
      loadedFromUnrecoverableCorruption = true;
      const emptyS = emptyStore();
      delete emptyS.transactions;
      delete emptyS.viet_qr_raw_uploads;
      return addSplitGetters(emptyS);
    }
  }
}

function writeFileDurable(filePath, data) {
  // NOTE: fs.writeSync(fd, buffer) is a thin wrapper over the write(2)
  // syscall and is NOT guaranteed to write the whole buffer in one call --
  // on a flaky/mounted filesystem a short write can silently truncate large
  // content mid-way. fs.writeFileSync() DOES loop internally until the
  // entire buffer is written (this is exactly the bug that corrupted this
  // data file more than once), so we use that for the actual content write,
  // then separately reopen the file just to fsync it to force a real flush.
  fs.writeFileSync(filePath, data);
  const fd = fs.openSync(filePath, "r+");
  try {
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
}

function save(store) {
  // Chi Nhan, 2026-08-06: tach store thanh 3 file de giam RAM startup:
  // - store.json: chi chua config (banks, users, momo, zvp, viet_qr_store_names...) ~7MB
  // - transactions.json: mang giao dich ~95MB
  // - viet_qr_raw.json: du lieu VietQR raw ~36MB
  // Moi file duoc ghi qua pattern tmp+rename+verify nhu truoc (dam bao toan ven).
  // Luu y: KHONG dung store.transactions hoac store.viet_qr_raw_uploads o day
  // (se kich hoat getter, doc file tu dia), thay vao do dung truc tiep
  // cachedTransactions/cachedVietQrRaw (du lieu dang o trong bo nho).

  // Xay dung config object (tat ca key tru transactions va viet_qr_raw_uploads).
  const configObj = {};
  for (const key of Object.keys(store)) {
    if (key === "transactions" || key === "viet_qr_raw_uploads") continue;
    configObj[key] = store[key];
  }
  const mainData = JSON.stringify(configObj);
  const txData = JSON.stringify(
    cachedTransactions !== null ? cachedTransactions : []
  );
  const vqrData = JSON.stringify(
    cachedVietQrRaw !== null
      ? cachedVietQrRaw
      : { bidv7704: [], bidv77020: [], mb11521268: [] }
  );

  function writeDurableWithRetry(filePath, data) {
    const expectedBytes = Buffer.byteLength(data, "utf8");
    let lastErr = null;
    for (let attempt = 1; attempt <= 3; attempt++) {
      const tmpFile =
        filePath +
        ".tmp-" +
        process.pid +
        "-" +
        Date.now() +
        "-" +
        Math.random().toString(36).slice(2);
      try {
        writeFileDurable(tmpFile, data);
        fs.renameSync(tmpFile, filePath);
        const actualBytes = fs.statSync(filePath).size;
        if (actualBytes !== expectedBytes) {
          throw new Error(
            "Ghi file khong du so byte (mong doi " +
              expectedBytes +
              ", thuc te " +
              actualBytes +
              ")"
          );
        }
        return; // thanh cong
      } catch (e) {
        lastErr = e;
        console.error(
          "[store] save() " + path.basename(filePath) + " attempt " + attempt + " that bai:",
          e.message
        );
        try { fs.unlinkSync(tmpFile); } catch (_) { /* bo qua */ }
      }
    }
    throw new Error(
      "Khong the luu " +
        path.basename(filePath) +
        " an toan sau nhieu lan thu: " +
        (lastErr && lastErr.message)
    );
  }

  writeDurableWithRetry(DATA_FILE, mainData);
  writeDurableWithRetry(TRANSACTIONS_FILE, txData);
  writeDurableWithRetry(VIET_QR_RAW_FILE, vqrData);

  cachedStore = store; // cap nhat cache sau khi ghi thanh cong
}

function nextId(store, collection) {
  // Chi Nhan, 2026-07-30: phat hien BIDV77021 va BIDV8681 (2 ngan hang khac
  // nhau) bi trung id=16 -- rat co the do ban offline (may Nhan) va ban online
  // (Railway) moi ben tu tao ngan hang moi doc lap, moi ben giu 1 ban
  // "cachedStore.seq.banks" RIENG trong bo nho (xem ghi chu cachedStore o
  // tren), nen ca 2 tinh ra CUNG 1 "so tiep theo" roi luu de len nhau qua
  // buoc dong bo/sao luu -- khien trang Danh sach giao dich loc theo 1 tai
  // khoan bi cong nham ca giao dich cua tai khoan kia vao (xem
  // fixDuplicateBidv8681BankId ben duoi cho lan nay cu the). De tranh trung
  // id THEM LAN NUA sau nay (khong chi rieng banks, ca users/transactions),
  // doi chieu voi ID LON NHAT dang thuc su co trong mang tuong ung truoc khi
  // tang, thay vi chi tin tuong con dem luu rieng (co the bi lech giua 2 tien
  // trinh) -- an toan vi ten mang y het ten collection ("banks" ->
  // store.banks, "transactions" -> store.transactions, "users" -> store.users).
  const arr = store[collection];
  if (Array.isArray(arr) && arr.length > 0) {
    const maxExisting = arr.reduce((m, x) => (x && typeof x.id === "number" && x.id > m ? x.id : m), 0);
    if (maxExisting > (store.seq[collection] || 0)) {
      store.seq[collection] = maxExisting;
    }
  }
  store.seq[collection] = (store.seq[collection] || 0) + 1;
  return store.seq[collection];
}

// Chi Nhan, 2026-07-30: "sao tôi chọn 77021 nó nhẩy ra tài khoản khác vậy ...
// cộng gì tới 90 mấy triệu vậy" -- BIDV77021 (tao 28/7) va BIDV8681 (tao 29/7,
// "Tài khoản chi") vo tinh bi trung id=16. Vi store.transactions chi luu
// "bank_id" (so), ca 2 tai khoan bi GOP CHUNG lam 1 tren moi trang loc/tong
// theo ngan hang -- vd ngay 29/7, sao ke that cua 77021 chi co 50.170.000d
// nhung he thong cong ra 94.657.521d vi bi cong them 1 giao dich that su cua
// BIDV8681 (44.507.521d, "FSS TT HTKD ...", KHONG lien quan VietQR). Da doi
// chieu THU CONG toan bo 361 giao dich dang nam duoi id=16 (tu 1/7 den 29/7):
// giao dich nao co nhac so tai khoan "8690077021" trong mo ta la CUA
// BIDV77021 that (71.583 dong, hau het la thu QR + 1 phi quan ly TK); con lai
// (30 thu + 331 chi, toan chi phi/tam ung/thanh toan nha cung cap -- dung
// dang giao dich cua 1 tai khoan CHI, khop voi vai tro "Tài khoản chi" cua
// BIDV8681) la cua BIDV8681 that. Tach BIDV8681 sang ID MOI + chuyen dung
// 361 giao dich do theo, giu nguyen id=16 cho BIDV77021.
function fixDuplicateBidv8681BankId(store) {
  const b77021 = store.banks.find((b) => b.name === "BIDV77021");
  const b8681 = store.banks.find((b) => b.name === "BIDV8681");
  if (!b77021 || !b8681 || b77021.id !== b8681.id) return false; // da tach roi hoac khong (con) trung id
  const oldId = b8681.id;
  const newId = store.banks.reduce((m, b) => (b.id > m ? b.id : m), 0) + 1;
  b8681.id = newId;
  let moved = 0;
  store.transactions.forEach((t) => {
    if (t.bank_id === oldId && !(t.description || "").includes("8690077021")) {
      t.bank_id = newId;
      moved += 1;
    }
  });
  console.log(
    `[fix] Da tach BIDV8681 khoi bi trung ID voi BIDV77021 (id cu=${oldId}) -> id moi=${newId}, da chuyen ${moved} giao dich ve dung BIDV8681.`
  );
  return true;
}

const SEED_CHT_NOP_TIEN_ROWS = [
  {"noiDungNopTien":"KH989KVCMB0001","maCongTrinh":"AM LBIEN KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0002","maCongTrinh":"AM HP KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0003","maCongTrinh":"LOTTE LPH KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0004","maCongTrinh":"AE HUE KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0005","maCongTrinh":"AM HP KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0006","maCongTrinh":"AM LBIEN KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0007","maCongTrinh":"FARM TIMES CITY","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0008","maCongTrinh":"AM LB KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0009","maCongTrinh":"FUNZONE IPH KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0011","maCongTrinh":"AE HUE KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0012","maCongTrinh":"AM HP KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0013","maCongTrinh":"AE HUE KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0014","maCongTrinh":"FZ DNANG KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0015","maCongTrinh":"KID FARM MM MARKET DA NANG","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0016","maCongTrinh":"AM HP KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0017","maCongTrinh":"FZ DNANG KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0018","maCongTrinh":"AE HUE KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH989KVCMB0019","maCongTrinh":"AM HP KVCN","phapNhan":"KH cũ","sheetName":"KVC MB KH989"},
  {"noiDungNopTien":"KH705KVCMB0001","maCongTrinh":"KVC TIMES","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0002","maCongTrinh":"KVC ROYAL","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0003","maCongTrinh":"KVC AE HUE","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0004","maCongTrinh":"KVC ROYAL","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0005","maCongTrinh":"AE HP KVC","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0006","maCongTrinh":"LOTTE VINH KVC","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0007","maCongTrinh":"KVC ROYAL","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMB0008","maCongTrinh":"SAVICO PHN","phapNhan":"KH mới","sheetName":"KVC MB KH705"},
  {"noiDungNopTien":"KH705KVCMN0001","maCongTrinh":"KVC ESTELLA","phapNhan":"KH mới","sheetName":"KVC MN KH705"},
  {"noiDungNopTien":"KH705KVCMN0002","maCongTrinh":"FARM LOTTE PHAN THIET","phapNhan":"KH mới","sheetName":"KVC MN KH705"},
  {"noiDungNopTien":"KH705KVCMN0003","maCongTrinh":"AM TP KVCM","phapNhan":"KH mới","sheetName":"KVC MN KH705"},
  {"noiDungNopTien":"KH705KVCMN0004","maCongTrinh":"FARM LOTTE NHA TRANG","phapNhan":"KH mới","sheetName":"KVC MN KH705"},
  {"noiDungNopTien":"KH705KVCMN0005","maCongTrinh":"TUTU MN AEON MALL TÂN AN","phapNhan":"KH mới","sheetName":"KVC MN KH705"},
  {"noiDungNopTien":"KH989KVCMN0001","maCongTrinh":"LOTTE GO VAP KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0002","maCongTrinh":"AM TP KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0003","maCongTrinh":"AE BT KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0004","maCongTrinh":"AM BD KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0005","maCongTrinh":"AM BD KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0006","maCongTrinh":"FZ DIY SORA","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0007","maCongTrinh":"KVC LOTTE VUNG TAU","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0008","maCongTrinh":"SC VIVO KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0009","maCongTrinh":"AM TP KVCM","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0010","maCongTrinh":"FUNFEST SCVIVO","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0010","maCongTrinh":"FUNFEST SCVIVO","phapNhan":"KH cũ","sheetName":"KVC MN KH989"},
  {"noiDungNopTien":"KH705MTDMB0001","maCongTrinh":"AE LBIEN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0002","maCongTrinh":"0","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0003","maCongTrinh":"0","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0004","maCongTrinh":"IPH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0005","maCongTrinh":"BIG C HGUOM PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0006","maCongTrinh":"BIG C LTT PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0007","maCongTrinh":"BIG C TL PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0008","maCongTrinh":"0","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0009","maCongTrinh":"LOTTY FRIENDS PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0010","maCongTrinh":"MLINH PLAZA PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0011","maCongTrinh":"MIPEC LB PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0012","maCongTrinh":"NSTV BTL PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0013","maCongTrinh":"NSTV OCP PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0014","maCongTrinh":"NSTV SAVICO PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0015","maCongTrinh":"NSTV ROY PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0016","maCongTrinh":"VINKE-TCUNG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0017","maCongTrinh":"TĐBS PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0018","maCongTrinh":"THE GARDEN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0019","maCongTrinh":"IPH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0020","maCongTrinh":"SAVICO PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0021","maCongTrinh":"RAP CPQG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0022","maCongTrinh":"VC BA TRIEU PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0023","maCongTrinh":"VC BTL PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0024","maCongTrinh":"VC METROPOLIS (LIEU GIAI) PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0025","maCongTrinh":"VC NCT PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0026","maCongTrinh":"OCP PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0027","maCongTrinh":"OCP PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0028","maCongTrinh":"VC PHAM HUNG (SKYLAKE)","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0029","maCongTrinh":"VC PNT PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0030","maCongTrinh":"VC ROY PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0031","maCongTrinh":"VC SMART PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0032","maCongTrinh":"VC TIMES PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0033","maCongTrinh":"VC TDH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0034","maCongTrinh":"VINKE-TCUNG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0035","maCongTrinh":"AE HP PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0036","maCongTrinh":"AE HP PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0037","maCongTrinh":"GO HP PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0038","maCongTrinh":"BENH VIEN BAI CHAY PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0039","maCongTrinh":"GO HA LONG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0040","maCongTrinh":"JP-KVC- P SUN HL","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0041","maCongTrinh":"VC HA LONG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0042","maCongTrinh":"GO NINH BINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0043","maCongTrinh":"GO THAI NGUYEN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0044","maCongTrinh":"NSTV GO TN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0045","maCongTrinh":"NSTV VC TN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0046","maCongTrinh":"VC THAI NGUYEN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0047","maCongTrinh":"GO BAC GIANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0048","maCongTrinh":"VC BAC GIANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0049","maCongTrinh":"TAM CHUC PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0050","maCongTrinh":"GO HA NAM PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0051","maCongTrinh":"NSTV PHU LY PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0052","maCongTrinh":"VC PHU LY PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0053","maCongTrinh":"GO THAI BINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0054","maCongTrinh":"NSTV THAI BINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0055","maCongTrinh":"VC THAI BINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0056","maCongTrinh":"CITY HUB VINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0057","maCongTrinh":"GO VINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0058","maCongTrinh":"LOTTE VINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0059","maCongTrinh":"VINH CENTER PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0060","maCongTrinh":"POSH AE HUE","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0061","maCongTrinh":"GO HUE PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0062","maCongTrinh":"VC HUE PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0063","maCongTrinh":"COOP DNANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0064","maCongTrinh":"GO DA NANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0065","maCongTrinh":"LOTTE ĐN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0066","maCongTrinh":"MIKAZUKI DN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0067","maCongTrinh":"0","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0068","maCongTrinh":"VC DN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0069","maCongTrinh":"BA NA HILL PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0070","maCongTrinh":"VINPEARL HOI AN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0071","maCongTrinh":"P GOLD COAST NHA TRANG (RSM)","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0072","maCongTrinh":"GO NTRANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0073","maCongTrinh":"NHA TRANG CENTER PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0074","maCongTrinh":"VC MAXI TN NT PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0075","maCongTrinh":"VC T.PHU N.TRANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0076","maCongTrinh":"VPNT PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0077","maCongTrinh":"JP-POSH VW NHA TRANG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0078","maCongTrinh":"MOC CHAU PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0079","maCongTrinh":"SUN FANSIPAN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0080","maCongTrinh":"JP-KVC-POSH SUN FANSIPAN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0081","maCongTrinh":"NSTV TUYEN QUANG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0082","maCongTrinh":"VC TUYEN QUANG PNH","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0083","maCongTrinh":"GO THANH HOA PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0084","maCongTrinh":"VC THANH HOA PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0085","maCongTrinh":"GO NAM DINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0086","maCongTrinh":"NSTV NAM DINH PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0087","maCongTrinh":"P-JP-PF SUN CAT BA","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0088","maCongTrinh":"GO HAI DUONG PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0089","maCongTrinh":"VC VU YEN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0090","maCongTrinh":"JP POSH VW VU YEN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0091","maCongTrinh":"MM MARKET DA NANG MTD","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0092","maCongTrinh":"MM MARKET DA NANG MTD","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0093","maCongTrinh":"GO VIET TRI PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0094","maCongTrinh":"VC VIET TRI PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0095","maCongTrinh":"JP VC BA TRIEU","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0096","maCongTrinh":"JP IPH","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0097","maCongTrinh":"JP AE LBIEN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0098","maCongTrinh":"JP VC ROY","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0099","maCongTrinh":"P-JP-PF SUN CAT BA","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0100","maCongTrinh":"JP POSH VW VU YEN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0101","maCongTrinh":"JP GO ĐA NANG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0102","maCongTrinh":"JP BA NA HILL","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0103","maCongTrinh":"JP-KVC- P SUN HL","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0104","maCongTrinh":"JP CTQT HLONG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0105","maCongTrinh":"JP-KVC-POSH SUN FANSIPAN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0106","maCongTrinh":"SUN FANSIPAN PHN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0107","maCongTrinh":"JP VINPEARL NT","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0108","maCongTrinh":"JP GOLD COAST NHA TRANG (RSM)","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0109","maCongTrinh":"JP NT CENTER","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0110","maCongTrinh":"JP VINPEARL HOI AN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0111","maCongTrinh":"MGG IPH (PICK FUN)","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0112","maCongTrinh":"TĐBS PF","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0113","maCongTrinh":"PF TIME CITY","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0114","maCongTrinh":"PF TAM CHUC","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0115","maCongTrinh":"P-JP-PF SUN CAT BA","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0116","maCongTrinh":"MGG BA NA HILL","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0117","maCongTrinh":"JP AE HUE","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0118","maCongTrinh":"POSH MB KUBO GO THĂNG LONG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0119","maCongTrinh":"POSH MB ECOPARK VINH","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0120","maCongTrinh":"POSH MB KUBO GO NAM ĐỊNH","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0121","maCongTrinh":"POSH MB KUBO GO HẢI DƯƠNG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0122","maCongTrinh":"POSH MB KUBO GO LONG BIÊN","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0123","maCongTrinh":"POSH MB KUBO GO ĐÀ NẴNG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0124","maCongTrinh":"POSH MB KUBO GO NINH BÌNH","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0125","maCongTrinh":"POSH MB KUBO GO VIỆT TRÌ","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0126","maCongTrinh":"POSH GO VĨNH PHÚC","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0127","maCongTrinh":"POSH KUBO VĨNH PHÚC","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0128","maCongTrinh":"POSH COOPMART VĨNH PHÚC","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH705MTDMB0129","maCongTrinh":"KUBO NHA TRANG","phapNhan":"KH mới","sheetName":"MTD MB KH705"},
  {"noiDungNopTien":"KH989MTDMB0001","maCongTrinh":"0","phapNhan":"KH cũ","sheetName":"MTĐ MB KH989"},
  {"noiDungNopTien":"KH989MTDMB0002","maCongTrinh":"SB VINH PHN","phapNhan":"KH cũ","sheetName":"MTĐ MB KH989"},
  {"noiDungNopTien":"KH989MTDMB0003","maCongTrinh":"SB CAM RANH PHN","phapNhan":"KH cũ","sheetName":"MTĐ MB KH989"},
  {"noiDungNopTien":"KH989MTDMB0004","maCongTrinh":"CHKQT CAM RANH","phapNhan":"KH cũ","sheetName":"MTĐ MB KH989"},
  {"noiDungNopTien":"KH989MTDMB0005","maCongTrinh":"JP SB NOI BAI","phapNhan":"KH cũ","sheetName":"MTĐ MB KH989"},
  {"noiDungNopTien":"KH989MTDMB0006","maCongTrinh":"CHKQT CAM RANH","phapNhan":"KH cũ","sheetName":"MTĐ MB KH989"},
  {"noiDungNopTien":"KH705MTDMN0001","maCongTrinh":"AE BINH TAN PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0002","maCongTrinh":"AM TP PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0003","maCongTrinh":"BV 175 PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0004","maCongTrinh":"BV UNG BUOU PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0005","maCongTrinh":"ESTELLA PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0006","maCongTrinh":"GIGAMALL PVD PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0007","maCongTrinh":"GO AU CO PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0008","maCongTrinh":"GO NTT PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0009","maCongTrinh":"GO TRUONG CHINH PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0010","maCongTrinh":"LOTTE Q7 (NSG) PHM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0011","maCongTrinh":"LOTTE GO VAP VR-PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0012","maCongTrinh":"POSH LOTTE PTHO","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0013","maCongTrinh":"SENSE CT PVĐ PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0014","maCongTrinh":"SC VIVO PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0015","maCongTrinh":"VC 3/2 JP-POSH","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0016","maCongTrinh":"JP-POSH GRAND PARK","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0017","maCongTrinh":"VC GV PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0018","maCongTrinh":"VC LVV PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0019","maCongTrinh":"VHANH MALL PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0020","maCongTrinh":"AE BD PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0021","maCongTrinh":"GO DI AN PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0022","maCongTrinh":"GO TDM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0023","maCongTrinh":"VC BIEN HOA PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0024","maCongTrinh":"SB PHU QUOC PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0025","maCongTrinh":"PQ SUN HTHOM PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0026","maCongTrinh":"GO BA RIA PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0027","maCongTrinh":"MM MARKET DA NANG MTD","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0028","maCongTrinh":"CON DAO AIRPORT PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0029","maCongTrinh":"GO BEN TRE PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0030","maCongTrinh":"SENSE BTRE PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0031","maCongTrinh":"GO MY THO PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0032","maCongTrinh":"GO TRA VINH","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0033","maCongTrinh":"GO CAN THO PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0034","maCongTrinh":"SB CAN THO PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0035","maCongTrinh":"SENSE CT CTHO PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0036","maCongTrinh":"GO BMT PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0037","maCongTrinh":"ZONE C KIEN GIANG PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0038","maCongTrinh":"SENSE CA MAU PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0039","maCongTrinh":"GO BAC LIEU PHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0040","maCongTrinh":"LOTTE PHAN THIET","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0041","maCongTrinh":"JP AE TAN PHU","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0042","maCongTrinh":"VC 3/2 JP-POSH","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0043","maCongTrinh":"PQ SUN HTHOM JPHCM","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0044","maCongTrinh":"GO NTRANG PHN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0045","maCongTrinh":"0","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0046","maCongTrinh":"AM HP KVCN","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0047","maCongTrinh":"POSH MN KUBO GO BÀ RỊA","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0048","maCongTrinh":"POSH MN KUBO GO CẦN THƠ","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0049","maCongTrinh":"POSH MN GALAXY KINH DƯƠNG VƯƠNG","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0050","maCongTrinh":"POSH MN GALAXY QUANG TRUNG","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0051","maCongTrinh":"POSH MN KUBO GO BIÊN HÒA","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH705MTDMN0052","maCongTrinh":"POSH MN KUBO GO BUÔN MÊ THUỘT","phapNhan":"KH mới","sheetName":"MTD MN KH705"},
  {"noiDungNopTien":"KH989MTDMN0001","maCongTrinh":"COOP PLAM PHCM","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0002","maCongTrinh":"COOP BDUONG PHCM","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0003","maCongTrinh":"PQ VINPERAL PHCM","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0004","maCongTrinh":"AE BT JP","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0005","maCongTrinh":"AE BD JP","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0006","maCongTrinh":"JP SORA BECAMEX","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0007","maCongTrinh":"JP VW PHÚ QUỐC","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989MTDMN0008","maCongTrinh":"CHKQT PHU QUOC","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
  {"noiDungNopTien":"KH989KVCMN0010","maCongTrinh":"FUNFEST SCVIVO","phapNhan":"KH cũ","sheetName":"MTĐ MN KH989"},
];

// ---- Seed default admin user + example bank on first run ----
(function seed() {
  const store = load();

  // Safety guard: if load() had to fall back to an empty store because the
  // real file on disk is corrupted and unrecoverable, do NOT run the normal
  // "first run" seeding-and-save logic -- that would permanently overwrite
  // the corrupted (but backed-up) file with a brand new empty one, losing
  // any chance of manual recovery. Just let the app run with an in-memory
  // empty store for this process; nothing gets written to disk until the
  // data file is fixed and the app restarted.
  if (loadedFromUnrecoverableCorruption) {
    console.error(
      "[store] Bo qua buoc seed mac dinh vi du lieu dang o trang thai loi khong the phuc hoi -- KHONG ghi de len file store.json."
    );
    return;
  }

  let changed = false;

  if (store.users.length === 0) {
    const defaultUsername = process.env.ADMIN_USERNAME || "admin";
    const defaultPassword = process.env.ADMIN_PASSWORD || "khadmin123";
    const hash = bcrypt.hashSync(defaultPassword, 10);
    store.users.push({
      id: nextId(store, "users"),
      username: defaultUsername,
      password_hash: hash,
      name: "Quan tri vien K&H",
      role: "admin",
      created_at: new Date().toISOString(),
    });
    changed = true;
    console.log(
      "[seed] Da tao tai khoan mac dinh -> username: " + defaultUsername + " / password: " + defaultPassword + " (doi ngay sau khi dang nhap lan dau)"
    );
  }

  // Luyen (2026-07-21): "quen mat khau, khong con dang nhap o dau ca" -- mat
  // khau duoc ma hoa 1 chieu nen KHONG the xem lai/khoi phuc, chi co the DAT
  // LAI. Co che "break-glass" nay cho phep tu dat lai mat khau 1 tai khoan MA
  // KHONG CAN dang nhap truoc: dat 2 bien moi truong RESET_PASSWORD_USERNAME
  // + RESET_PASSWORD_NEW (vd tren Railway: Settings > Variables), roi khoi
  // dong lai server (Railway tu redeploy khi luu bien moi) -- lan khoi dong
  // ke tiep se dat lai dung mat khau do (tao moi tai khoan Quan tri neu ten
  // dang nhap chua ton tai). BAT BUOC xoa 2 bien nay ngay sau khi dang nhap
  // lai duoc, neu khong moi lan server khoi dong lai se tiep tuc dat lai ve
  // dung mat khau do (khong an toan de lau dai).
  const resetUsername = (process.env.RESET_PASSWORD_USERNAME || "").trim();
  const resetPassword = process.env.RESET_PASSWORD_NEW || "";
  if (resetUsername && resetPassword) {
    let user = store.users.find((u) => u.username === resetUsername);
    if (!user) {
      user = {
        id: nextId(store, "users"),
        username: resetUsername,
        name: "Quan tri vien (tao qua RESET_PASSWORD)",
        role: "admin",
        created_at: new Date().toISOString(),
      };
      store.users.push(user);
    }
    user.password_hash = bcrypt.hashSync(resetPassword, 10);
    if (!user.role) user.role = "admin";
    changed = true;
    console.log(
      "[reset] Da dat lai mat khau cho tai khoan '" + resetUsername + "' theo bien moi truong RESET_PASSWORD_USERNAME/RESET_PASSWORD_NEW. " +
        "NHO XOA 2 BIEN NAY NGAY SAU KHI DANG NHAP LAI DUOC."
    );
  }

  if (fixDuplicateBidv8681BankId(store)) changed = true;

  if (store.banks.length === 0) {
    store.banks.push({
      id: nextId(store, "banks"),
      name: "Ngan hang mau 1",
      account_number: "0000000001",
      bank_name: "Vietcombank",
      opening_balance: 0,
      opening_date: new Date().toISOString().slice(0, 10),
      created_at: new Date().toISOString(),
    });
    changed = true;
  }

  // 3 tai khoan cua cong ty "KH Moi" (TNHH GIAI TRI K&H) -- them 1 lan, idempotent
  // (chi tao neu chua co ten nay), de cac trang Doi soat Momo/Viet QR/Chi phi
  // tim duoc ngan hang tuong ung ngay ca khi Luyen chua tu tao trong /banks.
  // So tai khoan lay tu file sao ke thuc te "Ngan hang KH moi.xlsx" chi lay.
  const khMoiBanks = [
    { name: "BIDV7701", account_number: "8620107701", bank_name: "BIDV" },
    { name: "BIDV7702", account_number: "8640107702", bank_name: "BIDV" },
    { name: "VP58888", account_number: "4552958888", bank_name: "VPBank" },
  ];
  khMoiBanks.forEach((b) => {
    if (!store.banks.some((x) => x.name === b.name)) {
      store.banks.push({
        id: nextId(store, "banks"),
        name: b.name,
        account_number: b.account_number,
        bank_name: b.bank_name,
        opening_balance: 0,
        opening_date: new Date().toISOString().slice(0, 10),
        company: "kh_moi",
        created_at: new Date().toISOString(),
      });
      changed = true;
    }
  });

  // Chi Nhan, 2026-07-30: "sao ngân hàng trả... vẫn thiếu 20k vậy là sao" --
  // sao ke BIDV77021 ngay 29/7 tai len lai bi thieu dung 1 giao dich that
  // (20.000d, "thu", tham chieu "083z36w-8Awg5iyRR"). Nguyen nhan SAU KHI tai
  // lai: bankStatementParser.js co chu dinh suy Thu/Chi tu DELTA cot "So du"
  // (khong doc truc tiep cot "Phat sinh No/Co") vi coi so du la "chan ly ngan
  // hang" dang tin cay hon -- nhung dong DAU TIEN cua 1 lan tai file phu
  // thuoc vao so du LUY KE ma he thong dang co san TRUOC ngay do (opening_balance
  // + tong Thu/Chi da luu); neu so du luy ke nay bi lech so voi ngan hang that
  // (o day lech ~2,37 ty, rat co the do sai sot lich su tu 21/7-28/7, can doi
  // chieu rieng voi sao ke day du giai doan do de tim dung nguyen nhan) thi
  // dong dau tien cua lan tai moi se bi tinh SAI ca loai (thu/chi) lan so tien
  // theo phan chenh lech do -- day chinh la ly do dong 20.000d nay bi ghi
  // nham thanh "chi" 2.374.997.880d thay vi "thu" 20.000d. Da doi chieu truc
  // tiep voi cot "Phat sinh Co" cua file goc (=20.000, "Phat sinh No"=0) nen
  // biet chac day la "thu" 20.000d that. Sua thang dong da bi ghi sai nay;
  // KHONG dong den goc re (chenh lech so du luy ke ~2,37 ty van con do, can
  // sao ke day du 21/7-28/7 tu Nhan de doi chieu rieng).
  function fixVib20260729ThuMisreadAsChi(store) {
    const t = store.transactions.find((x) => x.reference === "083z36w-8Awg5iyRR");
    if (!t) return false;
    if (t.type === "thu" && t.amount === 20000) return false; // da dung, khong can sua
    t.type = "thu";
    t.amount = 20000;
    return true;
  }
  if (fixVib20260729ThuMisreadAsChi(store)) changed = true;

  // Chi Nhan, 2026-07-30: doi chieu voi 2 file sao ke DAY DU CA THANG 7 (72.430
  // dong that) Nhan gui de tim goc re chenh lech ~2,37 ty noi tren. Ket qua: KHONG
  // phai loi cong don gi -- BIDV77021 tu truoc gio chi tung nhap duoc cac giao
  // dich "thu" (QR khach tra), con 846 dong that khac (chu yeu la cac dot
  // "CTNB BIDV021-681"/"BIDV021-VP888" -- ngan hang tu dong chuyen bot tien ve
  // BIDV8681/VP58888 dinh ky, tong ~2,35 ty) CHUA TUNG duoc tai len/nhap vao he
  // thong, vi truoc gio chi upload/dan 1 phan sao ke (thuong la loc rieng phan
  // "Co"). Rieng 1 dong DUY NHAT bi SAI ngay tu lan nhap dau tien (01/7, tham
  // chieu "0832DldY-8AFrfsJVU"): ghi nham "chi" 26.145.000d thay vi "thu"
  // 20.000d that (dung 1 kieu loi voi dong VIB o tren, nhung tu 8 ngay truoc khi
  // phat hien) -- sua thang dong nay; con 846 dong con thieu se duoc tu dong bo
  // sung khi Nhan tai lai 2 file sao ke day du thang 7 qua "Tai file sao ke truc
  // tiep tu ngan hang" (an toan, khong trung vi khop theo So tham chieu).
  function fixBidv77021Day1ThuMisreadAsChi(store) {
    const t = store.transactions.find((x) => x.reference === "0832DldY-8AFrfsJVU");
    if (!t) return false;
    if (t.type === "thu" && t.amount === 20000) return false; // da dung, khong can sua
    t.type = "thu";
    t.amount = 20000;
    return true;
  }
  if (fixBidv77021Day1ThuMisreadAsChi(store)) changed = true;

  // Luyen, 2026-07-31: "sao sao kê tháng 7 tôi chọn tk acb 1268 mà lại có
  // giao dịch của 123456" -- ngay 30/7 luc 09:12:30, co 1 dot nhap NHAM 13
  // giao dich (21/7-30/7, chu yeu la settlement Momo "CONG TY CP DICH VU DI
  // DONG TRUC TUYEN") vao bank_id=3 (ACB31268, kenh Zalo/VNPay/Payoo) trong
  // khi CA 13 dong nay da co san 1 ban SAO Y HET (cung so tham chieu/so tien/
  // mo ta) duoi bank_id=4 (BIDV123456) hoac bank_id=11 (BIDV8651, ve "thu"
  // cua 1 khoan CTNB noi bo) -- da doi chieu tung dong, xac nhan 100% la
  // trung, KHONG phai giao dich that cua ACB31268 (tai khoan nay khong lien
  // quan gi Momo). Xoa dung 13 dong nay khoi bank_id=3 theo SO THAM CHIEU
  // (khong dung filter ngay+ngan hang chung vi thang 7 con ~83 giao dich
  // THAT khac cua ACB31268 trong cung khoang ngay, xoa theo filter se xoa
  // nham). Luyen xac nhan xoa 2026-07-31 ("xóa cái trùng nhá").
  function removeAcb31268MomoMisimportBatch(store) {
    const badRefs = new Set([
      "0861EikM-8AlF3ofJp",
      "1131kcay-8Amhw7f5I",
      "086p9NG-8AmoYVNHl",
      "08619UpK-8AoLpOFn2",
      "0831EikM-8Apd6UJqS",
      "0861oTJc-8Apnt1Rp8",
      "0831eILI-8ApqsQcr9",
      "0001401a-8AqULJ5dA",
      "8682HcNQ-8AuCiVQQm",
      "086MpJG-8AuOdjypJ",
      "0861bGq4-8AvvYmLNq",
      "0862HcJS-8AxOSx0Xr",
      "086MpJG-8AyygZnnR",
    ]);
    const acb = store.banks.find((b) => b.name === "ACB31268");
    if (!acb) return false;
    const before = store.transactions.length;
    store.transactions = store.transactions.filter((t) => !(t.bank_id === acb.id && badRefs.has(t.reference)));
    return store.transactions.length !== before;
  }
  if (removeAcb31268MomoMisimportBatch(store)) changed = true;

  // Luyen, 2026-08-01: "cho chỗ nhập tay nữa đi đây gian mưới á LOTTE BAC
  // GIANG PHN đây á" -- gian moi "POSH Lotte Bac Giang" (Ma cua hang
  // TDT9Y5GDVP, kenh VietQR bidv77021, KH Moi) chua co Ma cong trinh trong
  // danh sach chuan (ma_cong_trinh_master.kh_moi) nen khong chon duoc trong
  // dropdown o muc "Ten diem ban chua co Ma cong trinh tuong ung". Them dong
  // moi vao danh sach chuan (giong cach dat ten cac gian PHN khac nhu "GO BAC
  // GIANG PHN"/"KUBO BAC GIANG PHN"/"VC BAC GIANG PHN" da co san) + gan luon
  // "POSH Lotte Bac Giang" -> "LOTTE BAC GIANG PHN" (viet_qr_ten_diem_master,
  // giong co che nut "Gan" tren trang lam thu cong).
  function seedLotteBacGiangGian(store) {
    let didChange = false;
    if (!store.ma_cong_trinh_master) store.ma_cong_trinh_master = {};
    if (!store.ma_cong_trinh_master.kh_moi) store.ma_cong_trinh_master.kh_moi = { rows: [] };
    if (!Array.isArray(store.ma_cong_trinh_master.kh_moi.rows)) store.ma_cong_trinh_master.kh_moi.rows = [];
    const rows = store.ma_cong_trinh_master.kh_moi.rows;
    const NEW_CODE = "LOTTE BAC GIANG PHN";
    if (!rows.some((r) => (r.maCongTrinh || "").trim().toUpperCase() === NEW_CODE)) {
      const maxStt = rows.reduce((mx, r) => {
        const n = parseInt(r.stt, 10);
        return Number.isFinite(n) && n > mx ? n : mx;
      }, 0);
      rows.push({
        maCongTrinh: NEW_CODE,
        stt: String(maxStt + 1),
        tenCongTrinh: NEW_CODE,
        loaiCongTrinh: "",
        tinhTrang: "Đang thực hiện",
        ngayBatDau: "",
        ngayKetThuc: "",
        duToan: "0",
        chuDauTu: "",
        chiNhanh: "CÔNG TY TNHH GIẢI TRÍ K&H",
        trangThai: "Đang sử dụng",
      });
      didChange = true;
    }
    if (!store.viet_qr_ten_diem_master) store.viet_qr_ten_diem_master = {};
    if (!store.viet_qr_ten_diem_master.bidv77021) store.viet_qr_ten_diem_master.bidv77021 = {};
    const key = "posh lotte bac giang"; // normText("POSH Lotte Bắc Giang")
    if (store.viet_qr_ten_diem_master.bidv77021[key] !== NEW_CODE) {
      store.viet_qr_ten_diem_master.bidv77021[key] = NEW_CODE;
      didChange = true;
    }
    return didChange;
  }
  if (seedLotteBacGiangGian(store)) changed = true;

  // Luyen, 2026-08-01: "tôi tải nhầm sao kê lên rồi bạn cx nạp lận à
  // 8670068681 mà tôi nạp nhầm vô tk MB 02865168 xóa cho tôi đi cái tôi mới
  // nạp lên á" -- upload nham file sao ke tai khoan BIDV8681 (8670068681)
  // vao bank_id=17 (MB02865168) ngay 2026-07-31. Xoa DUNG 6 dong cua lan
  // upload sai (id 191390-191395, cung created_at giay 2026-08-01T01:25:58,
  // ca 2 dong con nhac ro so tai khoan "8670068681" trong dien giai), KHONG
  // dung xoa theo bank+ngay+loai chung vi con 4 dong QR MB02865168 THAT hop
  // le cung ngay (id 188046-188049, tao tu hom truoc, 20.000d/dong, doanh
  // thu ve QR that cua chinh MB02865168) khong duoc dung vao.
  function removeMb02865168BidvMisimportBatch(store) {
    const badIds = new Set([191390, 191391, 191392, 191393, 191394, 191395]);
    const before = store.transactions.length;
    store.transactions = store.transactions.filter((t) => !badIds.has(t.id));
    return store.transactions.length !== before;
  }
  if (removeMb02865168BidvMisimportBatch(store)) changed = true;

  // Luyen, 2026-08-01: "check ngân hàng luôn xem có tiền khác ngoài viet qr
  // thì bỏ ra nha" -- phat hien 5 giao dich KHONG PHAI tien ve QR (khong co
  // dinh dang "@VA_V3BLC...VQR..." nhu giao dich VietQR that) dang bi
  // resolveGianGrossByBankRef gom nham vao gian mac dinh "AE HP PHN" (kenh
  // mb02865168), gay hien "Chua co HD" sai lech ~723tr: 191296 (393.081.840d,
  // "REM Tfr Ac:8670068681..." -- chuyen tien tu chinh TK BIDV8681 cua cong
  // ty, noi bo), 191301 (3.130.000d, "YOKIDS TT TIEN MUA GHE MASSAGE..."),
  // 191341 (78.549.000d, "S001...TT 70 doanh thu Game KH va 50 ghe Posh
  // T062026" -- doanh thu gop tu nguon khac, khong phai ve QR), 191386 +
  // 191389 (44.507.521d + 125.458.000d, "FSS...TT HTKD T6.2026..." = thanh
  // toan Hop tac kinh doanh, khong phai ve QR). Luyen xac nhan qua
  // AskUserQuestion: "Loại khỏi VietQR hết" (ca 5 khoan) -- dung dung co che
  // excludeFromVietQrRecon da co san (giong dong "ctnb" da tu dong loai truoc
  // do), tien VAN o lai Giao dich/Sao ke binh thuong, chi khong tinh vao
  // doanh thu QR nua.
  function seedExcludeNonVqrTxFromMb02865168(store) {
    const ids = new Set([191296, 191301, 191341, 191386, 191389]);
    let didChange = false;
    store.transactions.forEach((t) => {
      if (ids.has(t.id) && !t.excludeFromVietQrRecon) {
        t.excludeFromVietQrRecon = true;
        didChange = true;
      }
    });
    return didChange;
  }
  if (seedExcludeNonVqrTxFromMb02865168(store)) changed = true;

  // Luyen, 2026-08-01: "chỗ nội dung nộp tiền á có map với tên gian dựa vào 4
  // sheet này á KVC MN KVC MB MTD MN MTD MB" -- gui 4 file tham khao (KVC
  // MB/MN, MTD MB/MN), moi file 2 sheet (KH cũ "...KH989" + KH mới
  // "...KH705"). store.cht_nop_tien_map dang RONG (tinh nang tu task #54 da
  // co san UI upload rieng tren trang Xuat Hoa Don Ban Ra nhung Luyen chua tung
  // tai file nao qua do) -- nap san 236 ma "Nội dung nộp tiền" -> "mã công
  // trình misa thuế" doc duoc tu 4 file nay bang chinh parser/merge da co san
  // (parseChtNopTienMasterSheet/mergeChtNopTienMap, utils/zvpReconcile.js) de
  // Luyen khong phai tu tai lai qua UI. Idempotent: mergeChtNopTienMap chi
  // "added" khi ma chua ton tai, nen chay lai nhieu lan (vd sau khi Luyen tu
  // tai them file khac qua UI) khong lam gi them / khong ghi de sai.
  function seedChtNopTienMapFromKvcMtdFiles(store) {
    const rows = SEED_CHT_NOP_TIEN_ROWS;
    const { mergeChtNopTienMap } = require("./utils/zvpReconcile");
    const { map, added } = mergeChtNopTienMap(store.cht_nop_tien_map, rows);
    if (added > 0) {
      store.cht_nop_tien_map = map;
      if (!store.cht_nop_tien_uploads) store.cht_nop_tien_uploads = [];
      store.cht_nop_tien_uploads.push({
        id: (store.cht_nop_tien_uploads.length || 0) + 1,
        uploaded_at: new Date().toISOString(),
        file_name: "KVC MB.xlsx + KVC MN.xlsx + MTD MB.xlsx + MTD MN.xlsx (nạp sẵn)",
        sheetsParsed: [
          { sheetName: "KVC MB KH989", rows: 18 },
          { sheetName: "KVC MB KH705", rows: 8 },
          { sheetName: "KVC MN KH705", rows: 5 },
          { sheetName: "KVC MN KH989", rows: 11 },
          { sheetName: "MTD MB KH705", rows: 129 },
          { sheetName: "MTĐ MB KH989", rows: 6 },
          { sheetName: "MTD MN KH705", rows: 52 },
          { sheetName: "MTĐ MN KH989", rows: 9 },
        ],
        rowCount: rows.length,
      });
      return true;
    }
    return false;
  }
  if (seedChtNopTienMapFromKvcMtdFiles(store)) changed = true;

  // Luyen, 2026-08-01 (lan 5): "bạn đọc hợp đồng cho tôi xem ... để khoong
  // khớp hết vậy" -- trong luc dieu tra da phat hien 1 bug khac RIENG, nghiem
  // trong hon: khoa upsert luc upload "Hóa Đơn Đầu Vào" (routes/hoa-don-dau-
  // vao.js, route /upload) truoc gio dung ca "kyHieuHD" (Ky hieu hoa don) --
  // truong nay KHONG ON DINH giua cac lan xuat file khac nhau tu he thong hoa
  // don dien tu (co lan co gia tri "C26TNT", co lan lai RONG cho CUNG 1 hoa
  // don that). Ket qua: 1191/5786 dong (~20%) bi TRUNG THAT tren toan bo du
  // lieu (vd hoa don so 1199, 1187 cua "Go Nha Trang" -- xem chi tiet dieu
  // tra). Da sua khoa upsert (dung ngayHD thay kyHieuHD, xem route) de KHONG
  // TAO TRUNG MOI nua, nhung 1191 dong TRUNG DA CO SAN tu truoc van con --
  // seed nay TU DONG DON DEP 1 LAN moi khi server khoi dong: gop cac dong
  // trung (cung congTy+soHoaDon+ngayHD+dienGiai+soTien) ve LAI 1 dong duy
  // nhat, uu tien giu dong co kyHieuHD (thuong la ban ghi cu hon, day du hon),
  // dong thoi GOP LAI cac truong da nhap tay/lam giau tu dong (daHachToan,
  // daChiTien, gianHang, hinhThucHopTac, taiKhoanCo, maDoiTuongNCC,
  // tenHangHoaMisa, phanLoai, taiKhoanNo, ghiChu, linkHoaDon) -- neu 1 trong 2
  // ban trung da duoc Luyen tick "Da chi tien"/"Da hach toan" hoac dien gian
  // hang tay thi GIU LAI (khong mat), khong chi giu dong nao duoc chon lam
  // "chinh". Idempotent tu nhien: sau lan chay dau, moi khoa chi con 1 dong
  // nen cac lan sau khong tim thay gi de gop nua.
  function dedupeHoaDonDauVaoRows(store) {
    if (!Array.isArray(store.hoa_don_dau_vao) || store.hoa_don_dau_vao.length === 0) return false;
    const groups = new Map();
    store.hoa_don_dau_vao.forEach((r) => {
      const key = [r.congTy, r.soHoaDon, r.ngayHD, r.dienGiai, r.soTien].join("||");
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });
    const toRemoveIds = new Set();
    let mergedGroups = 0;
    groups.forEach((rows) => {
      if (rows.length < 2) return;
      // Uu tien giu dong co kyHieuHD (thuong la ban goc, xuat truoc); neu
      // nhieu/khong dong nao co, giu id nho nhat (cu hon).
      const withKy = rows.filter((r) => (r.kyHieuHD || "").trim());
      const primary = (withKy.length > 0 ? withKy : rows).slice().sort((a, b) => a.id - b.id)[0];
      const dupes = rows.filter((r) => r.id !== primary.id);
      const MERGE_FIELDS = [
        "gianHang",
        "hinhThucHopTac",
        "taiKhoanCo",
        "maDoiTuongNCC",
        "tenHangHoaMisa",
        "phanLoai",
        "taiKhoanNo",
        "ghiChu",
        "linkHoaDon",
        "kyHieuHD",
      ];
      dupes.forEach((d) => {
        MERGE_FIELDS.forEach((f) => {
          if (!primary[f] && d[f]) primary[f] = d[f];
        });
        // Co/false la lua chon THEO Y THUC (co that su lam) -- OR lai de
        // khong mat viec Luyen da tick tren BAT KY ban trung nao.
        if (d.daHachToan) primary.daHachToan = true;
        if (d.daChiTien) primary.daChiTien = true;
        toRemoveIds.add(d.id);
      });
      mergedGroups++;
    });
    if (toRemoveIds.size === 0) return false;
    store.hoa_don_dau_vao = store.hoa_don_dau_vao.filter((r) => !toRemoveIds.has(r.id));
    console.log(
      `[seed] Da don dep ${toRemoveIds.size} dong hoa don dau vao TRUNG THAT (${mergedGroups} nhom, gop lai con 1 dong/nhom).`
    );
    return true;
  }
  if (dedupeHoaDonDauVaoRows(store)) changed = true;

  // Luyen, 2026-08-03: "sao số dư lại âm 40 mấy triệu vậy trên soa kê số dư
  // đầu kì là 0 mà" -- phat hien qua vi du that: sao ke BIDV7701 hien so du
  // am 43.803.988d ngay sau giao dich dau tien cua thang 4, dù bank.opening_
  // balance = 0. Nguyen nhan: 2 giao dich "chi" (12.808.994d, 34.086.994d)
  // co ngay "1899-12-30" -- day la Excel epoch/serial 0, DAU HIEU KINH DIEN
  // cua 1 o ngay KHONG DOC DUOC trong file upload (truoc khi utils/
  // bankStatementParser.js co them guard isPlausibleYear() chan dieu nay tai
  // nguon, xem ham parseDateCell) -- khong phai giao dich that, chi la rac
  // con sot lai TU LAN UPLOAD 2026-07-17 (truoc khi parser duoc sua). Luyen
  // xac nhan qua sao ke ngan hang that: so du dau ky = 0.00d dung nhu he
  // thong da hien, KHONG co khoan "nam truoc" nao ca -- an toan xoa het
  // (khong phai chuyen vao so du dau ky). Quet CA 6 tai khoan bi anh huong
  // (BIDV7701 x2, BIDV7702 x2, BIDV123456, BIDV7704, BIDV8651), khong rieng
  // BIDV7701, vi day la loi parser CHUNG anh huong nhieu lan upload khac
  // nhau, khong phai rieng 1 tai khoan.
  function removeCorruptEpochDateTransactions(store) {
    const before = (store.transactions || []).length;
    store.transactions = (store.transactions || []).filter(
      (t) => !(t.date && t.date.slice(0, 7) === "1899-12")
    );
    const removed = before - store.transactions.length;
    if (removed > 0) {
      console.log(
        `[seed] Da xoa ${removed} giao dich rac (ngay "1899-12-xx", artifact cua loi parser cu truoc khi co guard nam hop le 2000-2100).`
      );
    }
    return removed > 0;
  }
  if (removeCorruptEpochDateTransactions(store)) changed = true;

  // Luyen, 2026-08-03: "tôi đang chọn đối soát 7702 sao lại có 7701 tỏng đây
  // nữa" -- phat hien qua sao ke chi tiet theo ngan hang, tab BIDV7702: nhieu
  // dong "REM 9901CI..." / "CTNB BIDV701-681" / "THU PHI QLTK ... TK
  // 8620107701" (8620107701 la SO TAI KHOAN THAT cua BIDV7701, khong phai
  // 7702) xuat hien CA duoi bank_id=14 (BIDV7702) LAN duoi bank_id=13
  // (BIDV7701, dung cho). Kiem tra: toan bo 14 dong duoi day o bank_id=14
  // (batch tai len ngay 2026-07-18) TRUNG KHOP TUYET DOI (cung ngay + cung
  // loai thu/chi + cung so tien + cung dien giai) voi 1 dong da co san o
  // bank_id=13 -- xac nhan day la du lieu BIDV7701 bi tai NHAM vao bank_id
  // cua BIDV7702 trong dot tai 07-18, lam Tong thu/chi cua BIDV7702 bi thoi
  // phong. Xoa CHINH XAC 14 id nay khoi bank_id=14 (giu nguyen ban goc dung o
  // bank_id=13). CO Y THUC bo qua 2 giao dich khac cung ngay 07-18 co CUNG so
  // tham chieu nhung SO TIEN KHAC nhau voi ban ghi o bank_id=13 (id 52451 ref
  // 9901CI260701000104456: 167.255.571d o bank14 vs 4.352.822d o bank13 cung
  // ngay 07-01; id 52475 ref 9901CI260715000084343: 3.798.682d o bank14 vs
  // 199.448.132d o bank13 cung ngay 07-15) -- day la XUNG DOT so lieu (cung so
  // tham chieu, khac so tien), KHONG PHAI trung lap don gian, can Luyen xac
  // nhan so nao dung truoc khi sua, nen KHONG dong den 2 giao dich nay.
  function removeMisfiledBidv7701TransactionsFromBidv7702(store) {
    const idsToRemove = new Set([
      52452, 52455, 52456, 52459, 52461, 52462, 52463, 52464, 52466, 52468,
      52470, 52472, 52473, 52474,
    ]);
    const before = (store.transactions || []).length;
    store.transactions = (store.transactions || []).filter((t) => !idsToRemove.has(t.id));
    const removed = before - store.transactions.length;
    if (removed > 0) {
      console.log(
        `[seed] Da xoa ${removed} giao dich BIDV7701 bi tai NHAM vao bank_id cua BIDV7702 (batch 2026-07-18, trung khop tuyet doi voi du lieu da co o BIDV7701).`
      );
    }
    return removed > 0;
  }
  if (removeMisfiledBidv7701TransactionsFromBidv7702(store)) changed = true;

  // Luyen, 2026-08-03: "bên KH mới mã công trình vũng tàu này đổi lại thành
  // VUNG TAU PHCM nhá" -- ma cong trinh chuan cua KH Moi cho diem nay VON DA
  // la "VUNG TAU PHCM" (ten hien thi "POSH LOTTE MART VUNG TAU", xem
  // store.ma_cong_trinh_master.kh_moi) nhung 2 dong invoice_diem_alias
  // "VUNG TAU PHCM" / "POSH LOTTE MART VUNG TAU" -> "KVC LOTTE VUNG TAU" (alias
  // DUNG CHUNG cho Momo/ZVP/VietQR, xem store.invoice_diem_alias) dang CHAY
  // SAU applyGianRedirectToInvoices trong buildChannelReconciliation
  // (routes/doisoat-vietqr.js), lam hoa don BIDV7702 vua duoc doi dung thanh
  // "VUNG TAU PHCM" lai bi doi NGUOC lai thanh "KVC LOTTE VUNG TAU" (ma cua MOT
  // gian KHAC hoan toan ben KH Cu, xem store.ma_cong_trinh_master.kh_cu) --
  // day chinh la nguyen nhan file xuat MISA bao "Công trình <KVC LOTTE VUNG
  // TAU> không có trong danh mục". Xac nhan KHONG co hoa don KH Cu (momo_invoices/
  // zvp_invoices.zalo) nao dung raw maDiem "VUNG TAU PHCM"/"POSH LOTTE MART VUNG
  // TAU" ca (chi dung thang "KVC LOTTE VUNG TAU" cua chinh no), nen 2 alias nay
  // khong phuc vu gi cho KH Cu, chi dang lam hong KH Moi -- xoa han.
  function fixVungTauPhcmInvoiceDiemAlias(store) {
    if (!store.invoice_diem_alias) return false;
    let didFix = false;
    if (store.invoice_diem_alias["VUNG TAU PHCM"] === "KVC LOTTE VUNG TAU") {
      delete store.invoice_diem_alias["VUNG TAU PHCM"];
      didFix = true;
    }
    if (store.invoice_diem_alias["POSH LOTTE MART VUNG TAU"] === "KVC LOTTE VUNG TAU") {
      delete store.invoice_diem_alias["POSH LOTTE MART VUNG TAU"];
      didFix = true;
    }
    if (didFix) {
      console.log(
        '[seed] Da xoa alias sai "VUNG TAU PHCM"/"POSH LOTTE MART VUNG TAU" -> "KVC LOTTE VUNG TAU" (dang lam hoa don BIDV7702/KH Moi bi doi nguoc ve ma sai).'
      );
    }
    return didFix;
  }
  if (fixVungTauPhcmInvoiceDiemAlias(store)) changed = true;

  // Cung ghi chu tren: rieng PHIA SETTLEMENT (tien ngan hang thuc ve, khong
  // phai hoa don) cua BIDV7702 con 1 nguon khac doc lap voi invoice_diem_alias
  // -- store.viet_qr_ten_diem_master.bidv7702 (bang "Ten diem ban" -> "Ma cong
  // trinh", dung boi resolveGianGrossByBankRef cho giai doan >= refMatchFrom,
  // xem utils/vietqrReconcile.js) -- key "lotte vung tau" dang tro SAI ve
  // "KVC LOTTE VUNG TAU" thay vi "VUNG TAU PHCM", khien tien VietQR that ve
  // van bi gan nham gian nay o ca phia doi soat (khong chi phia hoa don), du
  // da sua invoice_diem_alias o tren.
  function fixVungTauPhcmTenDiemMaster(store) {
    const tdm = store.viet_qr_ten_diem_master && store.viet_qr_ten_diem_master.bidv7702;
    if (!tdm) return false;
    let didFix = false;
    Object.keys(tdm).forEach((k) => {
      if (tdm[k] === "KVC LOTTE VUNG TAU") {
        tdm[k] = "VUNG TAU PHCM";
        didFix = true;
      }
    });
    if (didFix) {
      console.log(
        '[seed] Da sua viet_qr_ten_diem_master.bidv7702["lotte vung tau"] tu "KVC LOTTE VUNG TAU" sai thanh "VUNG TAU PHCM" dung (phia doi soat tien ve, khong phai phia hoa don).'
      );
    }
    return didFix;
  }
  if (fixVungTauPhcmTenDiemMaster(store)) changed = true;

  // Cung ghi chu tren: 1 hoa don rieng le trong viet_qr_invoices.bidv7702 (so
  // HD 10031, ngay 2026-07-20) co san mot loi go nham CU THE hon -- raw
  // maDiem da la "KVC LOTTE VUNG TAU" (khong phai qua alias) trong khi tenDiem
  // van dung la "POSH LOTTE MART VUNG TAU" giong 12 hoa don cung diem khac --
  // sua lai maDiem cho khop tenDiem.
  function fixStrayVungTauInvoiceMaDiem(store) {
    const invoices = store.viet_qr_invoices && store.viet_qr_invoices.bidv7702;
    if (!Array.isArray(invoices)) return false;
    let didFix = false;
    invoices.forEach((inv) => {
      if (inv.maDiem === "KVC LOTTE VUNG TAU" && inv.tenDiem === "POSH LOTTE MART VUNG TAU") {
        inv.maDiem = "POSH LOTTE MART VUNG TAU";
        didFix = true;
      }
    });
    if (didFix) {
      console.log('[seed] Da sua 1 hoa don BIDV7702 co maDiem go nham "KVC LOTTE VUNG TAU" ve dung "POSH LOTTE MART VUNG TAU".');
    }
    return didFix;
  }
  if (fixStrayVungTauInvoiceMaDiem(store)) changed = true;

  // Luyen, 2026-08-03: "xóa cái cấn trừ 7702 đi giờ tôi sẽ check từng gian sai
  // số lệch nhá" -- hoi lai xac nhan "chỉ tháng 6 thôi" (khong dong den thang
  // 7). Trang doi soat VietQR co nut "cross-match" tu dong cap gian bi xuat
  // chung 1 hoa don (vd CON DAO AIRPORT PHCM <-> SB CAN THO PHCM) -- 40 dong
  // viet_qr_manual_matches.bidv7702 voi key "2026-06-xx|..." deu co created_at
  // = hom nay (2026-08-03), tuc la vua duoc tao boi 1 lan bam nut cross-match
  // gan day, con 113 dong con lai (2026-07-xx) tao tu 07-20/07-30 (da duoc
  // Luyen xac nhan tu truoc) thi GIU NGUYEN dung yeu cau. Xoa CHINH XAC cac
  // key bat dau "2026-06" de Luyen tu kiem tra lai tung gian thang 6 tu dau
  // (khong bi so lieu da cap tu dong lam mo).
  function removeBidv7702JuneCrossMatches(store) {
    const mm = store.viet_qr_manual_matches && store.viet_qr_manual_matches.bidv7702;
    if (!mm) return false;
    let removed = 0;
    Object.keys(mm).forEach((key) => {
      if (key.startsWith("2026-06")) {
        delete mm[key];
        removed++;
      }
    });
    if (removed > 0) {
      console.log(
        `[seed] Da xoa ${removed} dong cap gian tu dong (cross-match) thang 6/2026 cua BIDV7702 theo yeu cau Luyen (giu nguyen thang 7).`
      );
    }
    return removed > 0;
  }
  if (removeBidv7702JuneCrossMatches(store)) changed = true;

  // Luyen, 2026-08-03: "hóa đơn ngày 3,4 này nè nếu hk có tách xuất dồn thì
  // hk cần cộng đâu xuất lẻ như các ngày thông thường á dựa vào ngày xuất
  // với Dịch vụ thu hộ xem nó xuất ngày mấy" -- doi chieu truc tiep voi file
  // hoa don goc "MỚi XHD.xlsx" (Luyen goi la chuan) xac nhan: ham
  // fixBidv7702Day3TaggedAsDay4 (utils/vietqrReconcile.js, chay lai MOI LAN
  // load()/moi lan vao trang doi soat vietqr) dang CHUYEN NHAM 29 hoa don
  // thang 6 (soHd 4827-4855, "Ngày HĐ" 05/06, tag goc DUNG la "MTD MN 4" =
  // doanh thu ngay 4) tu ngay 4 sang ngay 3 -- day chinh la nguyen nhan man
  // hinh Luyen gui: ngay 03/06 du +18.960.000d "Lệch Dữ liệu↔Hóa đơn", ngay
  // 04/06 thieu dung -19.320.000d (= tong 29 hoa don nay), tat ca gian ngay 4
  // hien "Chưa có HĐ". Ham do von chi dung 1 lan cho 1 lo hoa don CU (da xac
  // nhan boi Chi Nhan 2026-07-30, tong dung 20.000.000d = tien ngan hang
  // ngay 3) nhung khong gioi han pham vi nen tiep tuc "sua" ca cac hoa don
  // MOI, dan tag DUNG "MTD MN 4" sau nay (thang 5 va thang 6) -- ham nay da
  // duoc vo hieu hoa (xem utils/vietqrReconcile.js). O day chi khoi phuc lai
  // CHINH XAC 29 hoa don thang 6 (theo dung yeu cau "chỉ tháng 6 thôi") ve
  // dung ngay 4 nhu file hoa don goc; 20 hoa don thang 5 (soHd 1746-1765,
  // cung loi) CHUA dong den, cho Luyen xac nhan rieng vi ngoai pham vi thang
  // 6 da yeu cau.
  const BIDV7702_JUNE_DAY4_SOHD = [
    4827, 4828, 4829, 4830, 4831, 4832, 4833, 4834, 4835, 4836, 4837, 4838, 4839, 4840, 4841, 4842, 4843, 4844, 4845,
    4846, 4847, 4848, 4849, 4850, 4851, 4852, 4853, 4854, 4855,
  ];
  function restoreBidv7702JuneDay4Invoices(store) {
    const invoices = store.viet_qr_invoices && store.viet_qr_invoices.bidv7702;
    if (!Array.isArray(invoices)) return false;
    let restored = 0;
    invoices.forEach((inv) => {
      if (BIDV7702_JUNE_DAY4_SOHD.includes(inv.soHd) && inv.raw === "MTD MN 3") {
        inv.days = [4];
        inv.raw = "MTD MN 4";
        restored++;
      }
    });
    if (restored > 0) {
      console.log(
        `[seed] Da khoi phuc ${restored} hoa don thang 6/2026 BIDV7702 tu ngay 3 ve dung ngay 4 (theo file hoa don goc, chi tach bi chuyen nham boi fixBidv7702Day3TaggedAsDay4).`
      );
    }
    return restored > 0;
  }
  if (restoreBidv7702JuneDay4Invoices(store)) changed = true;

  // Luyen, 2026-08-03: "7702 á KH mới đổi KUBO BA RIA PHCM thành KUBO GO BA
  // RIA PHCM nhá" -- man hinh nhap "Thu tiền gửi từ Excel" bao 9 dong "Công
  // trình <KUBO BA RIA PHCM> không có trong danh mục" vi danh muc MISA cua
  // Luyen chi co ma "KUBO GO BA RIA PHCM" (dung voi ma_cong_trinh_master.
  // kh_moi hang stt 106, da dung tu truoc). Con 3 cho khac trong store van
  // dung nham ma cu (thieu "GO"): gian_mapping (key), invoice_diem_alias (vi
  // tri "POSH MN KUBO GO BÀ RỊA" tro sai ve ma cu), viet_qr_ten_diem_master.
  // bidv7702 (key "kubo ba ria" tro sai ve ma cu), va 1 dong viet_qr_manual_
  // matches.bidv7702 dan theo key cu. Doi tat ca ve dung "KUBO GO BA RIA
  // PHCM" (khong dong den viet_qr_invoices vi khong co hoa don nao dang dung
  // dung ma cu -- da la "KUBO GO BA RIA PHCM" hoac ten tho "POSH MN KUBO GO
  // BÀ RỊA" roi).
  function renameKuboBaRiaToKuboGoBaRiaPhcm(store) {
    const OLD_CODE = "KUBO BA RIA PHCM";
    const NEW_CODE = "KUBO GO BA RIA PHCM";
    let didFix = false;

    if (store.invoice_diem_alias && store.invoice_diem_alias["POSH MN KUBO GO BÀ RỊA"] === OLD_CODE) {
      store.invoice_diem_alias["POSH MN KUBO GO BÀ RỊA"] = NEW_CODE;
      didFix = true;
    }

    if (store.gian_mapping && Object.prototype.hasOwnProperty.call(store.gian_mapping, OLD_CODE)) {
      if (!Object.prototype.hasOwnProperty.call(store.gian_mapping, NEW_CODE)) {
        store.gian_mapping[NEW_CODE] = store.gian_mapping[OLD_CODE];
      }
      delete store.gian_mapping[OLD_CODE];
      didFix = true;
    }

    const tdm = store.viet_qr_ten_diem_master && store.viet_qr_ten_diem_master.bidv7702;
    if (tdm && tdm["kubo ba ria"] === OLD_CODE) {
      tdm["kubo ba ria"] = NEW_CODE;
      didFix = true;
    }

    const mm = store.viet_qr_manual_matches && store.viet_qr_manual_matches.bidv7702;
    if (mm) {
      Object.keys(mm).forEach((key) => {
        if (key.endsWith("|" + OLD_CODE)) {
          const newKey = key.slice(0, -OLD_CODE.length) + NEW_CODE;
          if (!mm[newKey]) mm[newKey] = mm[key];
          delete mm[key];
          didFix = true;
        }
      });
    }

    if (didFix) {
      console.log(
        `[seed] Da doi ma cong trinh "${OLD_CODE}" thanh "${NEW_CODE}" cho BIDV7702/KH Moi (gian_mapping, invoice_diem_alias, viet_qr_ten_diem_master, viet_qr_manual_matches).`
      );
    }
    return didFix;
  }
  if (renameKuboBaRiaToKuboGoBaRiaPhcm(store)) changed = true;

  // Chi Nhan, 2026-08-05: "check lại vn pay offline luôn nhá số vẫn lệch nè"
  // -- TK ACB31268 (tai khoan nhan tien Zalo App/VNPay/Payoo, xem ZVP_BANK_NAME
  // trong routes/doisoat-zvp.js) nhieu lan bi ghi NHAM 1 vai dong settlement
  // thanh loai "chi" thay vi "thu" (parser sao ke doc nham cot lam giao dich
  // "tien ve" bi hieu thanh "tien ra"), khien extractZvpSettlements (chi xet
  // t.type === "thu") LOAI HOAN TOAN cac dong nay khoi "Ngan hang" cua trang
  // Doi soat Zalo/VNPay/Payoo -- xac nhan qua 2 lan doi chieu truc tiep voi
  // sao ke goc (file "12131268_SAOKE_TK_20260729-20260805.xlsx": MOI dong deu
  // la "(+) tien gui vao", KHONG co dong "rut ra" nao ca, nen dong nao mang
  // mo ta VNPay/Payoo settlement ma dang ghi "chi" chac chan la sai). Da tung
  // sua truc tiep 2 lan nhung deu bi MAT lai ngay sau do -- ly do: server
  // dang chay tren may Chi Nhan (localhost:3000) giu san 1 ban store.json
  // CU trong bo nho (nap luc no khoi dong, TRUOC khi duoc sua), nen bat ky
  // luc nao no goi save() (vd luc Chi Nhan tai/dan them sao ke khac) la no
  // ghi de ban CU do len tren, xoa mat cac cho vua sua truc tiep tren dia --
  // dung y het co che da giai thich day du tai ZVP_GIAN_LIST_BAD_REDIRECTS
  // (routes/doisoat-zvp.js). Chuyen han sang seed tu dong CHAY LAI + SUA MOI
  // LAN load() (giong moi hang so *_DEFAULTS khac) de KHONG BAO GIO bi mat
  // nua, du server cu co ghi de bao nhieu lan.
  //
  // Rieng 2 dong sau la "phantom" -- SAI HOAN TOAN ve so tien (khong phai chi
  // sai loai thu/chi), sinh ra tu 1 lan doc sao ke bi loi truoc do, da co dong
  // DUNG thay the roi nen xoa han thay vi sua lai loai:
  //  - 213.884.659d ngay 03/08 (mo ta "...NGAY 31.07-02.08.26") -- dong dung
  //    thay the la 138.182.369d/thu (doi chieu dung GD so 3259 tren sao ke goc).
  //  - 332.537.280d ngay 29/07 (mo ta "...NGAY 28.07.26") -- dong dung da co
  //    san tu truoc la 19.739.748d/thu (GD so 3245), dong nay thua/sai hoan toan.
  const ZVP_SETTLEMENT_PHANTOM_ROWS = [
    { date: "2026-08-03", amount: 213884659, descIncludes: "NGAY 31.07-02.08.26" },
    { date: "2026-07-29", amount: 332537280, descIncludes: "NGAY 28.07.26" },
  ];
  function fixZvpSettlementChiToThu(store) {
    const bank = store.banks.find((b) => b.name === "ACB31268");
    if (!bank) return false;
    let didFix = false;

    const before = store.transactions.length;
    store.transactions = store.transactions.filter((t) => {
      if (t.bank_id !== bank.id) return true;
      const isPhantom = ZVP_SETTLEMENT_PHANTOM_ROWS.some(
        (p) => t.date === p.date && Number(t.amount) === p.amount && (t.description || "").includes(p.descIncludes)
      );
      return !isPhantom;
    });
    if (store.transactions.length !== before) didFix = true;

    const SETTLEMENT_PATTERN = /(DV\s+CTT\s+NGAY|DV\s+QR\s+OFFLINE\s+NGAY|PAYOO.*TT\s+TD\s+NGAY)/i;
    store.transactions.forEach((t) => {
      if (t.bank_id === bank.id && t.type === "chi" && SETTLEMENT_PATTERN.test(t.description || "")) {
        t.type = "thu";
        didFix = true;
      }
    });

    return didFix;
  }
  if (fixZvpSettlementChiToThu(store)) changed = true;

  // Luyen, 2026-08-05: "ngày 1 với ngày 2 tổng là 70.040 mà sao lại trên wed
  // có 70.020" -- doi chieu voi sao ke goc BIDV7702 (file
  // "20260805_SAOKE_8640107702_...xlsx"): CA 2 ngay 01-02/08/2026 chi co dong
  // "(+) Phat sinh co" (tien vao), KHONG co dong "Phat sinh no" (chi) nao ca
  // (0d rut ra ca 2 ngay). Nhung tren web dang hien them 1 dong "chi"
  // 57.700.000d ngay 01/08 (id 188474) -- doi chieu dung mo ta/so tham chieu
  // ("8681rDcq-8B1J2pmqe") voi dong DAU TIEN cua sao ke goc thi day PHAI la 1
  // dong "thu" 20.000d (VU DUC ANH, QR VQR26316D494VM8S...), khong phai "chi"
  // 57.700.000d -- ro rang la loi doc/nhap sao ke (sai CA loai VA so tien,
  // giong dung kieu loi da gap voi ACB31268 o tren). Sua lai dung theo sao ke
  // goc; chay lai + tu sua MOI LAN load() (khong sua truc tiep 1 lan) vi cung
  // co nguy co bi may Luyen ghi de lai neu chi sua file tren dia 1 lan, dung
  // co che da giai thich day du o fixZvpSettlementChiToThu ngay tren.
  function fixBidv7702VuDucAnhAmount(store) {
    const bank = store.banks.find((b) => b.name === "BIDV7702");
    if (!bank) return false;
    let didFix = false;
    const t = store.transactions.find(
      (t) => t.bank_id === bank.id && t.id === 188474 && t.reference === "8681rDcq-8B1J2pmqe"
    );
    if (t && (t.type !== "thu" || Number(t.amount) !== 20000)) {
      t.type = "thu";
      t.amount = 20000;
      didFix = true;
    }
    return didFix;
  }
  if (fixBidv7702VuDucAnhAmount(store)) changed = true;

  // Luyen, 2026-08-05: "hóa đơn của vũng tàu nè ngày 1 với ngày 2 á" -- hoa
  // don gan day cua diem Vung Tau ghi "Ma diem" la "POSH LOTTE MART VUNG TAU"
  // (ten day du tu phan mem ke toan), khac voi ma gian dang dung ben doanh
  // thu ("VUNG TAU PHCM"), nen TOAN BO hoa don loai nay (khong chi 2 ngay 01-
  // 02/08, ca cac ngay truoc do) bi hien "Chua co HD" oan du tien ve dung.
  // Luu qua invoice_diem_alias (co che co san, dung chung Momo/ZVP/VietQR) --
  // tu sua lai MOI LAN load() vi cung co nguy co bi may Luyen ghi de (da xac
  // nhan xay ra that voi ban sua truc tiep truoc do), giong het co che
  // fixZvpSettlementChiToThu o tren.
  function seedInvoiceDiemAliasVungTau(store) {
    if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
    if (store.invoice_diem_alias["POSH LOTTE MART VUNG TAU"] === "VUNG TAU PHCM") return false;
    store.invoice_diem_alias["POSH LOTTE MART VUNG TAU"] = "VUNG TAU PHCM";
    return true;
  }
  if (seedInvoiceDiemAliasVungTau(store)) changed = true;

  // Luyen, 2026-08-05: "2 hóa đơn của estella zalo app đối soát 1,2 đây nhá"
  // -- hoa don Zalo App cua Funzone-tau Estella dang di qua zvp_gian_list ve
  // ma "KVC ESTELLA", nhung doanh thu Online (zvp_online_product_map) cua
  // dung san pham nay lai dang ve ma "DIY ESTELLA KVC" -- 2 ma khac nhau cho
  // CUNG 1 diem nen hoa don khong bao gio khop duoc voi doanh thu (700k+
  // moi ky, xac nhan dung bang 349.000d/hoa don x 2). Doi lai ca 2 huong
  // (tenDiem "Funzone-tàu Estella" va fallback "KVC ESTELLA") ve thang
  // "DIY ESTELLA KVC" (ma dang dung ben doanh thu) de khop lai.
  function seedZvpGianListEstellaFix(store) {
    if (!store.zvp_gian_list) store.zvp_gian_list = [];
    let didFix = false;
    store.zvp_gian_list.forEach((g) => {
      if (g.tenDiem === "Funzone-tàu Estella" && g.maCongTrinh !== "DIY ESTELLA KVC") {
        g.maCongTrinh = "DIY ESTELLA KVC";
        didFix = true;
      }
    });
    const hasFallback = store.zvp_gian_list.some(
      (g) => g.tenDiem === "KVC ESTELLA" && g.maCongTrinh === "DIY ESTELLA KVC"
    );
    if (!hasFallback) {
      store.zvp_gian_list.push({ tenDiem: "KVC ESTELLA", maCongTrinh: "DIY ESTELLA KVC", isCse: false });
      didFix = true;
    }
    return didFix;
  }
  if (seedZvpGianListEstellaFix(store)) changed = true;

  // Luyen, 2026-08-05: "2 hóa đơn estella ngày 1 2 nè thêm vô cho tôi trên
  // web đi" -- sau khi sua zvp_gian_list o tren, DIY ESTELLA KVC van hien
  // "Chưa có HĐ" vi co 1 dong invoice_diem_alias CU: "DIY ESTELLA KVC" ->
  // "KVC ESTELLA" (nguoc chieu, khong dung cho hoa don thuc te nao ca -- da
  // kiem tra ca 3 kenh zalo/vnpay/payoo, khong co hoa don nao co maDiem =
  // "DIY ESTELLA KVC" that ca) -- dong alias nay chi bi kich hoat NHU 1 TAC
  // DUNG PHU khi reconcileZvpChannel doc lai maDiem SAU KHI applyGianRedirectToInvoices
  // da doi "KVC ESTELLA" -> "DIY ESTELLA KVC" (xem buildReconciliation), roi
  // dong alias nay lai doi NGUOC VE "KVC ESTELLA", tu xoa sach fix
  // seedZvpGianListEstellaFix o tren. Xoa han dong alias thua nay.
  function seedRemoveEstellaBadAlias(store) {
    if (!store.invoice_diem_alias) return false;
    if (store.invoice_diem_alias["DIY ESTELLA KVC"] === undefined) return false;
    delete store.invoice_diem_alias["DIY ESTELLA KVC"];
    return true;
  }
  if (seedRemoveEstellaBadAlias(store)) changed = true;

  // Luyen, 2026-08-05: "sao cái côn đảo ngày 1/08 vẫn có 330k vậy" -- fix
  // truoc do (them 1 dong QR tho bi thieu, ref "8681rDcq-8B1J2pmqe", 20.000đ,
  // ma cua hang "6PWLTSTSWP" ngay 2026-08-01) chi ghi THANG vao store.json,
  // KHONG phai seed function -- da bi may local cua Luyen (dang chay san voi
  // ban nho cu, chua co dong nay) ghi de mat khi luu bat ky thay doi nao
  // (dung y het pattern da gap nhieu lan trong session nay). Bien thanh seed
  // tu vá: kiem tra dong QR tho co ref nay chua, neu chua thi them 1 "upload"
  // gia (chi 1 dong) chua no -- an toan tuyet doi voi mergeRawRows (de-dup
  // theo vqrCode+date+amount+raw, dong nay la duy nhat nen khong trung ai).
  function seedBidv7702ConDaoMissingTx(store) {
    if (!store.viet_qr_raw_uploads || !store.viet_qr_raw_uploads.bidv7702) return false;
    const uploads = store.viet_qr_raw_uploads.bidv7702;
    const targetRef = "8681rDcq-8B1J2pmqe";
    const already = uploads.some((u) => (u.rows || []).some((r) => r.refCode === targetRef));
    if (already) return false;
    uploads.push({
      id: "seed-condao-fix-1",
      uploaded_at: "2026-08-05T00:00:00.000Z",
      file_name: "seed-fix: giao dich thieu SAN BAY CON DAO 20.000d (2026-08-01)",
      sheetName: "dữ liệu",
      rows: [
        {
          vqrCode: "VQR26317T5AN6NR",
          maCuaHang: "6PWLTSTSWP",
          amount: 20000,
          date: "2026-08-01",
          raw: "VQR26317T5AN6NR PaymentForOrder",
          refCode: targetRef,
        },
      ],
    });
    return true;
  }
  if (seedBidv7702ConDaoMissingTx(store)) changed = true;

  // Luyen, 2026-08-05: cung ly do nhu tren -- 2 mapping "Ten diem - Ma cong
  // trinh" cho gian CGV moi mo (Ly Chinh Thang / Pearl Plaza, BIDV7702) chi
  // ghi THANG vao store.json, co nguy co bi may local ghi de mat truoc khi
  // Luyen restart. Bien thanh seed de tu vá lai neu bi mat.
  function seedBidv7702CgvTenDiemMaster(store) {
    if (!store.viet_qr_ten_diem_master) return false;
    if (!store.viet_qr_ten_diem_master.bidv7702) store.viet_qr_ten_diem_master.bidv7702 = {};
    const map = store.viet_qr_ten_diem_master.bidv7702;
    let did = false;
    if (map["cgv ly chinh thang"] !== "POSH MN CGV LÝ CHÍNH THẮNG") {
      map["cgv ly chinh thang"] = "POSH MN CGV LÝ CHÍNH THẮNG";
      did = true;
    }
    if (map["cgv peal palaza"] !== "POSH MN CGV PEARL PLAZA") {
      map["cgv peal palaza"] = "POSH MN CGV PEARL PLAZA";
      did = true;
    }
    return did;
  }
  if (seedBidv7702CgvTenDiemMaster(store)) changed = true;

  // Luyen, 2026-08-05: "hóa đơn đây soa lại kh lưu dc á gán cho tôi luôn đi"
  // -- BIDV77020, gian "SB CAM RANH PHN" ngay 2026-08-01, hoa don 2552
  // (100.000đ). Cung nguy co bi ghi de mat nhu 2 fix tren -- bien thanh seed.
  function seedBidv77020SbCamRanhManualMatch(store) {
    if (!store.viet_qr_manual_matches) return false;
    if (!store.viet_qr_manual_matches.bidv77020) store.viet_qr_manual_matches.bidv77020 = {};
    const key = "2026-08-01|SB CAM RANH PHN";
    const cur = store.viet_qr_manual_matches.bidv77020[key];
    if (cur && Array.isArray(cur.invoiceNumbers) && cur.invoiceNumbers.includes("2552") && cur.amount === 100000) {
      return false;
    }
    store.viet_qr_manual_matches.bidv77020[key] = {
      invoiceNumbers: ["2552"],
      amount: 100000,
      grossAdjustment: 0,
      note: "",
      created_at: "2026-08-05T00:00:00.000Z",
    };
    return true;
  }
  if (seedBidv77020SbCamRanhManualMatch(store)) changed = true;

  // Luyen, 2026-08-05: "đối soát chưa khớp này là của gian SB CAM RANH PHN
  // này á map cho tôi vô các cửa hàng vô cái SB CAM RANH PHN này luôn á" --
  // phat hien: file "Ten diem - Ma cong trinh" nguon co dong Cam Ranh voi
  // GIA TRI COT MA CONG TRINH la chinh chu "Chưa khớp" (chu khong phai de
  // trong/thieu dong) -- gia tri rac nay bi nhap THANG vao
  // viet_qr_ten_diem_master lam 3 ten diem ban ("POSH Sân bay Cam ranh", "1
  // JP SB Cam Ranh.new", "POSH Sân bay Quốc Tế Cam Ranh") deu tro toi 1 "ma
  // cong trinh" ten la "Chưa khớp" (khong phai "chua map", ma la DA MAP NHUNG
  // map sai vao 1 chuoi rac) -- hien thanh 1 dong gian ten "Chưa khớp" tren
  // doi soat, khong bao gio khop hoa don. Anh huong CA 3 kenh dung chung file
  // nguon nay (bidv77021, mb02865168, bidv8613600999 -- xac nhan qua
  // seedFromBidv77021). Sua ve dung "SB CAM RANH PHN" nhu Luyen xac nhan; seed
  // de tu vá lai neu file nguon (van con gia tri rac) duoc tai lai sau nay.
  function seedFixChuaKhopCamRanhTenDiem(store) {
    if (!store.viet_qr_ten_diem_master) return false;
    const badKeys = ["posh san bay cam ranh", "1 jp sb cam ranh.new", "posh san bay quoc te cam ranh"];
    let did = false;
    Object.keys(store.viet_qr_ten_diem_master).forEach((ch) => {
      const master = store.viet_qr_ten_diem_master[ch];
      badKeys.forEach((k) => {
        if (master[k] === "Chưa khớp" || master[k] === "Chua khop") {
          master[k] = "SB CAM RANH PHN";
          did = true;
        }
      });
    });
    return did;
  }
  if (seedFixChuaKhopCamRanhTenDiem(store)) changed = true;

  // Xoa file rac MOI LAN KHOI DONG de giai phong Volume:
  // 1. .bak files: moi lan restore tao 1 file .bak 138MB, giu toi da 2 ban moi nhat.
  // 2. .tmp-* files: cac lan restore/save bi crash (OOM/ENOSPC) de lai file
  //    .tmp-PID-timestamp-... chua duoc xoa -- moi file co the nang 95-138MB.
  try {
    const allFiles = fs.readdirSync(DATA_DIR);

    // Xoa tat ca .tmp-* files (phat sinh khi process crash giua chung write)
    const tmpFiles = allFiles.filter((f) => /\.tmp[-.]/i.test(f));
    if (tmpFiles.length > 0) {
      tmpFiles.forEach((f) => { try { fs.unlinkSync(path.join(DATA_DIR, f)); } catch (_) {} });
      console.log("[seed] Da xoa " + tmpFiles.length + " file .tmp rac de giai phong Volume.");
    }

    // Xoa .bak files cu, giu 2 ban moi nhat
    const bakFiles = allFiles
      .filter((f) => /\.bak$/.test(f))
      .map((f) => ({ name: f, mtime: fs.statSync(path.join(DATA_DIR, f)).mtimeMs }))
      .sort((a, b) => b.mtime - a.mtime) // moi nhat truoc
      .slice(2) // giu 2 ban moi nhat, xoa phan con lai
      .map((f) => f.name);
    if (bakFiles.length > 0) {
      bakFiles.forEach((f) => { try { fs.unlinkSync(path.join(DATA_DIR, f)); } catch (_) {} });
      console.log("[seed] Da xoa " + bakFiles.length + " file .bak cu de giai phong Volume.");
    }
  } catch (_) {}

  if (changed) {
    // Bat loi save() (vd ENOSPC) de app khong crash -- du lieu van dung trong bo nho.
    try {
      save(store);
    } catch (e) {
      console.error("[seed] Khong the luu sau seed (co the Volume day?):", e.message,
        "-- App tiep tuc chay, du lieu DUNG trong bo nho nhung CHUA duoc ghi len dia.");
    }
  }
})();

function resetCache() {
  cachedStore = null;
  cachedTransactions = null;
  cachedVietQrRaw = null;
}

module.exports = { load, save, nextId, DATA_FILE, TRANSACTIONS_FILE, VIET_QR_RAW_FILE, resetCache };
