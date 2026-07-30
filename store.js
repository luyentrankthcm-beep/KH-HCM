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
    gian_mapping: {}, // { "MA CONG TRINH": "131" | "1388" | "SKIP" } -- shared across Momo AND Zalo/VNPay/Payoo
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
    const fresh = emptyStore();
    fs.writeFileSync(DATA_FILE, JSON.stringify(fresh));
    cachedStore = fresh;
    return cachedStore;
  }
  const raw = fs.readFileSync(DATA_FILE, "utf8");
  try {
    const parsed = JSON.parse(raw);
    const base = emptyStore();
    cachedStore = Object.assign({}, base, parsed, {
      seq: Object.assign({}, base.seq, parsed.seq || {}),
    });
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
      save(healed); // persist the healed version immediately so it stays fixed
      cachedStore = healed;
      return cachedStore;
    } catch (e2) {
      // CRITICAL: the file exists but we could not recover it. Do NOT return
      // an empty store silently and let the caller (seed()) persist that
      // empty store over the real (if corrupted) data on disk. Back up the
      // unreadable file first, then flag this so seed() below refuses to
      // auto-save a fresh/empty store on top of it.
      console.error("[store] KHONG THE tu phuc hoi du lieu:", e2.message);
      backupCorruptedFile();
      console.error(
        "[store] CANH BAO NGHIEM TRONG: file data/store.json bi loi va KHONG the tu phuc hoi. " +
          "App se chay tam voi du lieu RONG trong bo nho (KHONG ghi de len file that) de ban con co the yeu cau khoi phuc thu cong. " +
          "Kiem tra file .corrupted-*.bak trong thu muc data/."
      );
      loadedFromUnrecoverableCorruption = true;
      return emptyStore();
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
  // IMPORTANT: on some mounted filesystems, overwriting store.json in place
  // can leave stale trailing bytes from the previous (longer) version of
  // the file if the new content is shorter, or even silently truncate a
  // large write -- corrupting the JSON. To guard against this we (1) write
  // to a temp file (via writeFileDurable, which loops until fully written
  // and fsyncs), (2) rename it into place (atomic at the filesystem level --
  // the old file is fully replaced, not overwritten byte-by-byte), then
  // (3) verify the write landed fully before declaring success, retrying a
  // few times if not.
  // Chi Nhan, 2026-07-22: bo indent (null, 2) -- file da lon (30-40MB+), indent
  // lam file to hon dang ke va JSON.stringify cham hon (van la JSON hop le,
  // doc/phuc hoi lai binh thuong), gop voi cache o load() de giam toi da thoi
  // gian chan luong Node.js gay 502.
  // Chi Nhan, 2026-07-30: file da qua lon (~90MB+, se con lon them) -- buoc (3)
  // truoc day doc lai TOAN BO file roi JSON.parse lai de "kiem tra", tao ra 1
  // BAN SAO du lieu THU HAI trong bo nho cung luc voi `store` dang giu +
  // chuoi `data` vua stringify -- do la nguyen nhan chinh gay HET BO NHO khi
  // phuc hoi file sao luu tren Railway (do luong: ~930MB dinh RAM cho 1 file
  // 92MB, trong khi goi Railway chi co vai tram MB) -- day chinh la nguyen
  // nhan loi "502 Application failed to respond" khi bam Phuc hoi sao luu.
  // writeFileDurable() da fsync + dung writeFileSync (tu lap toi khi ghi HET
  // buffer, khong con rui ro "ghi thieu byte") nen chi can kiem tra NHE: so
  // sanh dung so byte da ghi thuc te tren dia voi so byte du kien, khong can
  // doc lai + parse lai toan bo noi dung.
  const data = JSON.stringify(store);
  const expectedBytes = Buffer.byteLength(data, "utf8");
  let lastErr = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    const tmpFile = DATA_FILE + ".tmp-" + process.pid + "-" + Date.now() + "-" + Math.random().toString(36).slice(2);
    try {
      writeFileDurable(tmpFile, data);
      fs.renameSync(tmpFile, DATA_FILE);
      const actualBytes = fs.statSync(DATA_FILE).size;
      if (actualBytes !== expectedBytes) {
        throw new Error(`Ghi file khong du so byte (mong doi ${expectedBytes}, thuc te ${actualBytes})`);
      }
      cachedStore = store; // cache updated only after a verified-successful write
      return; // success
    } catch (e) {
      lastErr = e;
      console.error("[store] save() attempt " + attempt + " failed/khong hop le:", e.message);
      try {
        fs.unlinkSync(tmpFile);
      } catch (_) {
        /* ignore cleanup error */
      }
    }
  }
  throw new Error(
    "Khong the luu du lieu an toan sau nhieu lan thu (store.save that bai): " +
      (lastErr && lastErr.message)
  );
}

function nextId(store, collection) {
  store.seq[collection] = (store.seq[collection] || 0) + 1;
  return store.seq[collection];
}

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

  if (changed) save(store);
})();

module.exports = { load, save, nextId, DATA_FILE };
