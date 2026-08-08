const express = require("express");
const fs = require("fs");
const multer = require("multer");
const { load, save, nextId, DATA_FILE, TRANSACTIONS_FILE, VIET_QR_RAW_FILE, resetCache } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 200 * 1024 * 1024 },
});

// Chi Nhan, 2026-08-05: route phuc-hoi can dung disk storage de tranh OOM tren
// Railway Trial (1GB RAM). Voi memory storage: buffer(138MB) + string(138MB) +
// parsed(~400MB) + stringify(138MB) = ~876MB vuot 1GB. Voi disk storage: chi
// can string(138MB) + parsed(~400MB) + stringify(138MB) = ~676MB -- nhung van
// co the sat gioi han. Giai phap tot nhat: copy file truc tiep vao DATA_FILE
// (khong JSON.parse trong request handler), reset cache, load() lai tu dia --
// luc nay app da tra ve response va khong con xu ly gi khac, peak RAM chi la
// app baseline (~200MB) + load toan bo file khi server khoi dong tiep theo.
const uploadToDisk = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, require("os").tmpdir()),
    filename: (req, file, cb) => cb(null, `kh-restore-incoming-${Date.now()}.json`),
  }),
  limits: { fileSize: 200 * 1024 * 1024 },
});

// Trang dung de: (1) tai ban sao luu toan bo du lieu hien tai (file JSON), va
// (2) phuc hoi du lieu tu 1 file sao luu da tai truoc do -- dung khi chuyen
// du lieu tu ban chay o may (offline) sang ban chay online, hoac khi can
// khoi phuc sau su co.
router.get("/he-thong/sao-luu", (req, res) => {
  const store = load();
  res.render("backup", {
    userName: req.session.userName,
    banks: [...store.banks].sort((a, b) => a.name.localeCompare(b.name)),
    error: null,
    success: null,
  });
});

router.get("/he-thong/sao-luu/tai-xuong", (req, res) => {
  const store = load();
  // Chi Nhan, 2026-07-22: du lieu da lon dan (nhieu nam lich su tai len), stringify
  // co indent (null, 2) lam file to hon dang ke va cham hon -- bo indent de nhanh/nhe
  // hon (van la JSON hop le, phuc hoi lai binh thuong), tranh timeout/502 khi tai xuong.
  const data = JSON.stringify(store);
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Content-Disposition", `attachment; filename=kh-bank-tracker-sao-luu-${stamp}.json`);
  res.send(data);
});

router.post("/he-thong/sao-luu/phuc-hoi", requireAdmin, uploadToDisk.single("file"), (req, res) => {
  const tmpPath = req.file ? req.file.path : null;
  try {
    if (!req.file || !tmpPath) throw new Error("Vui long chon 1 file sao luu (.json) de phuc hoi.");

    // Kiem tra cau truc so bo bang cach chi doc 20KB dau file
    const headBuf = Buffer.allocUnsafe(20000);
    const fd = fs.openSync(tmpPath, "r");
    const bytesRead = fs.readSync(fd, headBuf, 0, 20000, 0);
    fs.closeSync(fd);
    const head = headBuf.slice(0, bytesRead).toString("utf8");

    if (!head.trim().startsWith("{")) {
      throw new Error("File nay khong phai file JSON hop le (khong bat dau bang '{').");
    }
    // Chi Nhan, 2026-08-08: "transactions" nam gan cuoi file (sau ~14MB config)
    // nen 20KB dau KHONG BAO GIO chua tu khoa nay -- chi kiem tra "users"/"banks"
    // trong head; "transactions" duoc kiem tra sau khi parse toan bo file (dong 104).
    const requiredKeys = ["users", "banks"];
    const missing = requiredKeys.filter((k) => !head.includes(`"${k}"`));
    if (missing.length > 0) {
      throw new Error(
        `File nay khong dung cau truc file sao luu K&H Bank Tracker (thieu: ${missing.join(", ")}).`
      );
    }

    // Backup du lieu hien tai truoc khi ghi de (phong phuc hoi nham)
    try {
      const preRestoreBackupPath = DATA_FILE + ".before-phuc-hoi-" + Date.now() + ".bak";
      if (fs.existsSync(DATA_FILE)) fs.copyFileSync(DATA_FILE, preRestoreBackupPath);
    } catch (backupErr) {
      console.error("[backup] Khong the tao ban sao truoc khi phuc hoi:", backupErr.message);
    }

    // QUAN TRONG: Giai phong cache cu TRUOC khi parse file backup.
    // Neu khong: cachedTransactions cu (~300MB) + parse moi (~400MB) = ~700MB+
    // vuot gioi han 1GB Railway Trial -> OOM crash.
    resetCache();
    // Ep GC chay ngay (server dung --expose-gc) de giai phong bo nho cu
    // truoc khi parse file 138MB -- khong co buoc nay GC co the chua chay
    // va RAM van ~560MB truoc khi parse (+400MB = ~960MB -> OOM).
    if (typeof global.gc === "function") {
      global.gc();
      console.log("[backup] Da goi global.gc() truoc khi parse backup.");
    }

    const fileSize = fs.statSync(tmpPath).size;
    const parsed = JSON.parse(fs.readFileSync(tmpPath, "utf8"));
    try { fs.unlinkSync(tmpPath); } catch (_) {} // xoa file tam ngay sau khi parse xong

    if (!Array.isArray(parsed.users) || !Array.isArray(parsed.banks) || !Array.isArray(parsed.transactions)) {
      throw new Error("File sao luu khong hop le (thieu users/banks/transactions array).");
    }

    // Tach config (khong co 2 mang lon) va cac mang lon ra rieng
    const configObj = {};
    for (const key of Object.keys(parsed)) {
      if (key === "transactions" || key === "viet_qr_raw_uploads") continue;
      configObj[key] = parsed[key];
    }
    const txArr = parsed.transactions; // khong copy -- chi tham chieu
    const vqrObj =
      parsed.viet_qr_raw_uploads && typeof parsed.viet_qr_raw_uploads === "object" && !Array.isArray(parsed.viet_qr_raw_uploads)
        ? parsed.viet_qr_raw_uploads
        : { bidv7704: [], bidv77020: [], mb11521268: [] };

    // Ghi 3 file rieng (moi file de dat hon 95MB, tranh ENOSPC)
    function writeDurableSimple(filePath, data) {
      const tmpF = filePath + ".tmp-restore-" + Date.now();
      fs.writeFileSync(tmpF, data);
      fs.renameSync(tmpF, filePath);
    }
    writeDurableSimple(DATA_FILE, JSON.stringify(configObj));
    writeDurableSimple(TRANSACTIONS_FILE, JSON.stringify(txArr));
    writeDurableSimple(VIET_QR_RAW_FILE, JSON.stringify(vqrObj));

    // Reset cache va load lai tu file config nho -- baseline RAM sau restore: ~220MB
    resetCache();
    const freshStore = load();

    res.render("backup", {
      userName: req.session.userName,
      error: null,
      success:
        `Da phuc hoi du lieu thanh cong: ${freshStore.banks.length} ngan hang, ` +
        `${txArr.length} giao dich, ${freshStore.users.length} tai khoan dang nhap. ` +
        `File ${Math.round(fileSize / 1024 / 1024)}MB da duoc tach thanh 3 file nho ` +
        `(config/transactions/viet_qr_raw) de giam RAM khi khoi dong. Ban co the can dang nhap lai.`,
    });
  } catch (e) {
    if (tmpPath) { try { fs.unlinkSync(tmpPath); } catch (_) {} }
    res.render("backup", { userName: req.session.userName, error: e.message, success: null });
  }
});

// Chi Nhan, 2026-07-22: "cái onl á bạn giữa các trang khác còn phần ngân
// hàng 7702, momo KH cũ momo kh mới bạn cập nhật lại cái code off qua cho
// tôi" -- chi lam RO: day la DONG BO DU LIEU (khong phai code), tu ban chay
// o may cua Nhan (offline) len ban Railway (online), CHI cho phan Momo (ca
// KH Cu va KH Moi) + 1 kenh Viet QR cu the (vd bidv7702) -- GIU NGUYEN het
// cac trang/du lieu khac cua ban online (khong ghi de toan bo nhu phuc-hoi
// thuong). Nhan file sao luu toan bo tai xuong tu ban offline, nhung chi lay
// dung cac truong lien quan de ghi de len ban online dang chay.
router.post("/he-thong/sao-luu/dong-bo-mot-phan", requireAdmin, upload.single("file"), (req, res) => {
  try {
    if (!req.file) throw new Error("Vui long chon 1 file sao luu (.json) tai tu ban offline.");
    const text = req.file.buffer.toString("utf8");
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error("File nay khong phai file JSON hop le.");
    }
    const channels = (req.body.channels || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
    if (channels.length === 0) {
      throw new Error('Thieu kenh Viet QR can dong bo (vd "bidv7702").');
    }

    const store = load();
    const summary = [];

    ["momo_gross_uploads", "momo_invoices", "momo_moi_gross_uploads", "momo_moi_invoices"].forEach((key) => {
      if (Array.isArray(parsed[key])) {
        store[key] = parsed[key];
        summary.push(`${key}: ${parsed[key].length}`);
      }
    });

    const viChannelKeys = [
      "viet_qr_raw_uploads",
      "viet_qr_store_names",
      "viet_qr_store_uploads",
      "viet_qr_store_names_baseline",
      "viet_qr_invoices",
    ];
    channels.forEach((ch) => {
      viChannelKeys.forEach((key) => {
        if (parsed[key] && parsed[key][ch] !== undefined) {
          if (!store[key]) store[key] = {};
          store[key][ch] = parsed[key][ch];
        }
      });
      const cnt = (parsed.viet_qr_raw_uploads && parsed.viet_qr_raw_uploads[ch] || []).length;
      summary.push(`viet_qr[${ch}]: ${cnt} lan tai giao dich`);
    });

    save(store);
    res.render("backup", {
      userName: req.session.userName,
      error: null,
      success: `Da dong bo tu file offline (chi Momo KH Cu/Moi + kenh Viet QR: ${channels.join(", ")}): ${summary.join(
        "; "
      )}. Cac trang/du lieu khac cua ban online giu nguyen.`,
    });
  } catch (e) {
    res.render("backup", { userName: req.session.userName, error: e.message, success: null });
  }
});

// Chi Nhan, 2026-07-24: Luyen phat hien du lieu giao dich (sao ke ngan hang)
// cua BIDV7702 bi lech/trung lap trong luc debug lien tuc tai lai sao ke
// (bug "So chung tu bi nham So tham chieu" -- xem utils/bankStatementParser.js).
// Luyen co 1 file sao luu CU HON, sach hon (truoc khi cac lan tai lap gay
// loi), muon dung DUNG PHAN GIAO DICH cua 1 ngan hang trong 1 khoang ngay tu
// file do de thay the lai cho dung, KHONG dong den giao dich cua ngan hang
// khac hay ngay khac ngoai khoang da chon -- khac voi "dong-bo-mot-phan" o
// tren (chi xu ly Momo + kenh Viet QR, khong dong den transactions).
//
// Khop ngan hang qua TEN (khong phai id) vi id co the KHAC nhau giua file
// sao luu va du lieu dang chay (thu tu tao ngan hang co the khac giua 2 lan).
router.post("/he-thong/sao-luu/khoi-phuc-giao-dich", requireAdmin, upload.single("file"), (req, res) => {
  try {
    if (!req.file) throw new Error("Vui long chon 1 file sao luu (.json) de lay du lieu.");
    const { bank_id, from, to } = req.body;
    if (!bank_id) throw new Error("Vui long chon ngan hang can khoi phuc.");
    if (!from || !to) throw new Error("Vui long chon du ca ngay bat dau va ngay ket thuc.");
    if (from > to) throw new Error("Ngay bat dau phai truoc ngay ket thuc.");

    const text = req.file.buffer.toString("utf8");
    let backup;
    try {
      backup = JSON.parse(text);
    } catch (e) {
      throw new Error("File nay khong phai file JSON hop le.");
    }
    if (!Array.isArray(backup.banks) || !Array.isArray(backup.transactions)) {
      throw new Error("File nay khong dung cau truc file sao luu K&H Bank Tracker (thieu banks/transactions).");
    }

    const store = load();
    const liveBank = store.banks.find((b) => b.id === Number(bank_id));
    if (!liveBank) throw new Error("Khong tim thay ngan hang nay tren ban dang chay.");
    const backupBank = backup.banks.find((b) => b.name === liveBank.name);
    if (!backupBank) {
      throw new Error(`File sao luu nay khong co ngan hang "${liveBank.name}".`);
    }

    const backupRowsInRange = backup.transactions.filter(
      (t) => t.bank_id === backupBank.id && t.date >= from && t.date <= to
    );
    if (backupRowsInRange.length === 0) {
      throw new Error(`File sao luu khong co giao dich nao cua "${liveBank.name}" trong khoang ${from} den ${to}.`);
    }

    const before = store.transactions.length;
    store.transactions = store.transactions.filter(
      (t) => !(t.bank_id === liveBank.id && t.date >= from && t.date <= to)
    );
    const removedCount = before - store.transactions.length;

    backupRowsInRange.forEach((t) => {
      store.transactions.push({
        id: nextId(store, "transactions"),
        bank_id: liveBank.id,
        date: t.date,
        description: t.description,
        amount: t.amount,
        type: t.type,
        reference: t.reference || "",
        created_at: t.created_at || new Date().toISOString(),
        created_by: req.session.userName || "",
        restored_from_backup_at: new Date().toISOString(),
      });
    });
    save(store);

    res.render("backup", {
      userName: req.session.userName,
      error: null,
      success: `Đã khôi phục giao dịch "${liveBank.name}" từ ${from} đến ${to}: xoá ${removedCount} dòng cũ, nạp lại ${backupRowsInRange.length} dòng đúng từ file sao lưu.`,
    });
  } catch (e) {
    res.render("backup", { userName: req.session.userName, error: e.message, success: null });
  }
});

// Chi Nhan, 2026-07-24: Luyen chay 1 ban OFFLINE (may rieng, du lieu Viet QR
// day du hon -- vd raw_uploads/hoa don di ve xa hon) song song voi ban ONLINE
// (Railway, du lieu Viet QR co the thieu 1 doan ngay nao do chua kip tai
// len). "dong-bo-mot-phan" o tren THAY THE TOAN BO mang cua kenh (nguy hiem
// neu ban online da co du lieu MOI HON ban offline cho nhung ngay SAU thoi
// diem chup file offline -- se bi xoa mat). Muc nay MERGE (chi THEM phan con
// thieu, KHONG xoa/ghi de bat ky gi ban online da co san):
//  - viet_qr_raw_uploads[ch]: cong THEM cac "lan tai" (batch) cua file
//    offline vao mang hien co (gan id MOI de tranh trung id), giu nguyen cac
//    lan tai da co cua ban online -- viec khu trung LAP O TUNG DONG giao dich
//    (theo vqrCode/ngay) da co san trong mergeRawRows() luc doc, nen gop
//    them 1 batch cu KHONG lam nhan doi doanh thu.
//  - viet_qr_store_names[ch]: CHI dien vao nhung ma cua hang ban online CHUA
//    CO (khong ghi de len ma da co, tranh mat cap nhat gan day tren online).
//  - viet_qr_invoices[ch]: cong them hoa don chua co (khop trung theo
//    soHd|ngayHd|maDiem, giong het logic /upload-hoadon dang dung).
router.post("/he-thong/sao-luu/dong-bo-mot-phan-gop", requireAdmin, upload.single("file"), (req, res) => {
  try {
    if (!req.file) throw new Error("Vui long chon 1 file sao luu (.json) tai tu ban offline.");
    const text = req.file.buffer.toString("utf8");
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error("File nay khong phai file JSON hop le.");
    }
    const channels = (req.body.channels || "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
    if (channels.length === 0) {
      throw new Error('Thieu kenh Viet QR can gop (vd "bidv7702").');
    }

    const store = load();
    const summary = [];

    channels.forEach((ch) => {
      // 1) raw_uploads: cong them batch, gan id moi.
      if (!store.viet_qr_raw_uploads) store.viet_qr_raw_uploads = {};
      if (!store.viet_qr_raw_uploads[ch]) store.viet_qr_raw_uploads[ch] = [];
      const offlineBatches = (parsed.viet_qr_raw_uploads && parsed.viet_qr_raw_uploads[ch]) || [];
      let addedBatches = 0;
      let addedRows = 0;
      offlineBatches.forEach((batch) => {
        store.viet_qr_raw_uploads[ch].push({
          ...batch,
          id: nextId(store, "viet_qr_raw_uploads_seq"),
          file_name: `[gop tu offline] ${batch.file_name || ""}`,
        });
        addedBatches += 1;
        addedRows += (batch.rows || []).length;
      });

      // 2) store_names: chi dien vao ma chua co, khong ghi de.
      if (!store.viet_qr_store_names) store.viet_qr_store_names = {};
      if (!store.viet_qr_store_names[ch]) store.viet_qr_store_names[ch] = {};
      const offlineStoreNames = (parsed.viet_qr_store_names && parsed.viet_qr_store_names[ch]) || {};
      let addedStoreNames = 0;
      Object.keys(offlineStoreNames).forEach((maCuaHang) => {
        if (store.viet_qr_store_names[ch][maCuaHang] !== undefined) return;
        store.viet_qr_store_names[ch][maCuaHang] = offlineStoreNames[maCuaHang];
        addedStoreNames += 1;
      });

      // 3) invoices: cong them hoa don chua co, khop trung theo soHd|ngayHd|maDiem.
      if (!store.viet_qr_invoices) store.viet_qr_invoices = {};
      if (!store.viet_qr_invoices[ch]) store.viet_qr_invoices[ch] = [];
      const offlineInvoices = (parsed.viet_qr_invoices && parsed.viet_qr_invoices[ch]) || [];
      const existingInvKeys = new Set(store.viet_qr_invoices[ch].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
      let addedInvoices = 0;
      offlineInvoices.forEach((inv) => {
        const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
        if (existingInvKeys.has(k)) return;
        existingInvKeys.add(k);
        store.viet_qr_invoices[ch].push(inv);
        addedInvoices += 1;
      });

      summary.push(
        `${ch}: +${addedBatches} lần tải (${addedRows} dòng QR), +${addedStoreNames} mã cửa hàng mới, +${addedInvoices} hóa đơn mới`
      );
    });

    save(store);
    res.render("backup", {
      userName: req.session.userName,
      error: null,
      success: `Đã gộp thêm dữ liệu từ file offline cho kênh: ${summary.join("; ")}. Dữ liệu online đã có KHÔNG bị xoá/ghi đè.`,
    });
  } catch (e) {
    res.render("backup", { userName: req.session.userName, error: e.message, success: null });
  }
});

module.exports = router;
