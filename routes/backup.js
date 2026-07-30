const express = require("express");
const multer = require("multer");
const { load, save, nextId, DATA_FILE } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
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

router.post("/he-thong/sao-luu/phuc-hoi", requireAdmin, upload.single("file"), (req, res) => {
  try {
    if (!req.file) throw new Error("Vui long chon 1 file sao luu (.json) de phuc hoi.");
    let text = req.file.buffer.toString("utf8");
    req.file.buffer = null; // file da lon (~90MB+) -- giai phong buffer goc ngay khi da co chuoi text, tranh giu 2 ban sao cung luc gay OOM tren Railway
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error("File nay khong phai file JSON hop le -- kiem tra lai file da tai xuong o muc tren.");
    }
    text = null; // da parse xong, giai phong chuoi text (~90MB+) truoc khi goi save() (con phai stringify lai toan bo store)
    if (global.gc) global.gc(); // ep don rac ngay (server.js chay voi --expose-gc) truoc buoc save() ton bo nho nhat, giam dinh RAM tranh OOM/502 tren Railway
    // Kiem tra so bo day co dung la file sao luu cua he thong nay khong,
    // tranh phuc hoi nham 1 file JSON khac roi mat het du lieu that.
    const requiredKeys = ["users", "banks", "transactions"];
    const missing = requiredKeys.filter((k) => !Array.isArray(parsed[k]));
    if (missing.length > 0) {
      throw new Error(
        `File nay khong dung cau truc file sao luu K&H Bank Tracker (thieu: ${missing.join(", ")}).`
      );
    }
    save(parsed);
    res.render("backup", {
      userName: req.session.userName,
      error: null,
      success: `Da phuc hoi du lieu thanh cong: ${parsed.banks.length} ngan hang, ${parsed.transactions.length} giao dich, ${parsed.users.length} tai khoan dang nhap. Ban co the can dang nhap lai.`,
    });
  } catch (e) {
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
