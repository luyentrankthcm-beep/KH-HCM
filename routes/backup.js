const express = require("express");
const multer = require("multer");
const { load, save, DATA_FILE } = require("../store");
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
  res.render("backup", { userName: req.session.userName, error: null, success: null });
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
    const text = req.file.buffer.toString("utf8");
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error("File nay khong phai file JSON hop le -- kiem tra lai file da tai xuong o muc tren.");
    }
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

module.exports = router;
