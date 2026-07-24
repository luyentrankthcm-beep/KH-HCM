const express = require("express");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { buildAllFlatLines, buildAgingRows } = require("../utils/overviewAggregate");
const { parseMisaCongNoXlsx } = require("../utils/misaCongNo");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

// Cong no "nhap tay" (Luyen yeu cau 2026-07-17): nhung dong cong no chi da tu
// doi soat xong TU TRUOC (co san so hoa don theo tung gian) -- khong tinh tu
// dong nhu bang aging o tren (bang do CHI danh cho dong CHUA xong), day chi
// la so ghi lai de tra cuu, khong cong don vao aging/cac tong so tu dong.
function ensureManualShape(store) {
  if (!store.congno_manual_entries) store.congno_manual_entries = [];
}

// Doi soat MISA cong no (TK 131) vs Ngan hang (Luyen yeu cau 2026-07-17): Luyen
// tai len file MISA "So chi tiet cong no phai thu theo cong trinh" (194 cong
// trinh, toan cong ty) -- nhung app nay chi theo doi ngan hang cho ~52
// gian/cong trinh (KH Cu + KH Moi), va ten goi trong MISA KHONG khop tu dong
// voi ten "gian" trong app (2 he thong dat ten khac nhau hoan toan). Vi vay
// can 1 buoc anh xa THU CONG 1 lan: Luyen chon "cong trinh MISA nao" tuong
// ung voi tung "gian" cua app -- luu vao congno_misa_mapping, dung lai moi
// lan tai file MISA moi. CHI so sanh 52 gian app dang co (theo dung y Luyen
// chon o cau hoi lam ro), khong hien 194 cong trinh MISA day du.
function ensureMisaShape(store) {
  if (!store.misa_congno) store.misa_congno = null;
  if (!store.congno_misa_mapping) store.congno_misa_mapping = {};
}

// "Cong no" o day nghia la: cac dong doi soat CHUA xong -- hoac chua co hoa
// don, hoac co hoa don nhung con lech -- tinh theo so ngay da troi qua ke tu
// ngay giao dich/settlement, giong cach nhin "cong no phai thu" thong thuong
// (tien da qua ngan hang nhung chua "khop so" voi hoa don tuong ung). Dung
// LAI chinh ket qua doi soat cua tung kenh (Momo/Zalo/VNPay/Payoo/3 kenh Viet
// QR) qua utils/overviewAggregate.js, khong tinh toan rieng.
// Luyen, 2026-07-19: "chỗ công nợ chia ra 2 trang, 1 trang là Công Nợ NCC 1
// trang là Công nợ khách hàng" -- doi duong dan trang nay (cong no PHAI THU
// tu khach hang/kenh doanh thu) thanh /cong-no/khach-hang, giu nguyen redirect
// cu /cong-no -> day de link/bookmark cu khong bi vo.
router.get("/cong-no", (req, res) => res.redirect("/cong-no/khach-hang"));

router.get("/cong-no/khach-hang", (req, res) => {
  const store = load();
  ensureManualShape(store);
  ensureMisaShape(store);
  let error = req.query.error || null;
  const success = req.query.success || null;
  let aging = { rows: [], buckets: [], grandTotal: 0 };
  let channelOptions = [];
  let misaComparison = [];
  let misaCongTrinhOptions = [];

  const manualEntries = [...store.congno_manual_entries].sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  const manualTotal = manualEntries.reduce((s, r) => s + r.amount, 0);

  try {
    const flat = buildAllFlatLines(store);
    channelOptions = Array.from(
      new Map(flat.map((l) => [l.channelKey, l.channelLabel])).entries()
    ).map(([key, label]) => ({ key, label }));

    const today = new Date().toISOString().slice(0, 10);
    const fullAging = buildAgingRows(flat, today);

    const selectedChannel = req.query.channel || "";
    const selectedBucket = req.query.bucket || "";
    let rows = fullAging.rows;
    if (selectedChannel) rows = rows.filter((r) => r.channelKey === selectedChannel);
    if (selectedBucket) rows = rows.filter((r) => r.bucket === selectedBucket);

    aging = {
      rows,
      buckets: fullAging.buckets,
      grandTotal: rows.reduce((s, r) => s + r.amount, 0),
      selectedChannel,
      selectedBucket,
    };

    // Tong "chua thu" theo tung gian (dung LAI dung so tu bang aging o tren,
    // khong tinh rieng, de khong bao gio lech nhau) -- dung lam ve trai cho
    // doi soat MISA.
    const gianBankTotals = {};
    fullAging.rows.forEach((r) => {
      if (!r.gian) return;
      gianBankTotals[r.gian] = (gianBankTotals[r.gian] || 0) + r.amount;
    });
    const gianList = Object.keys(gianBankTotals).sort((a, b) => a.localeCompare(b));

    if (store.misa_congno) {
      misaCongTrinhOptions = store.misa_congno.order.slice().sort((a, b) => a.localeCompare(b));
    }

    misaComparison = gianList.map((gian) => {
      const chuaThu = gianBankTotals[gian] || 0;
      const mappedCongTrinh = store.congno_misa_mapping[gian] || "";
      let soDuMisa = null;
      let chenhLech = null;
      let trangThai = "Chưa chọn công trình MISA";
      if (!store.misa_congno) {
        trangThai = "Chưa tải file MISA";
      } else if (mappedCongTrinh && store.misa_congno.byCongTrinh[mappedCongTrinh]) {
        const m = store.misa_congno.byCongTrinh[mappedCongTrinh];
        soDuMisa = (m.soDuNo || 0) - (m.soDuCo || 0);
        chenhLech = soDuMisa - chuaThu;
        trangThai = Math.abs(chenhLech) <= 1000 ? "Khớp" : "Lệch";
      }
      return { gian, chuaThu, mappedCongTrinh, soDuMisa, chenhLech, trangThai };
    });
  } catch (e) {
    error = e.message;
    console.error("Loi tinh cong no:", e);
  }

  res.render("congno", {
    userName: req.session.userName,
    aging,
    channelOptions,
    error,
    success,
    manualEntries,
    manualTotal,
    misaInfo: store.misa_congno,
    misaComparison,
    misaCongTrinhOptions,
  });
});

// ---------- Cong no nhap tay (da doi soat xong tu truoc, co so hoa don theo
// gian) -- 4 cot co ban theo yeu cau Luyen: Gian, Ngay, So tien, So HD. ----------
router.post("/cong-no/thu-cong", requireDataEntry, (req, res) => {
  const store = load();
  ensureManualShape(store);
  try {
    const { gian, date, amount, invoiceNumbers } = req.body;
    if (!gian || !gian.trim()) throw new Error("Thiếu Gian.");
    if (!date) throw new Error("Thiếu Ngày.");
    const amt = Number(String(amount || "").replace(/[^\d.-]/g, ""));
    if (!amt || amt <= 0) throw new Error("Số tiền không hợp lệ.");
    store.congno_manual_entries.push({
      id: nextId(store, "congno_manual_seq") || Date.now(),
      gian: gian.trim(),
      date,
      amount: amt,
      invoiceNumbers: (invoiceNumbers || "").trim(),
      createdAt: new Date().toISOString(),
    });
    save(store);
    res.redirect("/cong-no/khach-hang?success=" + encodeURIComponent("Đã lưu công nợ (nhập tay)."));
  } catch (e) {
    res.redirect("/cong-no/khach-hang?error=" + encodeURIComponent(e.message));
  }
});

router.post("/cong-no/thu-cong/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureManualShape(store);
  store.congno_manual_entries = store.congno_manual_entries.filter((r) => String(r.id) !== req.params.id);
  save(store);
  res.redirect("/cong-no/khach-hang?success=" + encodeURIComponent("Đã xóa dòng công nợ nhập tay."));
});

// ---------- Doi soat MISA cong no (131) vs Ngan hang ----------
router.post("/cong-no/upload-misa", requireDataEntry, upload.single("file"), (req, res) => {
  const store = load();
  ensureMisaShape(store);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file MISA để tải lên.");
    const parsed = parseMisaCongNoXlsx(req.file.buffer);
    if (!parsed.order.length) throw new Error("Không đọc được công trình nào trong file -- kiểm tra lại đúng file MISA xuất 'Sổ chi tiết công nợ phải thu theo công trình'.");
    store.misa_congno = {
      fileName: req.file.originalname,
      uploadedAt: new Date().toISOString(),
      byCongTrinh: parsed.byCongTrinh,
      order: parsed.order,
      tongCong: parsed.tongCong,
    };
    save(store);
    res.redirect("/cong-no/khach-hang?success=" + encodeURIComponent(`Đã tải file MISA (${parsed.order.length} công trình). Chọn công trình MISA tương ứng cho từng gian bên dưới.`));
  } catch (e) {
    res.redirect("/cong-no/khach-hang?error=" + encodeURIComponent(e.message));
  }
});

router.post("/cong-no/misa-mapping", requireDataEntry, (req, res) => {
  const store = load();
  ensureMisaShape(store);
  try {
    let gians = req.body.gian || [];
    let cts = req.body.congtrinh || [];
    if (!Array.isArray(gians)) gians = [gians];
    if (!Array.isArray(cts)) cts = [cts];
    const mapping = {};
    gians.forEach((g, i) => {
      const ct = (cts[i] || "").trim();
      if (ct) mapping[g] = ct;
    });
    store.congno_misa_mapping = mapping;
    save(store);
    res.redirect("/cong-no/khach-hang?success=" + encodeURIComponent("Đã lưu ánh xạ Gian ↔ Công trình MISA."));
  } catch (e) {
    res.redirect("/cong-no/khach-hang?error=" + encodeURIComponent(e.message));
  }
});

module.exports = router;
