const express = require("express");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan, 2026-09-11: "sao số dư cuối kì không khớp trên sao kê vậy check
// lại cho tôi nhá" -- BIDV123456 hien 521.954.686d nhung sao ke that chi
// 24.308.501d (lech dung 497.646.185d). Goc re: route /banks/:id/so-du-dau-ky
// (them 2026-07-31) cho sua opening_balance/opening_date SAU KHI ngan hang da
// co san giao dich tu truoc (de "doi chieu lai voi sao ke that"), nhung ham
// nay van cong TAT CA store.transactions cua ngan hang do KHONG LOC theo ngay
// -- nen cac giao dich co ngay TRUOC opening_date bi CONG HAI LAN: 1 lan an
// trong opening_balance (vi opening_balance da la so du TINH DEN thoi diem
// opening_date), 1 lan nua khi vong for nay cong lai chinh cac giao dich do.
// Da xac nhan tren du lieu that: BIDV123456 co 11 giao dich Thang 1-3/2026
// truoc opening_date 02/04/2026 van con trong store.transactions, va BIDV8651
// (Chi Phi) con nghiem trong hon voi toi 664 giao dich truoc opening_date
// 01/07/2026. Fix: chi cong giao dich co t.date >= bank.opening_date (giao
// dich truoc do coi nhu DA duoc gop san vao opening_balance rong, khong cong
// lai nua). Ngan hang nao opening_date <= ngay giao dich dau tien thi khong
// bi anh huong gi (dieu kien luon dung).
function computeBalance(store, bankId) {
  const bank = store.banks.find((b) => b.id === Number(bankId));
  if (!bank) return null;
  const openingDate = bank.opening_date || "";
  let thu = 0;
  let chi = 0;
  for (const t of store.transactions) {
    if (t.bank_id === bank.id && (!openingDate || t.date >= openingDate)) {
      if (t.type === "thu") thu += t.amount;
      else chi += t.amount;
    }
  }
  return bank.opening_balance + thu - chi;
}

// Loc theo cong ty dang chon (nut chuyen KH Cu/KH Moi tren topbar) -- ngan
// hang khong co field "company" (du lieu cu, tao truoc khi tach cong ty) mac
// dinh coi la "kh_cu", giong dung 1 quy uoc voi routes/transactions.js's
// companyBanks() de 2 cho khong lech nhau. Truoc khi co fix nay, BIDV7701/
// BIDV7702 (rieng cua KH Moi) van hien ca khi dang xem "Cu" vi trang nay
// chua loc gi ca -- Luyen bao 2026-07-16.
function companyBanks(store, company) {
  return store.banks.filter((b) => (b.company || "kh_cu") === company);
}

router.get("/banks", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  const banks = companyBanks(store, activeCompany)
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((b) => ({ ...b, balance: computeBalance(store, b.id) }));
  res.render("banks", { banks, userName: req.session.userName, error: null });
});

router.post("/banks", requireAdmin, (req, res) => {
  const {
    name,
    account_number,
    bank_name,
    opening_balance,
    opening_date,
    khu_vuc,
    chi_nhanh,
    trung_gian_thu_ho,
    mien,
  } = req.body;
  const store = load();
  const activeCompany = getCompany(req);

  if (!name || !name.trim()) {
    const banks = companyBanks(store, activeCompany)
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((b) => ({ ...b, balance: computeBalance(store, b.id) }));
    return res.render("banks", {
      banks,
      userName: req.session.userName,
      error: "Vui long nhap ten tai khoan / ngan hang.",
    });
  }

  // Ngan hang moi them luon gan vao cong ty DANG XEM luc bam "Them ngan
  // hang" (giong cach cac trang doi soat khac gan company) -- nen doi dung
  // cong ty truoc khi them neu la tai khoan cua KH Moi.
  store.banks.push({
    id: nextId(store, "banks"),
    name: name.trim(),
    account_number: (account_number || "").trim(),
    bank_name: (bank_name || "").trim(),
    opening_balance: parseFloat(opening_balance || 0) || 0,
    opening_date: opening_date || new Date().toISOString().slice(0, 10),
    khu_vuc: (khu_vuc || "").trim(),
    chi_nhanh: (chi_nhanh || "").trim(),
    trung_gian_thu_ho: (trung_gian_thu_ho || "").trim(),
    company: activeCompany,
    // Luyen, 2026-07-23: "chia cho tôi thành 2 trang 1 trang Miền Nam và 1
    // trang miền Bắc ... liên kết link với ngân hàng tôi sẽ liên kết sau" --
    // trang Chi Phí Miền Nam/Miền Bắc dung truong nay (xem routes/chi-phi.js
    // quet-ngan-hang) de biet 1 tai khoan ngan hang thuoc mien nao. Mac dinh
    // "nam" (tat ca tai khoan hien co deu dang phuc vu Mien Nam) -- doi thanh
    // "bac" ngay tai day khi Luyen them/lien ket tai khoan ngan hang Mien Bac.
    mien: mien === "bac" ? "bac" : "nam",
    created_at: new Date().toISOString(),
  });
  save(store);
  res.redirect("/banks");
});

// Luyen, 2026-07-23: cho sua rieng truong "mien" cua 1 tai khoan da co san
// (khong can xoa/them lai) -- dung khi Luyen lien ket 1 tai khoan hien co cho
// hoat dong Mien Bac, hoac lo chon nham luc them moi.
router.post("/banks/:id/mien", requireAdmin, (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  const bank = store.banks.find((b) => b.id === id);
  if (bank) {
    bank.mien = req.body.mien === "bac" ? "bac" : "nam";
    save(store);
  }
  res.redirect("/banks");
});

// Luyen, 2026-07-31: "số dư làm gì có âm tài khoản á lấy số dư giống sao kê
// nha thay đổi cái đầu kì đi" -- truoc gio "So du dau ky" chi dat duoc LUC
// TAO MOI 1 ngan hang (POST /banks o tren), KHONG co cach nao sua lai sau do
// neu ghi sai/can doi chieu lai voi sao ke that (vd VPBANK9997 dang de 0 hoac
// sai, khien so du chay am khong thuc te) -- phai xoa han ngan hang do (mat
// het lich su giao dich) roi tao lai moi sua duoc, qua nguy hiem. Them route
// rieng sua THANG 2 truong nay (khong dung den xoa/tao lai), giong cach
// /banks/:id/mien da lam.
router.post("/banks/:id/so-du-dau-ky", requireAdmin, (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  const bank = store.banks.find((b) => b.id === id);
  if (bank) {
    const ob = parseFloat(req.body.opening_balance);
    bank.opening_balance = isNaN(ob) ? bank.opening_balance : ob;
    if (req.body.opening_date) bank.opening_date = req.body.opening_date;
    save(store);
  }
  res.redirect("/banks");
});

router.post("/banks/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  store.banks = store.banks.filter((b) => b.id !== id);
  store.transactions = store.transactions.filter((t) => t.bank_id !== id);
  save(store);
  res.redirect("/banks");
});

module.exports = { router, computeBalance };
