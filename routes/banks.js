const express = require("express");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

function computeBalance(store, bankId) {
  const bank = store.banks.find((b) => b.id === Number(bankId));
  if (!bank) return null;
  let thu = 0;
  let chi = 0;
  for (const t of store.transactions) {
    if (t.bank_id === bank.id) {
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

router.post("/banks/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  store.banks = store.banks.filter((b) => b.id !== id);
  store.transactions = store.transactions.filter((t) => t.bank_id !== id);
  save(store);
  res.redirect("/banks");
});

module.exports = { router, computeBalance };
