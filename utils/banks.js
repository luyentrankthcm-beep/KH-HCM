const express = require("express");
const { load, save, nextId } = require("../store");
const { requireLogin } = require("../middleware/auth");
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

router.post("/banks", (req, res) => {
  const {
    name,
    account_number,
    bank_name,
    opening_balance,
    opening_date,
    khu_vuc,
    chi_nhanh,
    trung_gian_thu_ho,
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
    created_at: new Date().toISOString(),
  });
  save(store);
  res.redirect("/banks");
});

router.post("/banks/:id/delete", (req, res) => {
  const store = load();
  const id = Number(req.params.id);
  store.banks = store.banks.filter((b) => b.id !== id);
  store.transactions = store.transactions.filter((t) => t.bank_id !== id);
  save(store);
  res.redirect("/banks");
});

module.exports = { router, computeBalance };
