const express = require("express");
const router = express.Router();
const { requireLogin } = require("../middleware/auth");
const { load, save } = require("../store");
const { getCompany } = require("../utils/companies");

router.use(requireLogin);

router.get("/dau-ra", (req, res) => {
  res.render("dau-ra", { userName: req.session.userName || req.session.user || "" });
});

// ── KV store cho trang "Đầu Ra" ──────────────────────────────────────────
// Chi Nhan, 2026-09-10: trang nay truoc gio luu Momo/MTT/Zalo rows... bang
// localStorage CUA TRINH DUYET -- Luyen bao mo web o 2 may/tai khoan khac
// nhau thay 1 cai co du lieu 1 cai khong, vi localStorage khong dong bo qua
// server. 3 route duoi day luu/doc cung 1 du lieu do TREN SERVER (theo tung
// cong ty kh_cu/kh_moi, xem store.dau_ra_kv trong store.js) de mo may nao
// / trinh duyet nao cung thay GIONG NHAU. Client (dau-ra.ejs) van giu
// localStorage nhu cu de doc/hien nhanh (khong doi logic parse/render), chi
// them buoc: (1) luc tai trang, keo du lieu server ve ghi de vao localStorage
// TRUOC khi cac ham cu doc localStorage nhu binh thuong; (2) moi lan cac ham
// cu ghi vao localStorage thi ALSO gui 1 ban len server qua route POST /kv.
router.get("/api/dau-ra/kv-all", (req, res) => {
  try {
    const store = load();
    const company = getCompany(req);
    if (!store.dau_ra_kv) store.dau_ra_kv = { kh_cu: {}, kh_moi: {} };
    if (!store.dau_ra_kv[company]) store.dau_ra_kv[company] = {};
    res.json({ ok: true, data: store.dau_ra_kv[company] });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

router.post("/api/dau-ra/kv", express.json({ limit: "30mb" }), (req, res) => {
  try {
    const { key, value } = req.body || {};
    if (!key) return res.json({ ok: false, error: "Thiếu key" });
    const store = load();
    const company = getCompany(req);
    if (!store.dau_ra_kv) store.dau_ra_kv = { kh_cu: {}, kh_moi: {} };
    if (!store.dau_ra_kv[company]) store.dau_ra_kv[company] = {};
    store.dau_ra_kv[company][key] = value;
    save(store);
    res.json({ ok: true });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

router.post("/api/dau-ra/kv-delete", express.json({ limit: "1mb" }), (req, res) => {
  try {
    const { key } = req.body || {};
    if (!key) return res.json({ ok: false, error: "Thiếu key" });
    const store = load();
    const company = getCompany(req);
    if (store.dau_ra_kv && store.dau_ra_kv[company]) delete store.dau_ra_kv[company][key];
    save(store);
    res.json({ ok: true });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

// API: tổng thu ngân hàng theo ngày cho 1 tài khoản
// GET /api/dau-ra/bank-daily?account=7701&from=2026-09-01&to=2026-09-30
// Trả về { "02-09-2026": 47586000, ... } — date theo DD-MM-YYYY để match client
router.get("/api/dau-ra/bank-daily", (req, res) => {
  try {
    const { account, from, to } = req.query;
    if (!account) return res.json({ ok: false, error: "Thiếu account" });

    const store   = load();
    const company = getCompany(req);

    // Tìm bank theo phần cuối số tài khoản (vd "7701")
    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith(account)
    );
    if (!bank) return res.json({ ok: false, error: `Không tìm thấy tài khoản *${account} trong ${company}` });

    // Lọc giao dịch "thu" của bank đó trong khoảng ngày
    const txns = store.transactions.filter(t => {
      if (t.bank_id !== bank.id) return false;
      if (t.type !== "thu") return false;
      if (from && t.date < from) return false;
      if (to   && t.date > to)   return false;
      return true;
    });

    // Gom theo ngày: store dùng YYYY-MM-DD → convert sang DD-MM-YYYY cho client
    const daily = {};
    for (const t of txns) {
      // t.date có thể là "2026-09-02" hoặc "02-09-2026" tuỳ cách nhập
      let ddmmyyyy = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const [y, m, d] = t.date.split("-");
        ddmmyyyy = `${d}-${m}-${y}`;
      }
      daily[ddmmyyyy] = (daily[ddmmyyyy] || 0) + Number(t.amount || 0);
    }

    res.json({ ok: true, bankName: bank.name, bankId: bank.id, daily });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

// API: lấy các đợt Momo trả về ngân hàng (parse diễn giải "tu DD/MM/YYYY den DD/MM/YYYY")
// GET /api/dau-ra/bank-momo?account=7701&from=2026-08-01&to=2026-09-30
// Trả về { ok, payments: [{ payDate:"DD-MM-YYYY", amount, fromDate:"DD-MM-YYYY", toDate:"DD-MM-YYYY" }] }
router.get("/api/dau-ra/bank-momo", (req, res) => {
  try {
    const { account, from, to } = req.query;
    if (!account) return res.json({ ok: false, error: "Thiếu account" });

    const store   = load();
    const company = getCompany(req);

    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith(account)
    );
    if (!bank) return res.json({ ok: false, error: `Không tìm thấy TK *${account} trong ${company}` });

    const pattern = /tu (\d{2})\/(\d{2})\/(\d{4}) den (\d{2})\/(\d{2})\/(\d{4})/i;
    const payments = [];

    for (const t of store.transactions) {
      if (t.bank_id !== bank.id) continue;
      if (t.type !== "thu") continue;
      // Momo settlement descriptions: "DI DONG TRUC TUYEN" hoặc "MoMo" trong diễn giải
      if (!/DI DONG TRUC TUYEN|MOMO|MoMo/i.test(t.description || "")) continue;
      const m = pattern.exec(t.description || "");
      if (!m) continue;

      const [, d1, mo1, y1, d2, mo2, y2] = m;

      // Lọc theo khoảng ngày thanh toán (payDate = t.date, format YYYY-MM-DD)
      if (from && t.date < from) continue;
      if (to   && t.date > to)   continue;

      // Convert t.date (YYYY-MM-DD) sang DD-MM-YYYY cho client
      let payDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const [y, m2, d] = t.date.split("-");
        payDate = `${d}-${m2}-${y}`;
      }

      payments.push({
        payDate,
        amount:   Number(t.amount || 0),
        fromDate: `${d1}-${mo1}-${y1}`,
        toDate:   `${d2}-${mo2}-${y2}`,
        desc:     (t.description || "").substring(0, 80),
      });
    }

    // Sắp xếp theo ngày thanh toán tăng dần
    payments.sort((a, b) => {
      const toISO = d => { const [dd, mm, yy] = d.split("-"); return `${yy}-${mm}-${dd}`; };
      return toISO(a.payDate).localeCompare(toISO(b.payDate));
    });

    res.json({ ok: true, payments });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});


// GET /api/dau-ra/bank-zalo?from=2026-08-15&to=2026-09-08
// Trả về các đợt ZaloPay trả về ACB1268 (diễn giải VNPAY TT 829168, không OFFLINE)
router.get("/api/dau-ra/bank-zalo", (req, res) => {
  try {
    const { from, to } = req.query;
    const store   = load();
    const company = getCompany(req);
    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith("1268")
    );
    if (!bank) return res.json({ ok: false, error: "Không tìm thấy TK *1268" });

    const payments = [];
    for (const t of store.transactions) {
      if (t.bank_id !== bank.id) continue;
      if (t.type !== "thu") continue;
      // Chỉ lấy đúng mẫu ZaloPay mini app: "VNPAY TT 829168 GIAIT...989 DV CTT NGAY"
      if (!/VNPAY\s+TT\s+829168.*GIAIT.*989.*DV\s+CTT\s+NGAY/i.test(t.description || "")) continue;
      if (/OFFLINE/i.test(t.description || "")) continue;
      if (from && t.date < from) continue;
      if (to   && t.date > to)   continue;

      // Parse "NGAY 14-16.08.26" hoặc "NGAY 03.09.26"
      const m = /NGAY\s+(\d{1,2})(?:-(\d{1,2}))?\.(\d{1,2})\.(\d{2,4})/i.exec(t.description || "");
      let fromDate = null, toDate = null;
      if (m) {
        const d1 = m[1].padStart(2,"0"), d2 = (m[2]||m[1]).padStart(2,"0");
        const mo = m[3].padStart(2,"0");
        const yr = m[4].length === 2 ? "20" + m[4] : m[4];
        fromDate = d1 + "-" + mo + "-" + yr;
        toDate   = d2 + "-" + mo + "-" + yr;
      }
      let payDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const p = t.date.split("-");
        payDate = p[2] + "-" + p[1] + "-" + p[0];
      }
      payments.push({ payDate, amount: Number(t.amount || 0), fromDate, toDate, desc: (t.description||"").substring(0,100) });
    }
    payments.sort((a, b) => {
      const iso = d => { if (!d) return ""; const p = d.split("-"); return p[2]+"-"+p[1]+"-"+p[0]; };
      return iso(a.payDate).localeCompare(iso(b.payDate));
    });
    res.json({ ok: true, payments });
  } catch(err) { res.json({ ok: false, error: err.message }); }
});

module.exports = router;
