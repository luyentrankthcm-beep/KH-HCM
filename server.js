require("dotenv").config();
const express = require("express");
const path = require("path");
const cookieSession = require("cookie-session");

require("./store"); // ensures data file + seed data exist before routes load

const { COMPANIES, getCompany } = require("./utils/companies");
const companyRoutes = require("./routes/company");
const authRoutes = require("./routes/auth");
const { router: bankRoutes } = require("./routes/banks");
const transactionRoutes = require("./routes/transactions");
const dashboardRoutes = require("./routes/dashboard");
const doisoatRoutes = require("./routes/doisoat");
const doisoatZvpRoutes = require("./routes/doisoat-zvp");
const doisoatVietQrRoutes = require("./routes/doisoat-vietqr");
const doisoatVnpayKhMoiRoutes = require("./routes/doisoat-vnpay-khmoi");
const baocaoRoutes = require("./routes/baocao");
const congnoRoutes = require("./routes/congno");
const doisoatChiPhiRoutes = require("./routes/doisoat-chiphi");
const congnoNccRoutes = require("./routes/congno-ncc");
const hoaDonDauVaoRoutes = require("./routes/hoa-don-dau-vao");
const hoaDonDauRaRoutes = require("./routes/hoa-don-dau-ra");
const phapDanhRoutes = require("./routes/phap-danh");
const chiPhiRoutes = require("./routes/chi-phi");
const tongHopRoutes = require("./routes/tong-hop");
const danhMucRoutes = require("./routes/danh-muc");
const backupRoutes = require("./routes/backup");
const usersRoutes = require("./routes/users");
const gmailOauthRoutes = require("./routes/gmail-oauth");

const app = express();
const PORT = process.env.PORT || 3000;

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

app.use(
  cookieSession({
    name: "kh_bank_session",
    keys: [process.env.SESSION_SECRET || "doi-chuoi-bi-mat-nay-truoc-khi-deploy"],
    maxAge: 7 * 24 * 60 * 60 * 1000, // 7 days
  })
);

// Cong ty (KH Cu / KH Moi) dang duoc chon luu trong session -- dat vao
// res.locals o day (khong phai tung route rieng le) de moi view (nav.ejs,
// va sau nay cac trang khac neu can) tu dong co san activeCompany/COMPANIES
// ma khong phai sua res.render() cua toan bo cac route hien co.
app.use((req, res, next) => {
  res.locals.activeCompany = getCompany(req);
  res.locals.COMPANIES = COMPANIES;
  res.locals.currentPath = req.originalUrl || req.path;
  // Chi Nhan (2026-07-21): phan quyen quan tri/chi xem -- xem
  // middleware/auth.js (requireAdmin) cho phan chan o backend (day moi la
  // cho THAT SU chan sua/xoa/tai len). isAdmin o day chi phuc vu UI (nav.ejs
  // an bot nut/form thao tac cho tai khoan "chi xem" de do roi mat, KHONG
  // phai lop bao ve chinh).
  res.locals.isAdmin = !!(req.session && req.session.role === "admin");
  // Luyen (2026-07-24): them quyen "Nhap lieu" -- userRole dung cho UI (badge
  // trong nav.ejs) de phan biet 3 quyen, KHONG phai lop bao ve chinh (xem
  // middleware/auth.js requireAdmin/requireDataEntry cho chan o backend).
  res.locals.userRole = (req.session && req.session.role) || "viewer";
  next();
});

// Luyen, 2026-08-03: "mỗi lần tôi chỉnh báo cáo hay chỉnh khớp công nợ thì
// giữ trang nguyên cập nhật thôi chứ cứ đưa lên đầu đẩy tới hiện tại thì nó
// lại mất thời gian chỉnh bộ lọc nx" -- hau het cac form POST (doi soat
// Momo/VietQR/ZVP/Cong no, Chi Phi...) sau khi luu/xoa deu redirect ve 1 URL
// "sach" (vd "/doi-soat/vietqr?success=...") MAT HET query string dang xem
// (channel, month, stmtMonth...), buoc Luyen phai chinh lai bo loc tu dau
// moi lan sua 1 dong. Thay vi sua tung res.redirect() rieng le (250+ cho
// trong toan bo routes/), chan CHUNG 1 lan o day: neu URL sap redirect toi
// co CUNG duong dan (pathname) voi trang vua gui form (Referer), tu dong GIU
// LAI moi query param CU (channel, month...) ma route KHONG chu dinh ghi de
// (vd "success"/"error" moi van uu tien gia tri route dat, chi bo sung them
// cac param con thieu). Ket hop voi views/partials/scroll-restore.ejs (nho
// + khoi phuc vi tri cuon trang) de giai quyet tron ven ca 2 y Luyen neu: bo
// loc VA vi tri dang xem deu giu nguyen sau khi luu/xoa.
app.use((req, res, next) => {
  const originalRedirect = res.redirect.bind(res);
  res.redirect = function (statusOrUrl, maybeUrl) {
    const hasStatus = typeof statusOrUrl === "number";
    const url = hasStatus ? maybeUrl : statusOrUrl;
    try {
      const referer = req.get("Referrer") || req.get("Referer");
      if (referer && typeof url === "string" && url.startsWith("/")) {
        const target = new URL(url, `${req.protocol}://${req.get("host")}`);
        const refererUrl = new URL(referer);
        if (target.pathname === refererUrl.pathname) {
          refererUrl.searchParams.forEach((value, key) => {
            if (!target.searchParams.has(key)) target.searchParams.set(key, value);
          });
          const mergedUrl = target.pathname + target.search;
          return hasStatus ? originalRedirect(statusOrUrl, mergedUrl) : originalRedirect(mergedUrl);
        }
      }
    } catch (e) {
      // URL/Referer khong hop le (vd redirect sang domain khac) -- giu nguyen hanh vi cu.
    }
    return hasStatus ? originalRedirect(statusOrUrl, maybeUrl) : originalRedirect(statusOrUrl);
  };
  next();
});

// QUAN TRONG: authRoutes phai duoc dang ky TRUOC companyRoutes. companyRoutes
// tu goi router.use(requireLogin) (giong moi router khac), va middleware do
// chay cho MOI request di qua no bat ke co khop route nao ben trong hay
// khong -- neu dat truoc authRoutes thi ngay ca GET /login (trang cong khai)
// cung se bi requireLogin chan va redirect vong lai chinh /login, khong ai
// dang nhap duoc. Dat sau authRoutes (cung vi tri nhu bankRoutes/... ben
// duoi) de /login/logout luon duoc xu ly truoc, giong quy uoc san co.

// TEMP: diagnostic + fix (xoa sau khi chay)
app.post("/temp-diag2026", (req, res) => {
  if ((req.body || {}).secret !== "diag2026x") return res.status(403).json({ error: "forbidden" });
  const { load: ld, save: sv } = require("./store");
  const store = ld();
  const action = req.body.action || "query";

  if (action === "query") {
    // Tim tat ca QR OFFLINE transactions thang 8
    const offline = store.transactions.filter(t =>
      t.type === "thu" && t.date >= "2026-08-01" && /QR.OFFLINE/i.test(t.description || "")
    );
    const allAug = store.transactions.filter(t => t.date >= "2026-08-01" && t.type === "thu")
      .map(t => {
        const bank = (store.banks || []).find(b => b.id === t.bank_id);
        return { id: t.id, bank: bank?.name, date: t.date, amount: t.amount, desc: (t.description || "").slice(0, 80) };
      });
    // Farm Times City in payoo uploads
    const payooUploads = store.zvp_payoo_uploads || [];
    const farmKeys = [];
    payooUploads.forEach(u => {
      Object.keys(u.grossByCode || {}).forEach(k => {
        if (/farm.times/i.test(k)) farmKeys.push({ uploadId: u.id, key: k, val: u.grossByCode[k] });
      });
    });
    // zvp_payoo_raw_tx farm
    const rawTx = store.zvp_payoo_raw_tx || {};
    const farmRaw = Object.entries(rawTx).filter(([k, v]) => /farm.times/i.test(v.gian || "")).map(([k, v]) => ({ key: k, ...v }));
    // SB PHU QUOC PHCM
    const allUploads = [
      ...(store.momo_uploads || []).map(u => ({ src: "momo", ...u })),
      ...(store.zvp_online_uploads || []).map(u => ({ src: "online", ...u })),
      ...(store.zvp_offline_uploads || []).map(u => ({ src: "offline", ...u })),
      ...(store.viet_qr_uploads || []).map(u => ({ src: "vietqr", ...u })),
    ];
    const sbKeys = [];
    allUploads.forEach(u => {
      Object.keys(u.grossByCode || {}).filter(k => /sb.phu.quoc/i.test(k) || /SB PHU/i.test(k))
        .forEach(k => sbKeys.push({ src: u.src, uploadId: u.id, key: k }));
    });
    return res.json({ offlineAugTx: offline.length, farmPayooUpload: farmKeys, farmRawTx: farmRaw, sbPhuQuoc: sbKeys, augThu: allAug.slice(0, 20) });
  }

  if (action === "delete_farm_payoo") {
    // Xoa Farm Times City khoi zvp_payoo_raw_tx
    const before = Object.keys(store.zvp_payoo_raw_tx || {}).length;
    const newRaw = {};
    Object.entries(store.zvp_payoo_raw_tx || {}).forEach(([k, v]) => {
      if (!/farm.times/i.test(v.gian || "")) newRaw[k] = v;
    });
    store.zvp_payoo_raw_tx = newRaw;
    // Rebuild uploads - xoa farm times khoi grossByCode
    (store.zvp_payoo_uploads || []).forEach(u => {
      if (!u.grossByCode) return;
      Object.keys(u.grossByCode).forEach(k => {
        if (/farm.times/i.test(k)) delete u.grossByCode[k];
      });
      if (u.netByCode) Object.keys(u.netByCode).forEach(k => {
        if (/farm.times/i.test(k)) delete u.netByCode[k];
      });
    });
    sv(store);
    const after = Object.keys(store.zvp_payoo_raw_tx || {}).length;
    return res.json({ ok: true, rawTxBefore: before, rawTxAfter: after });
  }

  if (action === "delete_tx_by_id") {
    const id = Number(req.body.txId);
    const before = store.transactions.length;
    store.transactions = store.transactions.filter(t => t.id !== id);
    sv(store);
    return res.json({ ok: true, deleted: before - store.transactions.length });
  }

  res.json({ error: "unknown action" });
});

// TEMP: fix Farm Times City payoo invoices + SB PHU QUOC PHCM (xoa sau khi chay)
app.post("/temp-fix2-aug2026", (req, res) => {
  if ((req.body || {}).secret !== "fix2aug2026") return res.status(403).json({ error: "forbidden" });
  const { load: ld, save: sv } = require("./store");
  const store = ld();
  const results = {};

  // 1. Xoa Farm Times City khoi zvp_invoices.payoo
  if (store.zvp_invoices && store.zvp_invoices.payoo) {
    const before = store.zvp_invoices.payoo.length;
    store.zvp_invoices.payoo = store.zvp_invoices.payoo.filter(
      (i) => !/farm.times/i.test(i.maDiem || "")
    );
    results.payooFarmRemoved = before - store.zvp_invoices.payoo.length;
    results.payooTotal = store.zvp_invoices.payoo.length;
  }

  // 2. Fix SB PHU QUOC PHCM -> CHKQT PHU QUOC trong mb11521268
  if (!store.viet_qr_store_code_override) store.viet_qr_store_code_override = {};
  if (!store.viet_qr_store_code_override.mb11521268) store.viet_qr_store_code_override.mb11521268 = {};
  store.viet_qr_store_code_override.mb11521268["SB PHU QUOC PHCM"] = "CHKQT PHU QUOC";
  // Dam bao gian_merge cung dung
  if (!store.viet_qr_gian_merge) store.viet_qr_gian_merge = {};
  store.viet_qr_gian_merge["SB PHU QUOC PHCM"] = { maCongTrinh: "CHKQT PHU QUOC", isCse: false };
  results.sbPhuQuocFixed = true;

  // 3. Bao cao offline QR transactions hien tai
  const offlineAug = store.transactions.filter(
    (t) => t.type === "thu" && t.date >= "2026-08-01" && /QR.OFFLINE/i.test(t.description || "")
  );
  results.offlineAugCount = offlineAug.length;
  results.offlineTxs = offlineAug.map((t) => ({
    id: t.id,
    date: t.date,
    amount: t.amount,
    desc: (t.description || "").slice(0, 60),
  }));

  sv(store);
  res.json({ ok: true, ...results });
});

app.use("/", authRoutes);

app.use("/", companyRoutes);
app.use("/", bankRoutes);
app.use("/", transactionRoutes);
app.use("/", dashboardRoutes);
app.use("/", doisoatRoutes);
app.use("/", doisoatZvpRoutes);
app.use("/", doisoatVietQrRoutes);
app.use("/", doisoatVnpayKhMoiRoutes);
app.use("/", baocaoRoutes);
app.use("/", congnoRoutes);
app.use("/", congnoNccRoutes);
app.use("/", hoaDonDauVaoRoutes);
app.use("/", hoaDonDauRaRoutes);
app.use("/", doisoatChiPhiRoutes);
app.use("/", phapDanhRoutes);
app.use("/", chiPhiRoutes);
app.use("/", tongHopRoutes);
app.use("/", danhMucRoutes);
app.use("/", backupRoutes);
app.use("/", usersRoutes);
app.use("/", gmailOauthRoutes);

app.use((req, res) => {
  res.status(404).send("Khong tim thay trang.");
});

app.listen(PORT, () => {
  console.log(`K&H Bank Tracker dang chay tai http://localhost:${PORT}`);
});
