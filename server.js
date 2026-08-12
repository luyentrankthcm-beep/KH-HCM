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
app.use("/", authRoutes);

// TEMP: cap nhat phap_danh id=137 -- XOA SAU KHI DUNG
app.use(express.json({ limit: "1mb" }));
app.post("/temp-update-pd137", (req, res) => {
  if (req.body.secret !== "khbank-pd137-2026") return res.status(403).json({ error: "forbidden" });
  const { load, save } = require("./store");
  const store = load();
  const pd = (store.phap_danh_hop_dong_thue || []).find(r => r.id === 137);
  if (!pd) return res.status(404).json({ error: "not found" });
  pd.tenDiemNoiBo = "AE GO BA RIA KVC";
  pd.maCongTrinh  = "AE GO BA RIA KVC";
  pd.updatedAt    = new Date().toISOString();
  save(store);
  res.json({ ok: true, tenDiemNoiBo: pd.tenDiemNoiBo, maCongTrinh: pd.maCongTrinh });
});
// END TEMP

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
app.use("/", backupRoutes);
app.use("/", usersRoutes);
app.use("/", gmailOauthRoutes);

app.use((req, res) => {
  res.status(404).send("Khong tim thay trang.");
});

app.listen(PORT, () => {
  console.log(`K&H Bank Tracker dang chay tai http://localhost:${PORT}`);
});
