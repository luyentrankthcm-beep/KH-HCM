const { load } = require("../store");

// Chi Nhan (2026-07-22): "offline bỏ cái mật khẩu luôn đi" -- CHI ap dung cho
// ban chay OFFLINE tren may cua Nhan, tuyet doi KHONG duoc bat bien nay tren
// Railway (ban online): dat DISABLE_AUTH=true trong file .env O MAY (khong
// commit file .env len git, va KHONG duoc tao bien nay trong Railway
// Settings > Variables) -- khi bat, moi trang tu dong dang nhap san bang tai
// khoan admin dau tien, khong can nhap mat khau nua.
const AUTH_DISABLED = process.env.DISABLE_AUTH === "true";

function autoLoginSession(req) {
  const store = load();
  const user = (store.users || []).find((u) => u.role === "admin") || (store.users || [])[0];
  if (!user) return null;
  req.session = req.session || {};
  req.session.userId = user.id;
  req.session.userName = user.name;
  req.session.role = user.role || "admin";
  req.session.sessionVersion = user.session_version || 0;
  return user;
}

// Chi Nhan (2026-07-22): "khi tôi đổi mk đăng xuất khỏi các đăng nhập cũ cho
// tôi nhá" -- session dung cookie-session (KHONG co session store phia
// server, toan bo du lieu session nam trong cookie ky ten o may nguoi dung),
// nen KHONG co danh sach "cac session dang mo" de server chu dong dang xuat
// tu xa. Cach lam: moi user co 1 "session_version" (so nguyen, tang len moi
// lan doi mat khau -- xem routes/auth.js POST /account/password va
// routes/users.js POST /:id/mat-khau). Luc dang nhap, session_version HIEN
// TAI duoc ghi vao cookie. Moi request sau do, so sanh voi session_version
// MOI NHAT trong store: neu KHAC (tuc mat khau da doi sau khi cookie nay
// duoc tao) thi coi nhu het han, buoc dang xuat -- ap dung cho MOI thiet
// bi/trinh duyet dang dang nhap bang mat khau CU, ke ca may khong dung de
// doi mat khau.
function requireLogin(req, res, next) {
  // Luyen 2026-08-31: bypass cho Claude Cowork sync (X-Internal-Key header)
  const INTERNAL_SYNC_KEY = process.env.INTERNAL_SYNC_KEY || "";
  if (INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY) {
    return next();
  }
  if (AUTH_DISABLED) {
    const user = autoLoginSession(req);
    if (user) return next();
  }
  if (req.session && req.session.userId) {
    const store = load();
    const user = (store.users || []).find((u) => u.id === req.session.userId);
    if (!user) {
      req.session = null;
      return res.redirect("/login");
    }
    const currentVersion = user.session_version || 0;
    if ((req.session.sessionVersion || 0) !== currentVersion) {
      req.session = null;
      return res.redirect("/login?error=" + encodeURIComponent("Mat khau vua duoc doi -- vui long dang nhap lai."));
    }
    return next();
  }
  return res.redirect("/login");
}

// Chi Nhan (2026-07-21): "1 tai khoan quan tri duoc cap nhat/xem/lam tat ca
// chinh sua, 1 tai khoan chi duoc xem chon bo loc khong duoc xoa hay tai len
// bat cu gi" -- chan TRUOC KHI route xu ly bat ky thao tac ghi/sua/xoa/tai
// len nao (moi router.post ngoai tru /login, /logout, /chon-cong-ty,
// /account/password da duoc gan middleware nay). Neu khong phai admin, quay
// ve trang truoc (Referer) kem thong bao loi thay vi thuc hien thao tac.
function requireAdmin(req, res, next) {
  if (AUTH_DISABLED) return next();
  // Luyen 2026-08-31: bypass cho Claude Cowork sync
  const INTERNAL_SYNC_KEY = process.env.INTERNAL_SYNC_KEY || "";
  if (INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY) return next();
  if (req.session && req.session.role === "admin") {
    return next();
  }
  const msg = "Tai khoan nay khong co quyen thao tac nay (chi Quan tri moi lam duoc -- xoa, khoa so, cau hinh TK/mapping, tai khoan/ngan hang...).";
  const back = req.get("Referer") || "/";
  const sep = back.includes("?") ? "&" : "?";
  return res.redirect(back + sep + "error=" + encodeURIComponent(msg));
}

// Luyen (2026-07-24): them quyen thu 3 "Nhap lieu" -- duoc them/sua/tai len
// du lieu nghiep vu (giao dich, doi soat Momo/Zalo-VNPay-Payoo/VietQR, chi
// phi, cong no...) NHUNG khong duoc xoa bat cu gi, va khong duoc dung toi
// cac trang cau hinh he thong (Tai khoan, Ngan hang, Phap danh hop dong,
// mapping TK Co/TK No, khoa so, an gian...) -- nhung trang/thao tac do van
// chi dung requireAdmin nhu cu (khong doi gi), chi NHUNG route ro rang la
// "nhap du lieu" (tai file len, them/sua 1 dong cu the, dien ma con thieu
// cho 1 dong/giao dich cu the) moi doi sang requireDataEntry o duoi day.
function requireDataEntry(req, res, next) {
  if (AUTH_DISABLED) return next();
  // Luyen 2026-08-31: bypass cho Claude Cowork sync
  const INTERNAL_SYNC_KEY = process.env.INTERNAL_SYNC_KEY || "";
  if (INTERNAL_SYNC_KEY && req.headers["x-internal-key"] === INTERNAL_SYNC_KEY) return next();
  if (req.session && (req.session.role === "admin" || req.session.role === "nhap_lieu")) {
    return next();
  }
  const msg = "Tai khoan nay chi duoc xem, khong co quyen nhap/sua/tai len du lieu.";
  const back = req.get("Referer") || "/";
  const sep = back.includes("?") ? "&" : "?";
  return res.redirect(back + sep + "error=" + encodeURIComponent(msg));
}

module.exports = { requireLogin, requireAdmin, requireDataEntry };
