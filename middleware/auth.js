const { load } = require("../store");

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
  if (req.session && req.session.role === "admin") {
    return next();
  }
  const msg = "Tai khoan nay chi duoc xem, khong co quyen chinh sua/xoa/tai len du lieu.";
  const back = req.get("Referer") || "/";
  const sep = back.includes("?") ? "&" : "?";
  return res.redirect(back + sep + "error=" + encodeURIComponent(msg));
}

module.exports = { requireLogin, requireAdmin };
