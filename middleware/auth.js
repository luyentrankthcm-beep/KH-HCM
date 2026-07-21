function requireLogin(req, res, next) {
  if (req.session && req.session.userId) {
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
