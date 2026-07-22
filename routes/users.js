const express = require("express");
const bcrypt = require("bcryptjs");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan (2026-07-21): "1 tai khoan quan tri duoc cap nhat/xem/lam tat ca
// chinh sua, 1 tai khoan chi duoc xem chon bo loc khong duoc xoa hay tai len
// bat cu gi". Trang nay CHI danh cho quan tri vien -- ke ca xem danh sach tai
// khoan (khong phai chi thao tac ghi) deu chan requireAdmin ngay tu GET, vi
// day la thong tin quan tri he thong (ai dang nhap duoc), khong phai du lieu
// nghiep vu can cho tai khoan "chi xem" tham khao.
router.use(requireAdmin);

function ensureRoles(store) {
  let changed = false;
  (store.users || []).forEach((u) => {
    if (!u.role) {
      u.role = "admin"; // tai khoan tao truoc khi co tinh nang phan quyen -- giu nguyen quyen dang co.
      changed = true;
    }
  });
  return changed;
}

router.get("/he-thong/nguoi-dung", (req, res) => {
  const store = load();
  if (ensureRoles(store)) save(store);
  res.render("users", {
    userName: req.session.userName,
    users: store.users,
    currentUserId: req.session.userId,
    ghiChu: store.ghi_chu_he_thong || "",
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Chi Nhan (2026-07-22): "nam phia duoi tai khoan dang nhap them khung note
// ghi chu cho toi" -- 1 khung ghi chu chung, dung chung cho MOI tai khoan
// quan tri (khong rieng cho tung nguoi), de ghi lai nhung dieu can nho ve he
// thong (vd mat khau tam, viec can lam...). Luu truc tiep vao store, khong
// gioi han do dai.
router.post("/he-thong/nguoi-dung/ghi-chu", (req, res) => {
  const store = load();
  store.ghi_chu_he_thong = req.body.ghiChu || "";
  save(store);
  res.redirect("/he-thong/nguoi-dung?success=" + encodeURIComponent("Da luu ghi chu."));
});

router.post("/he-thong/nguoi-dung", (req, res) => {
  const store = load();
  ensureRoles(store);
  try {
    const username = (req.body.username || "").trim();
    const name = (req.body.name || "").trim();
    const password = req.body.password || "";
    const role = req.body.role === "viewer" ? "viewer" : "admin";
    if (!username || !name) throw new Error("Vui long dien Ten dang nhap va Ten hien thi.");
    if (!password || password.length < 6) throw new Error("Mat khau phai tu 6 ky tu tro len.");
    if (store.users.some((u) => u.username.toLowerCase() === username.toLowerCase())) {
      throw new Error(`Ten dang nhap "${username}" da co roi, chon ten khac.`);
    }
    store.users.push({
      id: nextId(store, "users"),
      username,
      name,
      role,
      password_hash: bcrypt.hashSync(password, 10),
      created_at: new Date().toISOString(),
    });
    save(store);
    res.redirect("/he-thong/nguoi-dung?success=" + encodeURIComponent(`Da tao tai khoan "${username}" (${role === "admin" ? "Quan tri" : "Chi xem"}).`));
  } catch (e) {
    res.redirect("/he-thong/nguoi-dung?error=" + encodeURIComponent(e.message));
  }
});

router.post("/he-thong/nguoi-dung/:id/role", (req, res) => {
  const store = load();
  ensureRoles(store);
  try {
    const user = store.users.find((u) => String(u.id) === req.params.id);
    if (!user) throw new Error("Khong tim thay tai khoan.");
    const role = req.body.role === "viewer" ? "viewer" : "admin";
    // Khong cho tu ha quyen chinh minh xuong "chi xem" -- tranh tu khoa minh
    // khoi cac trang quan tri (vd chinh trang nay) ma khong con ai co the
    // vao lai doi nguoc lai (neu do la tai khoan quan tri duy nhat).
    if (String(user.id) === String(req.session.userId) && role !== "admin") {
      throw new Error("Khong the tu ha quyen chinh tai khoan dang dang nhap.");
    }
    const otherAdmins = store.users.filter((u) => u.role === "admin" && String(u.id) !== String(user.id));
    if (user.role === "admin" && role !== "admin" && otherAdmins.length === 0) {
      throw new Error("Day la tai khoan Quan tri duy nhat -- can giu lai it nhat 1 tai khoan Quan tri.");
    }
    user.role = role;
    save(store);
    res.redirect("/he-thong/nguoi-dung?success=" + encodeURIComponent(`Da doi quyen tai khoan "${user.username}".`));
  } catch (e) {
    res.redirect("/he-thong/nguoi-dung?error=" + encodeURIComponent(e.message));
  }
});

router.post("/he-thong/nguoi-dung/:id/mat-khau", (req, res) => {
  const store = load();
  try {
    const user = store.users.find((u) => String(u.id) === req.params.id);
    if (!user) throw new Error("Khong tim thay tai khoan.");
    const password = req.body.password || "";
    if (!password || password.length < 6) throw new Error("Mat khau moi phai tu 6 ky tu tro len.");
    user.password_hash = bcrypt.hashSync(password, 10);
    // Chi Nhan (2026-07-22): "khi tôi đổi mk đăng xuất khỏi các đăng nhập cũ
    // cho tôi nhá" -- ap dung ca khi ADMIN dat lai mat khau cho tai khoan
    // khac: tang session_version de moi noi tai khoan do dang dang nhap san
    // (bang mat khau CU) bi dang xuat o request ke tiep -- xem
    // middleware/auth.js requireLogin.
    user.session_version = (user.session_version || 0) + 1;
    save(store);
    res.redirect("/he-thong/nguoi-dung?success=" + encodeURIComponent(`Da dat lai mat khau cho tai khoan "${user.username}".`));
  } catch (e) {
    res.redirect("/he-thong/nguoi-dung?error=" + encodeURIComponent(e.message));
  }
});

router.post("/he-thong/nguoi-dung/:id/delete", (req, res) => {
  const store = load();
  try {
    const user = store.users.find((u) => String(u.id) === req.params.id);
    if (!user) throw new Error("Khong tim thay tai khoan.");
    if (String(user.id) === String(req.session.userId)) {
      throw new Error("Khong the tu xoa tai khoan dang dang nhap.");
    }
    ensureRoles(store);
    const otherAdmins = store.users.filter((u) => u.role === "admin" && String(u.id) !== String(user.id));
    if (user.role === "admin" && otherAdmins.length === 0) {
      throw new Error("Day la tai khoan Quan tri duy nhat -- can giu lai it nhat 1 tai khoan Quan tri.");
    }
    store.users = store.users.filter((u) => String(u.id) !== req.params.id);
    save(store);
    res.redirect("/he-thong/nguoi-dung?success=" + encodeURIComponent(`Da xoa tai khoan "${user.username}".`));
  } catch (e) {
    res.redirect("/he-thong/nguoi-dung?error=" + encodeURIComponent(e.message));
  }
});

module.exports = router;
