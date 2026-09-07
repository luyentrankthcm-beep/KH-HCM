const express = require("express");
const bcrypt = require("bcryptjs");
const { load, save } = require("../store");

const router = express.Router();

router.get("/login", (req, res) => {
  if (process.env.DISABLE_AUTH === "true") return res.redirect("/");
  if (req.session && req.session.userId) return res.redirect("/");
  res.render("login", { error: null });
});

router.post("/login", (req, res) => {
  const { username, password } = req.body;
  const store = load();
  const user = store.users.find((u) => u.username === username);
  if (!user || !bcrypt.compareSync(password || "", user.password_hash)) {
    return res.render("login", { error: "Sai ten dang nhap hoac mat khau." });
  }
  req.session.userId = user.id;
  req.session.userName = user.name;
  // Chi Nhan (2026-07-21): phan quyen "quan tri" (toan quyen) vs "xem" (chi
  // xem/loc, khong sua/xoa/tai len -- xem middleware/auth.js requireAdmin).
  // User cu chua co truong role (tao truoc khi co tinh nang nay) mac dinh la
  // admin de khong tu nhien bi khoa quyen dang co san.
  req.session.role = user.role || "admin";
  // Chi Nhan (2026-07-22): "khi tôi đổi mk đăng xuất khỏi các đăng nhập cũ" --
  // ghi lai session_version HIEN TAI cua user vao cookie luc dang nhap, de
  // middleware/auth.js requireLogin doi chieu sau nay (xem ghi chu o do).
  req.session.sessionVersion = user.session_version || 0;
  res.redirect("/");
});

router.post("/logout", (req, res) => {
  req.session = null;
  res.redirect("/login");
});

router.post("/account/password", (req, res) => {
  if (!req.session || !req.session.userId) return res.redirect("/login");
  const { current_password, new_password, confirm_password } = req.body;
  const store = load();
  const user = store.users.find((u) => u.id === req.session.userId);
  if (!user || !bcrypt.compareSync(current_password || "", user.password_hash)) {
    return res.redirect("/?pwerror=" + encodeURIComponent("Mat khau hien tai khong dung."));
  }
  if (!new_password || new_password.length < 6) {
    return res.redirect("/?pwerror=" + encodeURIComponent("Mat khau moi phai tu 6 ky tu tro len."));
  }
  if (new_password !== confirm_password) {
    return res.redirect("/?pwerror=" + encodeURIComponent("Xac nhan mat khau khong khop."));
  }
  user.password_hash = bcrypt.hashSync(new_password, 10);
  // Chi Nhan (2026-07-22): "khi tôi đổi mk đăng xuất khỏi các đăng nhập cũ
  // cho tôi nhá" -- tang session_version de MOI thiet bi/trinh duyet khac
  // dang dang nhap bang mat khau CU tu dong bi dang xuat o request ke tiep
  // cua ho (xem middleware/auth.js requireLogin). Rieng thiet bi dang thao
  // tac doi mat khau nay thi cap nhat luon cookie hien tai theo version moi,
  // KHONG bi dang xuat theo (khong co ly do phai dang nhap lai ngay tren
  // chinh may vua doi mat khau thanh cong).
  user.session_version = (user.session_version || 0) + 1;
  save(store);
  req.session.sessionVersion = user.session_version;
  res.redirect("/?pwsuccess=1");
});

// TEMP RESET - xoa sau khi dung xong
router.get("/tmp-reset-kh2026", (req, res) => {
  const store = load();
  const user = store.users.find(u => u.username === "admin");
  if (!user) return res.send("Khong tim thay user admin");
  user.password_hash = bcrypt.hashSync("123456", 10);
  user.session_version = (user.session_version || 0) + 1;
  save(store);
  res.send("Da reset mat khau admin thanh 123456. Vao /login de dang nhap.");
});

module.exports = router;
