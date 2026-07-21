const express = require("express");
const bcrypt = require("bcryptjs");
const { load, save } = require("../store");

const router = express.Router();

router.get("/login", (req, res) => {
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
  save(store);
  res.redirect("/?pwsuccess=1");
});

module.exports = router;
