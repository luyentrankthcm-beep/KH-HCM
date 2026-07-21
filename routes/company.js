const express = require("express");
const { requireLogin } = require("../middleware/auth");
const { COMPANIES } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

// Nut chuyen cong ty tren topbar (partials/head.ejs) POST vao day roi quay
// lai dung trang dang xem (redirectTo = req.path luc render, xem
// res.locals.currentPath trong server.js) -- doi cong ty khong lam mat
// trang dang xem, chi lam moi lai du lieu theo cong ty vua chon.
router.post("/chon-cong-ty", (req, res) => {
  const { company, redirectTo } = req.body;
  if (COMPANIES[company]) {
    req.session.company = company;
  }
  const safeRedirect = redirectTo && redirectTo.startsWith("/") ? redirectTo : "/";
  res.redirect(safeRedirect);
});

module.exports = router;
