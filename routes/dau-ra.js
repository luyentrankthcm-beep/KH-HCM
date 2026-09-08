const express = require("express");
const router = express.Router();
const { requireLogin } = require("../middleware/auth");

router.use(requireLogin);

router.get("/dau-ra", (req, res) => {
  res.render("dau-ra", { user: req.session.user });
});

module.exports = router;
