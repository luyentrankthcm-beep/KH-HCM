// Nhan, 2026-07-22: "thêm cho tôi 1 nút cập nhật tìm hóa đơn ... tìm trên
// gmail" -- ket noi Gmail that su (OAuth2) de nut "Cap nhat tim hoa don" o
// trang Chi Phi co the tu tim, khong can qua Claude nua.
//
// CAN 3 BIEN MOI TRUONG tren Railway (Variables cua service):
//   GOOGLE_CLIENT_ID       -- tu Google Cloud Console, OAuth client "Web application"
//   GOOGLE_CLIENT_SECRET   -- cung cho o do
//   GOOGLE_REDIRECT_URI    -- vd https://kh-hcm-production.up.railway.app/gmail/callback
//                             (PHAI khop CHINH XAC voi "Authorized redirect URIs"
//                             da khai bao luc tao OAuth client tren Google)
//
// Cach tao (lam 1 lan, chi Nhan tu lam vi lien quan tai khoan Google rieng):
//   1. console.cloud.google.com -> tao project moi.
//   2. "APIs & Services" -> "Enabled APIs" -> bat "Gmail API".
//   3. "APIs & Services" -> "OAuth consent screen" -> chon "External", dien
//      thong tin co ban, o "Test users" them chinh email dang dung cho web.
//   4. "APIs & Services" -> "Credentials" -> "Create Credentials" -> "OAuth
//      client ID" -> loai "Web application" -> "Authorized redirect URIs"
//      dien dung GOOGLE_REDIRECT_URI o tren.
//   5. Copy Client ID + Client Secret dan vao Railway Variables, deploy lai.
//
// Sau khi co du 3 bien, vao trang Chi Phi se thay nut "Ket noi Gmail" (thay
// vi bao chua cau hinh) -- bam vao, dang nhap + cho phep, xong la nut "Cap
// nhat tim hoa don" dung duoc.

const express = require("express");
const { load, save } = require("../store");
const { requireAdmin } = require("../middleware/auth");
const gmailApi = require("../utils/gmailApi");

const router = express.Router();

router.get("/gmail/connect", requireAdmin, (req, res) => {
  if (!gmailApi.isConfigured()) {
    return res.redirect(
      "/chi-phi?error=" +
        encodeURIComponent(
          "Chua cau hinh Google OAuth tren Railway (thieu GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET/GOOGLE_REDIRECT_URI). Xem huong dan trong routes/gmail-oauth.js."
        )
    );
  }
  res.redirect(gmailApi.buildAuthUrl());
});

router.get("/gmail/callback", requireAdmin, async (req, res) => {
  const { code, error } = req.query;
  if (error) {
    return res.redirect("/chi-phi?error=" + encodeURIComponent("Google bao loi: " + error));
  }
  if (!code) {
    return res.redirect("/chi-phi?error=" + encodeURIComponent("Thieu ma xac thuc tu Google."));
  }
  try {
    const tokens = await gmailApi.exchangeCodeForTokens(code);
    if (!tokens.refresh_token) {
      // Google chi tra refresh_token lan dau tien app duoc cap quyen (hoac
      // khi prompt=consent) -- neu thieu, bao chi thu ket noi lai (Google se
      // hoi lai man hinh cho phep vi prompt=consent da bat san trong
      // buildAuthUrl).
      throw new Error("Google khong tra ve refresh token -- thu bam 'Ket noi Gmail' lai lan nua.");
    }
    const store = load();
    store.gmail_oauth = {
      refresh_token: tokens.refresh_token,
      access_token: tokens.access_token,
      expiry_date: Date.now() + (tokens.expires_in || 3000) * 1000,
      connected_at: new Date().toISOString(),
    };
    save(store);
    res.redirect("/chi-phi?success=" + encodeURIComponent("Da ket noi Gmail thanh cong! Gio co the bam 'Cap nhat tim hoa don'."));
  } catch (e) {
    res.redirect("/chi-phi?error=" + encodeURIComponent("Loi ket noi Gmail: " + e.message));
  }
});

// Nhan, 2026-09-15: "toi lien ket voi 2 gg drive lan a" -- endpoint tam de
// kiem tra CHINH XAC tai khoan Google nao dang duoc he thong dung (chi 1 tai
// khoan duy nhat, xem ghi chu dau file utils/gmailApi.js), giup phan biet
// voi cac tai khoan khac ma Nhan co the da chia se file nham.
router.get("/gmail/whoami", requireAdmin, async (req, res) => {
  try {
    const store = load();
    const token = await gmailApi.getValidAccessToken(store);
    save(store);
    // Dung Gmail API "users/me/profile" (scope gmail.readonly da co san) thay
    // vi userinfo endpoint (can them scope "email"/"openid" chua xin) de lay
    // dia chi email cua tai khoan dang ket noi.
    const resp = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/profile", {
      headers: { Authorization: "Bearer " + token },
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(JSON.stringify(data));
    res.json({ success: true, email: data.emailAddress });
  } catch (e) {
    res.json({ success: false, error: e.message });
  }
});

module.exports = router;
