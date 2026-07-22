// Nhan, 2026-07-22: "thêm cho tôi 1 nút cập nhật tìm hóa đơn ... tìm trên
// gmail" -- module goi thang Gmail API bang OAuth2 (khong dung goi googleapis
// cho nhe, dung fetch nhu cac cho khac trong app nay). Chi 1 tai khoan Gmail
// duy nhat (ktkh.hcm@poshvn.com) nen luu token thang trong store.json
// (store.gmail_oauth), khong can quan ly nhieu user.
//
// Can 3 bien moi truong (dat tren Railway): GOOGLE_CLIENT_ID,
// GOOGLE_CLIENT_SECRET, GOOGLE_REDIRECT_URI (vd
// https://kh-hcm-production.up.railway.app/gmail/callback) -- lay tu Google
// Cloud Console (OAuth client id loai "Web application"), xem huong dan
// trong routes/gmail-oauth.js.

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID || "";
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || "";
const REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || "";
const SCOPE = "https://www.googleapis.com/auth/gmail.readonly";

function isConfigured() {
  return !!(CLIENT_ID && CLIENT_SECRET && REDIRECT_URI);
}

function buildAuthUrl(state) {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    redirect_uri: REDIRECT_URI,
    response_type: "code",
    scope: SCOPE,
    access_type: "offline",
    prompt: "consent",
    state: state || "",
  });
  return "https://accounts.google.com/o/oauth2/v2/auth?" + params.toString();
}

async function exchangeCodeForTokens(code) {
  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      redirect_uri: REDIRECT_URI,
      grant_type: "authorization_code",
    }),
  });
  const data = await resp.json();
  if (!resp.ok) throw new Error("Loi xac thuc Google: " + (data.error_description || data.error || resp.status));
  return data; // { access_token, refresh_token, expires_in, ... }
}

async function refreshAccessToken(refreshToken) {
  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      refresh_token: refreshToken,
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      grant_type: "refresh_token",
    }),
  });
  const data = await resp.json();
  if (!resp.ok) throw new Error("Loi lam moi token Google: " + (data.error_description || data.error || resp.status));
  return data; // { access_token, expires_in, ... } (khong co refresh_token moi)
}

// Tra ve access_token con hieu luc, tu lam moi (va luu lai store) neu can.
// `store` la object da load() tu store.js -- ham nay CO THE sua
// store.gmail_oauth va CALLER phai tu goi save(store) sau khi dung xong (de
// khop voi pattern load/save dang dung trong toan bo app).
async function getValidAccessToken(store) {
  if (!isConfigured()) throw new Error("Chua cau hinh GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET/GOOGLE_REDIRECT_URI tren Railway.");
  const tok = store.gmail_oauth;
  if (!tok || !tok.refresh_token) throw new Error("Chua ket noi Gmail -- vao trang Chi Phi bam 'Ket noi Gmail' truoc.");
  const now = Date.now();
  if (tok.access_token && tok.expiry_date && now < tok.expiry_date - 60000) {
    return tok.access_token;
  }
  const fresh = await refreshAccessToken(tok.refresh_token);
  store.gmail_oauth.access_token = fresh.access_token;
  store.gmail_oauth.expiry_date = now + (fresh.expires_in || 3000) * 1000;
  return store.gmail_oauth.access_token;
}

async function searchMessages(store, query, maxResults) {
  const token = await getValidAccessToken(store);
  const params = new URLSearchParams({ q: query, maxResults: String(maxResults || 10) });
  const resp = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages?" + params.toString(), {
    headers: { Authorization: "Bearer " + token },
  });
  const data = await resp.json();
  if (!resp.ok) throw new Error("Loi tim Gmail: " + (data.error && data.error.message ? data.error.message : resp.status));
  return data.messages || []; // [{id, threadId}, ...]
}

// De qui walk payload.parts de gom text/plain + text/html body (base64url).
function decodeBase64Url(b64url) {
  if (!b64url) return "";
  const b64 = b64url.replace(/-/g, "+").replace(/_/g, "/");
  return Buffer.from(b64, "base64").toString("utf8");
}

function extractBodies(payload) {
  let plain = "";
  let html = "";
  function walk(part) {
    if (!part) return;
    if (part.mimeType === "text/plain" && part.body && part.body.data) plain += decodeBase64Url(part.body.data);
    if (part.mimeType === "text/html" && part.body && part.body.data) html += decodeBase64Url(part.body.data);
    if (part.parts) part.parts.forEach(walk);
  }
  walk(payload);
  return { plain, html };
}

async function getMessage(store, id) {
  const token = await getValidAccessToken(store);
  const resp = await fetch(
    "https://gmail.googleapis.com/gmail/v1/users/me/messages/" + id + "?format=full",
    { headers: { Authorization: "Bearer " + token } }
  );
  const data = await resp.json();
  if (!resp.ok) throw new Error("Loi doc email Gmail: " + (data.error && data.error.message ? data.error.message : resp.status));
  const headers = {};
  ((data.payload && data.payload.headers) || []).forEach((h) => { headers[h.name.toLowerCase()] = h.value; });
  const { plain, html } = extractBodies(data.payload);
  return { id: data.id, threadId: data.threadId, headers, plain, html, internalDate: data.internalDate };
}

module.exports = { isConfigured, buildAuthUrl, exchangeCodeForTokens, getValidAccessToken, searchMessages, getMessage };
