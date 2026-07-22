// Nhan, 2026-07-22: logic tu dong tim hoa don qua Gmail cho 1 dong Chi Phi
// dang thieu (khong co Link hoa don). Cach lam mo phong dung y het cach
// Claude da tra cuu bang tay trong phien lam viec 2026-07-22 (tim theo ten
// NCC, doc noi dung email, khop theo SO TIEN chinh xac -- vi cac mau email
// hoa don dien tu (Coca-Cola, Estella/Keppelland, eHoadon.vn/Bkav, MISA
// meInvoice...) deu ghi so tien kem dau phay ngan cach hang nghin, vd
// "15,180,000" -- nen cach kiem tra don gian va chac chan nhat la kiem tra
// chuoi so tien (dinh dang kieu Viet) co xuat hien nguyen van trong noi dung
// email hay khong, thay vi co doan regex rieng cho tung nha cung cap.
//
// LUU Y: cac email "Debit note"/"Account Statement" (Lotte, Aeon, Vincom...)
// gop nhieu khoan vao 1 file, KHONG co so tien tung dong rieng trong noi
// dung email (so tien nam trong file PDF dinh kem ma web khong doc duoc) --
// voi cac truong hop nay se KHONG khop duoc theo so tien, chi co the ganh
// link email chung chung (khong lam tu dong, de nguoi dung tu gan tay nhu
// truoc gio, tranh gan sai).

const gmailApi = require("./gmailApi");

function formatVnAmount(n) {
  const num = Math.round(Number(n) || 0);
  return num.toLocaleString("en-US"); // "15,180,000" -- dung dau phay giong cac email hoa don dien tu
}

function stripHtml(html) {
  return (html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/\s+/g, " ")
    .trim();
}

const INVOICE_NO_PATTERNS = [
  /Số\s*hóa\s*đơn\s*:?\s*(\d{3,10})/i,
  /Invoice\s*No\.?\s*:?\s*(\d{3,10})/i,
  /\bSố\s*:\s*(\d{3,10})\b/i,
];

function extractInvoiceNo(text) {
  for (const re of INVOICE_NO_PATTERNS) {
    const m = text.match(re);
    if (m) return m[1];
  }
  return "";
}

function gmailPermalink(messageId) {
  return "https://mail.google.com/mail/u/0/#all/" + messageId;
}

// Tim + kop hoa don cho 1 dong Chi Phi (r: {id, ncc, soTien, ngay}).
// Tra ve { soHoaDon, linkHoaDon } hoac null neu khong tim thay gi phu hop.
// `store` phai la object da load() -- ham nay co the sua store.gmail_oauth
// (lam moi access token), KHONG tu goi save(); caller tu save() sau.
async function findInvoiceForRow(store, r) {
  const ncc = (r.ncc || "").trim();
  if (!ncc) return null;
  const ngay = r.ngay || "";
  // Chi tim trong khoang +-45 ngay quanh ngay ghi chi, vi hoa don thuong den
  // truoc/sau ngay ghi so vai tuan (dua theo cac vi du da thay: lech ~1-10
  // ngay).
  const base = ngay ? new Date(ngay) : new Date();
  const after = new Date(base.getTime() - 45 * 86400000);
  const before = new Date(base.getTime() + 20 * 86400000);
  const fmt = (d) => d.toISOString().slice(0, 10).replace(/-/g, "/");
  const query = `"${ncc}" after:${fmt(after)} before:${fmt(before)}`;

  let messages;
  try {
    messages = await gmailApi.searchMessages(store, query, 15);
  } catch (e) {
    throw e;
  }
  if (!messages.length) return null;

  const targetAmount = formatVnAmount(r.soTien);
  for (const m of messages) {
    let msg;
    try {
      msg = await gmailApi.getMessage(store, m.id);
    } catch (e) {
      continue;
    }
    const text = stripHtml(msg.html) || msg.plain || "";
    if (targetAmount.length >= 4 && text.includes(targetAmount)) {
      const invoiceNo = extractInvoiceNo(text);
      return { soHoaDon: invoiceNo, linkHoaDon: gmailPermalink(msg.id) };
    }
  }
  return null;
}

module.exports = { findInvoiceForRow, formatVnAmount, stripHtml, extractInvoiceNo, gmailPermalink };
