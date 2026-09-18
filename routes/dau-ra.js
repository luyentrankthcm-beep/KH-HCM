const express = require("express");
const router = express.Router();
const { requireLogin } = require("../middleware/auth");
const { load, save } = require("../store");
const { getCompany } = require("../utils/companies");
const gmailApi = require("../utils/gmailApi");

// ── Cấu hình từng gian: spreadsheetId, tab, vị trí cột (0-based) ──────────
// dateCol = cột ngày (A=0), tmCol = Tiền Mặt, ckCol = Chuyển Khoản/Momo
// Eco Farm: tab TỔNG, TM=BA(52), CK=AZ(51/Momo)
// TUTU/FZ/EV phần lớn: tab VÉ, TM=AI(34), CK=AJ(35)
// Tân Phú: tab VÉ, nhiều cột SP hơn → TM=AN(39), CK=AO(40)
// Estella: tab "BÁO CÁO", TM=AI(34), CK=AJ(35)
const GIAN_SHEETS_CONFIG = {
  farm_lottebt: { sheetId:'1fiFGmerqMGjKW8oxJBfN_KXVn7c7rbIOszzuZIavHLc', tab:'TỔNG',    dateCol:0, tmCol:52, ckCol:51 },
  tutu_lottegv: { sheetId:'1F3ATisma35uHIdhlKAckd16GlmTGMvRE_QwH4N16H0M',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  tutu_aeontp:  { sheetId:'1jGG9Po3WnsntEeXC53uHUUxi-5bcKx-jPuPWhN68Qps',  tab:'VÉ',      dateCol:0, tmCol:39, ckCol:40 },
  tutu_aeonbt:  { sheetId:'1Lts3GeoeozWrPlprGzANK_ksmd6h4XMnl5xMUMM_OL4',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  tutu_aeonbd:  { sheetId:'15brZmUqEYuuEkbp7AgQSAEKmbgMiNqey6CIScGO97cs',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  tutu_estella: { sheetId:'1cegFodLAXbtYdITdfGGd8m0weoOUVP6RgwGPScb9JMo',  tab:'BÁO CÁO', dateCol:0, tmCol:34, ckCol:35 },
  tutu_aeontan: { sheetId:'1SGtQ0Kvnvidr4ipxSP-AY5uwGwFEEoOrwFhawJP2scI',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  fz_lottebt:   { sheetId:'1Be1E0pBJlYATKpogjqMTyHbhMcOg3Rbca7qoUlvHBWk',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  fz_scvivo:    { sheetId:'1lQMEpf1OhVOEROY5kzM_cWp77fZgUl92cVn0A78SGvw',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  fz_aeontan:   { sheetId:'1Ej4iwtbLGu-WgHkjTjdIDZPz02lL4KLRE9ANYzpMP9c',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  ev_ghostbr:   { sheetId:'1JwtV9Mg-LS_3xIuuzSuc0aE4x23riepFiomt7q0HAPU',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  ev_fzgoan:    { sheetId:'1bdU9XExnBp8aMLTj9rAAXS1G3_Xi6m2jI3y0SqQN80s',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  fz_scvivo2:   { sheetId:'1ZaNpXiOrtpnHHZXdMZEFmBKw1eooAPmXqqls0N0Uap0',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  pinball_amtp: { sheetId:'1r-huBGhhy_K_nOyHr0iz4Rd3Bkfw0vyTDQEZbT_PrOo',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
};

function colIdxToLetter(n) {
  let s = ''; n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

// Route: fetch dữ liệu TM/CK từng ngày từ Google Sheets
router.get('/api/dau-ra/sheets-daily', requireLogin, async (req, res) => {
  const { gianId, month, year } = req.query;
  const cfg = GIAN_SHEETS_CONFIG[gianId];
  if (!cfg) return res.json({ ok:false, error:'Chưa cấu hình cột cho gian: ' + gianId });

  const store = load();
  if (!store.gmail_oauth || !store.gmail_oauth.refresh_token) {
    return res.json({ ok:false, needAuth:true, error:'Chưa kết nối Google. Vào trang Chi Phí → bấm "Kết nối Gmail" để xác thực.' });
  }

  try {
    const accessToken = await gmailApi.getValidAccessToken(store);
    save(store);

    const mm   = String(month).padStart(2,'0');
    const yyyy = String(year);
    const maxColLetter = colIdxToLetter(Math.max(cfg.tmCol, cfg.ckCol));
    const range = `${cfg.tab}!A:${maxColLetter}`;
    const url   = `https://sheets.googleapis.com/v4/spreadsheets/${cfg.sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;

    const resp = await fetch(url, { headers:{ Authorization:'Bearer ' + accessToken } });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({}));
      const msg = err.error?.message || resp.statusText;
      return res.json({ ok:false, error:'Google Sheets lỗi: ' + msg });
    }

    const data = await resp.json();
    const rows = data.values || [];

    // Lọc hàng theo tháng/năm: ô dateCol phải chứa "/MM/YYYY" (dd/mm/yyyy)
    const suffix = `/${mm}/${yyyy}`;
    const parseNum = s => {
      if (s === undefined || s === null || s === '') return null;
      const n = Number(String(s).replace(/[.\s]/g,'').replace(',','.'));
      return isNaN(n) || n === 0 ? null : Math.round(n);
    };

    const daily = [];
    rows.forEach(row => {
      const dateCell = String(row[cfg.dateCol] || '');
      if (!dateCell.includes(suffix)) return;
      const day = dateCell.split('/')[0].trim().replace(/\D/g,'');
      if (!day || isNaN(+day)) return;
      const ngay = day.padStart(2,'0') + '/' + mm;
      const tm = parseNum(row[cfg.tmCol]);
      const ck = parseNum(row[cfg.ckCol]);
      if (tm !== null || ck !== null) {
        daily.push({ ngay, tienMat:tm, chuyenKhoan:ck, dtKhac:0 });
      }
    });

    res.json({ ok:true, daily, tab:cfg.tab, tmCol:colIdxToLetter(cfg.tmCol), ckCol:colIdxToLetter(cfg.ckCol) });
  } catch(err) {
    res.json({ ok:false, error:err.message });
  }
});

router.use(requireLogin);

router.get("/dau-ra", (req, res) => {
  res.render("dau-ra", { userName: req.session.userName || req.session.user || "" });
});

router.get("/dau-ra-noi-bo", (req, res) => {
  res.render("dau-ra-noi-bo", { userName: req.session.userName || req.session.user || "" });
});

// ── KV store cho trang "Đầu Ra" ──────────────────────────────────────────
// Chi Nhan, 2026-09-10: trang nay truoc gio luu Momo/MTT/Zalo rows... bang
// localStorage CUA TRINH DUYET -- Luyen bao mo web o 2 may/tai khoan khac
// nhau thay 1 cai co du lieu 1 cai khong, vi localStorage khong dong bo qua
// server. 3 route duoi day luu/doc cung 1 du lieu do TREN SERVER (theo tung
// cong ty kh_cu/kh_moi, xem store.dau_ra_kv trong store.js) de mo may nao
// / trinh duyet nao cung thay GIONG NHAU. Client (dau-ra.ejs) van giu
// localStorage nhu cu de doc/hien nhanh (khong doi logic parse/render), chi
// them buoc: (1) luc tai trang, keo du lieu server ve ghi de vao localStorage
// TRUOC khi cac ham cu doc localStorage nhu binh thuong; (2) moi lan cac ham
// cu ghi vao localStorage thi ALSO gui 1 ban len server qua route POST /kv.
router.get("/api/dau-ra/kv-all", (req, res) => {
  try {
    const store = load();
    const company = getCompany(req);
    if (!store.dau_ra_kv) store.dau_ra_kv = { kh_cu: {}, kh_moi: {} };
    if (!store.dau_ra_kv[company]) store.dau_ra_kv[company] = {};
    res.json({ ok: true, data: store.dau_ra_kv[company] });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

router.post("/api/dau-ra/kv", express.json({ limit: "30mb" }), (req, res) => {
  try {
    const { key, value } = req.body || {};
    if (!key) return res.json({ ok: false, error: "Thiếu key" });
    const store = load();
    const company = getCompany(req);
    if (!store.dau_ra_kv) store.dau_ra_kv = { kh_cu: {}, kh_moi: {} };
    if (!store.dau_ra_kv[company]) store.dau_ra_kv[company] = {};
    store.dau_ra_kv[company][key] = value;
    save(store);
    res.json({ ok: true });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

router.post("/api/dau-ra/kv-delete", express.json({ limit: "1mb" }), (req, res) => {
  try {
    const { key } = req.body || {};
    if (!key) return res.json({ ok: false, error: "Thiếu key" });
    const store = load();
    const company = getCompany(req);
    if (store.dau_ra_kv && store.dau_ra_kv[company]) delete store.dau_ra_kv[company][key];
    save(store);
    res.json({ ok: true });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

// API: tổng thu ngân hàng theo ngày cho 1 tài khoản
// GET /api/dau-ra/bank-daily?account=7701&from=2026-09-01&to=2026-09-30
// Trả về { "02-09-2026": 47586000, ... } — date theo DD-MM-YYYY để match client
router.get("/api/dau-ra/bank-daily", (req, res) => {
  try {
    const { account, from, to } = req.query;
    if (!account) return res.json({ ok: false, error: "Thiếu account" });

    const store   = load();
    const company = getCompany(req);

    // Tìm bank theo phần cuối số tài khoản (vd "7701")
    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith(account)
    );
    if (!bank) return res.json({ ok: false, error: `Không tìm thấy tài khoản *${account} trong ${company}` });

    // Lọc giao dịch "thu" của bank đó trong khoảng ngày
    const txns = store.transactions.filter(t => {
      if (t.bank_id !== bank.id) return false;
      if (t.type !== "thu") return false;
      if (from && t.date < from) return false;
      if (to   && t.date > to)   return false;
      return true;
    });

    // Gom theo ngày: store dùng YYYY-MM-DD → convert sang DD-MM-YYYY cho client
    const daily = {};
    for (const t of txns) {
      // t.date có thể là "2026-09-02" hoặc "02-09-2026" tuỳ cách nhập
      let ddmmyyyy = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const [y, m, d] = t.date.split("-");
        ddmmyyyy = `${d}-${m}-${y}`;
      }
      daily[ddmmyyyy] = (daily[ddmmyyyy] || 0) + Number(t.amount || 0);
    }

    res.json({ ok: true, bankName: bank.name, bankId: bank.id, daily });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});

// API: lấy các đợt Momo trả về ngân hàng (parse diễn giải "tu DD/MM/YYYY den DD/MM/YYYY")
// GET /api/dau-ra/bank-momo?account=7701&from=2026-08-01&to=2026-09-30
// Trả về { ok, payments: [{ payDate:"DD-MM-YYYY", amount, fromDate:"DD-MM-YYYY", toDate:"DD-MM-YYYY" }] }
router.get("/api/dau-ra/bank-momo", (req, res) => {
  try {
    const { account, from, to } = req.query;
    if (!account) return res.json({ ok: false, error: "Thiếu account" });

    const store   = load();
    const company = getCompany(req);

    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith(account)
    );
    if (!bank) return res.json({ ok: false, error: `Không tìm thấy TK *${account} trong ${company}` });

    const pattern = /tu (\d{2})\/(\d{2})\/(\d{4}) den (\d{2})\/(\d{2})\/(\d{4})/i;
    const payments = [];

    for (const t of store.transactions) {
      if (t.bank_id !== bank.id) continue;
      if (t.type !== "thu") continue;
      // Momo settlement descriptions: "DI DONG TRUC TUYEN" hoặc "MoMo" trong diễn giải
      if (!/DI DONG TRUC TUYEN|MOMO|MoMo/i.test(t.description || "")) continue;
      const m = pattern.exec(t.description || "");
      if (!m) continue;

      const [, d1, mo1, y1, d2, mo2, y2] = m;

      // Lọc theo khoảng ngày thanh toán (payDate = t.date, format YYYY-MM-DD)
      if (from && t.date < from) continue;
      if (to   && t.date > to)   continue;

      // Convert t.date (YYYY-MM-DD) sang DD-MM-YYYY cho client
      let payDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const [y, m2, d] = t.date.split("-");
        payDate = `${d}-${m2}-${y}`;
      }

      payments.push({
        payDate,
        amount:   Number(t.amount || 0),
        fromDate: `${d1}-${mo1}-${y1}`,
        toDate:   `${d2}-${mo2}-${y2}`,
        desc:     (t.description || "").substring(0, 80),
      });
    }

    // Sắp xếp theo ngày thanh toán tăng dần
    payments.sort((a, b) => {
      const toISO = d => { const [dd, mm, yy] = d.split("-"); return `${yy}-${mm}-${dd}`; };
      return toISO(a.payDate).localeCompare(toISO(b.payDate));
    });

    res.json({ ok: true, payments });
  } catch (err) {
    res.json({ ok: false, error: err.message });
  }
});


// GET /api/dau-ra/bank-zalo?from=2026-08-15&to=2026-09-08
// Trả về các đợt ZaloPay trả về ACB1268 (diễn giải VNPAY TT 829168, không OFFLINE)
router.get("/api/dau-ra/bank-zalo", (req, res) => {
  try {
    const { from, to } = req.query;
    const store   = load();
    const company = getCompany(req);
    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith("1268")
    );
    if (!bank) return res.json({ ok: false, error: "Không tìm thấy TK *1268" });

    const payments = [];
    for (const t of store.transactions) {
      if (t.bank_id !== bank.id) continue;
      if (t.type !== "thu") continue;
      // Chỉ lấy đúng mẫu ZaloPay mini app: "VNPAY TT 829168 GIAIT...989 DV CTT NGAY"
      if (!/VNPAY\s+TT\s+829168.*GIAIT.*989.*DV\s+CTT\s+NGAY/i.test(t.description || "")) continue;
      if (/OFFLINE/i.test(t.description || "")) continue;
      if (from && t.date < from) continue;
      if (to   && t.date > to)   continue;

      // Parse "NGAY 28.08-02.09.26" (khac thang), "NGAY 14-16.08.26" (cung
      // thang), hoặc "NGAY 03.09.26" (1 ngày). Chi Nhan, 2026-09-10: Luyen
      // bao dot NH tra 03-09-2026 (83.671.918đ, dien giai "...NGAY
      // 28.08-02.09.26") hien "ZaloPay từ→đến: —" vi regex cu chi ho tro
      // dang "D1(-D2)?.MM.YY" (mot thang duy nhat cho ca 2 dau) -- khoang
      // ngay bang qua 2 thang khac nhau (28/8 -> 2/9) khong khop duoc.
      // Them 1 dang regex rieng (mCross) thu truoc, uu tien hon 2 dang cu.
      let fromDate = null, toDate = null;
      const desc = t.description || "";
      const mCross  = /NGAY\s+(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSame   = !mCross && /NGAY\s+(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSingle = !mCross && !mSame && /NGAY\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      if (mCross) {
        const [, d1, mo1, d2, mo2, y] = mCross;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo1.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo2.padStart(2, "0") + "-" + yr;
      } else if (mSame) {
        const [, d1, d2, mo, y] = mSame;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
      } else if (mSingle) {
        const [, d, mo, y] = mSingle;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = fromDate;
      }
      let payDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const p = t.date.split("-");
        payDate = p[2] + "-" + p[1] + "-" + p[0];
      }
      payments.push({ payDate, amount: Number(t.amount || 0), fromDate, toDate, desc: (t.description||"").substring(0,100) });
    }
    payments.sort((a, b) => {
      const iso = d => { if (!d) return ""; const p = d.split("-"); return p[2]+"-"+p[1]+"-"+p[0]; };
      return iso(a.payDate).localeCompare(iso(b.payDate));
    });
    res.json({ ok: true, payments });
  } catch(err) { res.json({ ok: false, error: err.message }); }
});

// GET /api/dau-ra/bank-vnpay?from=2026-08-15&to=2026-09-08
// Trả về các đợt VNPay (QR Offline, khác Zalo Mini App) trả về ACB1268 --
// cùng tài khoản, cùng tiền tố diễn giải với Zalo Pay nhưng có chữ "OFFLINE"
// (vd "...VNPAY TT 829168 GIAITTRIKH989 DV QR OFFLINE NGAY 14-16.08.26...")
router.get("/api/dau-ra/bank-vnpay", (req, res) => {
  try {
    const { from, to } = req.query;
    const store   = load();
    const company = getCompany(req);
    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith("1268")
    );
    if (!bank) return res.json({ ok: false, error: "Không tìm thấy TK *1268" });

    const payments = [];
    for (const t of store.transactions) {
      if (t.bank_id !== bank.id) continue;
      if (t.type !== "thu") continue;
      // Chỉ lấy đúng mẫu VNPay QR Offline: "VNPAY TT 829168 GIAIT...989 DV QR OFFLINE NGAY"
      if (!/VNPAY\s+TT\s+829168.*GIAIT.*989/i.test(t.description || "")) continue;
      if (!/OFFLINE/i.test(t.description || "")) continue;
      // Luyen, 2026-09-10: "mấy tỷ giữ vậy" -- tài khoản ACB1268 này còn nhận
      // NHIỀU giao dịch KHÁC của cùng đơn vị trung gian thanh toán "CT CP
      // GIAI PHAP THANH TOAN VIET NAM" (VNPay) dùng ĐÚNG mẫu diễn giải y hệt
      // "VNPAY TT 829168 GIAITRIKH989 DV QR OFFLINE NGAY..." (không phân
      // biệt được bằng text) nhưng KHÔNG liên quan tới 3-4 gian QR Offline
      // đang đối soát ở đây -- đã xác nhận bằng cách so khớp với tổng Net từ
      // BaoCaoPhi: vd 03/08 thực tế chỉ có 19.320.752đ (khớp tuyệt đối với
      // BaoCaoPhi) nhưng cùng ngày còn có 1 dòng khác 3.070.166.115đ (số dư
      // NH tăng thật, không phải lỗi đọc file) -- rõ ràng không phải doanh
      // thu của các gian QR Offline nhỏ này. Cả tháng 08/2026 tổng Net cao
      // nhất 1 ngày chỉ ~64 triệu, cả tháng cộng lại ~650 triệu -- nên loại
      // hẳn các dòng > 300 triệu vì chắc chắn không phải của các gian này.
      if (Number(t.amount || 0) > 300000000) continue;
      if (from && t.date < from) continue;
      if (to   && t.date > to)   continue;

      let fromDate = null, toDate = null;
      const desc = t.description || "";
      const mCross  = /NGAY\s+(\d{1,2})\.(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSame   = !mCross && /NGAY\s+(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSingle = !mCross && !mSame && /NGAY\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      if (mCross) {
        const [, d1, mo1, d2, mo2, y] = mCross;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo1.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo2.padStart(2, "0") + "-" + yr;
      } else if (mSame) {
        const [, d1, d2, mo, y] = mSame;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
      } else if (mSingle) {
        const [, d, mo, y] = mSingle;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = fromDate;
      }
      let payDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const p = t.date.split("-");
        payDate = p[2] + "-" + p[1] + "-" + p[0];
      }
      payments.push({ payDate, amount: Number(t.amount || 0), fromDate, toDate, desc: (t.description||"").substring(0,100) });
    }
    payments.sort((a, b) => {
      const iso = d => { if (!d) return ""; const p = d.split("-"); return p[2]+"-"+p[1]+"-"+p[0]; };
      return iso(a.payDate).localeCompare(iso(b.payDate));
    });
    res.json({ ok: true, payments });
  } catch(err) { res.json({ ok: false, error: err.message }); }
});

// GET /api/dau-ra/bank-payoo?from=2026-08-15&to=2026-09-08
// Trả về các đợt Payoo trả về ACB1268 -- diễn giải "CTY CPDV TRUC TUYEN CONG
// DONG VIET-PAYOO TT TD NGAY...". Luyen, 2026-09-10: mỗi ngày NH trả về ĐÚNG
// 2 giao dịch (1 QR + 1 Thẻ) -- có "QRCODE" trong diễn giải = QR, không có
// = Thẻ. Payoo dùng dấu "_" để nối khoảng ngày (khác Zalo/VNPay dùng "-"):
// "NGAY 28.08_02.09.2026" (khác tháng), "NGAY 04_06.09.2026" (cùng tháng),
// "NGAY 03.09.2026" (1 ngày).
router.get("/api/dau-ra/bank-payoo", (req, res) => {
  try {
    const { from, to } = req.query;
    const store   = load();
    const company = getCompany(req);
    const bank = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      String(b.account_number || b.accountNumber || "").endsWith("1268")
    );
    if (!bank) return res.json({ ok: false, error: "Không tìm thấy TK *1268" });

    const payments = [];
    for (const t of store.transactions) {
      if (t.bank_id !== bank.id) continue;
      if (t.type !== "thu") continue;
      const desc = t.description || "";
      if (!/PAYOO\s+TT\s+TD/i.test(desc)) continue;
      if (from && t.date < from) continue;
      if (to   && t.date > to)   continue;

      const loai = /QRCODE/i.test(desc) ? "QR" : "THE";

      let fromDate = null, toDate = null;
      // Chi Nhan, 2026-09-11: phat hien them 1 dang dien giai Payoo khac --
      // dung chu " DEN " thay vi dau "_" de noi khoang ngay (vd "TT TD NGAY
      // 14.08 DEN 16.08.2026", khac thang) -- khien ca 3 regex duoi deu
      // KHONG khop, fromDate/toDate = null, gay crash khi loc theo khoang
      // ngay (TypeError o ddmmToISO). Them regex rieng cho dang " DEN " nay.
      const mCrossDen = /NGAY\s+(\d{1,2})\.(\d{1,2})\s+DEN\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSameDen  = !mCrossDen && /NGAY\s+(\d{1,2})\s+DEN\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mCross  = !mCrossDen && !mSameDen && /NGAY\s+(\d{1,2})\.(\d{1,2})_(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSame   = !mCrossDen && !mSameDen && !mCross && /NGAY\s+(\d{1,2})_(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      const mSingle = !mCrossDen && !mSameDen && !mCross && !mSame && /NGAY\s+(\d{1,2})\.(\d{1,2})\.(\d{2,4})/i.exec(desc);
      if (mCrossDen) {
        const [, d1, mo1, d2, mo2, y] = mCrossDen;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo1.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo2.padStart(2, "0") + "-" + yr;
      } else if (mSameDen) {
        const [, d1, d2, mo, y] = mSameDen;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
      } else if (mCross) {
        const [, d1, mo1, d2, mo2, y] = mCross;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo1.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo2.padStart(2, "0") + "-" + yr;
      } else if (mSame) {
        const [, d1, d2, mo, y] = mSame;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d1.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = d2.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
      } else if (mSingle) {
        const [, d, mo, y] = mSingle;
        const yr = y.length === 2 ? "20" + y : y;
        fromDate = d.padStart(2, "0") + "-" + mo.padStart(2, "0") + "-" + yr;
        toDate   = fromDate;
      }
      let payDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const p = t.date.split("-");
        payDate = p[2] + "-" + p[1] + "-" + p[0];
      }
      payments.push({ payDate, amount: Number(t.amount || 0), fromDate, toDate, loai, desc: desc.substring(0,100) });
    }
    payments.sort((a, b) => {
      const iso = d => { if (!d) return ""; const p = d.split("-"); return p[2]+"-"+p[1]+"-"+p[0]; };
      return iso(a.payDate).localeCompare(iso(b.payDate));
    });
    res.json({ ok: true, payments });
  } catch(err) { res.json({ ok: false, error: err.message }); }
});

// GET /api/dau-ra/bank-vietqr?bank=bidv7704&from=2026-08-01&to=2026-08-31
// Luyen, 2026-09-11: "tiếp theo là viet qr kh cũ á... đối soát theo số mã
// tham chiếu đó với ngân hàng á trong sao có á bạn check nha lấy ngân hàng
// làm chuẩn á" -- trả về các giao dịch "thu" của 1 trong 3 TK VietQR
// (BIDV7704 / BIDV77020 / MB11521268) mà diễn giải có chứa mã VietQR nhúng
// sẵn (dạng "VQR" + 13 ký tự chữ/số, luôn theo sau bởi "PaymentForOrder") --
// đây là mã DUY NHẤT dùng chung giữa file giao dịch VietQR (cột "Nội dung
// TT") và sao kê ngân hàng (cột "Diễn giải") -- cột "Mã tham chiếu" của file
// giao dịch KHÔNG xuất hiện trong sao kê nên không dùng được để join trực
// tiếp. Ngân hàng được coi là chuẩn: client sẽ ưu tiên lấy amount ở đây thay
// vì amount tự khai trong file giao dịch khi có match theo mã VQR.
// Chi Nhan, 2026-09-12: Luyen bao "2 KH á là khác nhau hoàn toàn á 2 công ty
// khác nhau á ... đừng hiện các dữ liệu liên quan tới kh cũ nha" -- truoc gio
// route nay CHI biet 3 TK Viet QR cua KH Cu (bidv7704/bidv77020/mb11521268),
// nen khi dang xem KH Moi, tab Viet QR (dau-ra.ejs) van bi hien NHAM 3 tab
// nay. KH Moi co RIENG 4 TK Viet QR cua no (BIDV7702/BIDV77021/MB02865168/
// BIDV8613600999 -- xem BANK_COMPANY trong utils/bankCompany.js va CHANNELS
// trong routes/doisoat-vietqr.js, noi 7 kenh nay da duoc xac nhan tu truoc).
// Doi tu "doan so cuoi tai khoan" (VIETQR_BANK_SUFFIX cu) sang khop THANG
// theo TEN ngan hang that (giong y het cach routes/doisoat-vietqr.js dang
// lam, "b.name === bankName") -- an toan hon, khong lo do dai/so chu so tai
// khoan khac nhau giua cac TK gay nham lan.
const VIETQR_BANK_NAME = {
  bidv7704: "BIDV7704",
  bidv77020: "BIDV77020",
  mb11521268: "MB11521268",
  bidv7702: "BIDV7702",
  bidv77021: "BIDV77021",
  mb02865168: "MB02865168",
  bidv8613600999: "BIDV8613600999",
};
router.get("/api/dau-ra/bank-vietqr", (req, res) => {
  try {
    const { bank, from, to } = req.query;
    const bankName = VIETQR_BANK_NAME[bank];
    if (!bankName) return res.json({ ok: false, error: "Tham số bank không hợp lệ" });
    const store = load();
    const company = getCompany(req);
    const bankRow = store.banks.find(b =>
      (b.company || "kh_cu") === company &&
      b.name === bankName
    );
    if (!bankRow) return res.json({ ok: false, error: "Không tìm thấy tài khoản ngân hàng cho " + bank });

    // Chi Nhan, 2026-09-11: xac nhan tren du lieu that -- ma VQR cua BIDV
    // dai 13 ky tu (vd VQR26375334CLIBH) nhung cua MB11521268 chi dai 11 ky
    // tu (vd VQR26343E6DDF5) -- regex cung {13} lam mat het GD cua MB (0/0
    // khop). Doi sang do dai linh hoat 8-20 ky tu de an toan voi ca 2 dinh
    // dang, khop dung phan "VQR..." lien tuc truoc khoang trang.
    const vqrRe = /VQR[A-Za-z0-9]{8,20}/;
    // Luyen, 2026-09-11: "trong nội dung á VQR26375407F1W là mã đơn hàng á
    // nên bạn check theo mã đơn hàng" -- file giao dich xuat tu MB11521268
    // hau nhu KHONG nhung "Ma tham chieu" cua no lai TRUNG KHOP TUYET DOI voi
    // cot "BÚT TOÁN" tren sao ke that (vd "FT26254796100178") -- cot nay da
    // duoc utils/bankStatementParser.js doc san thanh t.reference (xem
    // REF_HEADER_PATTERNS_FALLBACK). Truoc gio route nay CHI tra ve dong co
    // ma VQR trong dien giai (bo qua het nhung dong khong co, vd "PaymentFor
    // Order" tron trui cua MB), khien phia client khong co gi de doi chieu
    // theo Ma tham chieu ca du du lieu van co san server-side. Gio: giu dong
    // neu co MA VQR trong dien giai HOAC co san t.reference (BÚT TOÁN/So
    // tham chieu) -- tra ca 2 truong cho client tu chon khoa khop phu hop.
    const transactions = [];
    for (const t of store.transactions) {
      if (t.bank_id !== bankRow.id) continue;
      if (t.type !== "thu") continue;
      if (from && t.date < from) continue;
      if (to && t.date > to) continue;
      const desc = t.description || "";
      const m = vqrRe.exec(desc);
      const reference = (t.reference || "").trim();
      if (!m && !reference) continue;
      let dispDate = t.date;
      if (/^\d{4}-\d{2}-\d{2}$/.test(t.date)) {
        const p = t.date.split("-");
        dispDate = p[2] + "-" + p[1] + "-" + p[0];
      }
      transactions.push({ date: dispDate, amount: Number(t.amount || 0), vqrCode: m ? m[0] : "", reference, desc: desc.substring(0, 150) });
    }
    transactions.sort((a, b) => {
      const iso = d => { if (!d) return ""; const p = d.split("-"); return p[2]+"-"+p[1]+"-"+p[0]; };
      return iso(a.date).localeCompare(iso(b.date));
    });
    // Chi Nhan, 2026-09-11: "cho tôi cái xuất misa nhá" -- can biet dung so
    // TK/ten ngan hang THAT cua tung tab (BIDV7704/BIDV77020/MB11521268) de
    // dien vao cot "Nộp vào TK"/"Mở tại ngân hàng" cua file xuat Misa, giong
    // cach lam voi Momo/ZVP -- lay thang tu bankRow da tim duoc o tren (dung
    // 1 lan tim, khong can them route rieng).
    res.json({
      ok: true,
      transactions,
      bankAccount: bankRow.account_number || bankRow.accountNumber || "",
      bankFullName: "Ngân hàng " + (bankRow.bank_name || bankRow.name || ""),
    });
  } catch (err) { res.json({ ok: false, error: err.message }); }
});

module.exports = router;
