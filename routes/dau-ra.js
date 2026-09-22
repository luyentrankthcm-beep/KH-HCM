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
  tutu_estella: { sheetId:'1cegFodLAXbtYdITdfGGd8m0weoOUVP6RgwGPScb9JMo',  tab:'BÁO CÁO', dateCol:0, tmCol:34, ckCol:35, multiplier:1000 },
  tutu_aeontan: { sheetId:'1SGtQ0Kvnvidr4ipxSP-AY5uwGwFEEoOrwFhawJP2scI',  tab:'VÉ',      dateCol:0, tmCol:34, ckCol:35 },
  fz_lottebt:   { sheetId:'1Be1E0pBJlYATKpogjqMTyHbhMcOg3Rbca7qoUlvHBWk',  tab:'TỔNG',    dateCol:0, tmCol:58, ckCol:47, ckCol2:52 }, // BG=TM(nhà bóng+bếp), AV+BA=Momo cả 2
  fz_scvivo:    { sheetId:'1lQMEpf1OhVOEROY5kzM_cWp77fZgUl92cVn0A78SGvw',  tab:'TỔNG',    dateCol:0, tmCol:58, tmCol2:63, ckCol:57, ckCol2:61 }, // BG+BL=TM(nhà bóng+bếp), BF+BJ=Momo cả 2
  fz_aeontan:   { sheetId:'1Ej4iwtbLGu-WgHkjTjdIDZPz02lL4KLRE9ANYzpMP9c',  tab:'VÉ',                  dateCol:0, tmCol:9,  ckCol:8  }, // VR Tân An: J=TM, I=CK
  ev_ghostbr:   { sheetId:'1JwtV9Mg-LS_3xIuuzSuc0aE4x23riepFiomt7q0HAPU',  tab:'BÁO CÁO TỔNG THÁNG {M}',    dateCol:0, tmCol:44, ckCol:45, ckCol2:46,
                  // gviz không tìm được tab bằng tên tiếng Việt phức tạp → dùng GID theo từng tháng
                  gidMap: { '7': 729775129, '8': 276666945, '9': 1907810456 } }, // Ghost Bà Rịa: AS=TM, AT+AU=CK
  ev_fzgoan:    { sheetId:'1bdU9XExnBp8aMLTj9rAAXS1G3_Xi6m2jI3y0SqQN80s',  tab:'Bảng báo cáo hàng ngày',     dateCol:0, tmCol:38, ckCol:37 }, // ADV Go An Lạc: AM=TM, AL=Momo
  fz_scvivo2:   { sheetId:'1ZaNpXiOrtpnHHZXdMZEFmBKw1eooAPmXqqls0N0Uap0',  tab:'VÉ',                  dateCol:0, tmCol:9,  ckCol:8  }, // VR SC Vivo: J=TM, I=CK
  pinball_amtp: { sheetId:'1r-huBGhhy_K_nOyHr0iz4Rd3Bkfw0vyTDQEZbT_PrOo',  tab:'BC tiền tháng {M}',  dateCol:0, tmCol:4,  ckCol:5  }, // Pinball: E=TM, F=CK, tab thay đổi theo tháng
};

function colIdxToLetter(n) {
  let s = ''; n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

// ── Helper: parse CSV đơn giản (không cần thư viện ngoài) ────────────────────
// Xử lý dấu nháy kép, dấu phẩy, xuống dòng trong field
function parseCSVtoRows(csvText) {
  const rows = [];
  let row = [], cur = '', inQ = false;
  for (let i = 0; i < csvText.length; i++) {
    const ch = csvText[i];
    if (inQ) {
      if (ch === '"' && csvText[i+1] === '"') { cur += '"'; i++; }
      else if (ch === '"') { inQ = false; }
      else { cur += ch; }
    } else {
      if (ch === '"') { inQ = true; }
      else if (ch === ',') { row.push(cur); cur = ''; }
      else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else if (ch !== '\r') { cur += ch; }
    }
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

// ── Helper: fetch CSV từ Google Sheets qua gviz/tq ───────────────────────────
async function _fetchGvizCSV(url, accessToken) {
  const resp = await fetch(url, {
    headers: { Authorization: 'Bearer ' + accessToken },
    redirect: 'follow',
  });
  if (!resp.ok) throw new Error('gviz/tq lỗi HTTP ' + resp.status);
  const text = await resp.text();
  if (text.trimStart().startsWith('<')) {
    throw new Error('Không đủ quyền đọc sheet hoặc tên tab sai. Vào Chi Phí → Kết nối Gmail để xác thực lại.');
  }
  return text;
}

// Fetch bằng tên tab (có thể thất bại nếu tên tab có ký tự đặc biệt)
async function fetchSheetCSVbyName(sheetId, tabName, accessToken) {
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(tabName)}&headers=0`;
  return _fetchGvizCSV(url, accessToken);
}

// Fetch bằng GID (đáng tin cậy hơn cho tab có tên tiếng Việt phức tạp)
async function fetchSheetCSVbyGid(sheetId, gid, accessToken) {
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}&headers=0`;
  return _fetchGvizCSV(url, accessToken);
}

// ── API: Quản lý gian tùy chỉnh ──────────────────────────────────────────────
router.get('/api/dau-ra/gian-custom', requireLogin, (req, res) => {
  const store = load();
  res.json({ ok: true, gians: store.kvc_custom_gians || [] });
});

router.post('/api/dau-ra/gian-custom', requireLogin, express.json(), (req, res) => {
  const { id, kh, maCT, khFull, link, sheetId, tab, tmCol, ckCol } = req.body;
  if (!id || !maCT || !sheetId || !tab) return res.json({ ok: false, error: 'Thiếu thông tin bắt buộc (id, maCT, sheetId, tab)' });
  const store = load();
  if (!store.kvc_custom_gians) store.kvc_custom_gians = [];
  const allIds = [...Object.keys(GIAN_SHEETS_CONFIG), ...store.kvc_custom_gians.map(g => g.id)];
  if (allIds.includes(id)) return res.json({ ok: false, error: 'ID gian đã tồn tại: ' + id });
  store.kvc_custom_gians.push({ id, kh: kh || 'mới', maCT, khFull: khFull || '', link: link || '', sheetId, tab, dateCol: 0, tmCol: parseInt(tmCol) || 34, ckCol: parseInt(ckCol) || 35, addedAt: new Date().toISOString() });
  save(store);
  res.json({ ok: true });
});

router.delete('/api/dau-ra/gian-custom/:id', requireLogin, (req, res) => {
  const store = load();
  if (!store.kvc_custom_gians) return res.json({ ok: false, error: 'Không có gian tùy chỉnh' });
  const before = store.kvc_custom_gians.length;
  store.kvc_custom_gians = store.kvc_custom_gians.filter(g => g.id !== req.params.id);
  if (store.kvc_custom_gians.length === before) return res.json({ ok: false, error: 'Không tìm thấy gian: ' + req.params.id });
  save(store);
  res.json({ ok: true });
});

// Route: lấy danh sách tab + GID của một spreadsheet (để debug/cấu hình)
router.get('/api/dau-ra/sheet-info', requireLogin, async (req, res) => {
  const { sheetId } = req.query;
  if (!sheetId) return res.json({ ok:false, error:'Cần tham số sheetId' });
  const store = load();
  if (!store.gmail_oauth?.refresh_token) return res.json({ ok:false, error:'Chưa xác thực Google' });
  try {
    const accessToken = await gmailApi.getValidAccessToken(store);
    save(store);
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties(title,sheetId)`;
    const resp = await fetch(url, { headers:{ Authorization:'Bearer ' + accessToken } });
    const data = await resp.json();
    if (!resp.ok) return res.json({ ok:false, error: data.error?.message || 'Lỗi không xác định' });
    res.json({ ok:true, sheets: (data.sheets || []).map(s => ({ title: s.properties.title, gid: s.properties.sheetId })) });
  } catch(e) {
    res.json({ ok:false, error: e.message });
  }
});

// Route: fetch dữ liệu TM/CK từng ngày từ Google Sheets
router.get('/api/dau-ra/sheets-daily', requireLogin, async (req, res) => {
  const { gianId, month, year } = req.query;
  // Tìm trong hardcoded config trước, rồi trong custom gians
  let cfg = GIAN_SHEETS_CONFIG[gianId];
  if (!cfg) {
    const storeCheck = load();
    const customGian = (storeCheck.kvc_custom_gians || []).find(g => g.id === gianId);
    if (customGian) cfg = { sheetId: customGian.sheetId, tab: customGian.tab, dateCol: customGian.dateCol || 0, tmCol: customGian.tmCol, ckCol: customGian.ckCol };
  }
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
    // Hỗ trợ tab name động: {M} = tháng không có số 0, {MM} = tháng có số 0
    const tabName = cfg.tab.replace('{MM}', mm).replace('{M}', String(+month));
    const maxColLetter = colIdxToLetter(Math.max(cfg.tmCol, cfg.tmCol2 ?? 0, cfg.ckCol, cfg.ckCol2 ?? 0));

    let rows = null;
    let fetchMethod = 'sheets-api';

    // ── Cách 1: Sheets API v4 (ưu tiên) ──────────────────────────────────────
    try {
      // Dùng dấu nháy đơn quanh tab name để Sheets API xử lý đúng khi tab có dấu cách/tiếng Việt
      const range = `'${tabName}'!A:${maxColLetter}`;
      const url   = `https://sheets.googleapis.com/v4/spreadsheets/${cfg.sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
      const resp  = await fetch(url, { headers:{ Authorization:'Bearer ' + accessToken } });
      if (resp.ok) {
        const data = await resp.json();
        rows = data.values || [];
      } else {
        const err = await resp.json().catch(() => ({}));
        const msg = err.error?.message || '';
        // Nếu lỗi KHÔNG phải "API not enabled" → trả lỗi ngay
        if (!/disabled|not been used/i.test(msg)) {
          return res.json({ ok:false, error:'Google Sheets lỗi: ' + msg });
        }
        // else: Sheets API chưa bật → thử fallback bên dưới
      }
    } catch (_) { /* mạng lỗi → thử fallback */ }

    // ── Cách 2: Google Visualization API (gviz/tq) ─────────────────────────────
    // Không cần Sheets API. Dùng GID (nếu có) hoặc tên tab. Scope drive.readonly là đủ.
    // Fallback khi: Sheets API bị chặn (!rows) hoặc trả về rỗng (rows.length < 2)
    if (!rows || rows.length < 2) {
      fetchMethod = 'gviz-tq';
      try {
        let csvText;
        // Ưu tiên: gidMap[month] > cfg.gid > tên tab
        // gidMap dùng cho sheet có tab đặt tên theo tháng (gviz fail khi tên có tiếng Việt phức tạp)
        const monthGid = cfg.gidMap ? cfg.gidMap[String(+month)] : undefined;
        const effectiveGid = monthGid !== undefined ? monthGid : cfg.gid;
        if (effectiveGid !== undefined) {
          csvText = await fetchSheetCSVbyGid(cfg.sheetId, effectiveGid, accessToken);
        } else {
          csvText = await fetchSheetCSVbyName(cfg.sheetId, tabName, accessToken);
        }
        rows = parseCSVtoRows(csvText);
      } catch (e) {
        return res.json({ ok:false, error:'Lấy dữ liệu Google Sheets thất bại: ' + e.message });
      }
    }

    // Lọc hàng theo tháng/năm — hỗ trợ nhiều format ngày khác nhau:
    //   "01/09/2026" (chuẩn)  |  "1/9/2026" (không padding)  |  "1/9" (không năm, chỉ ngày/tháng)
    const mNum = String(+month);  // tháng không padding, vd "9"
    const matchDate = dateCell => {
      if (!dateCell) return false;
      const s = dateCell.trim();
      if (s.includes(`/${mm}/${yyyy}`)) return true;   // 01/09/2026
      if (s.includes(`/${mNum}/${yyyy}`)) return true; // 1/9/2026
      // Format không có năm: "d/m" hoặc "dd/m" — chỉ khớp nếu đúng tháng
      if (/^\d{1,2}\/\d{1,2}$/.test(s)) {
        return s.split('/')[1] === mNum || s.split('/')[1] === mm;
      }
      return false;
    };
    // parseNum: xử lý số tiếng Việt: "4.560.000 đ", "102.550.000", "1,5"
    // Xóa mọi ký tự không phải số/dấu phẩy/dấu chấm/dấu trừ (bao gồm "đ", khoảng trắng)
    const parseNum = s => {
      if (s === undefined || s === null || s === '') return null;
      const cleaned = String(s).replace(/[^0-9,.\-]/g, '').replace(/\./g, '').replace(',', '.');
      const n = Number(cleaned);
      return isNaN(n) || n === 0 ? null : Math.round(n);
    };

    // Debug mode: trả về 5 rows đầu để kiểm tra column indices
    if (req.query.debug === '1') {
      const sample = rows.slice(0, 10).map((row, i) => ({
        rowIdx: i,
        dateCol: row[cfg.dateCol],
        tmCol: row[cfg.tmCol],
        ckCol: row[cfg.ckCol],
        cols: row.slice(Math.max(0, cfg.tmCol-2), cfg.tmCol+5),
      }));
      return res.json({ ok:true, tab:tabName, sample, totalRows:rows.length });
    }

    const daily = [];
    rows.forEach(row => {
      const dateCell = String(row[cfg.dateCol] || '');
      if (!matchDate(dateCell)) return;
      const day = dateCell.trim().split('/')[0].replace(/\D/g,'');
      if (!day || isNaN(+day)) return;
      const ngay = day.padStart(2,'0') + '/' + mm;
      const mul = cfg.multiplier || 1;
      const tmRaw1 = parseNum(row[cfg.tmCol]);
      const tmRaw2 = cfg.tmCol2 !== undefined ? parseNum(row[cfg.tmCol2]) : null;
      const ckRaw1 = parseNum(row[cfg.ckCol]);
      const ckRaw2 = cfg.ckCol2 !== undefined ? parseNum(row[cfg.ckCol2]) : null;
      // Nếu có tmCol2: cộng TM cả 2 khu (ví dụ fz_scvivo: BG nhà bóng + BL nhà bếp)
      const tmCombined = (tmRaw1 !== null || tmRaw2 !== null)
        ? ((tmRaw1 || 0) + (tmRaw2 || 0)) : null;
      const tm = tmCombined !== null ? Math.round(tmCombined * mul) : null;
      // Nếu có ckCol2: cộng cả 2 cột Momo (ví dụ fz_lottebt: AV nhà bóng + BA nhà bếp)
      const ckCombined = (ckRaw1 !== null || ckRaw2 !== null)
        ? ((ckRaw1 || 0) + (ckRaw2 || 0)) : null;
      const ck = ckCombined !== null ? Math.round(ckCombined * mul) : null;
      if (tm !== null || ck !== null) {
        daily.push({ ngay, tienMat:tm, chuyenKhoan:ck, dtKhac:0 });
      }
    });

    res.json({ ok:true, daily, tab:tabName, tmCol:colIdxToLetter(cfg.tmCol), ckCol:colIdxToLetter(cfg.ckCol), method:fetchMethod });
  } catch(err) {
    res.json({ ok:false, error:err.message });
  }
});

router.use(requireLogin);

// Nhan, 2026-09-22: "bạn ơi cho tôi mã công trình đầy đủ hk á" -- bang Doi
// soat ngay (Zalo/VNPay/Payoo) chi hien ma CT NGAN (AEON, FARM, FUNZONE...)
// suy tu tu khoa ten san pham, KHONG phai Ma CT day du (vd co ~23 chi nhanh
// AEON khac nhau trong danh muc nen khong the tu dong doan dung chi nhanh
// nao -- de trong/ngan con hon doan sai). Giai phap: cho DANH SACH GOI Y Ma
// CT day du (giong allGianSuggestions ben Chi Phi, lay tu
// phap_danh_hop_dong_thue.maCongTrinh) de Nhan tu chon/go dung khi map Ten
// SP -> Ma CT (man "Map Tên SP → CT"), thay vi go tay tu do de sai chinh ta.
router.get("/dau-ra", (req, res) => {
  const store = load();
  const maCTSet = new Set();
  (store.phap_danh_hop_dong_thue || []).forEach((r) => { if (r.maCongTrinh) maCTSet.add(r.maCongTrinh); });
  const allMaCTSuggestions = [...maCTSet].sort();
  res.render("dau-ra", {
    userName: req.session.userName || req.session.user || "",
    allMaCTSuggestions,
  });
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
// ── Helper: build _zaloSpMap seed từ zvp_online_product_map (server-side) ──────
// Gọi khi kh_cu không có zalo_sp_map_kh_cu để tránh Railway mất data khi restart.
// CSE products: maCT = maCongTrinh + " - CSE", non-CSE: maCT = maCongTrinh.
function buildZaloSpMapSeed(store) {
  const zvpMap = store.zvp_online_product_map || {};
  const seed = {};
  Object.entries(zvpMap).forEach(([tenSP, info]) => {
    const maCT = (info.maCongTrinh || '').trim();
    if (maCT) seed[tenSP] = info.isCse ? maCT + ' - CSE' : maCT;
  });
  return seed;
}

router.get("/api/dau-ra/kv-all", (req, res) => {
  try {
    const store = load();
    const company = getCompany(req);
    if (!store.dau_ra_kv) store.dau_ra_kv = { kh_cu: {}, kh_moi: {} };
    if (!store.dau_ra_kv[company]) store.dau_ra_kv[company] = {};
    const kvCo = store.dau_ra_kv[company];
    // Seed zalo_sp_map nếu chưa có (Railway restart hoặc cài mới)
    // Chỉ seed khi key hoàn toàn không tồn tại (không ghi đè data user đã lưu)
    const spMapKey = `zalo_sp_map_${company}`;
    if (!kvCo[spMapKey]) {
      const seed = buildZaloSpMapSeed(store);
      if (Object.keys(seed).length > 0) {
        kvCo[spMapKey] = JSON.stringify(seed);
        save(store); // lưu seed xuống file để các request sau không cần seed lại
      }
    }
    res.json({ ok: true, data: kvCo });
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
// chore: trigger deploy 1789978645
