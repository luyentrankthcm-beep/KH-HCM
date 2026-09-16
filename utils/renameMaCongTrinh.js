// Nhan, 2026-09-16: "các mã công trình bị thay đổi hết rồi á ... thay đổi
// toàn bộ mã của công trình cũ của từng kh khác nhau nhá rồi lưu cho tôi
// nhá" -- doi 1 ma cong trinh CU sang MOI cho 1 cong ty (kh_cu/kh_moi),
// CASCADE sang tat ca noi trong store.json co luu maCongTrinh lam DU LIEU
// (khong chi la key tra cuu tam thoi luc render).
//
// Da khao sat toan bo code (xem lich su chat) va liet ke ~20 vi tri persisted
// dung maCongTrinh. Ham nay xu ly TAT CA cac vi tri do noi viec doi la an
// toan/ro rang. KHONG xu ly (va PHAI canh bao nguoi dung o noi goi ham nay):
//   - dau_ra_kv (vi du key "momo_ct_map_<COMPANY>") -- blob JSON tuy y do
//     views/dau-ra.ejs tu dinh nghia, server khong biet schema ben trong.
//   - Cac map doi soat Zalo/VNPay/Payoo/VietQR luu trong localStorage TRINH
//     DUYET cua nguoi dung (khong phai server) -- ve nguyen tac KHONG THE
//     dong bo tu backend duoc.
//
// Ma cong trinh KH Cu / KH Moi la 2 tap ky tu HOAN TOAN khac nhau (2 phap
// nhan MISA khac nhau, xem ghi chu Luyen 2026-08-21 trong routes/danh-muc.js)
// nen voi cac bang "toan cuc" (khong tach truong company/congTy trong
// store), doi theo GIA TRI la an toan -- khong lo dung nham sang cong ty kia.
const { normCode } = require("./momoReconcile");

function eq(val, code) {
  if (val === null || val === undefined) return false;
  const s = String(val);
  return s.trim() === code.trim() || normCode(s) === normCode(code);
}

// Doi 1 cap (oldCode -> newCode) cho 1 cong ty. Sua truc tiep tren `store`
// (goi save(store) o noi goi ham nay sau khi xong TOAN BO batch).
// Tra ve { counts: {storeKey: soLuongDaSua}, warnings: [string] }.
function renameOne(store, company, oldCode, newCode) {
  const counts = {};
  const warnings = [];
  const bump = (key, n) => { if (n) counts[key] = (counts[key] || 0) + n; };

  // 1) Danh muc chinh (nguon cua tinh nang nay) -- routes/danh-muc.js
  const dmKey = "danh_muc_ma_cong_trinh_" + (company === "kh_moi" ? "moi" : "cu");
  if (Array.isArray(store[dmKey])) {
    let n = 0;
    store[dmKey].forEach((r) => {
      if (!eq(r.ma, oldCode)) return;
      const dup = store[dmKey].find((x) => x !== r && eq(x.ma, newCode));
      if (dup) {
        warnings.push(`Danh mục: mã mới "${newCode}" đã có sẵn -- bỏ qua đổi tên dòng "${r.ma}".`);
        return;
      }
      r.ma = newCode;
      r.updatedAt = new Date().toISOString();
      n++;
    });
    bump(dmKey, n);
  }

  // 2) ma_cong_trinh_master (danh muc "chuan" rieng, utils/maCongTrinh.js)
  if (store.ma_cong_trinh_master && store.ma_cong_trinh_master[company] && Array.isArray(store.ma_cong_trinh_master[company].rows)) {
    let n = 0;
    store.ma_cong_trinh_master[company].rows.forEach((r) => {
      if (eq(r.maCongTrinh, oldCode)) { r.maCongTrinh = newCode; n++; }
    });
    bump("ma_cong_trinh_master", n);
  }

  // 3) Hop dong thue gian (2 bang, dung chung 2 cong ty, loc theo congTy)
  ["phap_danh_hop_dong_thue", "phap_danh_hop_dong_thue_tong"].forEach((key) => {
    if (!Array.isArray(store[key])) return;
    let n = 0;
    store[key].forEach((r) => {
      const congTy = r.congTy || "kh_cu";
      if (congTy === company && eq(r.maCongTrinh, oldCode)) { r.maCongTrinh = newCode; n++; }
    });
    bump(key, n);
  });

  // 4) Doanh thu chia se theo diem + Tien thue theo NCC (routes/ho-so.js)
  ["dtcs_chiase_diem", "ho_so_tien_thue"].forEach((key) => {
    if (!Array.isArray(store[key])) return;
    let n = 0;
    store[key].forEach((r) => {
      const cty = r.company || "kh_cu";
      if (cty === company && eq(r.maCongTrinh, oldCode)) { r.maCongTrinh = newCode; n++; }
    });
    bump(key, n);
  });

  // 5) Hoa don dau ra -- CHU Y: ten field la maKH, khong phai maCongTrinh
  if (Array.isArray(store.hoa_don_dau_ra)) {
    let n = 0;
    store.hoa_don_dau_ra.forEach((r) => {
      const cty = r.company || "kh_cu";
      if (cty === company && eq(r.maKH, oldCode)) { r.maKH = newCode; n++; }
    });
    bump("hoa_don_dau_ra", n);
  }

  // 6) Doi tuong NCC master (import MISA) -- object keyed boi maDT
  if (store.ncc_doi_tuong_master && typeof store.ncc_doi_tuong_master === "object") {
    let n = 0;
    Object.values(store.ncc_doi_tuong_master).forEach((r) => {
      if (r && eq(r.maCongTrinh, oldCode)) { r.maCongTrinh = newCode; n++; }
    });
    bump("ncc_doi_tuong_master", n);
  }

  // ---- Cac bang "toan cuc" dung chung Momo/Zalo/VNPay/Payoo/VietQR ----

  // 7) gian_mapping / zvp_gian_mapping -- KEYED BOI chinh ma cong trinh
  ["gian_mapping", "zvp_gian_mapping"].forEach((key) => {
    if (!store[key] || typeof store[key] !== "object") return;
    const foundKey = Object.keys(store[key]).find((k) => eq(k, oldCode));
    if (foundKey) {
      if (!Object.prototype.hasOwnProperty.call(store[key], newCode)) {
        store[key][newCode] = store[key][foundKey];
      }
      delete store[key][foundKey];
      bump(key, 1);
    }
  });

  // 8) cua_hang_mapping -- value = { maCongTrinh, code, gian }
  if (store.cua_hang_mapping && typeof store.cua_hang_mapping === "object") {
    let n = 0;
    Object.values(store.cua_hang_mapping).forEach((v) => {
      if (v && eq(v.maCongTrinh, oldCode)) {
        v.maCongTrinh = newCode;
        if (eq(v.code, oldCode)) v.code = newCode;
        n++;
      }
    });
    bump("cua_hang_mapping", n);
  }

  // 9) invoice_diem_alias -- value = ma (co the co hau to __FF)
  if (store.invoice_diem_alias && typeof store.invoice_diem_alias === "object") {
    let n = 0;
    Object.keys(store.invoice_diem_alias).forEach((k) => {
      const v = store.invoice_diem_alias[k];
      if (typeof v !== "string") return;
      const hasFF = v.endsWith("__FF");
      const base = hasFF ? v.slice(0, -4) : v;
      if (eq(base, oldCode)) {
        store.invoice_diem_alias[k] = newCode + (hasFF ? "__FF" : "");
        n++;
      }
    });
    bump("invoice_diem_alias", n);
  }

  // 10) zvp_gian_list[] -- {tenDiem, maCongTrinh, isCse}
  if (Array.isArray(store.zvp_gian_list)) {
    let n = 0;
    store.zvp_gian_list.forEach((r) => { if (eq(r.maCongTrinh, oldCode)) { r.maCongTrinh = newCode; n++; } });
    bump("zvp_gian_list", n);
  }

  // 11) zvp_offline_diem_map / zvp_payoo_diem_map / zvp_online_product_map -- value.maCongTrinh
  ["zvp_offline_diem_map", "zvp_payoo_diem_map", "zvp_online_product_map"].forEach((key) => {
    if (!store[key] || typeof store[key] !== "object") return;
    let n = 0;
    Object.values(store[key]).forEach((v) => { if (v && eq(v.maCongTrinh, oldCode)) { v.maCongTrinh = newCode; n++; } });
    bump(key, n);
  });

  // 12) zvp_gian_master.rows[] -- {raw, maCongTrinh, thuoc, isCse}
  if (store.zvp_gian_master && Array.isArray(store.zvp_gian_master.rows)) {
    let n = 0;
    store.zvp_gian_master.rows.forEach((r) => { if (eq(r.maCongTrinh, oldCode)) { r.maCongTrinh = newCode; n++; } });
    bump("zvp_gian_master", n);
  }

  // 13) VietQR: ten_diem_master / store_code_override / ref_override -- {channel: {key: maCongTrinh}}
  ["viet_qr_ten_diem_master", "viet_qr_store_code_override", "viet_qr_ref_override"].forEach((key) => {
    if (!store[key] || typeof store[key] !== "object") return;
    let n = 0;
    Object.values(store[key]).forEach((channelMap) => {
      if (!channelMap || typeof channelMap !== "object") return;
      Object.keys(channelMap).forEach((k) => {
        if (eq(channelMap[k], oldCode)) { channelMap[k] = newCode; n++; }
      });
    });
    bump(key, n);
  });

  // 14) viet_qr_invoices[channel][] -- .maDiem
  if (store.viet_qr_invoices && typeof store.viet_qr_invoices === "object") {
    let n = 0;
    Object.values(store.viet_qr_invoices).forEach((list) => {
      if (!Array.isArray(list)) return;
      list.forEach((inv) => { if (inv && eq(inv.maDiem, oldCode)) { inv.maDiem = newCode; n++; } });
    });
    bump("viet_qr_invoices", n);
  }

  // 15) cht_nop_tien_map -- value.maCongTrinh
  if (store.cht_nop_tien_map && typeof store.cht_nop_tien_map === "object") {
    let n = 0;
    Object.values(store.cht_nop_tien_map).forEach((v) => { if (v && eq(v.maCongTrinh, oldCode)) { v.maCongTrinh = newCode; n++; } });
    bump("cht_nop_tien_map", n);
  }

  // 16) chi_phi_gian_alias / chi_phi_gian_override -- value = ma
  ["chi_phi_gian_alias", "chi_phi_gian_override"].forEach((key) => {
    if (!store[key] || typeof store[key] !== "object") return;
    let n = 0;
    Object.keys(store[key]).forEach((k) => {
      if (eq(store[key][k], oldCode)) { store[key][k] = newCode; n++; }
    });
    bump(key, n);
  });

  // 17) zvp_manual_matches -- KEY = "<date>|<code>[__FF]"
  if (store.zvp_manual_matches && typeof store.zvp_manual_matches === "object") {
    let n = 0;
    Object.keys(store.zvp_manual_matches).forEach((k) => {
      const parts = k.split("|");
      if (parts.length < 2) return;
      const codePart = parts[1];
      const hasFF = codePart.endsWith("__FF");
      const codeBase = hasFF ? codePart.slice(0, -4) : codePart;
      if (!eq(codeBase, oldCode)) return;
      const newKey = parts[0] + "|" + newCode + (hasFF ? "__FF" : "");
      if (Object.prototype.hasOwnProperty.call(store.zvp_manual_matches, newKey)) {
        warnings.push(`zvp_manual_matches: khóa "${newKey}" đã tồn tại -- bỏ qua "${k}".`);
        return;
      }
      store.zvp_manual_matches[newKey] = store.zvp_manual_matches[k];
      delete store.zvp_manual_matches[k];
      n++;
    });
    bump("zvp_manual_matches", n);
  }

  // 18) Don dep bang "virtual rename" de tranh xung dot voi ma moi da that
  // su duoc ghi vao du lieu goc o tren -- deu KEYED BOI ma cu.
  ["viet_qr_gian_merge", "ma_cong_trinh_display_alias"].forEach((key) => {
    if (!store[key] || typeof store[key] !== "object") return;
    const foundKey = Object.keys(store[key]).find((k) => eq(k, oldCode));
    if (foundKey && foundKey !== newCode) {
      if (!Object.prototype.hasOwnProperty.call(store[key], newCode)) {
        store[key][newCode] = store[key][foundKey];
      }
      delete store[key][foundKey];
      bump(key, 1);
    }
  });

  return { counts, warnings };
}

// mappings: [{oldCode, newCode}, ...]
// Tra ve { applied: [{oldCode,newCode,counts}], skipped: [string], totalCounts: {}, warnings: [] }
function applyRenameBatch(store, company, mappings) {
  const totalCounts = {};
  const allWarnings = [];
  const applied = [];
  const skipped = [];
  (mappings || []).forEach((m) => {
    const oldCode = String((m && m.oldCode) || "").trim();
    const newCode = String((m && m.newCode) || "").trim();
    if (!oldCode || !newCode) return;
    if (eq(oldCode, newCode)) { skipped.push(`"${oldCode}": mã mới giống mã cũ, bỏ qua.`); return; }
    const { counts, warnings } = renameOne(store, company, oldCode, newCode);
    const totalTouched = Object.values(counts).reduce((s, v) => s + v, 0);
    if (totalTouched === 0) {
      skipped.push(`Không tìm thấy mã "${oldCode}" ở bất kỳ đâu -- bỏ qua.`);
      return;
    }
    applied.push({ oldCode, newCode, counts });
    Object.entries(counts).forEach(([k, v]) => { totalCounts[k] = (totalCounts[k] || 0) + v; });
    allWarnings.push(...warnings);
  });
  return { applied, skipped, totalCounts, warnings: allWarnings };
}

module.exports = { applyRenameBatch };
