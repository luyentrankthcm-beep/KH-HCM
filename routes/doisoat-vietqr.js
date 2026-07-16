const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { load, save, nextId } = require("../store");
const { requireLogin } = require("../middleware/auth");
const {
  extractVqrCode,
  extractVietQrSettlements,
  parseVietQrRawWorkbook,
  parseCuaHangSheet,
  parseStoreExportSheet,
  parseInvoiceWorkbookByTag,
  buildGianCandidatesFromInvoices,
  resolveGianGross,
  reconcileVietQr,
  parseVietQrMnRawWorkbook,
  parseMaCuaHangAppSheet,
  resolveGianGrossPrefix,
  isoToDmy,
  displayCode,
  FF_SUFFIX,
} = require("../utils/vietqrReconcile");
const { mergeGianListWithMaster, applyGianRedirectToInvoices } = require("../utils/zvpReconcile");
const { getCompany } = require("../utils/companies");

const router = express.Router();
router.use(requireLogin);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 80 * 1024 * 1024 },
});

// Each Viet QR bank posts individual QR payments (no batch settlement), so
// the reconciliation is per calendar day per bank. Store.transactions'
// bank.name is the join key against store.banks; the tag pattern is how a
// shared invoice-list upload (same "MTT" file used by Momo/ZVP) tells this
// bank's invoices apart from the others' -- see utils/vietqrReconcile.js.
const CHANNELS = {
  bidv7704: { bankName: "BIDV7704", label: "BIDV 7704", tagPattern: /POSH\+JP\s*MB\s*\(7704\)/i, company: "kh_cu" },
  bidv77020: { bankName: "BIDV77020", label: "BIDV 77020", tagPattern: /POSH\+JP\s*MB\s*\(7020\)/i, company: "kh_cu" },
  mb11521268: { bankName: "MB11521268", label: "MB 11521268", tagPattern: /POSH\+JP\s*MB\s*\(268\)/i, company: "kh_cu" },
  // Cong ty "KH Moi" (TNHH GIAI TRI K&H) -- tai khoan Viet QR BIDV7702. Tag
  // hoa don la "MTD MN" (xac nhan tu Luyen 2026-07-16, nam tren sheet "ke ds
  // xuat HD MTT - 705" cua file MTT dung chung -- xem parseInvoiceWorkbookByTag
  // trong utils/vietqrReconcile.js). parseMode: "mn" danh dau kenh nay dung
  // rieng file export "VIETQR MN 7702.xlsx" (sheet "VIET QR" + "Ma Cua Hang
  // APP", khong co cot "Noi dung TT"/token VQR nhu 3 kenh kia) va khop gian
  // theo TIEN TO ten cua hang (vd "AMTP 01" -> "AMTP") thay vi fuzzy text,
  // vi ten cua hang o day la chu viet tat cua chinh Ma cong trinh.
  bidv7702: {
    bankName: "BIDV7702",
    label: "BIDV 7702",
    tagPattern: /MTD\s*MN/i,
    company: "kh_moi",
    parseMode: "mn",
  },
};
const CHANNEL_KEYS = Object.keys(CHANNELS);

const TKCO_VALUES = ["131", "1388", "SKIP"];
const UPDATED_NOTE = " Ket qua doi soat ben duoi da tu cap nhat theo du lieu moi.";

// Full gian merge/rename (Luyen-confirmed 2026-07-16): "JP SC VIVO" la CUNG
// 1 gian voi "SC VIVO KVCM", chia se doanh thu (CSE) tu dau -- hoa don ghi
// "JP SC VIVO" cho cac ngay dau, roi doi sang ghi "SC VIVO KVCM"/"FUNFEST
// SCVIVO" tu khoang 1-2 ngay sau (do do tre xuat hoa don). Merge permanent,
// ap dung ca hoa don cu (con ghi "JP SC VIVO") lan hoa don moi ve sau.
const GIAN_MERGE_DEFAULTS = {
  "JP SC VIVO": { maCongTrinh: "SC VIVO KVCM", isCse: true },
};

function ensureChannelShape(store) {
  if (!store.viet_qr_raw_uploads) store.viet_qr_raw_uploads = {};
  if (!store.viet_qr_store_names) store.viet_qr_store_names = {};
  if (!store.viet_qr_invoices) store.viet_qr_invoices = {};
  if (!store.viet_qr_manual_matches) store.viet_qr_manual_matches = {};
  if (!store.viet_qr_gian_merge) store.viet_qr_gian_merge = {};
  Object.keys(GIAN_MERGE_DEFAULTS).forEach((k) => {
    if (!store.viet_qr_gian_merge[k]) store.viet_qr_gian_merge[k] = GIAN_MERGE_DEFAULTS[k];
  });
  CHANNEL_KEYS.forEach((ch) => {
    if (!store.viet_qr_raw_uploads[ch]) store.viet_qr_raw_uploads[ch] = [];
    if (!store.viet_qr_store_names[ch]) store.viet_qr_store_names[ch] = {};
    if (!store.viet_qr_invoices[ch]) store.viet_qr_invoices[ch] = [];
    if (!store.viet_qr_manual_matches[ch]) store.viet_qr_manual_matches[ch] = {};
  });
}

// Merge raw-upload rows across multiple uploads the same way Momo/ZVP merge
// gross uploads: newest upload wins for any (date, ma cua hang) key it
// covers -- de-dup by (vqrCode) across uploads, newest first.
//
// IMPORTANT: some banks (verified on MB11521268 -- its "Noi dung TT" column
// is always literally "PaymentForOrder", no "VQR..." token at all) never
// produce a vqrCode, so EVERY row from that bank fell into a single flat
// "noCode" bucket with no de-dup at all. Luyen's workflow re-uploads a fresh
// CUMULATIVE "to date" export (not just the new days) each time, so any date
// covered by both the old and the new upload used to get counted TWICE --
// this was found causing exactly-2x gross vs invoice on every single-day
// line for MB11521268's "CHKQT CAM RANH" gian (confirmed against real data:
// every date in both uploads' overlapping range showed gross = 2 x invoice
// total, while dates only in the newer upload matched 1:1). Fixed by
// de-duping no-vqrCode rows PER DATE instead: the newest upload that
// contains ANY row for a given date fully replaces earlier uploads' rows
// for that same date, rather than appending both.
function mergeRawRows(uploads) {
  const sorted = [...uploads].sort((a, b) => new Date(a.uploaded_at) - new Date(b.uploaded_at));
  const byVqr = {};
  const noCodeByDate = {};
  for (const u of sorted) {
    const datesInThisUpload = new Set();
    for (const row of u.rows || []) {
      if (row.vqrCode) {
        byVqr[row.vqrCode] = row;
      } else if (row.date) {
        datesInThisUpload.add(row.date);
      }
    }
    // Newest upload wins: wipe any earlier noCode rows for a date this
    // upload also covers, then refill with this upload's own rows for it.
    datesInThisUpload.forEach((d) => {
      noCodeByDate[d] = [];
    });
    for (const row of u.rows || []) {
      if (!row.vqrCode && row.date) {
        noCodeByDate[row.date].push(row);
      }
    }
  }
  const noCode = Object.values(noCodeByDate).flat();
  return [...Object.values(byVqr), ...noCode];
}

function seedGianMappingDefaults(store, codes) {
  codes.forEach((c) => {
    if (!(c in store.gian_mapping)) {
      store.gian_mapping[c] = c.endsWith(FF_SUFFIX) ? "1388" : "131";
    }
  });
}

// Rows from the shared daily master "gian " sheet (uploaded on the Zalo/
// VNPay/Payoo page, "0. Tai bang gian tong hop hang ngay") whose "Thuoc"
// column matches THIS Viet QR channel's own invoice tag pattern (e.g.
// "POSH+JP MB (7704) 11,12" for BIDV7704) -- reusing the exact same regex
// already used to tell this channel's invoices apart, so a row only has to
// be tagged correctly once on that one shared sheet. Needed because
// buildGianCandidatesFromInvoices alone only knows the RAW "Ma diem tren
// misa thue" text straight off each invoice, which for a site not yet typed
// in with its true Misa code (e.g. "Sân bay Phú Quốc" instead of the real
// "CHKQT PHU QUOC") would otherwise show up AS ITSELF in the reconciliation
// table and get exported to MISA under the wrong Ma cong trinh.
function getMasterRowsForVietQrChannel(store, tagPattern) {
  const rows = (store.zvp_gian_master && store.zvp_gian_master.rows) || [];
  return rows.filter((r) => tagPattern.test(r.thuoc || ""));
}

function buildChannelReconciliation(store, channelKey) {
  ensureChannelShape(store);
  const cfg = CHANNELS[channelKey];
  const bank = store.banks.find((b) => b.name === cfg.bankName);
  if (!bank) {
    return { error: `Chua co ngan hang "${cfg.bankName}" trong he thong.` };
  }
  const txs = store.transactions.filter((t) => t.bank_id === bank.id);
  const settlements = extractVietQrSettlements(txs);

  const rawRows = mergeRawRows(store.viet_qr_raw_uploads[channelKey]);
  const storeNames = store.viet_qr_store_names[channelKey];
  let invoices = store.viet_qr_invoices[channelKey];

  // Full gian merge/rename (Luyen-confirmed, applies to ALL invoices past +
  // future -- e.g. "JP SC VIVO" turned out to be the SAME site as "SC VIVO
  // KVCM" (invoiced from some point on as "FUNFEST SCVIVO"), doanh thu chia
  // se (CSE) tu dau. Unlike a single mid-stream CSE-status change (handled
  // generically in reconcileVietQr's per-day fallback below), this is a
  // permanent identity merge onto a DIFFERENT ma cong trinh entirely, so it
  // has to rewrite maDiem before candidates are even built.
  const gianMerge = store.viet_qr_gian_merge || {};
  const mergeKeys = Object.keys(gianMerge);
  if (mergeKeys.length > 0) {
    invoices = invoices.map((inv) => {
      const hadFF = inv.maDiem && inv.maDiem.endsWith(FF_SUFFIX);
      const baseCode = hadFF ? inv.maDiem.slice(0, -FF_SUFFIX.length) : inv.maDiem;
      const merge = gianMerge[baseCode];
      if (!merge) return inv;
      // isCse: true/false FORCES that CSE status onto every invoice (full
      // identity merge -- e.g. JP SC VIVO, genuinely CSE tu dau). isCse null/
      // undefined means "rename only" -- keep each invoice's OWN CSE status
      // as-is (e.g. Sân bay Phú Quốc -> CHKQT PHU QUOC: just the Misa export
      // code changes, the day-7-onward CSE split -- handled separately by
      // reconcileVietQr's per-day fallback -- must NOT be flattened here).
      const targetIsCse = merge.isCse === true || merge.isCse === false ? merge.isCse : hadFF;
      const newMaDiem = targetIsCse ? merge.maCongTrinh + FF_SUFFIX : merge.maCongTrinh;
      if (newMaDiem === inv.maDiem) return inv;
      return { ...inv, maDiem: newMaDiem };
    });
  }

  let gianCandidates = buildGianCandidatesFromInvoices(invoices);
  // Master gian sheet wins on a name collision (it carries the real Misa
  // code, e.g. "CHKQT PHU QUOC"), same precedence rule as ZVP's Online
  // channel -- see getMasterRowsForVietQrChannel above.
  gianCandidates = mergeGianListWithMaster(gianCandidates, getMasterRowsForVietQrChannel(store, cfg.tagPattern));
  if (mergeKeys.length > 0) {
    // The physical POS/store name on the bank/QR side never changes just
    // because invoicing renamed the site -- keep the OLD name fuzzy-
    // matchable for resolveGianGross, now pointing at the merged code.
    gianCandidates = mergeGianListWithMaster(
      gianCandidates,
      mergeKeys.map((oldCode) => ({
        raw: oldCode,
        maCongTrinh: gianMerge[oldCode].maCongTrinh,
        isCse: gianMerge[oldCode].isCse,
      }))
    );
  }
  // BIDV7702/VietQR MN: khop gian theo tien to ten cua hang (vd "AMTP 01"
  // -> "AMTP" ung voi ma cong trinh "AM TP KVCM"), khong dung fuzzy text
  // matcher nhu 3 kenh kia -- xem ghi chu tai CHANNELS.bidv7702 o tren.
  const resolved =
    cfg.parseMode === "mn"
      ? resolveGianGrossPrefix(rawRows, storeNames, gianCandidates)
      : resolveGianGross(rawRows, storeNames, gianCandidates);

  // Redirect each invoice's own "Ma diem tren misa thue" through the SAME
  // gianCandidates map (keyed by invoice's own "Ten diem xuat hoa don"),
  // same mechanism as ZVP's applyGianRedirectToInvoices -- otherwise an
  // invoice whose Ma diem column is just the site's own raw name (e.g.
  // "SÂN BAY PHÚ QUỐC" instead of "CHKQT PHU QUOC") never lines up against
  // the settlement line above (which IS already correctly redirected via
  // the master sheet), permanently showing as "chua khop" even after the
  // master gian sheet fixes the code everywhere else.
  invoices = applyGianRedirectToInvoices(invoices, gianCandidates);

  seedGianMappingDefaults(store, resolved.codes);

  const manualMatches = store.viet_qr_manual_matches[channelKey] || {};
  const reconciled = reconcileVietQr(
    settlements,
    resolved,
    { invoices },
    store.gian_mapping,
    manualMatches,
    store.invoice_diem_alias
  );

  const allCodes = new Set();
  reconciled.forEach((r) => r.lines.forEach((l) => allCodes.add(l.code)));

  const knownCodesForAlias = new Set([...allCodes, ...Object.keys(store.gian_mapping || {})]);
  const invoiceDiemAlias = store.invoice_diem_alias || {};
  const unmatchedInvoiceCodesSet = new Set();
  invoices.forEach((inv) => {
    if (!inv.maDiem) return;
    if (knownCodesForAlias.has(inv.maDiem)) return;
    if (invoiceDiemAlias[inv.maDiem]) return;
    unmatchedInvoiceCodesSet.add(inv.maDiem);
  });

  return {
    reconciled,
    allCodes: Array.from(allCodes).sort(),
    unmappedStores: resolved.unmapped,
    invoiceDiemAlias,
    unmatchedInvoiceCodes: Array.from(unmatchedInvoiceCodesSet).sort(),
  };
}

router.get("/doi-soat/vietqr", (req, res) => {
  const store = load();
  ensureChannelShape(store);

  // Chi hien cac ngan hang thuoc cong ty dang chon (nut chuyen cong ty tren
  // topbar) -- KH Cu van thay ca 3 ngan hang nhu truoc gio, KH Moi thay
  // BIDV 7702. Cac route upload/xoa/xuat file van nhan moi channelKey hop
  // le (khong gioi han theo cong ty) vi Luyen la nguoi dung duy nhat.
  const activeCompany = getCompany(req);
  const activeKeys = CHANNEL_KEYS.filter((ch) => CHANNELS[ch].company === activeCompany);

  const built = {};
  activeKeys.forEach((ch) => {
    built[ch] = buildChannelReconciliation(store, ch);
  });

  const monthSet = new Set();
  activeKeys.forEach((ch) => {
    (built[ch].reconciled || []).forEach((r) => monthSet.add(r.settlementDate.slice(0, 7)));
  });
  const months = Array.from(monthSet).sort().reverse();
  const selectedMonth = req.query.month !== undefined ? req.query.month : months[0] || "";
  // Chon ngan hang de xem: 3 ngan hang doi soat rieng, cuon qua ca 3 de tim
  // 1 dong rat mat cong -- mac dinh van hien ca 3 ("Tat ca"), nhung Luyen co
  // the loc con 1 ngan hang cho de doi chieu.
  const selectedChannel = req.query.channel !== undefined ? req.query.channel : "";

  const reconciledByChannel = {};
  activeKeys.forEach((ch) => {
    const all = built[ch].reconciled || [];
    reconciledByChannel[ch] = selectedMonth ? all.filter((r) => r.settlementDate.slice(0, 7) === selectedMonth) : all;
  });

  const allCodes = new Set();
  activeKeys.forEach((ch) => (built[ch].allCodes || []).forEach((c) => allCodes.add(c)));

  // "Bang tong quan theo gian": 1 dong = 1 ma cong trinh, 1 cot = 1 ngay
  // (trong pham vi thang/ngan hang dang loc), de Luyen nhin duoc lech cua
  // CA THANG cung 1 luc thay vi cuon qua tung ngay rieng le. Bam vao 1 o de
  // xem chi tiet (so HD, doanh thu, lech...) -- van dung dung 1 nguon du
  // lieu (reconciledByChannel) voi bang "Ket qua doi soat" chi tiet ben
  // duoi, nen 2 bang luon khop nhau tuyet doi.
  const pivotByChannel = {};
  activeKeys.forEach((ch) => {
    const rows = reconciledByChannel[ch] || [];
    const dates = Array.from(new Set(rows.map((r) => r.settlementDate))).sort();
    const cellMap = {};
    const codesSet = new Set();
    rows.forEach((r) => {
      r.lines.forEach((l) => {
        codesSet.add(l.code);
        if (!cellMap[l.code]) cellMap[l.code] = {};
        cellMap[l.code][r.settlementDate] = l;
      });
    });
    const codes = Array.from(codesSet).sort();
    pivotByChannel[ch] = {
      dates,
      rows: codes.map((code) => {
        let sumGross = 0;
        let sumInvoiceTotal = 0;
        const cells = dates.map((d) => {
          const l = (cellMap[code] || {})[d];
          if (!l) return null;
          sumGross += l.gross;
          sumInvoiceTotal += l.invoiceTotal;
          let status;
          if (l.tkCo === "SKIP") status = "skip";
          else if (l.manualOverride) status = "ok";
          else if (l.invoiceNumbers.length === 0) status = "missing";
          else if (l.matched) status = "ok";
          else status = "diff";
          return {
            tkCo: l.tkCo,
            gross: l.gross,
            invoiceTotal: l.invoiceTotal,
            diff: l.diff,
            invoiceNumbers: l.invoiceNumbers,
            matched: l.matched,
            manualOverride: l.manualOverride,
            status,
          };
        });
        return {
          code,
          maCongTrinh: displayCode(code),
          isCse: code.endsWith(FF_SUFFIX),
          cells,
          sumGross,
          sumInvoiceTotal,
          sumDiff: sumInvoiceTotal - sumGross,
        };
      }),
    };
  });

  res.render("doisoat-vietqr", {
    userName: req.session.userName,
    channels: activeKeys.map((ch) => ({ key: ch, label: CHANNELS[ch].label, parseMode: CHANNELS[ch].parseMode || null })),
    selectedChannel,
    rawUploads: activeKeys.reduce((acc, ch) => {
      acc[ch] = store.viet_qr_raw_uploads[ch];
      return acc;
    }, {}),
    invoiceCounts: activeKeys.reduce((acc, ch) => {
      acc[ch] = store.viet_qr_invoices[ch].length;
      return acc;
    }, {}),
    reconciled: reconciledByChannel,
    pivotByChannel,
    built,
    months,
    selectedMonth,
    gianMapping: store.gian_mapping,
    allCodes: Array.from(allCodes).sort(),
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// ---------- Upload: raw QR export file (contains BOTH the transaction log sheet AND the "Cua hang" store-catalog sheet) ----------
router.post("/doi-soat/vietqr/upload-raw/:channel", upload.single("file"), (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");

    // BIDV7702/VietQR MN dung dinh dang file rieng ("VIETQR MN 7702.xlsx":
    // sheet "VIET QR" + "Ma Cua Hang APP") -- khong co cot "Noi dung TT"/
    // token VQR nhu 3 kenh kia nen phai dung parser rieng.
    const isMn = CHANNELS[channelKey].parseMode === "mn";
    const parsed = isMn ? parseVietQrMnRawWorkbook(req.file.buffer) : parseVietQrRawWorkbook(req.file.buffer);
    const storeMap = isMn ? parseMaCuaHangAppSheet(req.file.buffer) : parseCuaHangSheet(req.file.buffer);

    store.viet_qr_raw_uploads[channelKey].push({
      id: nextId(store, "viet_qr_raw_uploads_seq") || Date.now(),
      uploaded_at: new Date().toISOString(),
      file_name: req.file.originalname,
      sheetName: parsed.sheetName,
      rows: parsed.rows,
    });
    if (Object.keys(storeMap).length > 0) {
      store.viet_qr_store_names[channelKey] = Object.assign({}, store.viet_qr_store_names[channelKey], storeMap);
    }
    save(store);

    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(
          `Da nap "${parsed.sheetName}": ${parsed.rows.length} giao dich QR, ${Object.keys(storeMap).length} cua hang.${UPDATED_NOTE}`
        )
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: danh sach diem ban rieng ("store_export"), doc lap voi
// file giao dich QR -- cho phep Luyen nap truoc bang ten diem ban moi nhat
// TRUOC KHI co giao dich, thay vi phai cho co du lieu QR moi biet ten. Neu 1
// ma cua hang la MOI (chua tung thay) hoac DOI ten diem ban so voi lan
// truoc, canh bao ngay trong thong bao de Luyen kiem tra xem co phai gian
// moi can gan vao ngan hang dang up hay khong. ----------
router.post("/doi-soat/vietqr/upload-store/:channel", upload.single("file"), (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  try {
    if (!CHANNELS[channelKey]) throw new Error("Kenh khong hop le.");
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");

    const { sheetName, map } = parseStoreExportSheet(req.file.buffer);
    const existing = store.viet_qr_store_names[channelKey] || {};
    const newEntries = [];
    const changedEntries = [];
    Object.keys(map).forEach((maCuaHang) => {
      const prev = existing[maCuaHang];
      const next = map[maCuaHang];
      if (!prev) {
        newEntries.push(`${maCuaHang} -> ${next.tenDiemBan || next.matchText}`);
      } else if (next.tenDiemBan && prev.tenDiemBan !== next.tenDiemBan) {
        changedEntries.push(`${maCuaHang}: "${prev.tenDiemBan || prev.matchText}" -> "${next.tenDiemBan}"`);
      }
    });
    // Cong don: giu nguyen cac ma cua hang cu khong co trong file moi, chi
    // them moi/cap nhat nhung ma co trong file nay (khong xoa du lieu cu).
    store.viet_qr_store_names[channelKey] = Object.assign({}, existing, map);
    save(store);

    let msg = `Da nap "${sheetName}": ${Object.keys(map).length} diem ban (${CHANNELS[channelKey].label}).${UPDATED_NOTE}`;
    if (newEntries.length > 0) {
      msg += ` CANH BAO: ${newEntries.length} ma cua hang MOI, kiem tra xem co phai gian moi khong -- ${newEntries
        .slice(0, 8)
        .join("; ")}${newEntries.length > 8 ? `... va ${newEntries.length - 8} ma khac` : ""}.`;
    }
    if (changedEntries.length > 0) {
      msg += ` CANH BAO: ${changedEntries.length} ma cua hang DOI TEN diem ban so voi lan truoc -- ${changedEntries
        .slice(0, 8)
        .join("; ")}${changedEntries.length > 8 ? `... va ${changedEntries.length - 8} ma khac` : ""}.`;
    }
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Upload: danh sach hoa don dung chung (file MTT), gan 3 tag Viet QR cung luc ----------
router.post("/doi-soat/vietqr/upload-hoadon", upload.single("file"), (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    if (!req.file) throw new Error("Vui long chon 1 file de tai len.");
    const addedCounts = {};
    let sheetName = null;
    for (const ch of CHANNEL_KEYS) {
      const parsed = parseInvoiceWorkbookByTag(req.file.buffer, CHANNELS[ch].tagPattern);
      sheetName = parsed.sheetName;
      const existingKeys = new Set(store.viet_qr_invoices[ch].map((i) => `${i.soHd}|${i.ngayHd}|${i.maDiem}`));
      let added = 0;
      for (const inv of parsed.invoices) {
        const k = `${inv.soHd}|${inv.ngayHd}|${inv.maDiem}`;
        if (existingKeys.has(k)) continue;
        existingKeys.add(k);
        store.viet_qr_invoices[ch].push(inv);
        added++;
      }
      addedCounts[ch] = added;
    }
    save(store);
    const summary = CHANNEL_KEYS.map((ch) => `${CHANNELS[ch].label}: ${addedCounts[ch]}`).join(", ");
    res.redirect(
      "/doi-soat/vietqr?success=" +
        encodeURIComponent(`Da nap sheet "${sheetName}": them moi hoa don theo kenh -- ${summary}.${UPDATED_NOTE}`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vietqr/mapping", (req, res) => {
  const store = load();
  const body = req.body || {};
  for (const [key, val] of Object.entries(body)) {
    if (key.startsWith("tkco_")) {
      const code = key.slice("tkco_".length);
      if (TKCO_VALUES.includes(val)) store.gian_mapping[code] = val;
    }
  }
  save(store);
  res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da luu bang TK Co theo gian."));
});

// ---------- Alias: dung CHUNG voi Momo/ZVP (store.invoice_diem_alias) ----------
router.post("/doi-soat/vietqr/diem-alias", (req, res) => {
  const store = load();
  try {
    const { sourceCode, targetCode } = req.body;
    if (!sourceCode || !targetCode) throw new Error("Thieu ma diem tren hoa don hoac ma cong trinh de anh xa.");
    if (!store.invoice_diem_alias) store.invoice_diem_alias = {};
    store.invoice_diem_alias[sourceCode] = targetCode;
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" + encodeURIComponent(`Da anh xa "${sourceCode}" -> "${targetCode}". Ket qua doi soat da tu cap nhat.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vietqr/diem-alias/delete", (req, res) => {
  const store = load();
  const { sourceCode } = req.body;
  if (store.invoice_diem_alias) delete store.invoice_diem_alias[sourceCode];
  save(store);
  res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da xoa anh xa ma diem."));
});

router.post("/doi-soat/vietqr/upload-raw/:channel/:id/delete", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.params.channel;
  if (!CHANNELS[channelKey]) return res.redirect("/doi-soat/vietqr");
  store.viet_qr_raw_uploads[channelKey] = store.viet_qr_raw_uploads[channelKey].filter((u) => String(u.id) !== req.params.id);
  save(store);
  res.redirect("/doi-soat/vietqr");
});

router.post("/doi-soat/vietqr/invoices/clear", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  CHANNEL_KEYS.forEach((ch) => (store.viet_qr_invoices[ch] = []));
  save(store);
  res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da xoa toan bo hoa don Viet QR da nap (ca 3 kenh)."));
});

// ---------- Manual match: dong "Chua co HD" da xac nhan la co HD bu ----------
router.post("/doi-soat/vietqr/manual-match", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    const { channel, settlementDate, code, invoiceNumbers, amount, note, grossAdjustment } = req.body;
    if (!CHANNELS[channel]) throw new Error("Kenh khong hop le.");
    if (!settlementDate || !code) throw new Error("Thieu thong tin dong can danh dau.");
    const key = `${settlementDate}|${code}`;
    const invoiceList = (invoiceNumbers || "")
      .split(/[,\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const amt = amount ? Number(String(amount).replace(/[^\d]/g, "")) : null;
    // grossAdjustment: cong them vao DOANH THU cua dong nay (khac voi "amount"
    // -- chi doi so tien HD hien thi) -- dung khi 1 GD QR "khong co ma cua
    // hang" da duoc Luyen xac dinh la thuoc ve gian nay, de tong "Tinh tu du
    // lieu tai len" ca ngay cung khop dung voi Ngan hang, khong con hien
    // "Chenh lech" o dau ngay du moi gian rieng le da "Khop".
    const grossAdj = grossAdjustment ? Number(String(grossAdjustment).replace(/[^\d]/g, "")) : 0;
    store.viet_qr_manual_matches[channel][key] = {
      invoiceNumbers: invoiceList,
      amount: amt,
      grossAdjustment: grossAdj || 0,
      note: note || "",
      created_at: new Date().toISOString(),
    };
    save(store);
    res.redirect(
      "/doi-soat/vietqr?success=" + encodeURIComponent(`Da danh dau thu cong dong "${code}" ngay ${settlementDate}.`)
    );
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

router.post("/doi-soat/vietqr/manual-match/delete", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  try {
    const { channel, settlementDate, code } = req.body;
    if (store.viet_qr_manual_matches[channel]) {
      delete store.viet_qr_manual_matches[channel][`${settlementDate}|${code}`];
    }
    save(store);
    res.redirect("/doi-soat/vietqr?success=" + encodeURIComponent("Da xoa danh dau thu cong."));
  } catch (e) {
    res.redirect("/doi-soat/vietqr?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Export: "MISATHue [kenh]" ----------
// Same 28-column "Mau phieu thu tien gui de nhap vao AMIS Accounting" as
// Momo/ZVP, with the SAME correct dien giai wording ("Thu tien dich vu vui
// choi giai tri theo HD ...") -- this was the whole point of building this
// as a proper reconciliation channel instead of hand-editing each bank's
// Excel file: the export now always uses that wording, no manual fixing.
function buildExportRows(reconciledList, startNo, bankAccount, bankFullName, lyDoThu) {
  let seq = startNo;
  const rows = [];
  reconciledList
    .sort((a, b) => (a.settlementDate > b.settlementDate ? 1 : -1))
    .forEach((r) => {
      const exportableLines = r.lines.filter((l) => l.tkCo !== "SKIP");
      if (exportableLines.length === 0) return;

      const soCt = "NTTK" + String(seq).padStart(7, "0") + "/26";
      seq++;
      const ngayDmy = isoToDmy(r.settlementDate);
      exportableLines.forEach((l) => {
        const hdText = l.invoiceNumbers.length > 0 ? l.invoiceNumbers.join(", ") : "";
        // Hau to kenh (Luyen, 2026-07-16, sua lai 2026-07-16): "- KL" cho Viet
        // QR (thay vi "- QR") de phan biet voi Momo ("- MM") va VNPay/Zalo/
        // Payoo ("- VNP"/"- PAYOO") ngay tren file xuat Misa, dat truoc
        // "theo HD ..." neu co.
        const dienGiai = hdText
          ? `Thu tiền dịch vụ vui chơi giải trí - KL theo HĐ ${hdText}`
          : "Thu tiền dịch vụ vui chơi giải trí - KL";
        rows.push({
          "Ngày hạch toán (*)": ngayDmy,
          "Ngày chứng từ (*)": ngayDmy,
          "Số chứng từ (*)": soCt,
          "Mã đối tượng": "KL",
          "Tên đối tượng": "",
          "Địa chỉ": "",
          "Nộp vào TK": bankAccount,
          "Mở tại ngân hàng": bankFullName,
          "Lý do thu": lyDoThu,
          "Diễn giải lý do thu": dienGiai,
          "Mã nhân viên thu": "",
          "Diễn giải (hạch toán)": dienGiai,
          "TK Nợ (*)": 112,
          "TK Có (*)": l.tkCo,
          "Số tiền": l.gross,
          "Mã đối tượng (hạch toán)": "KL",
          "Số khế ước đi vay": "",
          "Số khế ước cho vay": "",
          "Mã khoản mục chi phí": "",
          "Mã đơn vị": "",
          "Mã đối tượng THCP": "",
          "Mã công trình": l.maCongTrinh,
          "Số đơn đặt hàng": "",
          "Số đơn mua hàng": "",
          "Số hợp đồng mua": "",
          "Số hợp đồng bán": "",
          "Mã thống kê": "",
          "CP không hợp lý": "",
          "Số HĐ khớp": hdText,
          "Tổng tiền HĐ khớp": l.invoiceTotal,
          "Doanh thu (VietQR)": l.gross,
          "Chênh lệch HĐ vs doanh thu": l.diff,
          "Trạng thái": l.invoiceNumbers.length === 0 ? "Chưa có HĐ" : l.matched ? "Khớp" : "Lệch",
        });
      });
    });
  return { rows, nextSeq: seq };
}

router.get("/doi-soat/vietqr/export.xlsx", (req, res) => {
  const store = load();
  ensureChannelShape(store);
  const channelKey = req.query.channel;
  if (!CHANNELS[channelKey]) return res.status(400).send("Kenh khong hop le.");

  const bank = store.banks.find((b) => b.name === CHANNELS[channelKey].bankName);
  if (!bank) return res.status(400).send(`Chua co ngan hang "${CHANNELS[channelKey].bankName}" trong he thong.`);

  const built = buildChannelReconciliation(store, channelKey);
  if (built.error) return res.status(400).send(built.error);

  let startNo = parseInt(req.query.start || "1", 10);
  if (isNaN(startNo) || startNo < 1) startNo = 1;
  const { rows } = buildExportRows(
    built.reconciled,
    startNo,
    bank.account_number,
    `Ngân hàng ${bank.bank_name}`,
    `Thu tiền khách hàng qua Viet QR (${CHANNELS[channelKey].label})`
  );

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, "MISAThue VietQR");
  const buf = XLSX.write(wb2, { type: "buffer", bookType: "xlsx" });
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=MISAThue-vietqr-${channelKey}.xlsx`);
  res.send(buf);
});

// Exposed so routes/dashboard.js (Tong quan / Cong no) can reuse the exact
// same per-channel reconciliation this page shows, without a second
// implementation. CHANNELS/CHANNEL_KEYS let the caller loop over all 3 Viet
// QR banks (bidv7704, bidv77020, mb11521268) without hard-coding the list
// twice.
router.buildChannelReconciliation = buildChannelReconciliation;
router.VIETQR_CHANNELS = CHANNELS;
router.VIETQR_CHANNEL_KEYS = CHANNEL_KEYS;

module.exports = router;
