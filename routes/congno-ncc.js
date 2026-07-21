const express = require("express");
const { load } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { normText } = require("../utils/chiphiReconcile");
const chiPhi = require("./doisoat-chiphi");

const router = express.Router();
router.use(requireLogin);

// Luyen, 2026-07-19: "Công Nợ NCC" -- trang moi, tach khoi "Cong No Khach
// Hang" (cong no PHAI THU tu khach hang qua cac kenh doanh thu). Day la
// cong no PHAI TRA cho NCC (nha cung cap): Luyen xac nhan qua AskUserQuestion
// se dua vao danh sach NCC + giao dich chi da gan ma NCC ben trang "Doi
// soat Chi phi" -- KHONG tinh rieng, tai su dung dung logic khop NCC/hoa
// don/UNC da co san o routes/doisoat-chiphi.js (export them qua
// module.exports.buildChannelChiPhi/... o cuoi file do) de khong bao gio
// lech so voi trang Chi phi.
//
// Cong thuc 1 NCC: Tong hoa don (tu chi_phi_invoice_list, khop theo MST/ten
// khong dau) - Tong da chi (sum debit cac dong Chi phi da gan dung ma NCC
// nay, ca 3 kenh/ngan hang) = Con no (duong = con no NCC, am = da chi VUOT
// hoa don/tam ung). Chi hien NCC nao co it nhat 1 trong 2 so # 0, tranh
// bang qua dai voi ~884 NCC trong danh muc goc.
function buildNccDebt(store) {
  chiPhi.ensureShape(store);
  const nccList = chiPhi.activeNccList(store);

  // Tong hoa don NCC theo ma NCC -- khop hoa don (co MST/ten nguoi ban) ve
  // dung NCC trong danh sach qua MST truoc (chac chan nhat), khong co MST
  // khop thi thu qua ten khong dau.
  const byMst = new Map();
  const byNameNorm = new Map();
  nccList.forEach((n) => {
    if (n.mst) byMst.set(normText(n.mst), n);
    const nameKey = normText(n.tenKhongDau || n.tenNCC || "");
    if (nameKey) byNameNorm.set(nameKey, n);
  });

  const invoiceTotalByNcc = {};
  (store.chi_phi_invoice_list || []).forEach((inv) => {
    let rec = inv.mstNguoiBan ? byMst.get(normText(inv.mstNguoiBan)) : null;
    if (!rec) rec = byNameNorm.get(normText(inv.tenNguoiBan || ""));
    if (!rec) return; // hoa don khong khop duoc NCC nao trong danh muc -- bo qua, khong tinh nham
    invoiceTotalByNcc[rec.maNCC] = (invoiceTotalByNcc[rec.maNCC] || 0) + (inv.tongTienThanhToan || 0);
  });

  // Tong da chi theo ma NCC -- gom lai tu ca 3 kenh Chi phi (dung lai
  // buildChannelChiPhi, chi loc dong da co maNCC).
  const revMap = chiPhi.buildRevenueStatusMap(store);
  const paidTotalByNcc = {};
  const paidCountByNcc = {};
  const errors = [];
  chiPhi.CHANNEL_KEYS.forEach((ch) => {
    const built = chiPhi.buildChannelChiPhi(store, ch, revMap);
    if (built.error) {
      errors.push(`${chiPhi.CHANNELS[ch].label}: ${built.error}`);
      return;
    }
    built.lines.forEach((l) => {
      if (!l.maNCC) return;
      paidTotalByNcc[l.maNCC] = (paidTotalByNcc[l.maNCC] || 0) + (l.debit || 0);
      paidCountByNcc[l.maNCC] = (paidCountByNcc[l.maNCC] || 0) + 1;
    });
  });

  const allCodes = new Set([...Object.keys(invoiceTotalByNcc), ...Object.keys(paidTotalByNcc)]);
  const rows = Array.from(allCodes)
    .map((maNCC) => {
      const rec = nccList.find((n) => n.maNCC === maNCC);
      const tongHoaDon = invoiceTotalByNcc[maNCC] || 0;
      const daChi = paidTotalByNcc[maNCC] || 0;
      return {
        maNCC,
        tenNCC: rec ? rec.tenNCC : maNCC,
        mst: rec ? rec.mst : "",
        tongHoaDon,
        daChi,
        soGiaoDich: paidCountByNcc[maNCC] || 0,
        conNo: tongHoaDon - daChi,
      };
    })
    .filter((r) => r.tongHoaDon !== 0 || r.daChi !== 0)
    .sort((a, b) => Math.abs(b.conNo) - Math.abs(a.conNo));

  const grandTotalHoaDon = rows.reduce((s, r) => s + r.tongHoaDon, 0);
  const grandTotalDaChi = rows.reduce((s, r) => s + r.daChi, 0);
  const grandTotalConNo = rows.reduce((s, r) => s + r.conNo, 0);

  return { rows, grandTotalHoaDon, grandTotalDaChi, grandTotalConNo, errors };
}

router.get("/cong-no/ncc", (req, res) => {
  const store = load();
  let error = null;
  let result = { rows: [], grandTotalHoaDon: 0, grandTotalDaChi: 0, grandTotalConNo: 0, errors: [] };
  try {
    result = buildNccDebt(store);
  } catch (e) {
    error = e.message;
    console.error("Loi tinh cong no NCC:", e);
  }
  res.render("congno-ncc", {
    userName: req.session.userName,
    error,
    result,
  });
});

module.exports = router;
