const express = require("express");
const { load } = require("../store");
const { requireLogin } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");
const hoaDonDauVao = require("./hoa-don-dau-vao");
const { normVN } = require("../utils/hoaDonDauVaoEnrich");

const router = express.Router();
router.use(requireLogin);

// Chi Nhan, 2026-07-30: "bạn map từ hóa đơn với ngân hàng chi phí có 2 tài
// khoản chi mỗi kh Kh cũ là VP9997 và 8651 á còn kh mưới là Vp5888 và 8681 á
// bạn sẽ đưa đối tượng là ncc dựa trên file tôi đã tải đối tượng ncc bên hóa
// đơn đầu vào bạn map với ngân hàng á có tên đối ứng trên ngân hàng á xem nó
// thanh toán hóa đơn nào xem đã thanh toán chưa dựa trên hóa đơn và ngân hàng
// để lấy ra ncc để có thể check công nợ theo từng hóa đơn ... làm bản offline
// trc cho tôi nhá" -- VIET LAI HOAN TOAN trang "Cong No NCC": truoc day dua
// vao "Doi soat Chi phi" (3 kenh rieng, hien dang RONG tren may nay vi ban
// nay chua tai file Chi Phi nao), gio doi chieu THANG: Hoa Don Dau Vao (da
// co san Ten NCC tu file tai o trang do) voi giao dich "chi" tren DUNG 4 tai
// khoan Chi duoc chi dinh (KH Cu: VPBANK9997 + BIDV8651 | KH Moi: VP58888 +
// BIDV8681), khop qua Ten doi ung (tren sao ke) chua ten NCC + dung so tien
// hoa don (sai lech <=1000d) -- giong nguyen tac khop da dung o route
// /hoa-don-dau-vao/cap-nhat-da-chi nhung THU HEP dung 4 tai khoan chi dinh
// (khong lay TAT CA ngan hang cua cong ty) de tranh nhan nham giao dich khac.
// Trang nay CHI TINH DE HIEN THI (khong ghi de store), khong dung chung co
// "daChiTien" cua tinh nang "Cap nhat da chi" cu (giu nguyen, khong anh
// huong nhau) -- moi hoa don duoc tinh lai doc lap ngay khi mo trang.
const TAI_KHOAN_CHI_THEO_CONG_TY = {
  kh_cu: ["VPBANK9997", "BIDV8651"],
  kh_moi: ["VP58888", "BIDV8681"],
};
const COMPANY_LABEL = { kh_cu: "KH Cũ", kh_moi: "KH Mới" };

// Doi chieu tung hoa don (da gom dong qua groupRowsByInvoice) voi giao dich
// chi tren dung 4 tai khoan chi cua dung cong ty hoa don do -- tra ve 1 dong
// / hoa don kem trang thai da chi (neu khop) + giao dich khop duoc.
function buildInvoiceDebt(store) {
  hoaDonDauVao.ensureShape(store);
  const banksByName = new Map((store.banks || []).map((b) => [b.name, b]));
  const banksById = new Map((store.banks || []).map((b) => [b.id, b]));
  const allRows = (store.hoa_don_dau_vao || []).map(hoaDonDauVao.ensureDefaults);

  const invoices = [];
  const missingInvoiceChi = [];
  ["kh_cu", "kh_moi"].forEach((company) => {
    const bankNames = TAI_KHOAN_CHI_THEO_CONG_TY[company] || [];
    const bankIds = new Set(
      bankNames
        .map((n) => banksByName.get(n))
        .filter(Boolean)
        .map((b) => b.id)
    );
    // Moi giao dich chi tren 4 tai khoan chi, sap xep gan ngay hoa don nhat
    // truoc de uu tien khop dung dot thanh toan khi co nhieu giao dich trung
    // ten + so tien (vd cung 1 NCC duoc tra tien hang thang so tien co dinh).
    const chiTx = (store.transactions || [])
      .filter((t) => t.type === "chi" && bankIds.has(t.bank_id) && (t.tenDoiUng || "").trim())
      .map((t) => ({
        date: t.date,
        amount: t.amount,
        tenDoiUngNorm: normVN(t.tenDoiUng),
        tenDoiUngGoc: t.tenDoiUng,
        bankName: (banksById.get(t.bank_id) || {}).name || "",
        used: false,
      }));

    const rows = allRows.filter((r) => r.congTy === company);
    const groups = hoaDonDauVao.groupRowsByInvoice(rows);

    // Khop hoa don co ngay gan nhat truoc (on dinh, de doan nhu con nguoi doi
    // chieu tay: hoa don thang nao thi tim giao dich chi gan thang do truoc).
    groups
      .slice()
      .sort((a, b) => (a.ngayHD || "").localeCompare(b.ngayHD || ""))
      .forEach((g) => {
        let matched = null;
        if (g.tenNCC && g.soTien) {
          const nccKey = normVN(g.tenNCC).slice(0, 12);
          let best = null;
          let bestDiffDays = Infinity;
          chiTx.forEach((t) => {
            if (t.used) return;
            if (Math.abs(t.amount - g.soTien) > 1000) return;
            if (!nccKey || !t.tenDoiUngNorm.includes(nccKey)) return;
            const diffDays = g.ngayHD && t.date ? Math.abs(new Date(t.date) - new Date(g.ngayHD)) : 0;
            if (diffDays < bestDiffDays) {
              best = t;
              bestDiffDays = diffDays;
            }
          });
          matched = best;
        }
        if (matched) matched.used = true;
        const daChi = !!matched;
        invoices.push({
          company,
          companyLabel: COMPANY_LABEL[company],
          idsCsv: g.idsCsv,
          soHoaDon: g.soHoaDon,
          kyHieuHD: g.kyHieuHD,
          ngayHD: g.ngayHD,
          tenNCC: g.tenNCC || "(chưa có tên NCC)",
          mstNCC: g.mstNCC,
          soTien: g.soTien,
          daChi,
          ngayChi: matched ? matched.date : "",
          nganHangChi: matched ? matched.bankName : "",
          conNo: daChi ? 0 : g.soTien,
        });
      });

    // Chi Nhan, 2026-08-01: "đa phần đi từ tk công ty là cần lấy hóa đơn á bạn
    // xthêm cho tôi các ncc mà nếu chi của ngân hàng mà hk thấy hóa đơn thêm
    // vô ln nhá tôi không thể quét dc gmail để gắn link háo đơn vô dc nè" --
    // OAuth Google chưa cấu hình trên Railway nên không tự quét Gmail lấy hóa
    // đơn được (xem lỗi trên trang Chi Phí), nên Luyến cần thấy rõ: giao dịch
    // "chi" nào trên đúng 2 TK Chi của công ty này ĐÃ rời khỏi ngân hàng
    // nhưng KHÔNG khớp được với hóa đơn nào ở trên (used vẫn false sau vòng
    // lặp khớp) -- nghĩa là tiền đã chi thật nhưng CHƯA có hóa đơn ghi nhận.
    // Tách RIÊNG khỏi `invoices`/Còn nợ (tiền này không phải nợ, chỉ là THIẾU
    // CHỨNG TỪ) để không làm sai lệch các số Còn nợ NCC đang có, chỉ thêm 1
    // bảng cảnh báo riêng liệt kê Tên đối ứng (coi như tên NCC tạm) để Luyến
    // biết cần đi lấy hóa đơn bổ sung cho khoản nào.
    chiTx
      .filter((t) => !t.used)
      .forEach((t) => {
        missingInvoiceChi.push({
          company,
          companyLabel: COMPANY_LABEL[company],
          date: t.date,
          amount: t.amount,
          tenNCC: t.tenDoiUngGoc || "(không rõ tên đối ứng)",
          bankName: t.bankName,
        });
      });
  });
  return { invoices, missingInvoiceChi };
}

// Gop cac hoa don cung 1 NCC (trong cung 1 cong ty) thanh 1 dong tong hop --
// de xem nhanh NCC nao dang no nhieu nhat, khong can doc tung hoa don. Chi
// Nhan, 2026-07-30: "cho xem chi tiết dưới cái đó luôn đi đừng có nhấn gì cx
// nhẩy lên đầu trang" -- moi dong tong hop giu LUON danh sach hoa don cua
// dung NCC do (r.invoices), de trang render san 1 bang chi tiet AN SAN ngay
// duoi dong tong hop, bam vao la hien/an tai cho bang JS (khong tai lai
// trang, khong nhay vi tri cuon) thay vi phai dieu huong sang URL loc rieng.
function summarizeByNcc(invoices) {
  const map = new Map();
  invoices.forEach((inv) => {
    const key = inv.company + "||" + normVN(inv.tenNCC);
    if (!map.has(key)) {
      map.set(key, {
        tenNCC: inv.tenNCC,
        company: inv.company,
        companyLabel: inv.companyLabel,
        tongHoaDon: 0,
        daChi: 0,
        conNo: 0,
        soHoaDon: 0,
        soChuaChi: 0,
        invoices: [],
      });
    }
    const s = map.get(key);
    s.tongHoaDon += inv.soTien;
    if (inv.daChi) s.daChi += inv.soTien;
    else s.soChuaChi++;
    s.conNo += inv.conNo;
    s.soHoaDon++;
    s.invoices.push(inv);
  });
  return Array.from(map.values())
    .filter((r) => r.tongHoaDon !== 0)
    .sort((a, b) => b.conNo - a.conNo)
    .map((r) => {
      r.invoices.sort((a, b) => {
        if (a.daChi !== b.daChi) return a.daChi ? 1 : -1;
        return (b.ngayHD || "").localeCompare(a.ngayHD || "");
      });
      return r;
    });
}

// Gop cac giao dich chi CHUA co hoa don theo Ten doi ung (coi nhu 1 "NCC tam"
// -- co the trung/khong dung chinh xac ten NCC chinh thuc vi day la Ten doi
// ung tren sao ke, khong phai ten da chuan hoa qua file Hoa Don Dau Vao) --
// cung cach trinh bay voi summarizeByNcc de Luyen quen mat, nhung KHONG cong
// vao Con no (xem ghi chu o buildInvoiceDebt).
function summarizeMissingInvoiceByNcc(missingInvoiceChi) {
  const map = new Map();
  missingInvoiceChi.forEach((m) => {
    const key = m.company + "||" + normVN(m.tenNCC);
    if (!map.has(key)) {
      map.set(key, { tenNCC: m.tenNCC, company: m.company, companyLabel: m.companyLabel, tongSoTien: 0, soLan: 0, items: [] });
    }
    const s = map.get(key);
    s.tongSoTien += m.amount;
    s.soLan++;
    s.items.push(m);
  });
  return Array.from(map.values())
    .sort((a, b) => b.tongSoTien - a.tongSoTien)
    .map((r) => {
      r.items.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
      return r;
    });
}

router.get("/cong-no/ncc", (req, res) => {
  const store = load();
  const activeCompany = getCompany(req);
  let error = null;
  let invoicesAllCompanies = [];
  let invoices = [];
  let nccSummary = [];
  let missingInvoiceChi = [];
  let missingInvoiceSummary = [];
  let availableMonths = [];
  const thangFilter = (req.query.thang || "").trim();
  try {
    const built = buildInvoiceDebt(store);
    invoicesAllCompanies = built.invoices;
    // Chi Nhan, 2026-07-30: "khi tôi chọn Kh mới hiện của kh mới tk kh mới hóa
    // đơn kh mới thôi còn khi tôi chọn kh cũ thì hiện ra kh cũ á" -- loc theo
    // DUNG cong ty dang xem o topbar (giong Hoa Don Dau Vao/Chi Phi/Phap
    // danh), khong con gop chung 2 cong ty + dropdown loc thu cong nhu ban
    // dau nua.
    invoices = invoicesAllCompanies.filter((i) => i.company === activeCompany);

    // Danh sach cac thang co hoa don (YYYY-MM), moi nhat truoc
    const monthSet = new Set();
    invoices.forEach((i) => { const m = (i.ngayHD || "").slice(0,7); if (m) monthSet.add(m); });
    availableMonths = [...monthSet].sort().reverse();

    // Loc theo thang neu co
    if (thangFilter) invoices = invoices.filter((i) => (i.ngayHD || "").slice(0,7) === thangFilter);

    nccSummary = summarizeByNcc(invoices);
    missingInvoiceChi = built.missingInvoiceChi.filter((m) => m.company === activeCompany);
    missingInvoiceSummary = summarizeMissingInvoiceByNcc(missingInvoiceChi);

    // Enrich nccSummary voi maNCC tu danh_muc_ma_nha_cung_cap
    const danhMucKey = activeCompany === "kh_moi" ? "danh_muc_ma_nha_cung_cap_moi" : "danh_muc_ma_nha_cung_cap_cu";
    const danhMucSrc = store[danhMucKey] || store.danh_muc_ma_nha_cung_cap || [];
    // Index: normalized(ten) → {ma, ten}
    const nccDMIndex = new Map();
    danhMucSrc.filter((r) => r.ma && r.ten).forEach((r) => {
      nccDMIndex.set(normVN(r.ten), { ma: r.ma, ten: r.ten });
    });
    nccSummary.forEach((s) => {
      const norm = normVN(s.tenNCC || "");
      // Khớp chính xác trước
      if (nccDMIndex.has(norm)) {
        s.maNCC = nccDMIndex.get(norm).ma;
        return;
      }
      // Khớp mờ: tìm danh mục entry mà tên chứa ít nhất 2 từ chung có nghĩa
      const words = norm.split(/\s+/).filter((w) => w.length >= 3);
      let bestScore = 0, bestMa = "";
      nccDMIndex.forEach((v, k) => {
        const score = words.filter((w) => k.includes(w)).length;
        if (score >= 2 && score > bestScore) { bestScore = score; bestMa = v.ma; }
      });
      if (bestMa) s.maNCC = bestMa;
    });
  } catch (e) {
    error = e.message;
    console.error("Loi tinh cong no NCC:", e);
  }

  const daChiFilter = req.query.daChi || ""; // "", "1" = da chi, "0" = chua chi
  const nccFilter = (req.query.ncc || "").trim();
  const page = Math.max(1, parseInt(req.query.page, 10) || 1);
  const PAGE_SIZE = 30;

  let filtered = invoices;
  if (daChiFilter === "1") filtered = filtered.filter((i) => i.daChi);
  else if (daChiFilter === "0") filtered = filtered.filter((i) => !i.daChi);
  if (nccFilter) {
    const nf = normVN(nccFilter);
    filtered = filtered.filter((i) => normVN(i.tenNCC || "").includes(nf));
  }
  filtered = filtered.slice().sort((a, b) => {
    if (a.daChi !== b.daChi) return a.daChi ? 1 : -1; // chua chi len truoc
    return (b.ngayHD || "").localeCompare(a.ngayHD || ""); // hoa don moi nhat truoc
  });

  const totalMatching = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalMatching / PAGE_SIZE));
  const currentPage = Math.min(page, totalPages);
  const pageRows = filtered.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);

  const qs = [];
  if (thangFilter) qs.push("thang=" + encodeURIComponent(thangFilter));
  if (daChiFilter) qs.push("daChi=" + encodeURIComponent(daChiFilter));
  if (nccFilter) qs.push("ncc=" + encodeURIComponent(nccFilter));
  const baseQs = qs.join("&");

  const grandTotalHoaDon = invoices.reduce((s, i) => s + i.soTien, 0);
  const grandTotalDaChi = invoices.reduce((s, i) => s + (i.daChi ? i.soTien : 0), 0);
  const grandTotalConNo = invoices.reduce((s, i) => s + i.conNo, 0);
  const soHoaDonChuaChi = invoices.filter((i) => !i.daChi).length;
  const grandTotalMissingInvoice = missingInvoiceChi.reduce((s, m) => s + m.amount, 0);

  res.render("congno-ncc", {
    userName: req.session.userName,
    error,
    nccSummary,
    rows: pageRows,
    totalMatching,
    tongSoHoaDon: invoices.length,
    currentPage,
    totalPages,
    baseQs,
    thangFilter,
    availableMonths,
    daChiFilter,
    nccFilter,
    grandTotalHoaDon,
    grandTotalDaChi,
    grandTotalConNo,
    soHoaDonChuaChi,
    missingInvoiceSummary,
    grandTotalMissingInvoice,
    soLanMissingInvoice: missingInvoiceChi.length,
  });
});

module.exports = router;
