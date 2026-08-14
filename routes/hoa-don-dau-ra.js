// Luyen, 2026-08-12: "thêm cho tôi thêm 1 trang hóa đơn đầu ra nữa á"
// Trang quan ly hoa don dau ra (hoa don ban hang, xuat cho khach).
// Luu trong store.hoa_don_dau_ra -- mang cac doi tuong {id, congTy, ...}.
const express = require("express");
const router = express.Router();
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { COMPANIES, getCompany } = require("../utils/companies");

function ensureShape(store) {
  if (!store.hoa_don_dau_ra) store.hoa_don_dau_ra = [];
}

// GET /hoa-don-dau-ra
router.get("/hoa-don-dau-ra", requireLogin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);

  // Bo loc
  const selectedMonth = req.query.thang || "";
  const selectedLoai = req.query.loai || ""; // "khach-le", "hop-dong", ...
  const searchQ = (req.query.q || "").trim().toLowerCase();

  let rows = store.hoa_don_dau_ra.filter(
    (r) => (r.congTy || "kh_cu") === activeCompany
  );

  // Loc thang
  if (selectedMonth) {
    rows = rows.filter((r) => (r.ngayHD || "").startsWith(selectedMonth));
  }
  // Loc loai
  if (selectedLoai) {
    rows = rows.filter((r) => (r.loaiHD || "") === selectedLoai);
  }
  // Tim kiem tu khoa
  if (searchQ) {
    rows = rows.filter((r) =>
      [r.soHD, r.tenKhachHang, r.dienGiai, r.maKH].some(
        (v) => v && String(v).toLowerCase().includes(searchQ)
      )
    );
  }

  // Sort: moi nhat len tren
  rows = [...rows].sort((a, b) => (b.ngayHD || "").localeCompare(a.ngayHD || ""));

  // Tong tien
  const tongTien = rows.reduce((s, r) => s + (Number(r.soTien) || 0), 0);
  const tongTienVAT = rows.reduce((s, r) => s + (Number(r.soTienVAT) || 0), 0);

  // Danh sach thang co du lieu (de dropdown bo loc)
  const allRows = store.hoa_don_dau_ra.filter(
    (r) => (r.congTy || "kh_cu") === activeCompany
  );
  const monthSet = new Set(allRows.map((r) => (r.ngayHD || "").slice(0, 7)).filter(Boolean));
  const months = [...monthSet].sort().reverse();

  res.render("hoa-don-dau-ra", {
    COMPANIES, activeCompany,
    userName: req.session.userName, isAdmin: req.session.isAdmin,
    userRole: req.session.userRole,
    rows, months, tongTien, tongTienVAT,
    selectedMonth, selectedLoai, searchQ: req.query.q || "",
    currentPath: req.path,
    success: req.query.success || "",
    error: req.query.error || "",
  });
});

// GET /hoa-don-dau-ra/export-misa -- xuat Excel dinh dang MISA hoa don ban ra
router.get("/hoa-don-dau-ra/export-misa", requireLogin, (req, res) => {
  const XLSX = require("xlsx");
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const tuNgay   = (req.query.tuNgay   || "").trim(); // "YYYY-MM-DD"
  const denNgay  = (req.query.denNgay  || "").trim(); // "YYYY-MM-DD"
  const selectedMonth = (req.query.thang || "").trim(); // "YYYY-MM" (fallback cu)

  // Helper: parse ngayHD -> { dt, dmyStr, ym }
  function parseNgay(s) {
    if (!s) return { dt: null, dmyStr: "", ym: "" };
    s = String(s).trim();
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
      const [d, m, y] = s.split("/");
      return { dt: new Date(+y, +m - 1, +d), dmyStr: s, ym: `${y}-${m}` };
    }
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      const dt = new Date(s);
      const dd = String(dt.getDate()).padStart(2, "0");
      const mm = String(dt.getMonth() + 1).padStart(2, "0");
      const yy = dt.getFullYear();
      return { dt, dmyStr: `${dd}/${mm}/${yy}`, ym: `${yy}-${mm}` };
    }
    return { dt: null, dmyStr: s, ym: "" };
  }

  // Build tap so HD tu doi soat VNPay/Payoo (manual_matches)
  const mm = store.vnpay_khmoi_manual_matches || {};
  const vnpayHDs = new Set();
  const payooHDs  = new Set();
  ["offline","online"].forEach(ch => {
    Object.values(mm[ch] || {}).forEach(entry => {
      (entry.invoiceNumbers || []).forEach(n => vnpayHDs.add(String(parseInt(n) || 0)));
    });
  });
  Object.values(mm.payoo || {}).forEach(entry => {
    (entry.invoiceNumbers || []).forEach(n => payooHDs.add(String(parseInt(n) || 0)));
  });

  // Luyen, 2026-08-13: map soHD -> kenh thu ho tu file MTT 11.08.xlsx
  // (cot "Dich vu thu ho", sheet "ke ds xuat HD MTT - 705")
  // vnpay/momo/payoo -- KL neu khong co trong bang nay
  const THU_HO_MAP = {
  "11585": "vnpay",
  "11586": "vnpay",
  "11587": "momo",
  "11588": "momo",
  "11589": "momo",
  "11590": "momo",
  "11730": "payoo",
  "11731": "payoo",
  "11732": "momo",
  "11733": "momo",
  "11734": "momo",
  "11735": "momo",
  "11736": "payoo",
  "11737": "payoo",
  "11738": "payoo",
  "11881": "vnpay",
  "11882": "vnpay",
  "11883": "vnpay",
  "11884": "vnpay",
  "11885": "vnpay",
  "11886": "vnpay",
  "11887": "vnpay",
  "11888": "vnpay",
  "11889": "momo",
  "11890": "momo",
  "11891": "momo",
  "11892": "momo",
  "12022": "payoo",
  "12023": "payoo",
  "12024": "payoo",
  "12025": "vnpay",
  "12026": "vnpay",
  "12027": "vnpay",
  "12028": "momo",
  "12029": "momo",
  "12030": "momo",
  "12031": "momo",
  "12171": "payoo",
  "12172": "payoo",
  "12173": "payoo",
  "12174": "payoo",
  "12175": "payoo",
  "12176": "payoo",
  "12177": "vnpay",
  "12178": "vnpay",
  "12179": "vnpay",
  "12180": "momo",
  "12181": "momo",
  "12182": "momo",
  "12183": "momo",
  "12319": "payoo",
  "12320": "payoo",
  "12321": "payoo",
  "12322": "payoo",
  "12323": "vnpay",
  "12324": "vnpay",
  "12325": "vnpay",
  "12326": "momo",
  "12327": "momo",
  "12328": "momo",
  "12329": "momo",
  "12465": "payoo",
  "12466": "payoo",
  "12467": "payoo",
  "12468": "payoo",
  "12469": "payoo",
  "12470": "payoo",
  "12472": "vnpay",
  "12473": "vnpay",
  "12474": "momo",
  "12475": "momo",
  "12476": "momo",
  "12477": "momo",
  "12616": "payoo",
  "12617": "payoo",
  "12618": "payoo",
  "12619": "payoo",
  "12620": "payoo",
  "12621": "payoo",
  "12622": "payoo",
  "12624": "vnpay",
  "12625": "vnpay",
  "12626": "momo",
  "12627": "momo",
  "12628": "momo",
  "12629": "momo",
  "12777": "payoo",
  "12778": "payoo",
  "12779": "payoo",
  "12780": "payoo",
  "12781": "payoo",
  "12782": "payoo",
  "12783": "payoo",
  "12785": "vnpay",
  "12786": "vnpay",
  "12787": "momo",
  "12788": "momo",
  "12789": "momo",
  "12790": "momo",
  "12934": "payoo",
  "12935": "payoo",
  "12936": "payoo",
  "12937": "payoo",
  "12938": "payoo",
  "12939": "payoo"
  };

  function mapKH(row, soIntStr) {
    // Uu tien 1: MTT thu ho map (chinh xac nhat, trich tu file xuat HD)
    const thuHo = THU_HO_MAP[soIntStr];
    if (thuHo === "vnpay") return "VN PAY0102182292";
    if (thuHo === "payoo") return "DONGVIET0305458683";
    if (thuHo === "momo")  return "TRỰC TUYẾN0305289153";
    // Uu tien 2: doi soat VNPay/Payoo (manual_matches)
    if (vnpayHDs.has(soIntStr)) return "VN PAY0102182292";
    if (payooHDs.has(soIntStr))  return "DONGVIET0305458683";
    // Fallback: ghiChu keywords
    const g = (row.ghiChu || "").toLowerCase();
    const t = (row.tenKhachHang || "").toUpperCase();
    if (g.includes("momo")) return "TRỰC TUYẾN0305289153";
    if (g.includes("vnpay") || g.includes("zalo")) return "VN PAY0102182292";
    if (g.includes("payoo") || g.includes("dong viet")) return "DONGVIET0305458683";
    if (t.includes("YOKIDS")) return "YOKIDS0801365316";
    if (t.includes("KIWOOZA")) return "KIWOOZA0315850120";
    if (t.includes("HOA SEN")) return "HOA SEN3700381324";
    return "KL";
  }

// Luyen, 2026-08-13: nguong doi Ky hieu HD THAT (doi chieu tu file goc
      // "baocaochitiet.xlsx" -- cot Ky hieu chuyen tu "1C26TYY" sang "1C26MKH"
      // dung tai hoa don so 11585, KHONG phai 5000 nhu code cu (sai, gay xuat
      // nham Ky hieu cho hang nghin hoa don so 5000-11584). Trung voi moc 11585
      // da dung san trong THU_HO_MAP o tren.
  function getKyHieu(soHD) {
    const n = parseInt(String(soHD).replace(/^0+/, "")) || 0;
    return n < 11585 ? "1C26TYY" : "1C26MKH";
  }

  let rows = store.hoa_don_dau_ra.filter(r => (r.congTy || "kh_cu") === activeCompany);

  // Loc theo khoang ngay (uu tien) hoac theo thang (fallback)
  const dtTu  = tuNgay  ? new Date(tuNgay)  : null;
  const dtDen = denNgay ? new Date(denNgay) : null;
  if (dtTu || dtDen) {
    rows = rows.filter(r => {
      const { dt } = parseNgay(r.ngayHD);
      if (!dt) return false;
      const d = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
      if (dtTu  && d < new Date(dtTu.getFullYear(),  dtTu.getMonth(),  dtTu.getDate()))  return false;
      if (dtDen && d > new Date(dtDen.getFullYear(), dtDen.getMonth(), dtDen.getDate())) return false;
      return true;
    });
  } else if (selectedMonth) {
    rows = rows.filter(r => parseNgay(r.ngayHD).ym === selectedMonth);
  }

  rows = [...rows].sort((a, b) => {
    const na = parseInt(String(a.soHD || "").replace(/^0+/, "")) || 0;
    const nb = parseInt(String(b.soHD || "").replace(/^0+/, "")) || 0;
    return na - nb;
  });

  // 8 dong header theo mau MISA
  const MISA_HEADERS = [
    ["FILE MẪU CHỨNG TỪ BÁN HÀNG TRONG NƯỚC ĐỂ NHẬP VÀO PHẦN MỀM AMIS ACCOUNTING"],
    ["Hướng dẫn:"],
    ["- Điền dữ liệu vào các cột tương ứng trên file này"],
    ["- Các cột có dấu (*) là những cột bắt buộc"],
    ["- Nếu muốn nhập nhiều thông tin hơn người dùng có thể tải mẫu đầy đủ/hoặc tự thêm cột trên mẫu cơ bản"],
    ["- Các dòng dữ liệu phía dưới chỉ là ví dụ minh họa"],
    [null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,"Chi tiết hàng tiền",null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,"Chi tiết giá vốn"],
    ["Phương thức thanh toán","Kiêm phiếu xuất kho","Lập kèm hóa đơn","Đã lập hóa đơn","Ngày hạch toán (*)","Ngày chứng từ (*)","Số chứng từ (*)","Số phiếu xuất","Mẫu số HĐ","Ký hiệu HĐ","Số hóa đơn","Ngày hóa đơn","Khách hàng","Địa chỉ","Mã số thuế","Người nộp","Nộp vào TK","Diễn giải/Lý do nộp","Lý do xuất","Mã nhân viên bán hàng","Số chứng từ kèm theo (Phiếu thu)","Số chứng từ kèm theo (Phiếu xuất)","Hạn thanh toán","Mã hàng (*)","Tên hàng","Là dòng ghi chú","Hàng khuyến mại","Chiết khấu thương mại","TK Tiền/Chi phí/Nợ (*)","TK Doanh thu/Có (*)","ĐVT","Số lượng","Đơn giá","Thành tiền","Tỷ lệ CK (%)","Tiền chiết khấu","TK chiết khấu","% thuế GTGT","Tiền thuế GTGT","TK thuế GTGT","HH không TH trên tờ khai thuế GTGT","Mã khoản mục chi phí","Mã đơn vị","Mã đối tượng THCP","Mã công trình","Số đơn đặt hàng","Số hợp đồng bán","Mã thống kê","CP không hợp lý","Mã kho","Vị trí","TK giá vốn","TK Kho","Đơn giá vốn","Tiền vốn","Hàng hóa giữ hộ/bán hộ","Tên khách hàng"],
  ];

  const dataRows = rows.map(row => {
    const { dmyStr } = parseNgay(row.ngayHD);
    const soInt = parseInt(String(row.soHD || "").replace(/^0+/, "")) || 0;
    const kyHieu = getKyHieu(soInt);
    const ngay = dmyStr;
    let mm = "08", yy = "26";
    if (dmyStr && dmyStr.length === 10) { mm = dmyStr.slice(3,5); yy = dmyStr.slice(8,10); }
    const soCT = `BH${mm}-${String(soInt).padStart(6,"0")}/${yy}`;
    const maKH = mapKH(row, String(soInt));
    const soTien = Number(row.soTien) || 0;
    const soVAT  = Number(row.soTienVAT) || 0;
    const thueVAT = String(row.thueVAT || "8").replace("%","").trim();
    const maCT = row.maKH || "";
    const tenKH = (row.tenKhachHang && !["Bán cho người tiêu dùng",""].includes(row.tenKhachHang))
      ? row.tenKhachHang : "";

    const r = new Array(57).fill(null);
    r[0]="Chưa thu tiền"; r[1]="Không"; r[2]="Có"; r[3]="Đã lập";
    r[4]=ngay; r[5]=ngay; r[6]=soCT;
    r[9]=kyHieu; r[10]=soInt; r[11]=ngay; r[12]=maKH;
    r[17]=`Dịch vụ vui chơi giải trí theo HĐ ${soInt} ký hiệu ${kyHieu}`;
    r[23]="KVC"; r[24]="Dịch vụ vui chơi giải trí";
    r[28]=131; r[29]=5113; r[30]="Kỳ"; r[31]=1;
    r[32]=soTien; r[33]=soTien;
    r[37]=thueVAT; r[38]=soVAT; r[39]=33311;
    r[44]=maCT;
    if (tenKH) r[56]=tenKH;
    return r;
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([...MISA_HEADERS, ...dataRows]);
  let sheetName = "Ban hang trong nuoc";
  if (tuNgay && denNgay) sheetName = `${tuNgay.slice(5).replace("-","")}--${denNgay.slice(5).replace("-","")}`;
  else if (selectedMonth) sheetName = selectedMonth.replace("-","");
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
  let fname = "HoaDon_BanRa_MISA_TatCa.xlsx";
  if (tuNgay && denNgay) fname = `HoaDon_BanRa_MISA_${tuNgay}_${denNgay}.xlsx`;
  else if (selectedMonth) fname = `HoaDon_BanRa_MISA_${selectedMonth.replace("-","_")}.xlsx`;

  res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buf);
});

// POST /hoa-don-dau-ra/them -- them moi 1 dong
router.post("/hoa-don-dau-ra/them", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);

  const row = {
    id: nextId(store),
    congTy: activeCompany,
    ngayHD: (req.body.ngayHD || "").trim(),
    soHD: (req.body.soHD || "").trim(),
    maKH: (req.body.maKH || "").trim(),
    tenKhachHang: (req.body.tenKhachHang || "").trim(),
    loaiHD: (req.body.loaiHD || "").trim(),
    dienGiai: (req.body.dienGiai || "").trim(),
    soTien: Number(String(req.body.soTien || "0").replace(/[^\d.-]/g, "")) || 0,
    soTienVAT: Number(String(req.body.soTienVAT || "0").replace(/[^\d.-]/g, "")) || 0,
    thueVAT: (req.body.thueVAT || "").trim(),
    ghiChu: (req.body.ghiChu || "").trim(),
    createdAt: new Date().toISOString(),
  };

  store.hoa_don_dau_ra.push(row);
  save(store);
  res.redirect("/hoa-don-dau-ra?success=Đã thêm hóa đơn " + encodeURIComponent(row.soHD || String(row.id)));
});

// POST /hoa-don-dau-ra/sua/:id -- cap nhat 1 dong
router.post("/hoa-don-dau-ra/sua/:id", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const id = Number(req.params.id);
  const row = store.hoa_don_dau_ra.find((r) => r.id === id);
  if (!row) return res.redirect("/hoa-don-dau-ra?error=Không tìm thấy hóa đơn id=" + id);

  row.ngayHD      = (req.body.ngayHD || "").trim();
  row.soHD        = (req.body.soHD || "").trim();
  row.maKH        = (req.body.maKH || "").trim();
  row.tenKhachHang = (req.body.tenKhachHang || "").trim();
  row.loaiHD      = (req.body.loaiHD || "").trim();
  row.dienGiai    = (req.body.dienGiai || "").trim();
  row.soTien      = Number(String(req.body.soTien || "0").replace(/[^\d.-]/g, "")) || 0;
  row.soTienVAT   = Number(String(req.body.soTienVAT || "0").replace(/[^\d.-]/g, "")) || 0;
  row.thueVAT     = (req.body.thueVAT || "").trim();
  row.ghiChu      = (req.body.ghiChu || "").trim();
  row.updatedAt   = new Date().toISOString();

  save(store);
  res.redirect("/hoa-don-dau-ra?success=Đã cập nhật hóa đơn " + encodeURIComponent(row.soHD || String(row.id)));
});

// POST /hoa-don-dau-ra/xoa/:id
router.post("/hoa-don-dau-ra/xoa/:id", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const id = Number(req.params.id);
  const idx = store.hoa_don_dau_ra.findIndex((r) => r.id === id);
  if (idx === -1) return res.redirect("/hoa-don-dau-ra?error=Không tìm thấy hóa đơn id=" + id);
  const removed = store.hoa_don_dau_ra.splice(idx, 1)[0];
  save(store);
  res.redirect("/hoa-don-dau-ra?success=Đã xóa hóa đơn " + encodeURIComponent(removed.soHD || String(removed.id)));
});

module.exports = router;
