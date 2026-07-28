const XLSX = require("xlsx");

// Chi Nhan, 2026-07-28: bo cac ham "lam giau" du lieu cho Hoa Don Dau Vao --
// doi chieu voi 3 file/1 Google Sheet Chi Nhan cung cap de tu dien: Gian
// Hang + Tai khoan Co (1388/331), Ma doi tuong NCC, Ten hang hoa (mau Misa),
// va phan loai/Tai khoan No tu Dien giai. TAT CA deu la "goi y tu dong" --
// Chi Nhan tu xem/sua lai tung dong qua nut sua truc tiep tren bang, khong co
// gi bi khoa cung.

function removeDiacritics(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}

function normVN(s) {
  return removeDiacritics(s).toLowerCase().replace(/\s+/g, " ").trim();
}

// ---------- 1) Danh sach nha cung cap (upload) ----------
// Cot: STT, Ma nha cung cap, Ten nha cung cap, Dia chi, So tien no, Ma so
// thue/CCCD chu ho, ..., Chi nhanh (ten cong ty phap nhan dang xuat -- dung
// de biet danh sach nay thuoc KH Cu hay KH Moi khi Chi Nhan tai len).
function parseNccListWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let headerIdx = -1;
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = (grid[r] || []).map((c) => normVN(c));
    if (row.some((c) => c.includes("ma nha cung cap")) && row.some((c) => c.includes("ten nha cung cap"))) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx === -1) throw new Error('Không nhận diện được file: cần có cột "Mã nhà cung cấp" và "Tên nhà cung cấp".');
  const header = grid[headerIdx].map((c) => normVN(c));
  const col = {
    ma: header.findIndex((c) => c.includes("ma nha cung cap")),
    ten: header.findIndex((c) => c.includes("ten nha cung cap")),
    mst: header.findIndex((c) => c.includes("ma so thue")),
    // Luyen, 2026-07-28: dung == chinh xac (khong phai .includes) -- cot "Là
    // Tổng công ty/chi nhánh" (index truoc do) CUNG chua chuoi con "chi
    // nhanh", .includes() se khop nham cot do truoc cot "Chi nhánh" that.
    chiNhanh: header.findIndex((c) => c === "chi nhanh"),
  };
  const rows = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const ten = row[col.ten];
    if (!ten) continue;
    rows.push({
      maNCC: col.ma !== -1 ? String(row[col.ma] || "").trim() : "",
      tenNCC: String(ten).trim(),
      mst: col.mst !== -1 ? String(row[col.mst] || "").trim() : "",
      chiNhanh: col.chiNhanh !== -1 ? String(row[col.chiNhanh] || "").trim() : "",
    });
  }
  return { rows };
}

// ---------- 2) Danh sach hang hoa, dich vu (upload) ----------
// Cot: STT, Ma, Ten, Tinh chat (Hang hoa/Dich vu), Nhom VTHH, Don vi tinh
// chinh, So luong ton, Gia tri ton, TK Kho, TK Doanh thu.
function parseHangHoaWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let headerIdx = -1;
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = (grid[r] || []).map((c) => normVN(c));
    if (row.some((c) => c === "ma") && row.some((c) => c === "ten")) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx === -1) throw new Error('Không nhận diện được file: cần có cột "Mã" và "Tên".');
  const header = grid[headerIdx].map((c) => normVN(c));
  const col = {
    ma: header.findIndex((c) => c === "ma"),
    ten: header.findIndex((c) => c === "ten"),
    tinhChat: header.findIndex((c) => c.includes("tinh chat")),
    donViTinh: header.findIndex((c) => c.includes("don vi tinh")),
  };
  const rows = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const ten = row[col.ten];
    if (!ten || String(ten).trim().length < 3) continue; // bo qua ten qua ngan (de tranh khop nham hang loat)
    rows.push({
      ma: col.ma !== -1 ? String(row[col.ma] || "").trim() : "",
      ten: String(ten).trim(),
      tinhChat: col.tinhChat !== -1 ? String(row[col.tinhChat] || "").trim() : "",
      donViTinh: col.donViTinh !== -1 ? String(row[col.donViTinh] || "").trim() : "",
    });
  }
  // Uu tien khop TEN DAI truoc (cu the hon) khi nhieu ten cung la substring cua Dien giai.
  rows.sort((a, b) => b.ten.length - a.ten.length);
  return { rows };
}

// ---------- 3) Danh sach cac gian (Google Sheet, nhieu tab) ----------
// Cot thuc te (theo anh chup man hinh Chi Nhan gui 2026-07-28, sheet "HÀ NỘI
// MTD"): STT, Pháp nhân (KH mới/KH cũ), Khu Vực, Dịch vụ, Mã Điểm Nội Bộ, Mã
// Điểm Thuê, Địa điểm, Tên khách hàng, MST khách hàng, Hình Thức Hợp Tác
// (CSE/Tiền thuê), Ghi Chú, Link hợp đồng. "Tên khách hàng" o day chinh la
// BEN CHO THUE (vd Sun World, Vincom...) -- trung voi "Tên người bán"/NCC
// tren hoa don dau vao khi do la hoa don thue mat bang/tien dien cua gian.
function parseGianSheetOneTab(grid, sheetLabel) {
  let headerIdx = -1;
  for (let r = 0; r < Math.min(grid.length, 10); r++) {
    const row = (grid[r] || []).map((c) => normVN(c));
    if (row.some((c) => c.includes("ma diem noi bo")) && row.some((c) => c.includes("ten khach hang"))) {
      headerIdx = r;
      break;
    }
  }
  if (headerIdx === -1) return [];
  const header = grid[headerIdx].map((c) => normVN(c));
  const col = {
    phapNhan: header.findIndex((c) => c.includes("phap nhan")),
    maDiemNoiBo: header.findIndex((c) => c.includes("ma diem noi bo")),
    maDiemThue: header.findIndex((c) => c.includes("ma diem thue")),
    tenKhachHang: header.findIndex((c) => c.includes("ten khach hang")),
    mstKhachHang: header.findIndex((c) => c.includes("mst khach hang")),
    hinhThucHopTac: header.findIndex((c) => c.includes("hinh thuc hop tac")),
  };
  const rows = [];
  for (let r = headerIdx + 1; r < grid.length; r++) {
    const row = grid[r] || [];
    const tenKhachHang = row[col.tenKhachHang];
    if (!tenKhachHang) continue;
    const phapNhanRaw = normVN(col.phapNhan !== -1 ? row[col.phapNhan] : "");
    rows.push({
      sheetLabel,
      congTy: phapNhanRaw.includes("moi") ? "kh_moi" : phapNhanRaw.includes("cu") ? "kh_cu" : "",
      gianHang: col.maDiemNoiBo !== -1 ? String(row[col.maDiemNoiBo] || "").trim() : "",
      maDiemThue: col.maDiemThue !== -1 ? String(row[col.maDiemThue] || "").trim() : "",
      tenKhachHang: String(tenKhachHang).trim(),
      mstKhachHang: col.mstKhachHang !== -1 ? String(row[col.mstKhachHang] || "").trim() : "",
      hinhThucHopTac: col.hinhThucHopTac !== -1 ? String(row[col.hinhThucHopTac] || "").trim() : "",
    });
  }
  return rows;
}

function parseGianSheetWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "buffer" });
  let rows = [];
  const sheetsRead = [];
  wb.SheetNames.forEach((name) => {
    const ws = wb.Sheets[name];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const tabRows = parseGianSheetOneTab(grid, name);
    if (tabRows.length > 0) {
      rows = rows.concat(tabRows);
      sheetsRead.push(`${name}: ${tabRows.length}`);
    }
  });
  if (rows.length === 0) {
    throw new Error(
      `Không đọc được sheet nào có đủ cột "Mã Điểm Nội Bộ"/"Tên khách hàng" (đã thấy các tab: ${wb.SheetNames.join(", ")}).`
    );
  }
  return { rows, sheetsRead };
}

// ---------- Ghep 1 ten/MST NCC voi 1 danh sach (nha cung cap hoac cac gian) ----------
// Uu tien khop theo MST (chac chan nhat, MST khong trung), khong co/khong
// khop thi thu khop theo ten (chuan hoa, so sanh CHUA het) -- CHUA khop chinh
// xac (bao gom mot phan) thi coi la khong khop, tra ve null (Chi Nhan tu dien
// tay, KHONG doan dai de tranh gan nham).
function matchByMstOrName(mst, ten, list, mstField, nameField) {
  const mstNorm = String(mst || "").trim();
  if (mstNorm) {
    const byMst = list.find((r) => String(r[mstField] || "").trim() === mstNorm);
    if (byMst) return byMst;
  }
  const tenNorm = normVN(ten);
  if (!tenNorm) return null;
  const byNameExact = list.find((r) => normVN(r[nameField]) === tenNorm);
  if (byNameExact) return byNameExact;
  // Khop 1 chieu: ten tren hoa don la chuoi con cua ten trong danh sach hoac
  // nguoc lai (nhieu file MISA rut gon/them "CHI NHANH..." khac nhau).
  const candidates = list.filter((r) => {
    const rn = normVN(r[nameField]);
    return rn && (rn.includes(tenNorm) || tenNorm.includes(rn));
  });
  if (candidates.length === 1) return candidates[0];
  return null;
}

// Chi Nhan, 2026-07-28: "check trên unc ra tên gian" -- khi khong khop duoc
// Gian Hang qua Google Sheet (theo Ten NCC/MST), thu do NOI DUNG cua 1 dong
// UNC (da khop truoc do theo Ten NCC + So tien, xem matchUncForPayment ben
// chiphiReconcile.js) xem co chua Ma Diem Noi Bo hoac Ma Diem Thue cua gian
// nao trong danh sach gian (cung cong ty) khong -- CHI dien khi khop DUY
// NHAT 1 gian (theo ca 2 ma), tranh doan nham nhu cac ham match khac o day.
function matchGianViaUncContent(noiDungUnc, gianForCompany) {
  const noiDungNorm = normVN(noiDungUnc);
  if (!noiDungNorm || noiDungNorm.length < 3) return null;
  const matched = new Set();
  for (const g of gianForCompany || []) {
    const codes = [g.gianHang, g.maDiemThue].map((c) => String(c || "").trim()).filter((c) => c.length >= 3);
    const hit = codes.some((c) => noiDungNorm.includes(normVN(c)));
    if (hit) matched.add(g);
  }
  const list = Array.from(matched);
  return list.length === 1 ? list[0] : null;
}

// ---------- Phan loai tu Dien giai + so tien ----------
// Chi Nhan, 2026-07-28: "nếu nó số lượng đếm được mua vào bán ra thì phân
// làm hàng hóa 156 cái nào cccd thì phân vào 153 cái nào là tài sản thì đưa
// vào 242 cái nào là dịch vụ thì đưa vào 154 hay các chi phí mua văn phòng
// phẩm hay các xăng dầu hay phí ngân hàng thì đưa vào đầu 642". Day la GOI Y
// tot nhat co the tu tu khoa -- KHONG the chinh xac 100% (ranh gioi Tai
// san/CCDC/Dich vu/642 phu thuoc xet doan ke toan thuc te), Chi Nhan tu sua
// lai tung dong qua nut sua tren bang.
const TAI_SAN_KEYWORDS = [
  "may lanh", "dieu hoa", "camera", "may tinh", "laptop", "tu lanh",
  "may photocopy", "may in", "he thong am thanh", "thiet bi", "may moc",
  "xe oto", "o to", "macbook", "ipad", "iphone", "apple watch",
];
const CCDC_KEYWORDS = ["ccdc", "cong cu dung cu", "dung cu", "ban ghe", "airpods", "tai nghe"];
const CHI_PHI_642_KEYWORDS = [
  "xang dau", "xang", "dau nhot", "van phong pham", "vpp", "phi ngan hang",
  "phi chuyen khoan", "phi quan ly tai khoan", "phi thuong nien the", "in an",
  "photo", "cuoc phi", "lai vay",
];
const DICH_VU_KEYWORDS = [
  "tien dien", "tien nuoc", "phi thue", "thue gian", "thue mat bang",
  "phi quan ly", "bao hiem", "internet", "cuoc vien thong", "tu van",
  "bao tri", "bao duong", "van chuyen", "dao tao", "kiem toan", "quang cao",
  "hoa hong", "ve ", "dich vu", "phi dich vu",
];
const TAI_SAN_THRESHOLD = 30000000; // quy dinh TSCD: gia tri >= 30 trieu VA thoi gian su dung > 1 nam

function classifyPhanLoai(dienGiai, soTienTruocThue) {
  const t = normVN(dienGiai);
  const amount = Number(soTienTruocThue) || 0;
  const isTaiSanCandidate = TAI_SAN_KEYWORDS.some((k) => t.includes(k));
  if (isTaiSanCandidate && amount >= TAI_SAN_THRESHOLD) {
    return { phanLoai: "Tài sản", taiKhoanNo: "242" };
  }
  if (CCDC_KEYWORDS.some((k) => t.includes(k)) || isTaiSanCandidate) {
    return { phanLoai: "CCDC", taiKhoanNo: "153" };
  }
  if (CHI_PHI_642_KEYWORDS.some((k) => t.includes(k))) {
    return { phanLoai: "Chi phí QLDN", taiKhoanNo: "642" };
  }
  if (DICH_VU_KEYWORDS.some((k) => t.includes(k))) {
    return { phanLoai: "Dịch vụ", taiKhoanNo: "154" };
  }
  return { phanLoai: "", taiKhoanNo: "" };
}

function matchTenHangHoa(dienGiai, hangHoaList) {
  const t = normVN(dienGiai);
  if (!t) return "";
  const found = hangHoaList.find((h) => h.ten && t.includes(normVN(h.ten)));
  return found ? found.ten : "";
}

module.exports = {
  normVN,
  parseNccListWorkbook,
  parseHangHoaWorkbook,
  parseGianSheetWorkbook,
  matchByMstOrName,
  matchGianViaUncContent,
  classifyPhanLoai,
  matchTenHangHoa,
};
