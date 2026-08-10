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
    // Chi Nhan, 2026-07-29: "tài khoản là 156 đi ... tôi có tải bảng hệ thống
    // tài khoản chia theo gian cho tôi rồi ấy" -- file "Danh sách hàng hóa,
    // dịch vụ" (MISA) THUONG co san 2 cot nay (da ghi chu tu 2026-07-28 nhung
    // chua doc): "TK Kho" (tai khoan No khi nhap kho -- 156/152/153 tuy loai
    // hang) va "TK Doanh thu". Doc THEM 2 cot nay, khong bat buoc phai co
    // (khong throw neu thieu) -- dung lam TAI KHOAN NO DUNG NHAT cho tung mat
    // hang cu the (uu tien HON keyword doan trong classifyPhanLoai) khi khop
    // duoc dong hang hoa nao qua matchTenHangHoa/matchHangHoaRecord.
    tkKho: header.findIndex((c) => c.includes("tk kho")),
    tkDoanhThu: header.findIndex((c) => c.includes("tk doanh thu")),
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
      tkKho: col.tkKho !== -1 ? String(row[col.tkKho] || "").trim() : "",
      tkDoanhThu: col.tkDoanhThu !== -1 ? String(row[col.tkDoanhThu] || "").trim() : "",
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

// Chi Nhan, 2026-07-29: "đi tìm từ 3 unc gg sheet á lấy ra gian cho tôi" --
// 3 Google Sheet "UNC" (ĐI ỦY NHIỆM CHI KVC + MTĐ MN / Tạo lệnh UNC KVC MB /
// TẠO LỆNH UNC MTĐ MB) DA duoc dong bo san vao store.chi_phi (trang "Chi
// Phí", routes/chi-phi.js) voi cac cot gian/ncc/soHoaDon/soTien day du (1735/
// 2119 dong kh_moi da co san "gian" luc kiem tra 2026-07-29) -- day moi la
// nguon UNC that Chi Nhan dang nhac toi (KHAC voi chi_phi_unc_list, la file
// UNC rieng Chi Nhan tu tai len o trang Doi soat Chi phi, thuong con trong).
// Uu tien khop theo SO HOA DON (chinh xac nhat, vi cot soHoaDon co san tren
// ca 2 ben), fallback theo Ten NCC + So tien (nhu matchUncForPayment). CHI
// tra ve khi khop DUY NHAT 1 gian, tranh doan nham.
function matchGianViaChiPhiLedger(row, chiPhiForCompany) {
  if (!chiPhiForCompany || chiPhiForCompany.length === 0) return null;
  const soHD = String(row.soHoaDon || "").trim();
  if (soHD) {
    const bySoHD = chiPhiForCompany.filter((c) => String(c.soHoaDon || "").trim() === soHD);
    const distinctGian = new Set(bySoHD.map((c) => String(c.gian || "").trim()).filter(Boolean));
    if (distinctGian.size === 1) return Array.from(distinctGian)[0];
  }
  const tenNorm = normVN(row.tenNCC);
  const amount = row.soTien || 0;
  if (tenNorm && amount) {
    let candidates = chiPhiForCompany.filter(
      (c) => Math.abs((c.soTien || 0) - amount) < 1000 && normVN(c.ncc) === tenNorm
    );
    if (candidates.length === 0) {
      candidates = chiPhiForCompany.filter((c) => {
        if (Math.abs((c.soTien || 0) - amount) >= 1000) return false;
        const cn = normVN(c.ncc);
        return cn && (cn.includes(tenNorm) || tenNorm.includes(cn));
      });
    }
    const distinctGian = new Set(candidates.map((c) => String(c.gian || "").trim()).filter(Boolean));
    if (distinctGian.size === 1) return Array.from(distinctGian)[0];
  }
  return null;
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
  "photo", "cuoc phi", "lai vay", "sim", "cuoc dien thoai", "thue bao",
  "vien thong", "internet", "phi thuong nien", "phi duy tri",
  // Chi Nhan, 2026-07-29: them tu du lieu that (kiem tra 2026-07-29) -- phi
  // giao dich phan mem/vi dien tu, phi marketing: KHONG tao doanh thu truc
  // tiep tu vui choi giai tri, dung 1 y voi "Khác"/642 Chi Nhan yeu cau.
  "phi giao dich", "phan mem", "phi marketing", "quang cao",
];
// Chi Nhan, 2026-07-29: "tiền điện tiền nước đưa vào phí lưu trú nhá" -- tach
// rieng tien dien/nuoc ra khoi nhom Dich vu chung (truoc do gom chung voi
// thue mat bang/phi quan ly...), van cung TK No 154 (chi phi truc tiep de
// duy tri gian dang thue) nhung PHAN LOAI ten rieng cho de doi soat/xem xet.
const PHI_LUU_TRU_KEYWORDS = ["tien dien", "tien nuoc"];
const DICH_VU_KEYWORDS = [
  "phi thue", "thue gian", "thue mat bang", "phi quan ly",
  "bao hiem", "tu van", "bao tri", "bao duong", "van chuyen",
  "dao tao", "kiem toan", "hoa hong", "ve ", "dich vu", "phi dich vu",
  // Luyen, 2026-08-10: cong ty vui choi giai tri KHONG gia cong/xay dung --
  // bat ky dien giai co cac tu khoa nay luon la DICH VU (TK 154), KHONG PHAI
  // hang hoa ban ra (du ten vat lieu xuat hien trong dien giai, vd "thi cong
  // decal formex" = dich vu thi cong, khong phai ban formex cho khach).
  "thi cong", "lap dat", "sua chua", "gia cong", "nhan cong",
  "thiet ke", "gia lap", "boc ep", "dan decal", "cat decal",
  "cat be tong", "son ", "han ", "khoan ", "dien nuoc",
];
// Chi Nhan, 2026-07-29: "các nvl phải qua chế biến mới bỏ vào 154 nhá còn kem
// hay cái nào đếm được bán liền không qua chế biến thì đưa vô 156 nhá" --
// GOI Y tu khoa cho truong hop KHONG khop duoc voi danh muc hang hoa (uploaded
// hang-hoa list, xem matchHangHoaRecord) -- vd hang dong lanh/nguyen lieu tuoi
// song can nau/chien/hap truoc khi ban (154, "dang che" -- chua thanh pham),
// khac voi hang dong goi/che bien san co the ban thang cho khach (156).
const NVL_CHE_BIEN_KEYWORDS = [
  "vien hai san", "tom vien", "ca vien", "muc vien", "muc xoan", "cha ca",
  "cha hai san", "dau hu ca", "tam bot", "hai san dong lanh", "thit song",
  "thit dong lanh", "ca dong lanh", "rau hon hop", "rau cu dong lanh",
  "nguyen lieu", "gia vi", "bot chien gion", "khoai tay dong lanh",
  // Chi Nhan, 2026-07-29: them tu du lieu that -- xot/sot dung de che bien
  // mon an (khong ban rieng cho khach), tinh la NVL nhu gia vi.
  "xot ", "sot ",
];
// Chi Nhan, 2026-07-29: mo rong tu du lieu that (kiem tra 2026-07-29, cac
// mat hang nay KHONG khop duoc voi danh muc hang hoa Chi Nhan da tai truoc
// do nen truoc day roi ve "Khác"/642 sai -- nuoc uong/banh keo/snack DA co
// ban tai cac gian FARM PT-... la hang hoa BAN THANG cho khach, khong lien
// quan che bien) -- day CHI la GOI Y bo sung, chinh xac nhat van la khop
// dung danh muc hang hoa Chi Nhan tai len (xem matchHangHoaRecord).
const HANG_HOA_BAN_LIEN_KEYWORDS = [
  "kem", "banh keo", "banh ", "keo ", "do uong", "nuoc ngot", "nuoc suoi",
  "nuoc uong", "nuoc ep", "sinh to", "sting", "coca", "pepsi", "tra xanh",
  "nuoc tang luc", "sua ", "la vie", "aquafina", "dasani", "lay's", "lays",
  "oishi", "poca", "oreo", "chocopie", "xuc xich", "snack", "do choi",
  "luu niem", "ao thun", "quan short", "non", "tui xach", "moc khoa", "sticker",
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
  if (PHI_LUU_TRU_KEYWORDS.some((k) => t.includes(k))) {
    return { phanLoai: "Phí lưu trú", taiKhoanNo: "154" };
  }
  if (DICH_VU_KEYWORDS.some((k) => t.includes(k))) {
    return { phanLoai: "Dịch vụ", taiKhoanNo: "154" };
  }
  if (NVL_CHE_BIEN_KEYWORDS.some((k) => t.includes(k))) {
    return { phanLoai: "NVL chế biến", taiKhoanNo: "154" };
  }
  if (HANG_HOA_BAN_LIEN_KEYWORDS.some((k) => t.includes(k))) {
    return { phanLoai: "Hàng hóa", taiKhoanNo: "156" };
  }
  // Chi Nhan, 2026-07-29: "có nhiều cái không liên quan tới dịch vụ vui chơi
  // giải trí không tạo ra doanh thu thì đưa vào tên Khác nhá đưa vào 642 theo
  // lý lẽ của bạn" -- LUC DAU thu mac dinh MOI dong con lai (khong khop nhom
  // nao) thanh "Khác"/642, nhung kiem tra tren du lieu that (2026-07-29) cho
  // thay nhieu mat hang THAT SU la hang hoa ban tai cac gian (nuoc uong,
  // snack, sua...) chi vi CHUA co trong danh muc hang hoa Chi Nhan tai len
  // nen bi roi xuong day va bi gan SAI thanh 642 -- rui ro hon la de trong.
  // Vi vay: CHI gan "Khác"/642 khi Dien giai chua 1 tu khoa trong
  // CHI_PHI_642_KEYWORDS o tren (da mo rong them phi giao dich/phan mem/
  // marketing tu du lieu that) -- con lai KHONG khop duoc nhom nao thi de
  // TRONG (nhu truoc), Chi Nhan tu xem/gan tay hoac tai them danh muc hang
  // hoa day du hon (mat hang/mã trong file se tu khop chinh xac, xem
  // matchHangHoaRecord) thay vi doan sai qua tu khoa.
  return { phanLoai: "", taiKhoanNo: "" };
}

// Chi Nhan, 2026-07-29: "tên hàng hóa là mã hàng hóa á bim bim cũng vậy á nên
// bạn check lại tên nha có thể tên khác" -- truoc day CHI khop theo "Tên" dai
// (substring), nhieu mat hang trong danh muc (dac biet snack/bimbim) ghi theo
// "Mã" (vd "TEXAS53G10X10") thay vi ten day du, hoac ten trong Dien giai viet
// tat/khac thu tu voi ten trong danh muc nen khong khop duoc -- THEM buoc 2:
// thu khop theo Ma (bo het khoang trang ca 2 ben) truoc khi bo cuoc. Tra ve CA
// BAN GHI (khong chi ten) de lay them tkKho dung lam Tai khoan No.
function matchHangHoaRecord(dienGiai, hangHoaList) {
  const t = normVN(dienGiai);
  if (!t) return null;
  const byTen = hangHoaList.find((h) => h.ten && t.includes(normVN(h.ten)));
  if (byTen) return byTen;
  const tNoSpace = t.replace(/\s+/g, "");
  const byMa = hangHoaList.find((h) => h.ma && h.ma.length >= 4 && tNoSpace.includes(normVN(h.ma).replace(/\s+/g, "")));
  return byMa || null;
}

function matchTenHangHoa(dienGiai, hangHoaList) {
  const found = matchHangHoaRecord(dienGiai, hangHoaList);
  return found ? found.ten : "";
}

// Chi Nhan, 2026-07-29: dung CHUNG 1 cho 3 cho goi (enrichNewRow, upload-
// hang-hoa, cap-nhat-phan-loai ben routes/hoa-don-dau-vao.js) -- danh muc
// hang hoa CHINH CHI Chi Nhan da tai len co san cot "Tính chất" (Hàng
// hóa/Dịch vụ, kiem tra du lieu that 2026-07-29 thay co san, KHONG phai doan)
// -- mat hang tinh chat "Dịch vụ" (vd "THU PHI SMS...") thi KHONG duoc gan
// thang 156 (do la hang hoa ban ra), phai roi ve doan qua tu khoa
// (classifyPhanLoai) nhu binh thuong de ra dung 154/642/... CHI mat hang
// tinh chat "Hàng hóa" (hoac khong ghi ro tinh chat) moi mac dinh 156.
function classifyFromHangHoaMatch(matched, dienGiai, soTienTruocThue) {
  const isDichVu = normVN(matched.tinhChat).includes("dich vu");
  // Luyen, 2026-08-10: du danh muc ghi tinh chat "Hang hoa" (vd formex, decal)
  // nhung neu dien giai ro rang la DICH VU THI CONG (co cac tu khoa nhu "thi
  // cong", "lap dat", "sua chua"...) thi danh gia la dich vu, KHONG phai hang
  // hoa ban ra. Ten vat lieu xuat hien trong dien giai vi la vat lieu thi cong,
  // khong phai mat hang ban cho khach (vd "thi cong decal formex" != ban formex).
  const t = normVN(dienGiai);
  const isDichVuByKeyword = DICH_VU_KEYWORDS.some((k) => t.includes(k));
  if (isDichVu || isDichVuByKeyword) {
    const classified = classifyPhanLoai(dienGiai, soTienTruocThue);
    return {
      tenHangHoaMisa: matched.ten,
      phanLoai: classified.phanLoai || "Dịch vụ (theo danh mục)",
      taiKhoanNo: matched.tkKho || classified.taiKhoanNo || "154",
    };
  }
  return {
    tenHangHoaMisa: matched.ten,
    phanLoai: "Hàng hóa (theo danh mục)",
    taiKhoanNo: matched.tkKho || "156",
  };
}

module.exports = {
  normVN,
  parseNccListWorkbook,
  parseHangHoaWorkbook,
  parseGianSheetWorkbook,
  matchByMstOrName,
  matchGianViaUncContent,
  matchGianViaChiPhiLedger,
  classifyPhanLoai,
  matchTenHangHoa,
  matchHangHoaRecord,
  classifyFromHangHoaMatch,
};
