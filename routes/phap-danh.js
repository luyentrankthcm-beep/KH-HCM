const express = require("express");
const XLSX = require("xlsx");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin, requireDataEntry } = require("../middleware/auth");
const { getCompany, COMPANIES } = require("../utils/companies");
const { parseHopDongHcmWorkbook, isNccRow } = require("../utils/hopDongHcmParser");
const { parseAmount } = require("../utils/parse");
const { parseGianSheetWorkbook, normVN } = require("../utils/hoaDonDauVaoEnrich");
const { removeDiacritics } = require("../utils/rentPaymentMatcher");

const router = express.Router();
router.use(requireLogin);
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 80 * 1024 * 1024 } });

// Luyen, 2026-07-19: muc "Phap danh" moi -- 2 trang hop dong. Luyen xac nhan
// qua AskUserQuestion: tao khung trang RONG truoc (bang + nut them dong),
// bo sung them cot/truong sau khi Luyen mo ta chi tiet. Cot hien co la bo
// truong pho bien nhat cho tung loai hop dong, khong phai danh sach cuoi
// cung -- de sua/them cot sau nay khong kho (chi 1 object moi dong, EJS doc
// truc tiep tu field, khong co schema cung).
function ensureShape(store) {
  if (!store.phap_danh_hop_dong_thue) store.phap_danh_hop_dong_thue = [];
  if (!store.phap_danh_hop_dong_ncc) store.phap_danh_hop_dong_ncc = [];
  // Chi Nhan, 2026-07-31: ghi chu tay cho trang "Doi soat doanh thu chia se
  // theo thang" (xem route ben duoi) -- key la thang "YYYY-MM", value la
  // chuoi ghi chu tu do Chi Nhan nhap, doc lap voi so lieu tu dong tinh.
  if (!store.phap_danh_doanhthu_chiase_ghichu) store.phap_danh_doanhthu_chiase_ghichu = {};
  // Chi Nhan, 2026-07-31: bang NHAP TAY rieng cho loai "xuat_hoa_don" (ben
  // cho thue xuat hoa don dua tren so Chi Nhan tu bao cao cho ho -- KHONG the
  // tu tinh tu doi soat nhu loai "giu_tien" o tren).
  if (!store.phap_danh_doanhthu_chiase_xuathoadon) store.phap_danh_doanhthu_chiase_xuathoadon = [];
  // Nhan, 2026-08-13: "Hợp Đồng Thuê Gian Tổng" -- muc moi trong dropdown
  // Phap danh, nap tu Google Sheet "Bản sao của Danh sách các gian Luyến.xlsx"
  // (4 tab KVC MN/KVC MB/MTD MN/MTD MB, xem tmp_import_hopdong_tong.js da chay
  // 1 lan de nap 246 dong). Moi dong = 1 gian, sheetKey phan biet 4 tab, congTy
  // "kh_cu"/"kh_moi"/"ca_2" (dong "Cả 2" hien o CA 2 cong ty khi loc, xem route
  // ben duoi) hoac "" khi sheet ghi "Không tìm thấy"/rong (chua xac dinh duoc,
  // KHONG doan).
  if (!store.phap_danh_hop_dong_thue_tong) store.phap_danh_hop_dong_thue_tong = [];
  fillSnowAeBinhDuongFromContract(store);
}

// Chi Nhan, 2026-07-30: "có mà đọc file lấy ra cho tôi đi" -- gian "SNow AE
// Binh Duong" (Nha Tuyet Binh Duong, id=255, nhap tay 30/7) con thieu het cac
// truong ngay/tien vi luc them chua co hop dong. Nhan chup man hinh 1 phan
// hop dong (Google Doc lien ket san o linkHopDongChuaDuDau) -- doc duoc:
// Ngay Bat Dau Kinh Doanh/Tinh Gia Thue 05/08/2026, Ngay Het Han 06/09/2026,
// Gia Thue = 50.000.000d/thang HOAC 10% Tong doanh thu (ap dung muc nao cao
// hon), Tien Coc Dam Bao 50.000.000d (truoc ngay ban giao). KHONG thay "So
// HD" hay "Ngay ky" trong phan anh chup gui -- de trong, can Nhan bo sung
// rieng. Chi dien vao neu truong DANG TRONG (khong ghi de neu da tu sua tay).
function fillSnowAeBinhDuongFromContract(store) {
  const g = store.phap_danh_hop_dong_thue.find((x) => x.gian === "SNow AE Bình Dương");
  if (!g) return false;
  let changed = false;
  const fill = (key, value) => {
    if (!g[key] && value) {
      g[key] = value;
      changed = true;
    }
  };
  fill("ngayBatDauHD", "2026-08-05");
  fill("ngayHetHanHD", "2026-09-06");
  fill("thoiHanHopDong", "05/08/2026-06/09/2026");
  fill("tienCocDamBao", 50000000);
  fill(
    "tongTienThueThangHCM",
    "50.000.000đ/tháng (giá thuê cơ sở) HOẶC 10% Tổng doanh thu trong Thời hạn thuê -- áp dụng mức nào cao hơn (theo hợp đồng)"
  );
  if (!g.ghiChu) {
    g.ghiChu =
      "Con thieu So HD va Ngay ky (khong thay trong anh chup hop dong Nhan gui) -- can bo sung tay.";
    changed = true;
  }
  if (changed) save(store);
  return changed;
}

// ---------- Hop Dong Thue Gian Hang ----------
// Luyen, 2026-07-20: mo rong schema ban dau (chi 6 truong don gian) thanh cau
// truc day du hon sau khi Luyen gui 2 file MTD MN.xlsx / KVC MN.xlsx (danh
// sach gian that cua POSH-may tu dong va KVC-khu vui choi, tach san theo
// KH705=KH Moi / KH989=KH Cu). Da import 67 dong tu 2 file nay lam du lieu
// khoi tao (xem source="import MTD MN.xlsx / KVC MN.xlsx 2026-07-20"); form
// them tay ben duoi van dung DE THEM GIAN MOI sau nay khi Luyen doc duoc tu
// hop dong. "gian" luon uu tien "ten diem xuat hoa don" (cung 1 gia tri dung
// de khop giao dich doi soat VietQR/ZVP -- xem bug fix cot F ngay 2026-07-19)
// de 2 he thong dong bo voi nhau. "thoiHanHopDong" giu nguyen dang text tu do
// (KHONG ep ve 2 o ngay bat dau/ket thuc) vi du lieu thuc te co rat nhieu
// dinh dang gia han/phu luc khac nhau, ep ve ngay se de doc sai.
// Luyen, 2026-07-20: bo sung du lieu hop dong that tu Google Sheet "THEO DOI
// HD HN-HCM" (sheet HCM, 787 dong) -- khop 52/67 gian qua ma diem rut gon
// (vd 50AMBT/JPAMBT/TTAMBT deu chung 1 mat bang vat ly "AMBT") va khop ten/
// tu khoa dia diem. Cac truong moi deu LAY TU SHEET (khong phai Luyen nhap
// tay) nen giu rieng, khong ghi de len cac truong nhap tay (gian, benChoThue,
// thoiHanHopDong, tienThueThang...) de tranh mat du lieu Luyen da co san.
function ensureThueDefaults(row) {
  return Object.assign(
    {
      loaiHinh: "",
      congTy: "kh_cu",
      maDiemMisa: "",
      tenDiemNoiBo: "",
      khuVuc: "",
      gian: "",
      maCongTrinh: "",
      maKH: "",
      benChoThue: "",
      mstBenChoThue: "",
      hinhThucHopTac: "",
      trangThaiHoatDong: "",
      thoiHanHopDong: "",
      tienThueThang: 0,
      ghiChu: "",
      soHopDongHCM: "",
      ngayKyHD: "",
      ngayBatDauHD: "",
      ngayHetHanHD: "",
      baoHanHD: "",
      tienCocDamBao: "",
      tienCocThiCong: "",
      tongTienThueThangHCM: "",
      linkHopDongDuDau: "",
      linkHopDongChuaDuDau: "",
      sheetHCMThamKhao: "",
      sheetHCMDiaDiemGoc: "",
      // Luyen, 2026-07-20: them tu anh chup dieu khoan "Thanh toan" trong hop
      // dong goc (khong co san trong sheet HCM, Luyen gui rieng tung anh chup
      // trang hop dong) -- ghi tu do, khong tach truong rieng vi moi hop dong
      // co the co dieu khoan khac nhau (co thu doanh so hay khong, lai suat
      // qua han, phi may tinh tien...).
      dieuKhoanThanhToan: "",
      // Luyen, 2026-07-20: "cac gian can xem ky hop dong la gian doanh thu ma
      // ben mall dang giu tien roi moi cuoi thang tra ve qua tai khoan cho
      // minh phan da tru tien thue voi phi dich vu" -- danh dau rieng cac
      // gian nay (khac voi kieu thu tien truc tiep qua VietQR/POS cong ty
      // binh thuong) de Luyen doi chieu ky hon khi nhan tien cuoi thang.
      hinhThucThuTien: "",
      // Luyen, 2026-07-23: "có tàu bình tân với nhà ma bình dương hay ghost
      // bình dương cũng là doanh thu chia sẻ, cho thêm thủ công cũng được" --
      // khong phai gian doanh thu chia se nao cung tu doan duoc tu van ban
      // hop dong (dieuKhoanThanhToan/hinhThucThuTien o tren) -- them 1 co tich
      // chon TAY de Luyen tu danh dau bat ky gian nao la doanh thu chia se,
      // dung cho ca cot Tai Khoan (131/1388) ben trang Chi Phi (xem
      // isDoanhThuChiaSeRecord trong utils/rentPaymentMatcher.js).
      doanhThuChiaSe: false,
      // Chi Nhan, 2026-07-31: "nó sẽ chia ra làm 2 loại doanh thu chia sẻ" --
      // rieng cac gian doanhThuChiaSe=true, phan biet THEM 2 kieu quan he khac
      // han nhau ve mat ke toan: "giu_tien" (mac dinh -- ben cho thue GIU
      // TIEN, giua thang gui file doi chieu, cuoi thang TRU tien thue+phi roi
      // CHUYEN KHOAN phan con lai -- tu dong tinh duoc tu doi soat + giao dich
      // NH) vs "xuat_hoa_don" (ben cho thue XUAT HOA DON dua tren so MINH tu
      // bao cao doanh thu cho ho -- KHONG the tu dong tinh tu doi soat, Chi
      // Nhan se tu liet ke gui rieng, xem trang "Đối Soát Doanh Thu Chia Sẻ").
      loaiChiaSe: "giu_tien",
      // Chi Nhan, 2026-07-31: "chia làm 4 loại MTD miền nam KVC miền nam MTD
      // miền bắc KVC miền bắc trong cái doanh thu chia sẻ này" -- "MTD"="Máy
      // tự động"/"KVC"="Khu vui chơi" da co san qua truong loaiHinh o tren,
      // CHI thieu chieu Nam/Bac -- them truong "mien" rieng (doc lap voi
      // "khuVuc" cu, dang la text tu do "HCM"/rong, khong phai Nam/Bac ro
      // rang), cung quy uoc "nam"/"bac" nhu store.chi_phi.mien. TAT CA gian
      // doanh thu chia se hien co deu o khuVuc "HCM" nen mac dinh "nam", giong
      // cach lam voi Chi Phi truoc do.
      mien: "nam",
      // Luyen, 2026-07-23 (lan 2): "có thể viết tắc á" -- ten gian ben Chi Phi
      // thuong la ma viet tat rieng cua Luyen (vd "TÀU BT", "GHOST AMBD") ma
      // khac han ten day du trong hop dong nay nen he thong khong tu khop
      // duoc -- them truong nay de Luyen TU DIEN cac ma viet tat (cach nhau
      // boi dau phay) cho tung hop dong, dung khop CHINH XAC ben Chi Phi
      // (xem findContractForGianText trong utils/rentPaymentMatcher.js).
      aliasGian: "",
    },
    row
  );
}

// Luyen, 2026-07-20: "xem gian nao con hoat dong gian nao het han cho bo loc
// theo thang di vi la event co khi chay co 1 thang thoi" -- suy ra trang
// thai tu ngayHetHanHD (uu tien, lay tu sheet HCM that) hoac baoHanHD ("qua
// han"); gian chua co ngay het han thi coi la "Khong ro" (khong ep buoc phai
// con/het han khi chua co du lieu, tranh bao sai).
function computeTrangThaiHD(r, todayStr) {
  // Luyen, 2026-07-20: bug -- code cu kiem tra baoHanHD ("qua han" copy tho tu
  // sheet, co the la trang thai CU truoc khi co phu luc gia han) TRUOC ngay
  // het han thuc te, lam nhieu gian da duoc gia han (ngayHetHanHD trong
  // tuong lai) van bi hien "Het han" sai (vd Sense CT Can Tho, VC 3/2).
  // Dung nhu comment ban dau da noi: ngayHetHanHD phai UU TIEN khi co du lieu;
  // baoHanHD chi la fallback khi chua co ngay het han ro rang.
  if (r.ngayHetHanHD) {
    return r.ngayHetHanHD < todayStr ? "Hết hạn" : "Còn hoạt động";
  }
  if (r.baoHanHD && /qua han|quá hạn/i.test(r.baoHanHD)) return "Hết hạn";
  return "Không rõ";
}

// Loc theo thang (query "thang", dang "YYYY-MM"): gian nao co khoang thoi
// gian hop dong [ngayBatDauHD, ngayHetHanHD] giao voi thang do thi giu lai.
// Gian CHUA co du lieu ngay (chi co thoiHanHopDong dang text tu do, khong ep
// parse) thi LUON giu lai -- khong du du lieu de loai bo dung, thay vi loc
// nham mat gian that.
function monthOverlaps(r, thang) {
  if (!thang) return true;
  if (!r.ngayBatDauHD && !r.ngayHetHanHD) return true;
  const monthStart = thang + "-01";
  const monthEnd = thang + "-31"; // so sanh chuoi ISO yyyy-mm-dd, du 28-31 deu <= "-31"
  const bd = r.ngayBatDauHD || "0000-00-00";
  const hh = r.ngayHetHanHD || "9999-99-99";
  return bd <= monthEnd && hh >= monthStart;
}

// Luyen, 2026-07-25: "ưu tiên hiển thị tháng 7" -- mac dinh loc theo THANG
// HIEN TAI ngay khi vao trang (thay vi hien het ca 100+ gian moi lan vao),
// nhung van cho xem "Tat ca" qua 1 lua chon ro rang. Vi input HTML type=month
// khong co gia tri nao dai dien cho "khong loc" (rong = chua chon, khong
// phai "tat ca"), dung sentinel "all" o query string de phan biet 2 truong
// hop: KHONG co "thang" trong URL (vao trang lan dau, tab moi...) -> mac dinh
// thang nay; co "thang=all" ro rang (bam nut "Bỏ lọc tháng") -> tat ca.
function resolveThangFilter(req) {
  const currentMonthStr = new Date().toISOString().slice(0, 7);
  const raw = req.query.thang;
  if (raw === "all") return "";
  return raw || currentMonthStr;
}

// Luyen, 2026-08-01 (lan 6): "tôi đang chọn Hợp đồng thuê gian của miền nam
// cơ mà soa lại có cả miền bắc vào đấy" -- truoc gio trang Nam KHONG loc theo
// "mien" (hien TAT CA gian bat ke Nam/Bac), khong sao khi CHUA co du lieu
// Mien Bac that; nay Luyen da tu cap nhat xong trang Mien Bac (gian nhu KVC
// ROYAL/KVC TIMES/LOTTE BAC GIANG KVC... co mien="bac") nen bi lan CHUNG vao
// trang Nam. Them dieu kien loc r.mien !== "bac" (gian cu chua co truong mien
// se mac dinh "nam" qua ensureThueDefaults, khong bi anh huong).
function isMienNamRow(r) {
  return r.mien !== "bac";
}

router.get("/phap-danh/hop-dong-thue-gian-hang", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const loaiFilter = req.query.loai || "";
  const thangFilter = resolveThangFilter(req);
  const todayStr = new Date().toISOString().slice(0, 10);
  let rows = store.phap_danh_hop_dong_thue
    .map(ensureThueDefaults)
    .filter((r) => (r.congTy || "kh_cu") === activeCompany && isMienNamRow(r));
  if (loaiFilter) rows = rows.filter((r) => r.loaiHinh === loaiFilter);
  if (thangFilter) {
    rows = rows.filter((r) => monthOverlaps(r, thangFilter));
    // Luyen, 2026-08-24: khi loc theo thang, bo cac HD da het han truoc dau thang
    rows = rows.filter((r) => {
      if (!r.ngayHetHanHD) return true;
      return r.ngayHetHanHD >= thangFilter + "-01";
    });
  }
  rows.forEach((r) => {
    r.trangThaiHD = computeTrangThaiHD(r, todayStr);
  });
  rows.sort((a, b) => (a.loaiHinh === b.loaiHinh ? a.gian.localeCompare(b.gian) : a.loaiHinh.localeCompare(b.loaiHinh)));
  const counts = {
    "Khu vui chơi": store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && isMienNamRow(r) && r.loaiHinh === "Khu vui chơi"
    ).length,
    "Máy tự động": store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && isMienNamRow(r) && r.loaiHinh === "Máy tự động"
    ).length,
    "Ghế": store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && isMienNamRow(r) && r.loaiHinh === "Ghế"
    ).length,
  };
  res.render("phapdanh-hopdong-thue", {
    userName: req.session.userName,
    rows,
    loaiFilter,
    thangFilter,
    counts,
    totalForCompany: store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && isMienNamRow(r)
    ).length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Luyen, 2026-07-21: "có nút xuất excel cho tôi nhá" -- xuat danh sach gian
// dang xem (theo cong ty + bo loc loai hinh/thang dang chon tren man hinh,
// giong cach lam voi trang giao dich /transactions/export.xlsx) ra file
// Excel. Xuat toan bo cot dang hien thi tren bang (ke ca cac truong tu sheet
// HCM) de chi doi chieu/luu tru duoc ben ngoai web, khong can them cot moi.
router.get("/phap-danh/hop-dong-thue-gian-hang/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const loaiFilter = req.query.loai || "";
  const thangFilter = req.query.thang || "";
  const todayStr = new Date().toISOString().slice(0, 10);
  let rows = store.phap_danh_hop_dong_thue
    .map(ensureThueDefaults)
    .filter((r) => (r.congTy || "kh_cu") === activeCompany && isMienNamRow(r));
  if (loaiFilter) rows = rows.filter((r) => r.loaiHinh === loaiFilter);
  if (thangFilter) rows = rows.filter((r) => monthOverlaps(r, thangFilter));
  rows.forEach((r) => {
    r.trangThaiHD = computeTrangThaiHD(r, todayStr);
  });
  rows.sort((a, b) => (a.loaiHinh === b.loaiHinh ? a.gian.localeCompare(b.gian) : a.loaiHinh.localeCompare(b.loaiHinh)));

  const exportRows = rows.map((r) => ({
    "Loại hình": r.loaiHinh,
    Gian: r.gian,
    "Tên điểm nội bộ": r.tenDiemNoiBo,
    "Mã công trình": r.maCongTrinh,
    "Mã KH": r.maKH,
    "Bên cho thuê": r.benChoThue,
    "MST bên cho thuê": r.mstBenChoThue,
    "Hình thức HT": r.hinhThucHopTac,
    "Hình thức thu tiền": r.hinhThucThuTien,
    "Trạng thái HĐ": r.trangThaiHD,
    "Thời hạn hợp đồng": r.thoiHanHopDong,
    "Tiền thuê/tháng": r.tienThueThang,
    "Số HĐ (sheet HCM)": r.soHopDongHCM,
    "Ngày ký HĐ": r.ngayKyHD,
    "Ngày bắt đầu HĐ": r.ngayBatDauHD,
    "Ngày hết hạn HĐ": r.ngayHetHanHD,
    "Tiền cọc đảm bảo": r.tienCocDamBao,
    "Tiền cọc thi công": r.tienCocThiCong,
    "Tiền thuê/tháng (sheet HCM)": r.tongTienThueThangHCM,
    "Điều khoản thanh toán": r.dieuKhoanThanhToan,
    "Link HĐ đủ dấu": r.linkHopDongDuDau,
    "Link HĐ chưa đủ dấu": r.linkHopDongChuaDuDau,
    "Ghi chú": r.ghiChu,
  }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Thue gian hang");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader(
    "Content-Disposition",
    `attachment; filename=hop-dong-thue-gian-hang-${activeCompany}.xlsx`
  );
  res.send(buf);
});

router.post("/phap-danh/hop-dong-thue-gian-hang", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const {
      loaiHinh,
      gian,
      tenDiemNoiBo,
      maCongTrinh,
      maKH,
      benChoThue,
      mstBenChoThue,
      hinhThucHopTac,
      trangThaiHoatDong,
      thoiHanHopDong,
      tienThueThang,
      ghiChu,
      linkHopDongChuaDuDau,
    } = req.body;
    if (!gian || !gian.trim()) throw new Error("Thiếu Gian/Mặt bằng.");
    const amt = tienThueThang ? Number(String(tienThueThang).replace(/[^\d]/g, "")) : 0;
    store.phap_danh_hop_dong_thue.push({
      id: nextId(store, "phap_danh_hop_dong_thue_seq") || Date.now(),
      loaiHinh: loaiHinh || "",
      congTy: activeCompany,
      maDiemMisa: "",
      tenDiemNoiBo: (tenDiemNoiBo || "").trim(),
      khuVuc: "",
      gian: gian.trim(),
      maCongTrinh: (maCongTrinh || "").trim(),
      maKH: (maKH || "").trim(),
      benChoThue: (benChoThue || "").trim(),
      mstBenChoThue: (mstBenChoThue || "").trim(),
      hinhThucHopTac: (hinhThucHopTac || "").trim(),
      trangThaiHoatDong: (trangThaiHoatDong || "").trim(),
      thoiHanHopDong: (thoiHanHopDong || "").trim(),
      tienThueThang: amt,
      ghiChu: (ghiChu || "").trim(),
      // Luyen, 2026-07-21: "thêm chỗ link hợp đồng chưa có dấu cập nhật cho
      // tôi nhá" -- gian nhap tay (chua co trong sheet HCM) van can luu duoc
      // link ban chua ky/dong dau ngay luc them, thay vi phai doi den khi
      // gian do xuat hien tren Google Sheet HCM. linkHopDongDuDau van chi
      // dien qua co che khop tu dong voi sheet HCM (khong co o form nay) --
      // khi ban co dau xuat hien tren sheet, gian se duoc khop lai va dien
      // linkHopDongDuDau rieng.
      linkHopDongChuaDuDau: (linkHopDongChuaDuDau || "").trim(),
      createdAt: new Date().toISOString(),
      source: "nhap tay",
    });
    save(store);
    res.redirect("/phap-danh/hop-dong-thue-gian-hang?success=" + encodeURIComponent("Đã lưu hợp đồng."));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-thue-gian-hang?error=" + encodeURIComponent(e.message));
  }
});

router.post("/phap-danh/hop-dong-thue-gian-hang/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  store.phap_danh_hop_dong_thue = store.phap_danh_hop_dong_thue.filter((r) => String(r.id) !== req.params.id);
  save(store);
  res.redirect("/phap-danh/hop-dong-thue-gian-hang?success=" + encodeURIComponent("Đã xóa hợp đồng."));
});

// Chi Nhan, 2026-07-29: "cái nào mà thuê gian hàng thì cho vô hợp đồng thuê
// gian còn lại của hcm cho qua bên ncc" -- nhieu dong duoc import tu dong tu
// Google Sheet ("Thuộc BP" trong sheet khong dung "Khác"/"Mua bán vocher" nen
// lot qua bo loc NCC cu, xem utils/hopDongHcmParser.js) thuc ra la hop dong
// NCC that (van tai, xay dung, thanh toan, thuc pham...), khong phai thue mat
// bang/gian hang. Truoc day KHONG CO route nao chuyen 1 dong giua 2 danh sach
// -- chi co xoa (mat het du lieu) hoac them tay lai tu dau ben NCC. Route nay
// chuyen NGUYEN VEN 1 dong sang phap_danh_hop_dong_ncc (anh xa truong tuong
// duong), roi xoa khoi phap_danh_hop_dong_thue, giu lai toan bo thong tin da
// co (ten, MST, so hop dong, ngay, tien, link...) thay vi phai nhap lai tay.
router.post("/phap-danh/hop-dong-thue-gian-hang/:id/chuyen-sang-ncc", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const idx = store.phap_danh_hop_dong_thue.findIndex((r) => String(r.id) === req.params.id);
  if (idx === -1) {
    return res.redirect(thueGianListPath(req) + "?error=" + encodeURIComponent("Không tìm thấy hợp đồng này."));
  }
  const r = ensureThueDefaults(store.phap_danh_hop_dong_thue[idx]);
  const nccRec = ensureNccDefaults({
    id: nextId(store, "phap_danh_hop_dong_ncc_seq") || Date.now(),
    tenNCC: r.gian || r.tenDiemNoiBo || "",
    tenDayDuNCC: r.benChoThue || "",
    mstNCC: r.mstBenChoThue || "",
    diaChiNCC: "",
    congTy: r.congTy || "",
    hangHoaMua: "",
    noiDung: "",
    phanLoaiHD: r.hinhThucHopTac || "",
    soHopDong: r.soHopDongHCM || "",
    ngayKy: r.ngayKyHD || "",
    ngayBatDauHD: r.ngayBatDauHD || "",
    ngayHetHan: r.ngayHetHanHD || "",
    baoHanHD: r.baoHanHD || "",
    giaTriHopDong: r.tienThueThang || 0,
    giaTriHopDongRaw: r.tongTienThueThangHCM || "",
    linkHopDong: r.linkHopDongChuaDuDau || r.linkHopDongDuDau || "",
    soTKNH: "",
    khachMoTaiNH: "",
    khuVuc: r.khuVuc || "",
    chiTiet: [],
    ghiChu: [
      r.ghiChu || "",
      `Chuyển từ "Hợp đồng thuê gian hàng" sang NCC ngày ${new Date().toISOString().slice(0, 10)} (không phải hợp đồng thuê mặt bằng/gian hàng).`,
    ]
      .filter(Boolean)
      .join(" -- "),
  });
  if (!store.phap_danh_hop_dong_ncc) store.phap_danh_hop_dong_ncc = [];
  store.phap_danh_hop_dong_ncc.push(nccRec);
  store.phap_danh_hop_dong_thue.splice(idx, 1);
  save(store);
  res.redirect(
    thueGianListPath(req) + "?success=" + encodeURIComponent(`Đã chuyển "${nccRec.tenNCC}" sang Hợp đồng NCC.`)
  );
});

// Luyen, 2026-07-23: "có tàu bình tân với nhà ma bình dương hay ghost bình
// dương cũng là doanh thu chia sẻ ... cho thêm thủ công cũng được" -- tich
// chon tay 1 gian la "doanh thu chia se" (dung cho cot Tai Khoan 131/1388 ben
// trang Chi Phi), khong phu thuoc vao viec dieu khoan hop dong co ghi ro hay
// khong. Giu nguyen loai/thang filter dang xem khi redirect ve.
// Luyen, 2026-08-01: cac route hanh dong tren tung dong (doanh-thu-chia-se,
// loai-chia-se, mien, alias-gian, chuyen-sang-ncc) dung CHUNG cho ca 2 trang
// Nam/Bac (cung thao tac tren cung mang store.phap_danh_hop_dong_thue, chi
// khac o trang hien thi) -- dung field an "listPath=mien-bac" tren form de
// biet redirect ve trang nao, mac dinh ve trang Nam (form Nam khong gui field
// nay, giu nguyen hanh vi cu).
function thueGianListPath(req) {
  return req.body && req.body.listPath === "mien-bac"
    ? "/phap-danh/hop-dong-thue-gian-hang-mien-bac"
    : "/phap-danh/hop-dong-thue-gian-hang";
}

function redirectBackToThueGianList(req, res, extra) {
  const qs = [];
  if (req.body.loai) qs.push("loai=" + encodeURIComponent(req.body.loai));
  if (req.body.thang) qs.push("thang=" + encodeURIComponent(req.body.thang));
  Object.entries(extra || {}).forEach(([k, v]) => qs.push(k + "=" + encodeURIComponent(v)));
  res.redirect(thueGianListPath(req) + (qs.length ? "?" + qs.join("&") : ""));
}

router.post("/phap-danh/hop-dong-thue-gian-hang/:id/doanh-thu-chia-se", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.phap_danh_hop_dong_thue.find((x) => String(x.id) === req.params.id);
  if (!r) return redirectBackToThueGianList(req, res, { error: "Không tìm thấy gian này." });
  r.doanhThuChiaSe = req.body.doanhThuChiaSe === "1";
  save(store);
  redirectBackToThueGianList(req, res, { success: "Đã cập nhật doanh thu chia sẻ." });
});

// Chi Nhan, 2026-07-31: 2 truong phan loai rieng cho gian doanh thu chia se
// (xem ensureThueDefaults) -- "loaiChiaSe" (giu tien cuoi thang chuyen khoan
// vs doi tac xuat hoa don theo bao cao) va "mien" (Nam/Bac, ket hop voi
// loaiHinh co san = MTD/KVC ra du 4 nhom Chi Nhan yeu cau). Dung select tu
// dong submit (giong cac bo loc thang khac trong app), khong can nut Luu rieng.
router.post("/phap-danh/hop-dong-thue-gian-hang/:id/loai-chia-se", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.phap_danh_hop_dong_thue.find((x) => String(x.id) === req.params.id);
  if (!r) return redirectBackToThueGianList(req, res, { error: "Không tìm thấy gian này." });
  r.loaiChiaSe = req.body.loaiChiaSe === "xuat_hoa_don" ? "xuat_hoa_don" : "giu_tien";
  save(store);
  redirectBackToThueGianList(req, res, { success: "Đã cập nhật loại chia sẻ." });
});

router.post("/phap-danh/hop-dong-thue-gian-hang/:id/mien", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.phap_danh_hop_dong_thue.find((x) => String(x.id) === req.params.id);
  if (!r) return redirectBackToThueGianList(req, res, { error: "Không tìm thấy gian này." });
  r.mien = req.body.mien === "bac" ? "bac" : "nam";
  save(store);
  redirectBackToThueGianList(req, res, { success: "Đã cập nhật miền." });
});

// Luyen, 2026-07-23 (lan 2): "có thể viết tắc á" -- cho Luyen tu dien cac ma
// viet tat (vd "TAU BT, TÀU BT") de Chi Phi khop CHINH XAC gian ma khong doan
// nham (xem findContractForGianText trong utils/rentPaymentMatcher.js).
router.post("/phap-danh/hop-dong-thue-gian-hang/:id/alias-gian", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.phap_danh_hop_dong_thue.find((x) => String(x.id) === req.params.id);
  if (!r) return redirectBackToThueGianList(req, res, { error: "Không tìm thấy gian này." });
  r.aliasGian = (req.body.aliasGian || "").trim();
  save(store);
  redirectBackToThueGianList(req, res, { success: "Đã cập nhật mã viết tắt." });
});

// Luyen, 2026-07-21 (lan 3): "làm luôn cho thuê gian hàng nhá" -- nut doc
// thang tu CUNG Google Sheet "Tổng hợp HCM-HN" da dung cho Hop Dong NCC (dung
// chung 1 link, khac o cho LOC NGUOC LAI: tat ca dong KHONG phai NCC -- tuc
// isNccRow() === false -- la dong thue gian/mat bang, xem chu thich isNccRow
// trong utils/hopDongHcmParser.js).
//
// Khac voi NCC (khoa ghep chinh xac la cap (diaDiem, congTy)), trang nay
// KHONG co khoa ghep dang tin cay tuong tu voi cac gian DA CO san (52/67 gian
// duoc ghep 1 LAN DUY NHAT qua ma diem rut gon + tu khoa dia diem -- viec lam
// tay/xet doan, khong phai ham xac dinh). De AN TOAN, KHONG ghi de nham cac
// truong Luyen tu nhap tay (benChoThue, thoiHanHopDong, tienThueThang, dieu
// khoan thanh toan, hinh thuc thu tien...):
//   - Khop lai voi gian DA CO qua khoa CHINH XAC "soHopDongHCM" (da duoc dien
//     tu lan ghep dau tien) + congTy -- neu khop, CHI cap nhat cac truong lay
//     tu sheet (ngay ky/bat dau/het han, tien coc, link hop dong...), tuyet
//     doi khong dung vao cac truong nhap tay.
//   - Khong khop duoc (soHopDong nay CHUA co gian nao luu) -- Luyen xac nhan
//     qua AskUserQuestion "Thêm cả gian mới chưa có": tao THEM 1 dong MOI, chi
//     dien cac truong lay tu sheet, de trong loaiHinh/benChoThue/thoiHanHopDong/
//     tienThueThang... cho Luyen tu dien tay + doi ten "Gian" cho dung quy uoc
//     hien dang dung (sheet ghi ten/ma tho, chua chac trung ten Luyen quen goi).
//   - Dong khong co soHopDong (khong the khop/khong the tao moi an toan) --
//     BO QUA, bao so luong trong thong bao ket qua.
function thueGianGroupKey(diaDiem, congTy) {
  return (diaDiem || "(không có địa điểm)") + "||" + (congTy || "");
}

function buildThueGianFieldsFromRows(diaDiem, groupRows, sourceLabel) {
  const first = groupRows[0];
  let ngayHetHanMax = "";
  let baoHanCuaMax = "";
  groupRows.forEach((r) => {
    if (r.ngayHetHan && r.ngayHetHan > ngayHetHanMax) {
      ngayHetHanMax = r.ngayHetHan;
      baoHanCuaMax = r.baoHanHD;
    }
  });
  const tienCocDamBao = groupRows.map((r) => r.tienCocDamBao).find((v) => v !== null && v !== undefined && v !== "");
  const tienCocThiCong = groupRows.map((r) => r.tienCocThiCong).find((v) => v !== null && v !== undefined && v !== "");
  let tongTienThueThangHCM = groupRows
    .map((r) => r.tongTienThueThang)
    .find((v) => v !== null && v !== undefined && v !== "");
  // Cell dang "Số tiền" duoc dinh dang trong Google Sheet (vd " 397.200.000 ")
  // xuat ra CSV van la CHUOI co khoang trang thua o 2 dau -- trim lai cho gon,
  // KHONG ep ve so (giu nguyen dinh dang hien co cua truong tham khao nay).
  if (typeof tongTienThueThangHCM === "string") tongTienThueThangHCM = tongTienThueThangHCM.trim();
  return {
    soHopDongHCM: first.soHopDong || "",
    ngayKyHD: first.ngayKyHD || "",
    ngayBatDauHD: first.ngayBatDauHD || "",
    ngayHetHanHD: ngayHetHanMax || first.ngayHetHan || "",
    baoHanHD: baoHanCuaMax || first.baoHanHD || "",
    tienCocDamBao: tienCocDamBao !== undefined ? tienCocDamBao : "",
    tienCocThiCong: tienCocThiCong !== undefined ? tienCocThiCong : "",
    tongTienThueThangHCM: tongTienThueThangHCM !== undefined ? tongTienThueThangHCM : "",
    linkHopDongDuDau: groupRows.map((r) => r.linkHDduDau).find((v) => v) || "",
    linkHopDongChuaDuDau: groupRows.map((r) => r.linkHDchuaDuDau).find((v) => v) || "",
    sheetHCMDiaDiemGoc: diaDiem,
    sheetHCMThamKhao: `${sourceLabel}, ${groupRows.length} dòng nguồn`,
  };
}

router.post("/phap-danh/hop-dong-thue-gian-hang/cap-nhat-tu-sheet", requireAdmin, async (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const resp = await fetch(NCC_SHEET_CSV_URL);
    if (!resp.ok) {
      throw new Error(
        `Không đọc được Google Sheet (mã lỗi ${resp.status}). Kiểm tra lại sheet đã chia sẻ "Bất kỳ ai có link đều xem được" chưa, hoặc link/gid có bị đổi không.`
      );
    }
    const buf = Buffer.from(await resp.arrayBuffer());
    const { rows } = parseHopDongHcmWorkbook(buf);
    const thueGianRows = rows.filter((r) => !isNccRow(r));
    if (thueGianRows.length === 0) {
      throw new Error('Không tìm thấy dòng nào không phải NCC (Thuộc BP khác "Khác"/"Mua bán vocher") trong Google Sheet.');
    }

    const groups = new Map();
    let skippedNoSoHopDong = 0;
    thueGianRows.forEach((r) => {
      if (!r.soHopDong) {
        skippedNoSoHopDong++;
        return;
      }
      const key = thueGianGroupKey(r.diaDiem, r.congTy);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });

    // Khoa ghep AN TOAN voi gian da co: soHopDongHCM + congTy (KHONG dung ten
    // gian/dia diem de tranh doan nham, xem chu thich phia tren).
    const existingBySoHopDong = new Map();
    store.phap_danh_hop_dong_thue.forEach((r) => {
      if (r.soHopDongHCM) existingBySoHopDong.set(r.soHopDongHCM + "||" + (r.congTy || "kh_cu"), r);
    });

    const SOURCE_TAG = "GG Sheet " + new Date().toISOString().slice(0, 10);
    let addedNew = 0;
    let updatedExisting = 0;
    const addedNames = [];
    const updatedNames = [];

    groups.forEach((groupRows) => {
      const diaDiem = groupRows[0].diaDiem || "(không có địa điểm)";
      const soHopDong = groupRows[0].soHopDong;
      const congTy = groupRows.map((r) => r.congTy).find((v) => v) || "kh_cu";
      const fields = buildThueGianFieldsFromRows(diaDiem, groupRows, SOURCE_TAG);
      // Luyen, 2026-07-25: "mình đang là người thuê địa điểm của họ ... các
      // hợp đồng đã có thì chắc chắn nó sẽ có NCC rồi thì thêm vào cho tôi
      // với" -- sheet nguon (cot tenKH/mstKH) da co san ten/MST cua BEN CHO
      // THUE (chu khong phai KH cua minh, giong het cach hieu o Hop Dong NCC
      // ben duoi), nhung truoc gio khong duoc dien vao benChoThue/mstBenChoThue
      // (2 truong nay truoc gio CHI dien tay). Tu nay tu dong dien tu sheet
      // NHUNG CHI khi con trong (khong ghi de neu Luyen da tu sua tay khac di).
      const benChoThueSheet = groupRows.map((r) => r.tenKH).find((v) => v) || "";
      const mstBenChoThueSheet = groupRows.map((r) => r.mstKH).find((v) => v) || "";
      // Suy loaiHinh tu cot soGheMayDienTich (vd "02 ghe" -> "Ghe", "01 may" -> "May tu dong")
      const soGheMay = groupRows.map((r) => r.soGheMayDienTich).find((v) => v) || "";
      const detectedLoaiHinh = loaiHinhFromSoGheMay(soGheMay, "");
      const matchKey = soHopDong + "||" + congTy;
      const existing = existingBySoHopDong.get(matchKey);
      if (existing) {
        // Gian da co, khop dung qua soHopDongHCM -- CHI ghi de cac truong lay
        // tu sheet, TUYET DOI khong dung vao cac truong Luyen tu nhap tay.
        Object.assign(existing, fields);
        if (!existing.benChoThue && benChoThueSheet) existing.benChoThue = benChoThueSheet;
        if (!existing.mstBenChoThue && mstBenChoThueSheet) existing.mstBenChoThue = mstBenChoThueSheet;
        // Cap nhat loaiHinh neu dang de trong va detect duoc tu cot soGheMay
        if (!existing.loaiHinh && detectedLoaiHinh) existing.loaiHinh = detectedLoaiHinh;
        updatedExisting++;
        updatedNames.push(existing.gian || diaDiem);
      } else {
        // Chua co gian nao khop -- tao MOI, de trong cac truong can nhap tay
        // (Luyen xac nhan qua AskUserQuestion 2026-07-21: "Thêm cả gian mới chưa có").
        store.phap_danh_hop_dong_thue.push({
          id: nextId(store, "phap_danh_hop_dong_thue_seq") || Date.now(),
          loaiHinh: detectedLoaiHinh || "",
          congTy,
          maDiemMisa: "",
          tenDiemNoiBo: "",
          khuVuc: "",
          gian: diaDiem,
          maCongTrinh: "",
          maKH: "",
          benChoThue: benChoThueSheet,
          mstBenChoThue: mstBenChoThueSheet,
          hinhThucHopTac: "",
          trangThaiHoatDong: "",
          thoiHanHopDong: "",
          tienThueThang: 0,
          ghiChu: `Import tự động từ ${SOURCE_TAG} -- CHỊ CẦN TỰ ĐIỀN: Loại hình, Thời hạn HĐ, Tiền thuê/tháng, và đổi lại tên "Gian" cho đúng quy ước (sheet ghi tên/mã thô).`,
          dieuKhoanThanhToan: "",
          hinhThucThuTien: "",
          ...fields,
          createdAt: new Date().toISOString(),
          source: SOURCE_TAG,
        });
        addedNew++;
        addedNames.push(diaDiem);
      }
    });

    save(store);
    let msg = `Đã đọc thẳng từ Google Sheet: ${thueGianRows.length} dòng thuê gian (${groups.size} địa điểm khớp được). `;
    if (updatedExisting > 0) {
      msg += `Đã cập nhật các trường từ sheet (ngày ký/hết hạn, tiền cọc, link...) cho ${updatedExisting} gian đã có: ${updatedNames.join(", ")}. `;
    }
    if (addedNew > 0) {
      msg += `Thêm ${addedNew} gian MỚI (cần tự điền Loại hình/Bên cho thuê/Thời hạn/Tiền thuê, đổi tên Gian cho đúng): ${addedNames.join(", ")}. `;
    }
    if (skippedNoSoHopDong > 0) {
      msg += `Bỏ qua ${skippedNoSoHopDong} dòng không có Số Hợp Đồng trong sheet (không đủ dữ liệu để ghép an toàn).`;
    }
    res.redirect("/phap-danh/hop-dong-thue-gian-hang?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-thue-gian-hang?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Hop Dong Thue Gian Hang MIEN BAC ----------
// Luyen, 2026-08-01: "có thêm tôi 1 trang là hợp đồng thuê gian hàng miền
// bắc ... chia từ gg sheet [Danh sách các gian] để có thể lấy ra tên gian
// với hợp đồng bên cạnh của 2 sheet Hà Nội MTD và HN KVC á nếu thiếu á bạn
// sẽ lấy trên link giống miền nam [...] sheet HN-Chị Nhung". Khac voi trang
// Nam o tren (nguon CHINH la sheet hop dong "THEO DOI HD HN-HCM", suy gian tu
// dia diem), trang nay nguon CHINH la sheet "Danh sách các gian" (cung file
// dung o Hoa Don Dau Vao, xem parseGianSheetWorkbook trong
// utils/hoaDonDauVaoEnrich.js) -- loc rieng 2 tab "HÀ NỘI MTD"/"HÀ NỘI KVC"
// lay danh sach gian That + "Mã Điểm Thuê" (dung lam "gian" xuyen suot app)
// + ten/MST ben cho thue co san. BO SUNG THEM (neu khop duoc) chi tiet hop
// dong (so HD, ngay ky/het han, tien coc, link...) tu tab "HN-Chị Nhung"
// (cung 1 file voi sheet Hop Dong NCC, khac gid) qua parseHopDongHcmWorkbook
// voi opts.anyRegion=true -- CHUA the xac nhan quy uoc cot "KV" cua tab nay
// (tuong tu ket noi Google bi chan boi buoc xac minh tai khoan khac, xem
// utils/hopDongHcmParser.js) nen dung an toan nhat la bo qua han loc do, Luyen
// kiem tra ky ket qua lan chay dau tien (neu 0 dong hoac sai gian, bao lai).
const GIAN_SHEET_MIEN_BAC_XLSX_URL =
  process.env.HOA_DON_DAU_VAO_GIAN_SHEET_URL ||
  "https://docs.google.com/spreadsheets/d/1Fd5v128o6eVuzHqtYzV5Eef62YtW6x5X/export?format=xlsx";
const NCC_SHEET_MIEN_BAC_CSV_URL =
  process.env.HOP_DONG_MIEN_BAC_SHEET_CSV_URL ||
  "https://docs.google.com/spreadsheets/d/1Kh_IDjW580UFwyrB1ESqdAZvMLLe5sffLrjRUxjS-DI/export?format=csv&gid=480440647";

function isHaNoiSheetLabel(label) {
  return normVN(label).includes("ha noi");
}

function loaiHinhFromSheetLabel(label) {
  const n = normVN(label);
  if (n.includes("mtd")) return "Máy tự động";
  if (n.includes("kvc")) return "Khu vui chơi";
  return "";
}
// Luyen, 2026-08-12: tu dong suy loaiHinh tu cot "So ghe/ May tu dong/ Dien tich"
// tren Google Sheet HCM -- CGV ghe o day se co gia tri nhu "02 ghe", may ATM
// se co "01 may"... Neu khong co thong tin nay thi fallback ve sheet label.
function loaiHinhFromSoGheMay(soGheMayDienTich, sheetLabelFallback) {
  const v = normVN(soGheMayDienTich || "");
  if (v.includes("ghe")) return "Ghế";
  if (v.includes("may")) return "Máy tự động";
  return loaiHinhFromSheetLabel(sheetLabelFallback || "");
}

// Khoa ghep AN TOAN gian mien Bac da co: "gian" (= Ma Diem Thue, nguon goc
// xac dinh gian nay, khong nhu Nam dung soHopDongHCM vi nguon Nam la hop dong)
// + congTy.
function mienBacGroupKey(gian, congTy) {
  return (gian || "") + "||" + (congTy || "");
}

router.get("/phap-danh/hop-dong-thue-gian-hang-mien-bac", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const loaiFilter = req.query.loai || "";
  const thangFilter = resolveThangFilter(req);
  const todayStr = new Date().toISOString().slice(0, 10);
  let rows = store.phap_danh_hop_dong_thue
    .map(ensureThueDefaults)
    .filter((r) => (r.congTy || "kh_cu") === activeCompany && r.mien === "bac");
  if (loaiFilter) rows = rows.filter((r) => r.loaiHinh === loaiFilter);
  if (thangFilter) rows = rows.filter((r) => monthOverlaps(r, thangFilter));
  rows.forEach((r) => {
    r.trangThaiHD = computeTrangThaiHD(r, todayStr);
  });
  rows.sort((a, b) => (a.loaiHinh === b.loaiHinh ? a.gian.localeCompare(b.gian) : a.loaiHinh.localeCompare(b.loaiHinh)));
  const counts = {
    "Khu vui chơi": store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && r.mien === "bac" && r.loaiHinh === "Khu vui chơi"
    ).length,
    "Máy tự động": store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && r.mien === "bac" && r.loaiHinh === "Máy tự động"
    ).length,
  };
  res.render("phapdanh-hopdong-thue-mienbac", {
    userName: req.session.userName,
    rows,
    loaiFilter,
    thangFilter,
    counts,
    totalForCompany: store.phap_danh_hop_dong_thue.filter(
      (r) => (r.congTy || "kh_cu") === activeCompany && r.mien === "bac"
    ).length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.get("/phap-danh/hop-dong-thue-gian-hang-mien-bac/export.xlsx", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const loaiFilter = req.query.loai || "";
  const thangFilter = req.query.thang || "";
  const todayStr = new Date().toISOString().slice(0, 10);
  let rows = store.phap_danh_hop_dong_thue
    .map(ensureThueDefaults)
    .filter((r) => (r.congTy || "kh_cu") === activeCompany && r.mien === "bac");
  if (loaiFilter) rows = rows.filter((r) => r.loaiHinh === loaiFilter);
  if (thangFilter) rows = rows.filter((r) => monthOverlaps(r, thangFilter));
  rows.forEach((r) => {
    r.trangThaiHD = computeTrangThaiHD(r, todayStr);
  });
  rows.sort((a, b) => (a.loaiHinh === b.loaiHinh ? a.gian.localeCompare(b.gian) : a.loaiHinh.localeCompare(b.loaiHinh)));

  const exportRows = rows.map((r) => ({
    "Loại hình": r.loaiHinh,
    Gian: r.gian,
    "Tên điểm nội bộ": r.tenDiemNoiBo,
    "Mã công trình": r.maCongTrinh,
    "Mã KH": r.maKH,
    "Bên cho thuê": r.benChoThue,
    "MST bên cho thuê": r.mstBenChoThue,
    "Hình thức HT": r.hinhThucHopTac,
    "Hình thức thu tiền": r.hinhThucThuTien,
    "Trạng thái HĐ": r.trangThaiHD,
    "Thời hạn hợp đồng": r.thoiHanHopDong,
    "Tiền thuê/tháng": r.tienThueThang,
    "Số HĐ (sheet HN)": r.soHopDongHCM,
    "Ngày ký HĐ": r.ngayKyHD,
    "Ngày bắt đầu HĐ": r.ngayBatDauHD,
    "Ngày hết hạn HĐ": r.ngayHetHanHD,
    "Tiền cọc đảm bảo": r.tienCocDamBao,
    "Tiền cọc thi công": r.tienCocThiCong,
    "Tiền thuê/tháng (sheet HN)": r.tongTienThueThangHCM,
    "Điều khoản thanh toán": r.dieuKhoanThanhToan,
    "Link HĐ đủ dấu": r.linkHopDongDuDau,
    "Link HĐ chưa đủ dấu": r.linkHopDongChuaDuDau,
    "Ghi chú": r.ghiChu,
  }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Thue gian hang MB");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader(
    "Content-Disposition",
    `attachment; filename=hop-dong-thue-gian-hang-mien-bac-${activeCompany}.xlsx`
  );
  res.send(buf);
});

router.post("/phap-danh/hop-dong-thue-gian-hang-mien-bac", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const {
      loaiHinh,
      gian,
      tenDiemNoiBo,
      maCongTrinh,
      maKH,
      benChoThue,
      mstBenChoThue,
      hinhThucHopTac,
      trangThaiHoatDong,
      thoiHanHopDong,
      tienThueThang,
      ghiChu,
      linkHopDongChuaDuDau,
    } = req.body;
    if (!gian || !gian.trim()) throw new Error("Thiếu Gian/Mặt bằng.");
    const amt = tienThueThang ? Number(String(tienThueThang).replace(/[^\d]/g, "")) : 0;
    store.phap_danh_hop_dong_thue.push({
      id: nextId(store, "phap_danh_hop_dong_thue_seq") || Date.now(),
      loaiHinh: loaiHinh || "",
      congTy: activeCompany,
      maDiemMisa: "",
      tenDiemNoiBo: (tenDiemNoiBo || "").trim(),
      khuVuc: "",
      gian: gian.trim(),
      maCongTrinh: (maCongTrinh || "").trim(),
      maKH: (maKH || "").trim(),
      benChoThue: (benChoThue || "").trim(),
      mstBenChoThue: (mstBenChoThue || "").trim(),
      hinhThucHopTac: (hinhThucHopTac || "").trim(),
      trangThaiHoatDong: (trangThaiHoatDong || "").trim(),
      thoiHanHopDong: (thoiHanHopDong || "").trim(),
      tienThueThang: amt,
      ghiChu: (ghiChu || "").trim(),
      linkHopDongChuaDuDau: (linkHopDongChuaDuDau || "").trim(),
      mien: "bac",
      createdAt: new Date().toISOString(),
      source: "nhap tay",
    });
    save(store);
    res.redirect("/phap-danh/hop-dong-thue-gian-hang-mien-bac?success=" + encodeURIComponent("Đã lưu hợp đồng."));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-thue-gian-hang-mien-bac?error=" + encodeURIComponent(e.message));
  }
});

router.post("/phap-danh/hop-dong-thue-gian-hang-mien-bac/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  store.phap_danh_hop_dong_thue = store.phap_danh_hop_dong_thue.filter((r) => String(r.id) !== req.params.id);
  save(store);
  res.redirect("/phap-danh/hop-dong-thue-gian-hang-mien-bac?success=" + encodeURIComponent("Đã xóa hợp đồng."));
});

router.post("/phap-danh/hop-dong-thue-gian-hang-mien-bac/cap-nhat-tu-sheet", requireAdmin, async (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const gianResp = await fetch(GIAN_SHEET_MIEN_BAC_XLSX_URL);
    if (!gianResp.ok) {
      throw new Error(`Không đọc được Google Sheet "Danh sách các gian" (mã lỗi ${gianResp.status}).`);
    }
    const gianBuf = Buffer.from(await gianResp.arrayBuffer());
    const { rows: allGianRows, sheetsRead } = parseGianSheetWorkbook(gianBuf);
    const gianRowsMienBac = allGianRows.filter((r) => isHaNoiSheetLabel(r.sheetLabel) && r.maDiemThue);
    if (gianRowsMienBac.length === 0) {
      throw new Error(
        `Không tìm thấy dòng nào ở 2 tab Hà Nội có "Mã Điểm Thuê" (đã đọc được các tab: ${sheetsRead.join(", ")}).`
      );
    }

    // Chi tiet hop dong bo sung (so HD, ngay, tien coc, link...) -- KHONG bat
    // buoc, loi/khong doc duoc van tao/cap nhat gian binh thuong, chi thieu
    // phan bo sung nay (Luyen tu dien tay sau).
    let hdDetailRows = [];
    let hdSheetWarning = "";
    try {
      const hdResp = await fetch(NCC_SHEET_MIEN_BAC_CSV_URL);
      if (hdResp.ok) {
        const hdBuf = Buffer.from(await hdResp.arrayBuffer());
        const { rows } = parseHopDongHcmWorkbook(hdBuf, { anyRegion: true });
        hdDetailRows = rows.filter((r) => !isNccRow(r));
      } else {
        hdSheetWarning = `Không đọc được sheet "HN-Chị Nhung" (mã lỗi ${hdResp.status}) -- vẫn tạo/cập nhật gian, chỉ thiếu phần chi tiết hợp đồng bổ sung.`;
      }
    } catch (e2) {
      hdSheetWarning = `Lỗi đọc sheet "HN-Chị Nhung" (${e2.message}) -- vẫn tạo/cập nhật gian, chỉ thiếu phần chi tiết hợp đồng bổ sung.`;
    }

    function findHdDetailMatch(mst, ten) {
      const mstNorm = (mst || "").replace(/\D/g, "");
      if (mstNorm) {
        const byMst = hdDetailRows.find((r) => (r.mstKH || "").replace(/\D/g, "") === mstNorm);
        if (byMst) return byMst;
      }
      const tenNorm = normVN(ten);
      if (!tenNorm) return null;
      const byName = hdDetailRows.filter(
        (r) => normVN(r.tenKH) === tenNorm || (r.diaDiem && normVN(r.diaDiem).includes(tenNorm))
      );
      return byName.length === 1 ? byName[0] : null;
    }

    const existingByGian = new Map();
    store.phap_danh_hop_dong_thue.forEach((r) => {
      if (r.mien === "bac" && r.gian) existingByGian.set(mienBacGroupKey(r.gian, r.congTy || "kh_cu"), r);
    });

    const SOURCE_TAG = "GG Sheet (Miền Bắc) " + new Date().toISOString().slice(0, 10);
    let addedNew = 0;
    let updatedExisting = 0;
    let matchedHdDetail = 0;
    const addedNames = [];
    const updatedNames = [];

    gianRowsMienBac.forEach((g) => {
      const congTy = g.congTy || "kh_cu";
      const key = mienBacGroupKey(g.maDiemThue, congTy);
      const hd = findHdDetailMatch(g.mstKhachHang, g.tenKhachHang);
      if (hd) matchedHdDetail++;
      const hdFields = hd
        ? {
            soHopDongHCM: hd.soHopDong || "",
            ngayKyHD: hd.ngayKyHD || "",
            ngayBatDauHD: hd.ngayBatDauHD || "",
            ngayHetHanHD: hd.ngayHetHan || "",
            baoHanHD: hd.baoHanHD || "",
            tienCocDamBao: hd.tienCocDamBao !== null && hd.tienCocDamBao !== undefined ? hd.tienCocDamBao : "",
            tienCocThiCong: hd.tienCocThiCong !== null && hd.tienCocThiCong !== undefined ? hd.tienCocThiCong : "",
            tongTienThueThangHCM:
              hd.tongTienThueThang !== null && hd.tongTienThueThang !== undefined
                ? String(hd.tongTienThueThang).trim()
                : "",
            linkHopDongDuDau: hd.linkHDduDau || "",
            linkHopDongChuaDuDau: hd.linkHDchuaDuDau || "",
            sheetHCMThamKhao: `HN-Chị Nhung, khớp qua ${g.mstKhachHang ? "MST" : "tên"}`,
          }
        : {};

      const existing = existingByGian.get(key);
      if (existing) {
        Object.assign(existing, hdFields);
        if (!existing.tenDiemNoiBo && g.gianHang) existing.tenDiemNoiBo = g.gianHang;
        if (!existing.benChoThue && g.tenKhachHang) existing.benChoThue = g.tenKhachHang;
        if (!existing.mstBenChoThue && g.mstKhachHang) existing.mstBenChoThue = g.mstKhachHang;
        if (!existing.hinhThucHopTac && g.hinhThucHopTac) existing.hinhThucHopTac = g.hinhThucHopTac;
        if (!existing.loaiHinh && loaiHinhFromSheetLabel(g.sheetLabel)) {
          existing.loaiHinh = loaiHinhFromSheetLabel(g.sheetLabel);
        }
        updatedExisting++;
        updatedNames.push(existing.gian);
      } else {
        store.phap_danh_hop_dong_thue.push({
          id: nextId(store, "phap_danh_hop_dong_thue_seq") || Date.now(),
          loaiHinh: loaiHinhFromSheetLabel(g.sheetLabel),
          congTy,
          maDiemMisa: "",
          tenDiemNoiBo: g.gianHang || "",
          khuVuc: "",
          gian: g.maDiemThue,
          maCongTrinh: "",
          maKH: "",
          benChoThue: g.tenKhachHang || "",
          mstBenChoThue: g.mstKhachHang || "",
          hinhThucHopTac: g.hinhThucHopTac || "",
          trangThaiHoatDong: "",
          thoiHanHopDong: "",
          tienThueThang: 0,
          ghiChu: hd
            ? `Import tự động từ ${SOURCE_TAG} (khớp được chi tiết hợp đồng từ HN-Chị Nhung).`
            : `Import tự động từ ${SOURCE_TAG} -- CHỊ CẦN TỰ ĐIỀN: Thời hạn HĐ, Tiền thuê/tháng (chưa khớp được chi tiết hợp đồng từ HN-Chị Nhung).`,
          dieuKhoanThanhToan: "",
          hinhThucThuTien: "",
          mien: "bac",
          ...hdFields,
          createdAt: new Date().toISOString(),
          source: SOURCE_TAG,
        });
        addedNew++;
        addedNames.push(g.maDiemThue);
      }
    });

    save(store);
    let msg = `Đã đọc từ Google Sheet "Danh sách các gian" (${sheetsRead.join(
      ", "
    )}): ${gianRowsMienBac.length} gian miền Bắc. Khớp được chi tiết hợp đồng (số HĐ, ngày, tiền cọc...) cho ${matchedHdDetail} gian từ tab "HN-Chị Nhung". `;
    if (updatedExisting > 0) msg += `Đã cập nhật ${updatedExisting} gian đã có: ${updatedNames.join(", ")}. `;
    if (addedNew > 0) msg += `Thêm ${addedNew} gian MỚI: ${addedNames.join(", ")}. `;
    if (hdSheetWarning) msg += hdSheetWarning;
    res.redirect("/phap-danh/hop-dong-thue-gian-hang-mien-bac?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-thue-gian-hang-mien-bac?error=" + encodeURIComponent(e.message));
  }
});

// ---------- Hop Dong NCC ----------
// Luyen, 2026-07-20: mo rong tu khung rong ban dau -- da import 48 hop dong
// NCC that tu Google Sheet "THEO DOI HD HN-HCM" (cac dong co "Thuoc BP" =
// "Khac" hoac "Mua ban vocher", tuc KHONG phai hop dong thue gian hang --
// day la hop dong mua hang hoa/dich vu/van chuyen/thi cong... voi nha cung
// cap). Moi hop dong = 1 dong (gop HD goc + phu luc/BBNT/BBBG cung 1 NCC +
// cung cong ty vao 1 ban ghi, chi tiet tung dong nguon nam trong "chiTiet").
// Cong ty (kh_cu/kh_moi) lay truc tiep tu cot "Cong ty" co san trong sheet
// (K&H moi (705) -> kh_moi, K&H cu (989) -> kh_cu); day cung khop voi quy tac
// Luyen dua ra ("ten cong ty ky HD co chu 'Dich vu' la KH Cu, khong co la KH
// Moi") de doi chieu khi can nhap tay hop dong moi sau nay.
function ensureNccDefaults(row) {
  return Object.assign(
    {
      tenNCC: "",
      tenDayDuNCC: "",
      mstNCC: "",
      diaChiNCC: "",
      congTy: "",
      hangHoaMua: "",
      noiDung: "",
      phanLoaiHD: "",
      soHopDong: "",
      ngayKy: "",
      ngayBatDauHD: "",
      ngayHetHan: "",
      baoHanHD: "",
      giaTriHopDong: 0,
      giaTriHopDongRaw: "",
      linkHopDong: "",
      soTKNH: "",
      khachMoTaiNH: "",
      khuVuc: "",
      chiTiet: [],
      ghiChu: "",
      // Nhan, 2026-09-17: them de dien "Đại diện"/"Chức vụ" Bên A vao Bien Ban
      // (xem utils/chungTuNccDoc.js) -- CHI dung khi hop dong NCC da map hoac
      // duoc them thu cong tu trang Ho So > Hoa Don NCC (nut "+ Thêm hợp đồng
      // NCC"), khong tu doan/lay tu noi khac.
      daiDien: "",
      chucVu: "",
    },
    row
  );
}

function computeTrangThaiNCC(r, todayStr) {
  if (r.ngayHetHan) return r.ngayHetHan < todayStr ? "Hết hạn" : "Còn hoạt động";
  if (r.baoHanHD && /qua han|quá hạn/i.test(r.baoHanHD)) return "Hết hạn";
  return "Không rõ";
}

router.get("/phap-danh/hop-dong-ncc", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const todayStr = new Date().toISOString().slice(0, 10);
  let rows = store.phap_danh_hop_dong_ncc.map(ensureNccDefaults);
  rows = rows.filter((r) => !r.congTy || r.congTy === activeCompany);
  rows.forEach((r) => {
    r.trangThaiNCC = computeTrangThaiNCC(r, todayStr);
  });
  rows.sort((a, b) => (a.ngayHetHan < b.ngayHetHan ? 1 : a.ngayHetHan > b.ngayHetHan ? -1 : a.tenNCC.localeCompare(b.tenNCC)));

  // Luyen, 2026-08-24: build hoaDonDaThanhToan dong tu chi_phi -- join theo
  // tenNCC (normalize lowercase trim), lay cac dong daChi=true, tra them
  // linkHoaDon + mien de modal co the mo file HĐ truc tiep va link chi phi
  // chi hien 1 ban ghi (soloId).
  const norm = (v) => (v || "").trim().toLowerCase();
  // Build lookup: tenNCC -> [chi_phi records]
  const chiPhiByNcc = {};
  (store.chi_phi || []).forEach((cp) => {
    const key = norm(cp.ncc);
    if (!key) return;
    if (!chiPhiByNcc[key]) chiPhiByNcc[key] = [];
    chiPhiByNcc[key].push(cp);
  });
  rows.forEach((r) => {
    const key = norm(r.tenNCC);
    const matched = chiPhiByNcc[key] || [];
    r.hoaDonDaThanhToan = matched
      .filter((cp) => cp.daChi)
      .sort((a, b) => (a.ngay < b.ngay ? 1 : -1))
      .map((cp) => ({
        ngay: cp.ngay || "",
        soHoaDon: cp.soHoaDon || "",
        soTien: cp.soTien || 0,
        chiPhiId: cp.id,
        linkHoaDon: cp.linkHoaDon || "",
        mien: cp.mien || "nam",
      }));
  });

  res.render("phapdanh-hopdong-ncc", {
    userName: req.session.userName,
    rows,
    totalForCompany: rows.length,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Luyen, 2026-07-21: "có cái nút tự cập nhật cho tôi nhá" -- upload lai file
// Google Sheet "THEO DÕI HĐ HN-HCM" (export xlsx) de nap hop dong NCC MOI
// (thuoc BP = "Khac"/"Mua ban vocher") ma sheet co nhung chua co trong 48 hop
// dong da import truoc do, VA phat hien phu luc/HD moi cho cac NCC DA CO.
//
// Luyen, 2026-07-21 (lan 2): "cập nhật LẠI TẤT CẢ các trường cho dòng cũ theo
// sheet mới nhất" -- doi tu che do "chi append, khong bao gio ghi de" (ban
// dau de tranh lap lai bug AMBD/Binh Duong khop nham) sang GHI DE TOAN BO cho
// moi NCC/hop dong da co MOI LAN upload, vi Luyen xac nhan van muon the du
// biet co rui ro mat sua tay giua 2 lan sheet chua kip cap nhat. Van GIU
// NGUYEN co che nhom theo CAP (diaDiem, congTy) (khong chi diaDiem) de tranh
// lap lai chinh bug AMBD do (2 NCC KHAC NHAU trung ten dia diem o 2 cong ty).
// Luyen, 2026-07-21: BUG FIX -- ban dau nhom/khop CHI theo diaDiem (tenNCC),
// nhung du lieu goc co nhieu diaDiem TRUNG TEN nhung la 2 NCC KHAC NHAU
// cho 2 cong ty (vd "Nguyen Phuoc" vua co hop dong rieng cho KH Moi vua
// co hop dong rieng cho KH Cu). Neu chi khop theo diaDiem, Map se bi va
// cham (ghi de) va co the gan nham/gop nham chi tiet cua cong ty nay vao
// ho so cua cong ty kia -- da phat hien loi nay lam contaminate 7/9 ho so
// trong lan chay dau tien (Nguyen Phuoc, Bluecom, Toan Phat, Lam Sang,
// AIKIA, Tien Phat, Tra Sua Gia) va da revert thu cong ve dung du lieu
// truoc do. Tu nay nhom/khop theo CAP (diaDiem, congTy) de tach rieng 2
// cong ty ngay ca khi trung ten dia diem.
function nccGroupKey(diaDiem, congTy) {
  return (diaDiem || "(không có địa điểm)") + "||" + (congTy || "");
}

// Xay dung toan bo cac truong tong hop tu 1 nhom groupRows (dung chung
// cho ca NCC moi va NCC da co, vi gio ca 2 truong hop deu GHI DE giong het
// nhau -- chi khac o cho tao moi object hay ghi vao object da co san).
function buildNccFieldsFromRows(diaDiem, groupRows) {
  const first = groupRows[0];
  let ngayHetHanMax = "";
  let baoHanCuaMax = "";
  groupRows.forEach((r) => {
    if (r.ngayHetHan && r.ngayHetHan > ngayHetHanMax) {
      ngayHetHanMax = r.ngayHetHan;
      baoHanCuaMax = r.baoHanHD;
    }
  });
  const soTienList = groupRows.map((r) => r.soTienHD).filter((v) => v !== null && v !== undefined && v !== "");
  const giaTriHopDong = soTienList.reduce((sum, v) => (typeof v === "number" ? sum + v : sum), 0);
  const linkHopDong = groupRows.map((r) => r.linkHDduDau).find((v) => v) || groupRows.map((r) => r.linkHDchuaDuDau).find((v) => v) || "";
  const congTy = groupRows.map((r) => r.congTy).find((v) => v) || "";
  return {
    tenNCC: diaDiem,
    tenDayDuNCC: first.tenKH,
    mstNCC: first.mstKH,
    diaChiNCC: first.diaChiKH,
    congTy,
    hangHoaMua: first.noiDungHD,
    noiDung: first.noiDungHD,
    phanLoaiHD: first.phanLoaiHD,
    soHopDong: first.soHopDong,
    ngayKy: first.ngayKyHD,
    ngayBatDauHD: first.ngayBatDauHD,
    ngayHetHan: ngayHetHanMax || first.ngayHetHan,
    baoHanHD: baoHanCuaMax || first.baoHanHD,
    giaTriHopDong,
    giaTriHopDongRaw: soTienList.join(" | "),
    linkHopDong,
    soTKNH: groupRows.map((r) => r.soTKNHkhach).find((v) => v) || "",
    khachMoTaiNH: groupRows.map((r) => r.khachMoTaiNH).find((v) => v) || "",
    khuVuc: "HCM",
    chiTiet: groupRows.map((r) => ({
      noiDung: r.noiDungHD,
      phanLoaiHD: r.phanLoaiHD,
      soHopDong: r.soHopDong,
      ngayKy: r.ngayKyHD,
      ngayBatDau: r.ngayBatDauHD,
      ngayHetHan: r.ngayHetHan,
      baoHanHD: r.baoHanHD,
      soTienHD: r.soTienHD,
      tongTienThang: "",
      link: r.linkHDduDau || r.linkHDchuaDuDau || "",
    })),
  };
}

// Luyen, 2026-07-21 (lan 2): "cập nhật LẠI TẤT CẢ các trường cho dòng cũ theo
// sheet mới nhất" -- doi tu che do "chi append, khong bao gio ghi de" (ban
// dau de tranh lap lai bug AMBD/Binh Duong khop nham) sang GHI DE TOAN BO cho
// moi NCC/hop dong da co MOI LAN cap nhat (upload file hoac doc thang tu
// sheet), vi Luyen xac nhan van muon the du biet co rui ro mat sua tay giua
// 2 lan sheet chua kip cap nhat.
// Luyen, 2026-07-21 (lan 3): tach logic nay thanh ham dung chung cho CA route
// upload file VA route moi doc thang tu Google Sheet link (xem ben duoi) --
// "muốn khi tôi nhấn cập nhật ... hợp đồng nó sẽ cập nhật từ gg sheet cho tôi".
function applyNccRowsToStore(store, nccRows, sourceLabel) {
  const groups = new Map();
  nccRows.forEach((r) => {
    const key = nccGroupKey(r.diaDiem, r.congTy);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });

  const existingByKey = new Map();
  store.phap_danh_hop_dong_ncc.forEach((r) => existingByKey.set(nccGroupKey(r.tenNCC, r.congTy), r));

  let addedNew = 0;
  let updatedExisting = 0;
  const addedNames = [];
  const updatedNames = [];

  groups.forEach((groupRows, key) => {
    const diaDiem = groupRows[0].diaDiem || "(không có địa điểm)";
    const existing = existingByKey.get(key);
    const fields = buildNccFieldsFromRows(diaDiem, groupRows);
    if (!existing) {
      // NCC hoan toan moi
      store.phap_danh_hop_dong_ncc.push({
        id: nextId(store, "phap_danh_hop_dong_ncc_seq") || Date.now(),
        ...fields,
        ghiChu: `Import tự động từ ${sourceLabel} (Thuộc BP: Khác/Mua bán vocher), ${groupRows.length} dòng nguồn -- chị kiểm tra lại các trường tổng hợp, dữ liệu lấy tự động chưa qua rà soát tay như 48 hợp đồng ban đầu.`,
        createdAt: new Date().toISOString(),
        source: sourceLabel,
      });
      addedNew++;
      addedNames.push(diaDiem);
    } else {
      // NCC da co -- Luyen xac nhan 2026-07-21 muon GHI DE lai toan bo cac
      // truong (ke ca link/han/gia tri/chi tiet) theo dung sheet moi nhat
      // moi lan bam "Cap nhat hop dong", giu nguyen id/createdAt goc.
      Object.assign(existing, fields);
      existing.ghiChu = `Đã cập nhật lại từ ${sourceLabel} (Thuộc BP: Khác/Mua bán vocher), ${groupRows.length} dòng nguồn.`;
      existing.source = sourceLabel;
      updatedExisting++;
      updatedNames.push(diaDiem);
    }
  });

  return { groupsSize: groups.size, addedNew, updatedExisting, addedNames, updatedNames };
}

function buildNccResultMessage(prefix, nccRowsLength, result) {
  let msg = `${prefix} ${nccRowsLength} dòng NCC (${result.groupsSize} địa điểm). `;
  if (result.addedNew > 0) msg += `Thêm ${result.addedNew} NCC mới: ${result.addedNames.join(", ")}. `;
  if (result.updatedExisting > 0) msg += `Đã cập nhật lại ${result.updatedExisting} NCC đã có theo sheet mới nhất: ${result.updatedNames.join(", ")}. `;
  if (result.addedNew === 0 && result.updatedExisting === 0) msg += "Không có gì mới so với dữ liệu hiện tại.";
  return msg;
}

router.post("/phap-danh/hop-dong-ncc/upload", requireAdmin, upload.single("file"), (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    if (!req.file) throw new Error("Vui lòng chọn 1 file để tải lên.");
    const { rows } = parseHopDongHcmWorkbook(req.file.buffer);
    const nccRows = rows.filter(isNccRow);
    if (nccRows.length === 0) {
      throw new Error('Không tìm thấy dòng nào có "Thuộc BP" = Khác/Mua bán vocher trong file này.');
    }
    const UPLOAD_TAG = "upload " + req.file.originalname + " " + new Date().toISOString().slice(0, 10);
    const result = applyNccRowsToStore(store, nccRows, UPLOAD_TAG);
    save(store);
    const msg = buildNccResultMessage(`Đã quét từ "${req.file.originalname}":`, nccRows.length, result);
    res.redirect("/phap-danh/hop-dong-ncc?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-ncc?error=" + encodeURIComponent(e.message));
  }
});

// Luyen, 2026-07-21 (lan 3): "muốn khi tôi nhấn cập nhật ... hợp đồng nó sẽ
// cập nhật từ gg sheet cho tôi á" -- doc THANG tu Google Sheet qua export CSV
// (khong can tai file ve roi tai len nua). Luyen da dong y doi quyen chia se
// sheet "THEO DÕI HĐ HN-HCM" sang "Bất kỳ ai có link đều xem được" de server
// doc duoc ma khong can dang nhap Google. Neu Luyen doi sang sheet/gid khac
// sau nay, chi can dat bien moi truong HOP_DONG_NCC_SHEET_CSV_URL tren
// Railway (khong can sua code).
const NCC_SHEET_CSV_URL =
  process.env.HOP_DONG_NCC_SHEET_CSV_URL ||
  "https://docs.google.com/spreadsheets/d/1Kh_IDjW580UFwyrB1ESqdAZvMLLe5sffLrjRUxjS-DI/export?format=csv&gid=299256887";

router.post("/phap-danh/hop-dong-ncc/cap-nhat-tu-sheet", requireAdmin, async (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const resp = await fetch(NCC_SHEET_CSV_URL);
    if (!resp.ok) {
      throw new Error(
        `Không đọc được Google Sheet (mã lỗi ${resp.status}). Kiểm tra lại sheet đã chia sẻ "Bất kỳ ai có link đều xem được" chưa, hoặc link/gid có bị đổi không.`
      );
    }
    const buf = Buffer.from(await resp.arrayBuffer());
    const { rows } = parseHopDongHcmWorkbook(buf);
    const nccRows = rows.filter(isNccRow);
    if (nccRows.length === 0) {
      throw new Error('Không tìm thấy dòng nào có "Thuộc BP" = Khác/Mua bán vocher trong Google Sheet.');
    }
    const SOURCE_TAG = "GG Sheet " + new Date().toISOString().slice(0, 10);
    const result = applyNccRowsToStore(store, nccRows, SOURCE_TAG);
    save(store);
    const msg = buildNccResultMessage("Đã đọc thẳng từ Google Sheet:", nccRows.length, result);
    res.redirect("/phap-danh/hop-dong-ncc?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-ncc?error=" + encodeURIComponent(e.message));
  }
});

router.post("/phap-danh/hop-dong-ncc", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  try {
    const { tenNCC, tenTrenHD, noiDung, hangHoaMua, phanLoaiHD, soHopDong, ngayKy, ngayHetHan, giaTriHopDong, linkHopDong, ghiChu, daiDien, chucVu } = req.body;
    if (!tenNCC || !tenNCC.trim()) throw new Error("Thiếu Tên NCC.");
    const amt = giaTriHopDong ? Number(String(giaTriHopDong).replace(/[^\d]/g, "")) : 0;
    store.phap_danh_hop_dong_ncc.push({
      id: nextId(store, "phap_danh_hop_dong_ncc_seq") || Date.now(),
      tenNCC: tenNCC.trim(),
      tenTrenHD: (tenTrenHD || "").trim(),
      congTy: activeCompany,
      hangHoaMua: (hangHoaMua || "").trim(),
      noiDung: (noiDung || "").trim(),
      phanLoaiHD: (phanLoaiHD || "").trim(),
      soHopDong: (soHopDong || "").trim(),
      ngayKy: ngayKy || "",
      ngayHetHan: ngayHetHan || "",
      giaTriHopDong: amt,
      linkHopDong: (linkHopDong || "").trim(),
      chiTiet: [],
      ghiChu: (ghiChu || "").trim(),
      daiDien: (daiDien || "").trim(),
      chucVu: (chucVu || "").trim(),
      createdAt: new Date().toISOString(),
      source: "nhap tay",
    });
    save(store);
    res.redirect("/phap-danh/hop-dong-ncc?success=" + encodeURIComponent("Đã lưu hợp đồng."));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-ncc?error=" + encodeURIComponent(e.message));
  }
});

router.post("/phap-danh/hop-dong-ncc/:id/delete", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  store.phap_danh_hop_dong_ncc = store.phap_danh_hop_dong_ncc.filter((r) => String(r.id) !== req.params.id);
  save(store);
  res.redirect("/phap-danh/hop-dong-ncc?success=" + encodeURIComponent("Đã xóa hợp đồng."));
});

// Chi Nhan, 2026-07-31: "chõo pháp danh thêm 1 chõo lưu đối soát doanh thu chi
// sẻ theo tháng cho tôi" -- hoi qua AskUserQuestion, Chi Nhan chon "Cả 2": (1)
// bao cao TU DONG, (2) THEM 1 o ghi chu nhap tay rieng cho tung thang (luu
// vao store, doc lap voi so tu dong).
//
// QUAN TRONG: ban dau dinh loc theo "TK Co = 1388" tren dong doi soat, nhung
// Luyen da doi quy uoc nay tu 2026-07-17 ("đổi xuất ra 1388 thành 131 hết" --
// xem utils/vietqrReconcile.js) nen KHONG CON dong doi soat nao mang TK Co
// 1388 nua (da kiem tra thuc te: 0/4932 dong). "Doanh thu chia se" GIO CHI
// con song qua co "doanhThuChiaSe"/dieuKhoanThanhToan tren TUNG HOP DONG THUE
// GIAN (Phap danh, xem isDoanhThuChiaSeRecord) -- dung LAI chinh xac may nay
// (giong cach routes/chi-phi.js computeTaiKhoanChiPhi dang lam cho Chi Phi),
// loc cac dong doi soat theo GIAN thuoc 1 hop dong duoc danh dau nay, roi cong
// theo thang.
// Chi Nhan, 2026-07-31: "nó sẽ chia ra làm 2 loại doanh thu chia sẻ ... cũng
// chia làm 4 loại MTD miền nam KVC miền nam MTD miền bắc KVC miền bắc" -- viet
// lai hoan toan trang nay lam 2 phan rieng biet:
//   Loai 1 "giu_tien": ben cho thue GIU TIEN, cuoi thang tru phi roi chuyen
//     khoan phan con lai -- TU DONG tinh duoc tu doi soat (giong cach lam cu),
//     pivot theo thang x 4 nhom (MTĐ/KVC lay tu loaiHinh san co, ghep voi Nam/
//     Bac tu truong "mien" moi -- xem nhomLabelFor).
//   Loai 2 "xuat_hoa_don": ben cho thue XUAT HOA DON dua tren so Chi Nhan tu
//     bao cao cho ho -- KHONG tu tinh duoc tu doi soat (khong phai tien ve
//     thang ngan hang), Chi Nhan TU LIET KE va luu vao bang rieng
//     (store.phap_danh_doanhthu_chiase_xuathoadon).
function nhomLabelFor(rec) {
  if (!rec) return "Chưa phân loại";
  const loai = rec.loaiHinh === "Khu vui chơi" ? "KVC" : rec.loaiHinh === "Máy tự động" ? "MTĐ" : "";
  const mienLabel = rec.mien === "bac" ? "Miền Bắc" : rec.mien === "nam" ? "Miền Nam" : "";
  if (!loai || !mienLabel) return "Chưa phân loại";
  return loai + " " + mienLabel;
}

router.get("/phap-danh/doanh-thu-chia-se", (req, res) => {
  const store = load();
  ensureShape(store);
  const { buildAllFlatLines } = require("../utils/overviewAggregate");
  const { isDoanhThuChiaSeRecord, buildGianAliasIndex, findContractForGianText } = require("../utils/rentPaymentMatcher");

  const gianList = store.phap_danh_hop_dong_thue || [];
  const aliasIndex = buildGianAliasIndex(gianList);
  const gianDoanhThuChiaSe = gianList.filter(isDoanhThuChiaSeRecord);
  const gianNameSet = new Set(gianDoanhThuChiaSe.map((r) => r.gian));
  function findRec(gianText) {
    if (gianNameSet.has(gianText)) return gianList.find((r) => r.gian === gianText);
    const rec = findContractForGianText(gianText, gianList, aliasIndex);
    return rec && isDoanhThuChiaSeRecord(rec) ? rec : null;
  }

  let giuTien = { nhoms: [], monthRows: [] };
  try {
    const flat = buildAllFlatLines(store).filter((l) => {
      const rec = findRec(l.gian);
      return !!rec && (rec.loaiChiaSe || "giu_tien") === "giu_tien";
    });
    const byMonthNhom = {};
    const nhomSet = new Set();
    flat.forEach((l) => {
      const nhom = nhomLabelFor(findRec(l.gian));
      nhomSet.add(nhom);
      if (!byMonthNhom[l.month]) byMonthNhom[l.month] = {};
      byMonthNhom[l.month][nhom] = (byMonthNhom[l.month][nhom] || 0) + l.gross;
    });
    const nhoms = [...nhomSet].sort();
    const monthRows = Object.keys(byMonthNhom)
      .sort()
      .reverse()
      .map((month) => {
        const cells = nhoms.map((n) => byMonthNhom[month][n] || 0);
        return {
          month,
          cells,
          total: cells.reduce((s, v) => s + v, 0),
          ghiChu: store.phap_danh_doanhthu_chiase_ghichu[month] || "",
        };
      });
    giuTien = { nhoms, monthRows };
  } catch (e) {
    console.error("Loi tong hop doanh thu chia se (giu tien):", e);
  }

  const xuatHoaDonGianOptions = gianDoanhThuChiaSe.filter((r) => r.loaiChiaSe === "xuat_hoa_don");
  const xuatHoaDonRows = (store.phap_danh_doanhthu_chiase_xuathoadon || [])
    .map((r) => ({ ...r, nhom: nhomLabelFor(findRec(r.gian)) }))
    .sort((a, b) => (a.thang < b.thang ? 1 : -1));

  res.render("phapdanh-doanhthuchiase", {
    userName: req.session.userName,
    COMPANIES,
    giuTien,
    gianDoanhThuChiaSe,
    xuatHoaDonGianOptions,
    xuatHoaDonRows,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

router.post("/phap-danh/doanh-thu-chia-se/ghi-chu", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  const { thang, ghiChu } = req.body;
  if (!thang) {
    return res.redirect("/phap-danh/doanh-thu-chia-se?error=" + encodeURIComponent("Thiếu tháng."));
  }
  store.phap_danh_doanhthu_chiase_ghichu[thang] = (ghiChu || "").trim();
  save(store);
  res.redirect("/phap-danh/doanh-thu-chia-se?success=" + encodeURIComponent("Đã lưu ghi chú."));
});

// Chi Nhan, 2026-07-31: "họ xuất hóa đơn cho mình dựa trên số mình báo cáo
// doanh thu cho bên họ tôi sẽ liệt kê và gửi đối soát bạn lưu lại cho tôi" --
// bang nhap tay cho loai 2, Chi Nhan tu dien tung dong sau khi doi chieu voi
// ben cho thue.
router.post("/phap-danh/doanh-thu-chia-se/xuat-hoa-don", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const { thang, gian, doanhThuBaoCao, soHoaDon, ngayHoaDon, soTienHoaDon, ghiChu } = req.body;
    if (!thang) throw new Error("Thiếu tháng.");
    if (!gian || !gian.trim()) throw new Error("Thiếu gian.");
    store.phap_danh_doanhthu_chiase_xuathoadon.push({
      id: nextId(store, "phap_danh_doanhthu_chiase_xuathoadon_seq") || Date.now(),
      thang,
      gian: gian.trim(),
      doanhThuBaoCao: Math.abs(parseAmount(doanhThuBaoCao)),
      soHoaDon: (soHoaDon || "").trim(),
      ngayHoaDon: ngayHoaDon || "",
      soTienHoaDon: Math.abs(parseAmount(soTienHoaDon)),
      ghiChu: (ghiChu || "").trim(),
      createdAt: new Date().toISOString(),
    });
    save(store);
    res.redirect("/phap-danh/doanh-thu-chia-se?success=" + encodeURIComponent("Đã lưu dòng xuất hóa đơn."));
  } catch (e) {
    res.redirect("/phap-danh/doanh-thu-chia-se?error=" + encodeURIComponent(e.message));
  }
});

router.post("/phap-danh/doanh-thu-chia-se/xuat-hoa-don/:id/delete", requireDataEntry, (req, res) => {
  const store = load();
  ensureShape(store);
  store.phap_danh_doanhthu_chiase_xuathoadon = store.phap_danh_doanhthu_chiase_xuathoadon.filter(
    (r) => String(r.id) !== req.params.id
  );
  save(store);
  res.redirect("/phap-danh/doanh-thu-chia-se?success=" + encodeURIComponent("Đã xóa dòng."));
});

// ---------- Hop Dong Thue Gian Tong ----------
// Nhan, 2026-08-13: "thêm 1 mục mới trong dropdown Pháp danh ... gọi là Hợp
// Đồng Thuê Gian Tổng ... có 4 sheet KVC MN, KVC MB, MTD MN, MTD MB" -- khac
// voi 2 trang Thue Gian Hang / Thue Gian Hang Mien Bac o tren (nguon la sheet
// hop dong "THEO DOI HD HN-HCM" + "Danh sach cac gian", chi 1 danh sach cho
// Nam va 1 cho Bac), trang nay nguon la 1 file Google Sheet DUY NHAT "Bản sao
// của Danh sách các gian Luyến.xlsx" voi DUNG 4 tab (KVC MN/KVC MB/MTD MN/MTD
// MB), hien thi RIENG 4 tab qua query "sheet". Da nap san 246 dong qua
// tmp_import_hopdong_tong.js (23/33/60/130 dong -- xem bao cao gui Nhan).
const HOP_DONG_THUE_TONG_SHEETS = [
  { key: "kvc-mn", storeKey: "kvc_mn", label: "KVC MN" },
  { key: "kvc-mb", storeKey: "kvc_mb", label: "KVC MB" },
  { key: "mtd-mn", storeKey: "mtd_mn", label: "MTD MN" },
  { key: "mtd-mb", storeKey: "mtd_mb", label: "MTD MB" },
];

// Tinh trang thai mau cho 1 dong, so voi ngay he thong hien tai (todayStr,
// dang "YYYY-MM-DD"): "het_han" (do) neu ngayHetHanThue < hom nay; "het_han_
// thang_nay" (cam) neu ngayHetHanThue cung nam-thang voi hom nay VA >= hom
// nay; con lai (chua co ngay, hoac het han xa hon thang nay) khong to mau.
// Chi dung khi ngayHetHanThue da parse duoc ro rang (dang "dd/mm/yyyy -
// dd/mm/yyyy") -- gian chua ro thoi han (con lai phan lon o MTD MB) se KHONG
// bi to mau, dung y "khong doan" da noi trong yeu cau.
function computeTrangThaiThueTong(r, todayStr) {
  // Nhan, 2026-09-14: "có những hợp đồng mình hủy trước thời hạn do lỗ ...
  // chấm dứt hợp đồng ... tô nền mờ nhạt ... vẫn để ... cảnh báo chấm dứt" --
  // gian bi danh dau daChamDut (chu dong huy truoc han, KHONG phai tu nhien
  // het han) uu tien HON het_han/het_han_thang_nay, hien mo nhat + canh bao
  // rieng thay vi mau do/cam (khac ly do voi het han thong thuong).
  if (r.daChamDut) return "cham_dut";
  if (!r.ngayHetHanThue) return "";
  if (r.ngayHetHanThue < todayStr) return "het_han";
  if (r.ngayHetHanThue.slice(0, 7) === todayStr.slice(0, 7)) return "het_han_thang_nay";
  return "";
}

// Nhan, 2026-08-13 (lan 2): "lọc theo tháng nhá" -- them bo loc thang giong
// het cach lam voi trang "Thue Gian Hang" (monthOverlaps ben tren), nhung
// gian tong dung ten truong rieng ngayBatDauThue/ngayHetHanThue (khac
// ngayBatDauHD/ngayHetHanHD cua trang Thue Gian Hang) nen viet ham rieng.
// Gian CHUA parse duoc ngay (phan lon MTD MB, xem chu thich sheetKey/
// computeTrangThaiThueTong o tren) LUON duoc giu lai, khong bi loc mat.
function monthOverlapsThueTong(r, thang) {
  if (!thang) return true;
  if (!r.ngayBatDauThue && !r.ngayHetHanThue) return true;
  const monthStart = thang + "-01";
  const monthEnd = thang + "-31";
  const bd = r.ngayBatDauThue || "0000-00-00";
  const hh = r.ngayHetHanThue || "9999-99-99";
  return bd <= monthEnd && hh >= monthStart;
}

// Nhan, 2026-09-14: "cho tôi cột số thứ tự ... bộ lọc ở trên là còn hoạt
// động hay lọc tên khách hàng" -- them 2 bo loc moi: q (tim theo Ten khach
// hang, khong phan biet hoa/thuong) va trangThai (Tat ca / Con hoat dong /
// Da het han / Da cham dut). Ap dung SAU khi tinh trangThaiMau (trangThai
// phu thuoc gia tri nay), va ap dung DONG BO len ca counts (giong ly do da
// ghi chu thich o duoi) de so tren nut chon sheet luon khop voi bang ben
// duoi dang hien.
function applyGianExtraFilters(rows, q, trangThai) {
  let out = rows;
  if (q) {
    const qLow = q.trim().toLowerCase();
    if (qLow) out = out.filter((r) => (r.tenKhachHang || "").toLowerCase().includes(qLow));
  }
  if (trangThai === "hoat_dong") out = out.filter((r) => !r.trangThaiMau);
  else if (trangThai === "het_han") out = out.filter((r) => r.trangThaiMau === "het_han" || r.trangThaiMau === "het_han_thang_nay");
  else if (trangThai === "cham_dut") out = out.filter((r) => r.trangThaiMau === "cham_dut");
  return out;
}

router.get("/phap-danh/hop-dong-thue-gian-tong", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const sheetParam = HOP_DONG_THUE_TONG_SHEETS.some((s) => s.key === req.query.sheet)
    ? req.query.sheet
    : "kvc-mn";
  const sheetInfo = HOP_DONG_THUE_TONG_SHEETS.find((s) => s.key === sheetParam);
  const todayStr = new Date().toISOString().slice(0, 10);
  const thangFilter = resolveThangFilter(req);
  const qFilter = (req.query.q || "").trim();
  const trangThaiFilter = ["hoat_dong", "het_han", "cham_dut"].includes(req.query.trangThai) ? req.query.trangThai : "";

  let rows = store.phap_danh_hop_dong_thue_tong.filter(
    (r) => r.sheetKey === sheetInfo.storeKey && (r.congTy === activeCompany || r.congTy === "ca_2")
  );
  if (thangFilter) rows = rows.filter((r) => monthOverlapsThueTong(r, thangFilter));
  rows.forEach((r) => {
    r.trangThaiMau = computeTrangThaiThueTong(r, todayStr);
  });
  rows = applyGianExtraFilters(rows, qFilter, trangThaiFilter);

  // Nhan, 2026-08-13 (lan 3): "có 5 mà sao đếm ra 6 vậy" -- truoc do counts
  // tren cac nut chon sheet la TONG so gian (khong loc theo thang) trong khi
  // bang ben duoi DA loc theo thang, gay lech so nhin nham la bug (giong
  // sticky-header truoc do). Chi Nhan xac nhan qua AskUserQuestion: doi counts
  // loc THEO CUNG bo loc thang voi bang ben duoi de 2 so luon khop nhau.
  const counts = {};
  HOP_DONG_THUE_TONG_SHEETS.forEach((s) => {
    let sheetRows = store.phap_danh_hop_dong_thue_tong.filter(
      (r) => r.sheetKey === s.storeKey && (r.congTy === activeCompany || r.congTy === "ca_2")
    );
    if (thangFilter) sheetRows = sheetRows.filter((r) => monthOverlapsThueTong(r, thangFilter));
    sheetRows.forEach((r) => { r.trangThaiMau = computeTrangThaiThueTong(r, todayStr); });
    sheetRows = applyGianExtraFilters(sheetRows, qFilter, trangThaiFilter);
    counts[s.key] = sheetRows.length;
  });

  res.render("phapdanh-hopdong-thue-tong", {
    userName: req.session.userName,
    rows,
    sheets: HOP_DONG_THUE_TONG_SHEETS,
    sheetParam,
    counts,
    thangFilter,
    qFilter,
    trangThaiFilter,
    error: req.query.error || null,
    success: req.query.success || null,
  });
});

// Nhan, 2026-08-13 (lan 3): "làm nút Cập nhật từ Google Sheet trên web thật"
// -- truoc gio 246 dong chi nap 1 LAN qua script tam tren may cuc bo (KHONG
// co tren Railway vi data/ khong di theo git). Nut nay doc THANG tu chinh
// Google Sheet nguon (spreadsheet "Bản sao của Danh sách các gian Luyến.xlsx",
// id 1VdTxB5Tkh_QxCfmQezfaFYdvDE7lp1NH -- xac nhan dung link Nhan gui lan 2,
// TRUNG voi link 1 ve gid nen dung link nao cung ra cung du lieu) qua gviz
// CSV (giong cach lam voi Hop Dong NCC/Thue Gian Hang o tren), KHONG can dang
// nhap rieng vi sheet da chia se cong khai. Da doi chieu ket qua ham parse
// nay TRUNG KHOP 100% voi 246 dong nhap tay ban dau (23/33/60/130, ca
// congTy breakdown tung sheet) truoc khi dua vao code that.
const HOP_DONG_THUE_TONG_SPREADSHEET_ID =
  process.env.HOP_DONG_THUE_TONG_SHEET_ID || "1VdTxB5Tkh_QxCfmQezfaFYdvDE7lp1NH";
const HOP_DONG_THUE_TONG_GIDS = {
  kvc_mn: "254418876",
  mtd_mn: "337828892",
  kvc_mb: "11361918",
  mtd_mb: "1750812066",
};

// Parser CSV THAT SU (khong chi split theo dong) -- gviz CSV co the co 1 o
// chua xuong dong that trong dau ngoac kep (gap 1 dong o sheet MTD MB), neu
// chi split("\n") don gian se cat nham 1 dong thanh 2, lam sai lech so dem.
function parseGvizCsv(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i++;
        } else inQuotes = false;
      } else cur += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ",") {
      row.push(cur);
      cur = "";
    } else if (c === "\r") {
      // bo qua
    } else if (c === "\n") {
      row.push(cur);
      cur = "";
      rows.push(row);
      row = [];
    } else cur += c;
  }
  if (cur.length > 0 || row.length > 0) {
    row.push(cur);
    rows.push(row);
  }
  return rows.filter((r) => r.length > 1 || (r[0] || "").trim() !== "");
}

function buildGvizHeaderIndex(headerRow) {
  const idx = {};
  headerRow.forEach((h, i) => {
    const key = (h || "").trim();
    if (key && !(key in idx)) idx[key] = i;
  });
  return idx;
}

function gvizCol(rowArr, idx, name) {
  const i = idx[name];
  return i === undefined ? "" : (rowArr[i] || "").trim();
}

// "KH mới"/"KH cũ"/"Cả 2" (nhu cot Phap nhan tren sheet) -> "kh_moi"/"kh_cu"/
// "ca_2"; rong hoac gia tri la khac -> "" (chua xac dinh, KHONG doan).
function congTyFromPhapNhan(v) {
  const n = (v || "").trim();
  if (n === "KH mới") return "kh_moi";
  if (n === "KH cũ") return "kh_cu";
  if (n.toLowerCase().includes("cả 2") || n.toLowerCase().includes("ca 2")) return "ca_2";
  return "";
}

// "dd/mm/yyyy - dd/mm/yyyy" -> {start, end} dang ISO "yyyy-mm-dd". Khong
// khop duoc dinh dang (thoi han ghi tu do, hoac de trong) -> "" ca 2, KHONG doan.
function parseThoiHanThueRange(raw) {
  const m = (raw || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s*-\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return { start: "", end: "" };
  const p2 = (s) => String(s).padStart(2, "0");
  return {
    start: `${m[3]}-${p2(m[2])}-${p2(m[1])}`,
    end: `${m[6]}-${p2(m[5])}-${p2(m[4])}`,
  };
}

function stripLinkQuery(v) {
  return (v || "").split("?")[0];
}

// Nhan, 2026-08-13 (lan 3): moi tab (KVC MN/MB, MTD MN/MB) co SO CỘT KHAC
// NHAU (vd MTD co them "Bên giữ tiền"/"TK NỢ"/"TK CÓ", KVC MB khong co cot
// "Địa điểm"/"Ghi Chú") -- tra cot theo TEN HEADER (khong theo vi tri cot cu
// the) de doc dung du lieu moi sheet, khong bi lech cot.
async function fetchHopDongThueTongSheet(storeKey) {
  const gid = HOP_DONG_THUE_TONG_GIDS[storeKey];
  const url = `https://docs.google.com/spreadsheets/d/${HOP_DONG_THUE_TONG_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&gid=${gid}`;
  const resp = await fetch(url);
  if (!resp.ok) {
    throw new Error(
      `Không đọc được sheet ${storeKey} (mã lỗi ${resp.status}). Kiểm tra lại sheet đã chia sẻ "Bất kỳ ai có link đều xem được" chưa.`
    );
  }
  const text = await resp.text();
  const rows = parseGvizCsv(text);
  if (rows.length === 0) return [];
  const idx = buildGvizHeaderIndex(rows[0]);
  return rows
    .slice(1)
    .map((r) => {
      const thoiHanThueRaw = gvizCol(r, idx, "Thời hạn hợp đồng");
      const { start, end } = parseThoiHanThueRange(thoiHanThueRaw);
      return {
        sheetKey: storeKey,
        congTy: congTyFromPhapNhan(gvizCol(r, idx, "Pháp nhân")),
        congTyRaw: gvizCol(r, idx, "Pháp nhân"),
        khuVuc: gvizCol(r, idx, "Khu Vực"),
        dichVu: gvizCol(r, idx, "Dịch vụ"),
        tenNoiBo: gvizCol(r, idx, "Mã Điểm Nội Bộ"),
        maCongTrinh: gvizCol(r, idx, "Mã Điểm Thuế"),
        diaDiem: gvizCol(r, idx, "Địa điểm"),
        tenKhachHang: gvizCol(r, idx, "Tên khách hàng"),
        mstKhachHang: gvizCol(r, idx, "MST khách hàng"),
        hinhThucThue: gvizCol(r, idx, "Hình Thức Hợp Tác"),
        thoiHanThueRaw,
        ngayBatDauThue: start,
        ngayHetHanThue: end,
        ghiChu: gvizCol(r, idx, "Ghi Chú"),
        linkHopDong: stripLinkQuery(gvizCol(r, idx, "Link hợp đồng")),
        trangThaiRaw: gvizCol(r, idx, "còn hạn không"),
      };
    })
    .filter((r) => r.tenNoiBo || r.diaDiem || r.tenKhachHang); // bo dong rong hoan toan
}

router.post("/phap-danh/hop-dong-thue-gian-tong/cap-nhat-tu-sheet", requireAdmin, async (req, res) => {
  const store = load();
  ensureShape(store);
  try {
    const perSheetCounts = {};
    let allNewRows = [];
    for (const s of HOP_DONG_THUE_TONG_SHEETS) {
      const rows = await fetchHopDongThueTongSheet(s.storeKey);
      perSheetCounts[s.key] = rows.length;
      allNewRows = allNewRows.concat(rows);
    }
    // Nap lai TOAN BO 4 sheet 1 luc (nut chung, khong tach rieng tung sheet)
    // -- xoa het dong CU cua CA 4 sheetKey nay, thay bang du lieu MOI doc
    // thang tu Google Sheet, giu nguyen cac bang khac (phap_danh_hop_dong_thue...)
    // khong lien quan.
    const otherRows = store.phap_danh_hop_dong_thue_tong.filter(
      (r) => !HOP_DONG_THUE_TONG_SHEETS.some((s) => s.storeKey === r.sheetKey)
    );
    allNewRows.forEach((r) => {
      r.id = nextId(store, "phap_danh_hop_dong_thue_tong_seq") || Date.now();
    });
    store.phap_danh_hop_dong_thue_tong = otherRows.concat(allNewRows);
    save(store);
    const msg =
      `Đã cập nhật từ Google Sheet: ${allNewRows.length} dòng (` +
      HOP_DONG_THUE_TONG_SHEETS.map((s) => `${s.label} ${perSheetCounts[s.key]}`).join(", ") +
      `).`;
    res.redirect("/phap-danh/hop-dong-thue-gian-tong?success=" + encodeURIComponent(msg));
  } catch (e) {
    res.redirect("/phap-danh/hop-dong-thue-gian-tong?error=" + encodeURIComponent(e.message));
  }
});

// TEMP DEBUG: dump KH Mới record fields + NCC directory
// ?ncc=SAN+BAY to filter by tenNCC keyword
router.get("/phap-danh/unc-debug", (req, res) => {
  const store = load();
  const allR = store.ho_so_tien_thue || [];
  let rows = allR.filter(r => r.company === 'kh_moi');
  let rowsCu = allR.filter(r => !r.company || r.company === 'kh_cu');
  const nccFilter = (req.query.ncc || '').trim().toLowerCase();
  if (nccFilter) {
    rows = rows.filter(r => (r.tenNCC||'').toLowerCase().includes(nccFilter) || (r.maCongTrinh||'').toLowerCase().includes(nccFilter) || (r.gian||'').toLowerCase().includes(nccFilter));
    rowsCu = rowsCu.filter(r => (r.tenNCC||'').toLowerCase().includes(nccFilter) || (r.maCongTrinh||'').toLowerCase().includes(nccFilter));
  }
  const sample = (nccFilter ? rows : rows.slice(0, 5)).map(r => ({
    maNCC: r.maNCC, tenNCC: r.tenNCC, maSoThueNCC: r.maSoThueNCC,
    maCongTrinh: r.maCongTrinh, tenCongTrinh: r.tenCongTrinh, gian: r.gian,
    thang: r.thang, soHoaDon: r.soHoaDon, dienGiai: r.dienGiai, soTienTong: r.soTienTong,
    fields: Object.keys(r).join(', ')
  }));
  const sampleCu = (nccFilter ? rowsCu : rowsCu.slice(0, 5)).map(r => ({
    company: r.company, maNCC: r.maNCC, tenNCC: r.tenNCC, maSoThueNCC: r.maSoThueNCC,
    thang: r.thang, soHoaDon: r.soHoaDon, dienGiai: r.dienGiai, soTienTong: r.soTienTong
  }));
  const nccDir = (store.danh_muc_ma_nha_cung_cap_moi || []).slice(0, 10);
  const nccDirCu = (store.danh_muc_ma_nha_cung_cap_cu || []).slice(0, 5);
  res.json({ totalKhMoi: rows.length, totalKhCu: rowsCu.length, nccFilter, sample, sampleCu, nccDir, nccDirCu });
});

// Luyen, 2026-09-24: "thêm 1 tab dưới pháp nhân hợp đồng thuê gian là UNC tiền thuê"
// Luyen, 2026-09-24: "tách ra thành 2 loại 1 là kh mới 2 là kh cũ ... chia ra theo NCC
// mã số thuế mã NCC ở trong danh mục ncc ... hiển thị tên ncc với mst và mã ncc trước"
router.get("/phap-danh/unc-tien-thue", (req, res) => {
  try {
  const store = load();
  const allRows = store.ho_so_tien_thue || [];

  // Build NCC lookup from danh mục NCC (ma = MST, ten = tên đầy đủ)
  const nccLookup = new Map();
  (store.danh_muc_ma_nha_cung_cap_moi || []).forEach((n) => {
    if (n.ma) nccLookup.set(String(n.ma).trim(), { ma: String(n.ma).trim(), ten: n.ten || "" });
  });
  (store.danh_muc_ma_nha_cung_cap_cu || []).forEach((n) => {
    if (n.ma) nccLookup.set(String(n.ma).trim(), { ma: String(n.ma).trim(), ten: n.ten || "" });
  });
  (store.danh_muc_ma_nha_cung_cap || []).forEach((n) => {
    if (n.ma) nccLookup.set(String(n.ma).trim(), { ma: String(n.ma).trim(), ten: n.ten || "" });
  });
  (store.chi_phi_ncc_list || []).forEach((n) => {
    if (n.maNCC && !nccLookup.has(String(n.maNCC).trim())) {
      nccLookup.set(String(n.maNCC).trim(), { ma: String(n.maNCC).trim(), ten: n.tenNCC || "" });
    }
  });

  function buildGroups(rows) {
    const groupMap = new Map();
    rows.forEach((r) => {
      const key = r.maNCC || r.tenNCC || "(Không rõ NCC)";
      if (!groupMap.has(key)) {
        const mst = r.maSoThueNCC || "";
        const dmNcc = mst ? nccLookup.get(mst) : null;
        groupMap.set(key, {
          maNCC: key,
          tenNCC: (dmNcc && dmNcc.ten) || r.tenNCC || key,
          maSoThueNCC: mst,
          maDanhMuc: dmNcc ? dmNcc.ma : "",
          rows: [],
          tongTien: 0,
        });
      }
      const g = groupMap.get(key);
      g.rows.push(r);
      const amt = r.soTienTongRaw !== undefined ? Number(r.soTienTongRaw) : (r.soTienRaw !== undefined ? Number(r.soTienRaw) : 0);
      g.tongTien += amt || 0;
    });
    return Array.from(groupMap.values()).sort((a, b) => a.tenNCC.localeCompare(b.tenNCC, 'vi'));
  }

  const rowsKhCu  = allRows.filter((r) => !r.company || r.company === 'kh_cu');
  const rowsKhMoi = allRows.filter((r) => r.company === 'kh_moi');

  const groupsKhCu  = buildGroups(rowsKhCu);
  const groupsKhMoi = buildGroups(rowsKhMoi);

  // Sort rows inside each group by thang desc
  [...groupsKhCu, ...groupsKhMoi].forEach((g) => {
    g.rows.sort((a, b) => (b.thang || "").localeCompare(a.thang || ""));
  });

  // ---- KVC-MTD KH Mới: danh sách gian từ file Excel ----
  // Build lookup from KH Mới records by maCongTrinh → {mst, tenNCC, maNCC}
  const maCTInfo = new Map(); // maCongTrinh → {mst, tenNCC, maNCC}
  rowsKhMoi.forEach((r) => {
    const maCT = (r.maCongTrinh || "").trim().toUpperCase();
    const mst  = (r.maSoThueNCC || "").trim();
    if (maCT && !maCTInfo.has(maCT)) {
      maCTInfo.set(maCT, { mst, tenNCC: r.tenNCC || "", maNCC: r.maNCC || "" });
    }
    // Also store with the original case
    const maCTOrig = (r.maCongTrinh || "").trim();
    if (maCTOrig && !maCTInfo.has(maCTOrig)) {
      maCTInfo.set(maCTOrig, { mst, tenNCC: r.tenNCC || "", maNCC: r.maNCC || "" });
    }
  });

  // Also build tenNCC lookup as fallback
  const tenNccToMst = new Map();
  const tenNccToMaCT = new Map(); // tenNCC.upper → first maCongTrinh seen
  rowsKhMoi.forEach((r) => {
    const mst = (r.maSoThueNCC || "").trim();
    const tenKey = (r.tenNCC || "").trim().toUpperCase();
    if (mst) {
      [(r.tenNCC||""), (r.maNCC||"")].forEach(name => {
        if (name) tenNccToMst.set(name.trim().toUpperCase(), mst);
      });
    }
    // Collect distinct maCongTrinh per tenNCC
    if (tenKey && r.maCongTrinh) {
      const maCT = (r.maCongTrinh || "").trim();
      if (maCT && !tenNccToMaCT.has(tenKey)) tenNccToMaCT.set(tenKey, maCT);
    }
  });
  const nccTenToMst = new Map();
  (store.danh_muc_ma_nha_cung_cap_moi || []).forEach((n) => {
    if (n.ten && n.ma) nccTenToMst.set(String(n.ten).trim().toUpperCase(), String(n.ma).trim());
  });

  // Build invoice list per tenNCC for click-to-expand in KVC-MTD KH Mới tab
  const tenNccKeyToInvoices = new Map(); // tenNCC (short) → [invoice rows]
  rowsKhMoi.forEach((r) => {
    const key = (r.tenNCC || "").trim();
    if (!key) return;
    if (!tenNccKeyToInvoices.has(key)) tenNccKeyToInvoices.set(key, []);
    tenNccKeyToInvoices.get(key).push(r);
  });

  function lookupByMa(ma) {
    if (!ma) return null;
    return maCTInfo.get(ma.trim()) || maCTInfo.get(ma.trim().toUpperCase()) || null;
  }
  function lookupMst(tenNCC) {
    if (!tenNCC) return "";
    const key = tenNCC.trim().toUpperCase();
    return tenNccToMst.get(key) || nccTenToMst.get(key) || "";
  }

  // Static gian list — KVC-MTD KH Mới (from Book1.xlsx Sheet1)
  // tenNccKey = exact tenNCC short name used in ho_so_tien_thue records (for MST lookup)
  const KVC_MTD_MOI_GIAN = [
    { ma:'50AMBT',  ten:'POSH MN AEON MALL BÌNH TÂN',           tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI THÀNH PHỐ HỒ CHÍ MINH', chuThich:'Mới',                      tenNccKey:'AE BÌNH TÂN',          congTrinh:'AEON HỒ CHÍ MINH' },
    { ma:'50AMTP',  ten:'POSH MN AEON MALL TÂN PHÚ',            tenNCC:'CÔNG TY TNHH AEON VIỆT NAM',                                           chuThich:'Mới',                      tenNccKey:'AE TÂN PHÚ',           congTrinh:'AEON VIỆT NAM',        maCTHopDong:'AM TP PHCM',  tienThueKy:178875000 },
    { ma:'50BV175', ten:'POSH MN BỆNH VIỆN 175',                 tenNCC:'CÔNG TY TNHH QUẢNG CÁO VẠN THỊNH PHÁT',                               chuThich:'không có tiền thuê',       tenNccKey:'',                     congTrinh:'BỆNH VIỆN 175 (VẠN THỊNH PHÁT)' },
    { ma:'50BVUB',  ten:'POSH MN BỆNH VIỆN UNG BƯỚU HCM',       tenNCC:'CÔNG TY TNHH JD&M VIỆT NAM',                                          chuThich:'Mới',                      tenNccKey:'JD&M',                 congTrinh:'BỆNH VIỆN UNG BƯỚU (JD&M VIỆT NAM)' },
    { ma:'50CGVLCT',ten:'POSH MN CGV INTRESSCO LÝ CHÍNH THẮNG', tenNCC:'CÔNG TY TNHH CJ CGV VIỆT NAM - CHI NHÁNH QUẬN 3',                     chuThich:'Mới CSE',                  tenNccKey:'',                     congTrinh:'CGV VIỆT NAM CN QUẬN 3 (CGV INTRESSCO LÝ CHÍNH THẮNG)' },
    { ma:'50CGVLM', ten:'POSH MN CGV VINCOM LANDMARK',           tenNCC:'CÔNG TY TNHH CJ CGV VIỆT NAM - CHI NHÁNH BÌNH THẠNH 3',               chuThich:'Mới CSE',                  tenNccKey:'',                     congTrinh:'CGV VIỆT NAM CN BÌNH THẠCH 3 - CGV VC LANDMARK - CSE' },
    { ma:'50CGVPLZ',ten:'POSH MN CGV PEARL PLAZA',               tenNCC:'CÔNG TY TNHH CJ CGV VIỆT NAM - CHI NHÁNH VĂN THÁNH',                  chuThich:'Mới CSE',                  tenNccKey:'',                     congTrinh:'CGV VIỆT NAM CN VĂN THÁNH - CGV PEARL PLAZA - CSE' },
    { ma:'50CGVPVT',ten:'POSH MN CGV VINCOM PHAN VĂN TRỊ',      tenNCC:'CÔNG TY TNHH CJ CGV VIỆT NAM - CHI NHÁNH GÒ VẤP',                    chuThich:'Mới CSE',                  tenNccKey:'',                     congTrinh:'CGV VIỆT NAM CN GÒ VẤP - CGV VC PHAN VĂN TRỊ - CSE' },
    { ma:'50ETL',   ten:'POSH MN ESTELLA',                       tenNCC:'CÔNG TY TNHH LIÊN DOANH ESTELLA',                                     chuThich:'Mới',                      tenNccKey:'ESTELLA',              congTrinh:'LIÊN DOANH ESTELLA',   maCTHopDong:'ESTELLA PHN' },
    { ma:'50GAKDV', ten:'POSH MN GALAXY KINH DƯƠNG VƯƠNG',       tenNCC:'CÔNG TY CỔ PHẦN PHIM THIÊN NGÂN',                                     chuThich:'chung 1 hợp đồng',         tenNccKey:'THIÊN NGÂN',           congTrinh:'PHIM THIÊN NGÂN (GALAXY)', maCTHopDong:'GALAXY KINH DUONG VUONG PHCM' },
    { ma:'50GAQT',  ten:'POSH MN GALAXY QUANG TRUNG',            tenNCC:'CÔNG TY CỔ PHẦN PHIM THIÊN NGÂN',                                     chuThich:'chung 1 hợp đồng',         tenNccKey:'THIÊN NGÂN',           congTrinh:'PHIM THIÊN NGÂN (GALAXY)', maCTHopDong:'GALAXY QUANG TRUNG PHCM' },
    { ma:'50GGPVD', ten:'POSH MN GIGA PHẠM VĂN ĐỒNG',           tenNCC:'CÔNG TY CỔ PHẦN ĐẦU TƯ THƯƠNG MẠI DỊCH VỤ GIGAMALL VIỆT NAM',        chuThich:'Mới',                      tenNccKey:'GIGAMALL',             congTrinh:'GIGAMALL VIỆT NAM' },
    { ma:'50GOAC',  ten:'POSH MN GO ÂU CƠ',                     tenNCC:'CÔNG TY TNHH EB TÂN PHÚ',                                             chuThich:'Mới',                      tenNccKey:'EB TÂN PHÚ',           congTrinh:'GO ÂU CƠ (EB TÂN PHÚ)' },
    { ma:'50GONTT', ten:'POSH MN GO NGUYỄN THỊ THẬP',           tenNCC:'CÔNG TY TRÁCH NHIỆM HỮU HẠN THƯƠNG MẠI VÀ DỊCH VỤ SIÊU THỊ AN LẠC',  chuThich:'Mới 2 hd',                tenNccKey:'AN LẠC',               congTrinh:'GO AN LẠC (SIÊU THỊ AN LẠC)' },
    { ma:'50GOTC',  ten:'POSH MN GO TRƯỜNG CHINH',               tenNCC:'CÔNG TY TNHH ĐẦU TƯ BẤT ĐỘNG SẢN NEW PLAN',                          chuThich:'Mới 1 Hđ xuất 2 hóa đơn/tháng', tenNccKey:'TRƯỜNG CHINH',    congTrinh:'GO TRƯỜNG CHINH (NEW PLAN)' },
    { ma:'50LMNSG', ten:'POSH MN LOTTE MART NAM SÀI GÒN',        tenNCC:'CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM',                chuThich:'Mới',                      tenNccKey:'LOTTE VIỆT NAM',       congTrinh:'LOTTE VIỆT NAM',       maCTHopDong:'LOTTE Q7 (NSG) PHM' },
    { ma:'50LMPT',  ten:'POSH MN LOTTE MART PHÚ THỌ',           tenNCC:'CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM',                chuThich:'Mới',                      tenNccKey:'LOTTE VIỆT NAM',       congTrinh:'LOTTE VIỆT NAM',       maCTHopDong:'LOTTE PHU THO PHCM' },
    { ma:'50SCPVD', ten:'POSH MN SENSE CITY PHẠM VĂN ĐỒNG',     tenNCC:'CÔNG TY TNHH MTV SÀI GÒN - VĂN ĐỒNG',                               chuThich:'Mới',                      tenNccKey:'SENSE CITY PVĐ',       congTrinh:'SENSE PHẠM VĂN ĐỒNG (SÀI GÒN - VĂN ĐỒNG)' },
    { ma:'50SCVV',  ten:'POSH MN SC VIVO',                       tenNCC:'CÔNG TY CỔ PHẦN PHÁT TRIỂN KHU PHỨC HỢP THƯƠNG MẠI VIETSIN',         chuThich:'Mới thanh toán 2 lần',     tenNccKey:'VIETSIN VIVO',         congTrinh:'SC VIVO (PHỨC HỢP THƯƠNG MẠI VIETSIN)' },
    { ma:'50VC3/2', ten:'POSH MN VINCOM 3/2',                    tenNCC:'CHI NHÁNH TẠI THÀNH PHỐ HỒ CHÍ MINH - CÔNG TY TNHH VẬN HÀNH VINCOM RETAIL', chuThich:'chung 1 hợp đồng', tenNccKey:'VC CN HCM',           congTrinh:'CN HỒ CHÍ MINH VẬN HÀNH VINCOM RETAIL', maCTHopDong:'VC 3/2 JP-Posh' },
    { ma:'50VCGP',  ten:'POSH MN VINCOM GRAND PARK',             tenNCC:'CHI NHÁNH TẠI THÀNH PHỐ HỒ CHÍ MINH - CÔNG TY TNHH VẬN HÀNH VINCOM RETAIL', chuThich:'Mới',              tenNccKey:'VC CN HCM',           congTrinh:'CN HỒ CHÍ MINH VẬN HÀNH VINCOM RETAIL', maCTHopDong:'JP-POSH GRAND PARK' },
    { ma:'50VCGV',  ten:'POSH MN VINCOM GÒ VẤP',                tenNCC:'CHI NHÁNH TẠI THÀNH PHỐ HỒ CHÍ MINH - CÔNG TY TNHH VẬN HÀNH VINCOM RETAIL', chuThich:'Mới',              tenNccKey:'VC CN HCM',           congTrinh:'CN HỒ CHÍ MINH VẬN HÀNH VINCOM RETAIL', maCTHopDong:'VC GV PHCM' },
    { ma:'50VCLVV', ten:'POSH MN VINCOM LÊ VĂN VIỆT',           tenNCC:'CHI NHÁNH TẠI THÀNH PHỐ HỒ CHÍ MINH - CÔNG TY TNHH VẬN HÀNH VINCOM RETAIL', chuThich:'Mới',              tenNccKey:'VC CN HCM',           congTrinh:'CN HỒ CHÍ MINH VẬN HÀNH VINCOM RETAIL', maCTHopDong:'VC LVV PHCM' },
    { ma:'50VHM',   ten:'POSH MN VẠN HẠNH MALL',                tenNCC:'CÔNG TY CỔ PHẦN ĐẦU TƯ XÂY DỰNG BẮC BÌNH',                           chuThich:'Mới 1 Hđ xuất 2 hóa đơn/tháng', tenNccKey:'VẠN HẠNH MALL',  congTrinh:'VẠN HẠNH MALL (XÂY DỰNG BẮC BÌNH)' },
    { ma:'51AMBD',  ten:'POSH MN AEON MALL BÌNH DƯƠNG',         tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI BÌNH DƯƠNG',             chuThich:'Mới',                      tenNccKey:'AE BD',                congTrinh:'AEON BÌNH DƯƠNG' },
    { ma:'51CGVBD', ten:'POSH MN CGV BÌNH DƯƠNG SQUARE',        tenNCC:'CÔNG TY TNHH CJ CGV VIỆT NAM - CHI NHÁNH BÌNH DƯƠNG',                 chuThich:'Mới',                      tenNccKey:'',                     congTrinh:'CGV VIỆT NAM - CHI NHÁNH BÌNH DƯƠNG (SQUARE) - CSE' },
    { ma:'51GODA',  ten:'POSH MN GO DĨ AN',                     tenNCC:'CHI NHÁNH SỐ 2 CÔNG TY CP BẤT ĐỘNG SẢN VIỆT - NHẬT TẠI BÌNH DƯƠNG',  chuThich:'Mới 2 hợp đồng',           tenNccKey:'GO DI AN',             congTrinh:'SỐ 2 CÔNG TY CP BẤT ĐỘNG SẢN VIỆT - NHẬT TẠI BÌNH DƯƠNG' },
    { ma:'51GOTDM', ten:'POSH MN GO THỦ DẦU MỘT',              tenNCC:'CHI NHÁNH CÔNG TY CP BẤT ĐỘNG SẢN VIỆT - NHẬT TẠI BÌNH DƯƠNG',        chuThich:'Mới',                      tenNccKey:'GO TDM',               congTrinh:'GO THỦ DẦU MỘT (VIỆT - NHẬT TẠI BÌNH DƯƠNG)' },
    { ma:'52KBGBH', ten:'POSH MN KUBO GO BIÊN HÒA',             tenNCC:'CHI NHÁNH TẠI ĐỒNG NAI - CÔNG TY TNHH MỘT THÀNH VIÊN THƯƠNG MẠI HÀNG GIA DỤNG TỔNG HỢP', chuThich:'Mới', tenNccKey:'KUBO ĐỒNG NAI',       congTrinh:'KUBO GO ĐỒNG NAI (GIA DỤNG TỔNG HỢP)' },
    { ma:'52VCBH',  ten:'POSH MN VINCOM BIÊN HÒA',              tenNCC:'CHI NHÁNH TẠI TỈNH ĐỒNG NAI - CÔNG TY TNHH VẬN HÀNH VINCOM RETAIL',  chuThich:'Mới 1 hợp đồng',           tenNccKey:'VC CN ĐỒNG NAI',       congTrinh:'CN ĐỒNG NAI VẬN HÀNH VINCOM RETAIL (BIÊN HÒA)' },
    { ma:'53SBPQ',  ten:'POSH MN SÂN BAY PHÚ QUỐC',            tenNCC:'CÔNG TY TNHH QUẢNG CÁO SÂN BAY',                                      chuThich:'Mới 1 hợp đồng',           tenNccKey:'SÂN BAY',              congTrinh:'QUẢNG CÁO SÂN BAY',   maCTHopDong:'SB PHU QUOC PHN',   tienThueKy:255600000 },
    { ma:'53SWPQ',  ten:'POSH MN SUNWORLD PHÚ QUỐC',            tenNCC:'CHI NHÁNH CÔNG TY TNHH MẶT TRỜI PHÚ QUỐC TẠI HÒN THƠM',             chuThich:'Mới 1 hợp đồng CSE',       tenNccKey:'SUN H THƠM',           congTrinh:'MẶT TRỜI PHÚ QUỐC TẠI HÒN THƠM - CSE', maCTHopDong:'PQ SUN HTHOM PHCM' },
    { ma:'55GOBR',  ten:'POSH MN GO BÀ RỊA',                   tenNCC:'CHI NHÁNH CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN VIỆT - NHẬT TẠI BÀ RỊA',     chuThich:'Mới 2 hợp đồng',           tenNccKey:'GO BÀ RỊA',            congTrinh:'GO BÀ RỊA (VIỆT - NHẬT TẠI BÀ RỊA)' },
    { ma:'55GOBR',  ten:'POSH MN KUBO BÀ RỊA',                 tenNCC:'CHI NHÁNH BÀ RỊA - CÔNG TY TNHH MỘT THÀNH VIÊN THƯƠNG MẠI HÀNG GIA DỤNG TỔNG HỢP', chuThich:'Mới',      tenNccKey:'KUBO BÀ RỊA',          congTrinh:'KUBO GO BÀ RỊA (HÀNG GIA DỤNG TỔNG HỢP)' },
    { ma:'55SBCD',  ten:'POSH MN SÂN BAY CÔN ĐẢO',             tenNCC:'CÔNG TY TNHH QUẢNG CÁO SÂN BAY',                                      chuThich:'Mới',                      tenNccKey:'SÂN BAY',              congTrinh:'QUẢNG CÁO SÂN BAY',   maCTHopDong:'CON DAO AIRPORT PHCM', tienThueKy:48510000 },
    { ma:'56GOBT',  ten:'POSH MN GO BẾN TRE',                  tenNCC:'CHI NHÁNH CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN VIỆT-NHẬT TẠI BẾN TRE',       chuThich:'Mới 2 hợp đồng',           tenNccKey:'GO BẾN TRE',           congTrinh:'GO BẾN TRE (VIỆT - NHẬT TẠI BẾN TRE)' },
    { ma:'56SCBT',  ten:'POSH MN SENSE CITY BẾN TRE',           tenNCC:'CÔNG TY TNHH MTV THƯƠNG MẠI SÀI GÒN - BẾN TRE',                      chuThich:'Mới',                      tenNccKey:'SENSE BẾN TRE',        congTrinh:'SENSE BẾN TRE (SÀI GÒN - BẾN TRE)' },
    { ma:'57GOMT',  ten:'POSH MN GO MỸ THO',                   tenNCC:'CÔNG TY TRÁCH NHIỆM HỮU HẠN MỘT THÀNH VIÊN ĐẦU TƯ PHÁT TRIỂN NGUYỄN KIM TIỀN GIANG', chuThich:'Mới 2 hợp đồng', tenNccKey:'NGUYỄN KIM TIỀN GIANG', congTrinh:'GO MỸ THO (NGUYỄN KIM TIỀN GIANG)' },
    { ma:'58GOTV',  ten:'POSH MN GO TRÀ VINH',                 tenNCC:'CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN VÀ SIÊU THỊ BÁN LẺ ĐÔNG DƯƠNG TRÀ VINH', chuThich:'Mới',                    tenNccKey:'GO TRÀ VINH',          congTrinh:'GO TRÀ VINH (BÁN LẺ ĐÔNG DƯƠNG TRÀ VINH)' },
    { ma:'59CGVXK', ten:'POSH MN CGV VINCOM XUÂN KHÁNH',        tenNCC:'CÔNG TY TNHH CJ CGV VIỆT NAM - CHI NHÁNH CẦN THƠ 3',                 chuThich:'Mới',                      tenNccKey:'',                     congTrinh:'CGV VIỆT NAM - CHI NHÁNH CẦN THƠ (XUÂN KHÁNH) - CSE' },
    { ma:'59GNOV',  ten:'POSH MN GO CẦN THƠ',                  tenNCC:'CHI NHÁNH CÔNG TY CP BẤT ĐỘNG SẢN VIỆT - NHẬT TẠI CẦN THƠ',          chuThich:'Mới',                      tenNccKey:'GO C THƠ',             congTrinh:'GO CẦN THƠ (VIỆT - NHẬT TẠI CẦN THƠ)' },
    { ma:'59KBGCT', ten:'POSH MN KUBO GO CẦN THƠ',             tenNCC:'CHI NHÁNH TẠI CẦN THƠ - CÔNG TY TNHH MỘT THÀNH VIÊN THƯƠNG MẠI HÀNG GIA DỤNG TỔNG HỢP', chuThich:'Mới', tenNccKey:'KUBO CẦN THƠ',         congTrinh:'KUBO GO CẦN THƠ (HÀNG GIA DỤNG TỔNG HỢP)' },
    { ma:'59SBCT',  ten:'POSH MN SÂN BAY CẦN THƠ',             tenNCC:'CÔNG TY TNHH QUẢNG CÁO SÂN BAY',                                      chuThich:'Mới',                      tenNccKey:'SÂN BAY',              congTrinh:'QUẢNG CÁO SÂN BAY',   maCTHopDong:'SB CAN THO PHCM',   tienThueKy:134491500 },
    { ma:'59SCCT',  ten:'POSH MN SENSE CITY CẦN THƠ',           tenNCC:'CÔNG TY TNHH THƯƠNG MẠI SÀI GÒN CẦN THƠ',                            chuThich:'Mới',                      tenNccKey:'SENSE CITY CT',        congTrinh:'SENSE CẦN THƠ (SÀI GÒN CẦN THƠ)' },
    { ma:'60GOBMT', ten:'POSH MN GO BUÔN MÊ THUỘT',             tenNCC:'CHI NHÁNH CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN VIỆT - NHẬT TẠI BUÔN MA THUỘT', chuThich:'Mới',                  tenNccKey:'BMT',                  congTrinh:'GO BUÔN MA THUỘT (VIỆT - NHẬT TẠI BUÔN MA THUỘT)' },
    { ma:'60KBGBMT',ten:'POSH MN KUBO GO BUÔN MÊ THUỘT',       tenNCC:'CHI NHÁNH CÔNG TY TNHH MỘT THÀNH VIÊN THƯƠNG MẠI HÀNG GIA DỤNG TỔNG HỢP TẠI BUÔN MA THUỘT', chuThich:'Mới', tenNccKey:'KUBO BMT',         congTrinh:'KUBO GO BUÔN MA THUỘT (HÀNG GIA DỤNG TỔNG HỢP)' },
    { ma:'60YKBMT', ten:'POSH MN YOKIDS',                       tenNCC:'CÔNG TY CỔ PHẦN THƯƠNG MẠI VÀ DỊCH VỤ YOKIDS',                       chuThich:'mới họ giữ tiền',          tenNccKey:'',                     congTrinh:'YOKIDS' },
    { ma:'61ZCKG',  ten:'POSH MN ZONE C KIÊN GIANG',            tenNCC:'CÔNG TY CỔ PHẦN ĐÔNG HƯNG',                                           chuThich:'Mới',                      tenNccKey:'ĐÔNG HƯNG',            congTrinh:'ZONE C KIEN GIANG - ĐÔNG HƯNG' },
    { ma:'62SCCM',  ten:'POSH MN SENSE CITY CÀ MAU',            tenNCC:'CÔNG TY TNHH THƯƠNG MẠI DỊCH VỤ SÀI GÒN – CÀ MAU',                  chuThich:'Mới',                      tenNccKey:'SÀI GÒN – CÀ MAU',    congTrinh:'SENSE CÀ MAU (SÀI GÒN - CÀ MAU)' },
    { ma:'62GOBL',  ten:'POSH MN GO BẠC LIÊU',                 tenNCC:'CHI NHÁNH CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN VIỆT-NHẬT TẠI BẠC LIÊU',      chuThich:'Mới',                      tenNccKey:'GO BẠC LIÊU',          congTrinh:'GO BẠC LIÊU (VIỆT - NHẬT TẠI BẠC LIÊU)' },
    { ma:'63LTPT',  ten:'POSH MN LOTTE MART PHAN THIẾT',        tenNCC:'CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BÌNH THUẬN', chuThich:'Mới',             tenNccKey:'LOTTE BÌNH THUẬN',     congTrinh:'LOTTE VIỆT NAM - CHI NHÁNH BÌNH THUẬN', maCTHopDong:'LOTTE PHAN THIET' },
    { ma:'JPAMTP',  ten:'JP MN AEON MALL TÂN PHÚ',              tenNCC:'CÔNG TY TNHH AEON VIỆT NAM',                                          chuThich:'mới',                      tenNccKey:'AE TÂN PHÚ',           congTrinh:'AEON VIỆT NAM',        maCTHopDong:'JP AE TAN PHU', tienThueKy:44000000 },
    { ma:'JPSWPQ',  ten:'JP MN SUNWORLD PHÚ QUỐC',              tenNCC:'CHI NHÁNH CÔNG TY TNHH MẶT TRỜI PHÚ QUỐC TẠI HÒN THƠM',             chuThich:'mới CSE 50%-50%',          tenNccKey:'SUN H THƠM',           congTrinh:'MẶT TRỜI PHÚ QUỐC TẠI HÒN THƠM - CSE', maCTHopDong:'PQ SUN HTHOM JPHCM' },
    { ma:'FLMPT',   ten:'Farm Phan Thiết',                       tenNCC:'CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BÌNH THUẬN', chuThich:'mới',             tenNccKey:'LOTTE BÌNH THUẬN',     congTrinh:'LOTTE VIỆT NAM - CHI NHÁNH BÌNH THUẬN', maCTHopDong:'LOTTE PHAN THIET' },
    { ma:'TTAMTA',  ten:'TUTU MN AEON MALL TÂN AN',             tenNCC:'CÔNG TY TNHH AEON VIỆT NAM - CHI NHÁNH LONG AN',                     chuThich:'mới',                      tenNccKey:'',                     congTrinh:'AEON TÂN AN - AEON VIỆT NAM - CHI NHÁNH LONG AN' },
    { ma:'TTETL',   ten:'TUTU MN ESTELLA',                       tenNCC:'CÔNG TY TNHH LIÊN DOANH ESTELLA',                                     chuThich:'mới CSE',                  tenNccKey:'ESTELLA',              congTrinh:'LIÊN DOANH ESTELLA - CSE', maCTHopDong:'ESTELLA PHN' },
    { ma:'VRAMTA',  ten:'FZ MN VR AEON MALL TÂN AN',            tenNCC:'CÔNG TY TNHH AEON VIỆT NAM - CHI NHÁNH LONG AN',                     chuThich:'mới',                      tenNccKey:'',                     congTrinh:'AEON TÂN AN - AEON VIỆT NAM - CHI NHÁNH LONG AN' },
    { ma:'PBAMTP',  ten:'PINBALL MN AMTP',                       tenNCC:'CÔNG TY TNHH AEON VIỆT NAM',                                          chuThich:'mới',                      tenNccKey:'',                      congTrinh:'AEON VIỆT NAM' },
  ];

  // Group KVC-MTD KH Mới gian by NCC
  const kvcMoiGroupMap = new Map();
  KVC_MTD_MOI_GIAN.forEach((g) => {
    // Lookup MST using tenNccKey (exact short name from records) as primary method
    const keyUpper = g.tenNccKey ? g.tenNccKey.toUpperCase() : "";
    const mstByKey = keyUpper ? (tenNccToMst.get(keyUpper) || "") : "";
    const info = lookupByMa(g.ma);
    const mst = mstByKey || (info && info.mst) || lookupMst(g.tenNCC) || "";
    // Get maCongTrinh from records for display
    const maCongTrinhKhMoi = keyUpper ? (tenNccToMaCT.get(keyUpper) || "") : "";
    const tenNccDisplay = g.tenNCC || (info && info.tenNCC) || "(Không rõ NCC)";
    const key = tenNccDisplay;
    if (!kvcMoiGroupMap.has(key)) {
      kvcMoiGroupMap.set(key, { tenNCC: key, mst, tenNccKey: g.tenNccKey || "", gianList: [] });
    }
    // If this group now has mst and didn't before, update it
    if (mst && !kvcMoiGroupMap.get(key).mst) kvcMoiGroupMap.get(key).mst = mst;
    // Collect invoices with "tiền thuê" in description for this gian.
    // Matching strategy (applied when gian has maCTHopDong):
    //   1st: match by maCongTrinh (most accurate — exact field match)
    //   2nd: match by invoice total amount ≈ tienThueKy (for invoices without proper maCT)
    //   Fallback: show all tiền thuê from NCC (when no specific matches found)
    const allInvoices = g.tenNccKey ? (tenNccKeyToInvoices.get(g.tenNccKey) || []) : [];
    const candidateTT = allInvoices.filter(r => {
      const dg = (r.dienGiai || '').toLowerCase();
      return dg.includes('tiền thuê') || dg.includes('tien thue');
    });
    // Helper: parse Vietnamese number string "255.600.000" → 255600000
    const parseVnd = (s) => {
      if (!s) return 0;
      if (typeof s === 'number') return s;
      return parseInt(String(s).replace(/[^0-9]/g, ''), 10) || 0;
    };
    const tienThueInvoices = (() => {
      if (!g.maCTHopDong && !g.tienThueKy) return candidateTT; // no filter defined, show all
      const maCTKey = g.maCTHopDong ? g.maCTHopDong.trim().toUpperCase() : null;
      // Combine maCT match OR amount match (both active simultaneously)
      const matched = candidateTT.filter(r => {
        // Primary: exact maCongTrinh match
        if (maCTKey && (r.maCongTrinh || '').trim().toUpperCase() === maCTKey) return true;
        // Secondary: amount-based match (for invoices with missing/generic maCT)
        if (g.tienThueKy) {
          const inv = parseVnd(r.soTienTong || r.soTienTongRaw);
          return Math.abs(inv - g.tienThueKy) < 1000; // within 1000 VNĐ tolerance
        }
        return false;
      });
      if (matched.length > 0) return matched; // found specific invoices for this gian
      return candidateTT; // fallback: show all tiền thuê from NCC only if nothing matched
    })().sort((a, b) => (b.thang || '').localeCompare(a.thang || ''));
    kvcMoiGroupMap.get(key).gianList.push({ ...g, mst, maCTFound: g.congTrinh || '', tienThueInvoices });
  });
  const kvcMoiGroups = Array.from(kvcMoiGroupMap.values())
    .sort((a, b) => a.tenNCC.localeCompare(b.tenNCC, 'vi'));

  // ---- KVC-MTD KH Cũ: danh sách gian ----
  // Build invoice lookup from KH Cũ records (tenNCC short key → [invoice rows])
  const tenNccKeyToInvoicesCu = new Map();
  // Also build MST lookup from KH Cũ records (tenNCC short key.UPPER → MST)
  const tenNccToMstCu = new Map();
  rowsKhCu.forEach((r) => {
    const key = (r.tenNCC || "").trim();
    if (!key) return;
    if (!tenNccKeyToInvoicesCu.has(key)) tenNccKeyToInvoicesCu.set(key, []);
    tenNccKeyToInvoicesCu.get(key).push(r);
    const mst = (r.maSoThueNCC || "").trim();
    if (mst) tenNccToMstCu.set(key.toUpperCase(), mst);
  });

  // Static gian list — KVC-MTD KH Cũ (from user Excel)
  const KVC_MTD_CU_GIAN = [
    { ma:'TTAMBD',  ten:'TUTU MN AEON MALL BÌNH DƯƠNG',        tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI BÌNH DƯƠNG',                                                      tenNccKey:'AE BD',            congTrinh:'AEON BÌNH DƯƠNG' },
    { ma:'TTLMGV',  ten:'TUTU MN LOTTE MART GÒ VẤP',           tenNCC:'CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM',                                                          tenNccKey:'LOTTE VIỆT NAM',   congTrinh:'LOTTE VIỆT NAM - CHI NHÁNH GÒ VẤP' },
    { ma:'TTAMTP',  ten:'TUTU MN AEON MALL TÂN PHÚ',           tenNCC:'CÔNG TY TNHH AEON VIỆT NAM',                                                                                   tenNccKey:'AE TÂN PHÚ',       congTrinh:'AEON VIỆT NAM - AEON TÂN PHÚ' },
    { ma:'TTAMBT',  ten:'TUTU MN AEON MALL BÌNH TÂN',          tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI THÀNH PHỐ HỒ CHÍ MINH',                                         tenNccKey:'AE BÌNH TÂN',      congTrinh:'AEON HỒ CHÍ MINH - CSE',      tienThueKy: 33000000 },
    { ma:'FZVRBD',  ten:'FZ MN VR AEON MALL BÌNH DƯƠNG',       tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI BÌNH DƯƠNG',                                                      tenNccKey:'AE BD',            congTrinh:'AEON BÌNH DƯƠNG' },
    { ma:'FZNBVT',  ten:'FZ MN NHÀ BÓNG LOTTE MART VŨNG TÀU', tenNCC:'CÔNG TY CỔ PHẦN TRUNG TÂM THƯƠNG MẠI LOTTE VIỆT NAM - CHI NHÁNH BÀ RỊA VŨNG TÀU',                           tenNccKey:'LOTTE VŨNG TÀU',   congTrinh:'LOTTE VIỆT NAM - CHI NHÁNH BÀ RỊA VŨNG TÀU' },
    { ma:'FZFFVV',  ten:'FZ MN FUNFEST SC VIVO',                tenNCC:'CÔNG TY CỔ PHẦN PHÁT TRIỂN KHU PHỨC HỢP THƯƠNG MẠI VIETSIN',                                                   tenNccKey:'VIETSIN VIVO',     congTrinh:'SC VIVO (PHỨC HỢP THƯƠNG MẠI VIETSIN) - CSE' },
    { ma:'FZADVVV', ten:'FZ MN ADV SC VIVO',                    tenNCC:'CÔNG TY CỔ PHẦN PHÁT TRIỂN KHU PHỨC HỢP THƯƠNG MẠI VIETSIN',                                                   tenNccKey:'VIETSIN VIVO',     congTrinh:'SC VIVO (PHỨC HỢP THƯƠNG MẠI VIETSIN)',      tienThueKy: 14850000 },
    { ma:'FZADVTP', ten:'FZ MN ADV AEON MALL TÂN PHÚ',         tenNCC:'CÔNG TY TNHH AEON VIỆT NAM',                                                                                   tenNccKey:'AE TÂN PHÚ',       congTrinh:'AEON VIỆT NAM - AEON TÂN PHÚ' },
    { ma:'EVGHBD',  ten:'EVMN GHOST MN AEON MALL BÌNH DƯƠNG',  tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI BÌNH DƯƠNG',                                                      tenNccKey:'AE BD',            congTrinh:'AEON BÌNH DƯƠNG - CSE' },
    { ma:'EVSNTP',  ten:'EV MN SNOW MN AEON MALL TÂN PHÚ',     tenNCC:'CÔNG TY TNHH AEON VIỆT NAM',                                                                                   tenNccKey:'AE TÂN PHÚ',       congTrinh:'AEON VIỆT NAM - AEON TÂN PHÚ' },
    { ma:'EVGHBR',  ten:'EVMN GHOST MN GO BÀ RỊA',             tenNCC:'CHI NHÁNH CÔNG TY CỔ PHẦN BẤT ĐỘNG SẢN VIỆT- NHẬT TẠI BÀ RỊA',                                               tenNccKey:'BÀ RỊA',           congTrinh:'GO BÀ RỊA (VIỆT- NHẬT TẠI BÀ RỊA)' },
    { ma:'EVSNBD',  ten:'EVMN SNOW AEON MALL BÌNH DƯƠNG',       tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI BÌNH DƯƠNG',                                                      tenNccKey:'AE BD',            congTrinh:'AEON BÌNH DƯƠNG - CSE' },
    { ma:'EVADVGAL',ten:'EVENT MN ADV GO AN LẠC',               tenNCC:'CÔNG TY TRÁCH NHIỆM HỮU HẠN THƯƠNG MẠI VÀ DỊCH VỤ SIÊU THỊ AN LẠC',                                         tenNccKey:'SIÊU THỊ AN LẠC',  congTrinh:'GO AN LẠC (SIÊU THỊ AN LẠC)' },
    { ma:'50CMPL',  ten:'POSH MN COOPMART PHÚ LÂM',            tenNCC:'CÔNG TY TNHH MỘT THÀNH VIÊN SÀI GÒN CO.OP PHÚ LÂM',                                                           tenNccKey:'CO.OP PHÚ LÂM',    congTrinh:'COOPMART PHÚ LÂM' },
    { ma:'50CMBD',  ten:'POSH MN COOPMART BÌNH DƯƠNG',         tenNCC:'CHI NHÁNH LIÊN HIỆP HỢP TÁC XÃ THƯƠNG MẠI TP. HỒ CHÍ MINH - CO.OPMART BÌNH DƯƠNG 2',                        tenNccKey:'CO.OP BD',         congTrinh:'COOPMART BÌNH DƯƠNG' },
    { ma:'50VPPQ',  ten:'POSH MN VINPEARL PHÚ QUỐC',           tenNCC:'CHI NHÁNH KIÊN GIANG - CÔNG TY CỔ PHẦN VINPEARL',                                                             tenNccKey:'PHÚ QUỐC',         congTrinh:'VINWONDER PHÚ QUỐC (CỔ PHẦN VINPEARL)' },
    { ma:'JPAMBT',  ten:'JP MN AEON MALL BÌNH TÂN',            tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI THÀNH PHỐ HỒ CHÍ MINH',                                         tenNccKey:'AE BÌNH TÂN',      congTrinh:'AEON HỒ CHÍ MINH' },
    { ma:'JPAMBD',  ten:'JP MN AEON MALL BÌNH DƯƠNG',          tenNCC:'CHI NHÁNH CÔNG TY TNHH AEONMALL VIỆT NAM TẠI BÌNH DƯƠNG',                                                      tenNccKey:'AE BD',            congTrinh:'AEON BÌNH DƯƠNG' },
    { ma:'JPVWPQ',  ten:'JP MN VINWONDER PHÚ QUỐC',            tenNCC:'CHI NHÁNH KIÊN GIANG - CÔNG TY CỔ PHẦN VINPEARL',                                                             tenNccKey:'PHÚ QUỐC',         congTrinh:'VINWONDER PHÚ QUỐC (cổ phần Vinpearl)' },
    { ma:'JPSBPQ',  ten:'JP MN SÂN BAY PHÚ QUỐC',             tenNCC:'CÔNG TY CỔ PHẦN CẢNG HÀNG KHÔNG MẶT TRỜI- CHI NHÁNH CẢNG HÀNG KHÔNG QUỐC TẾ PHÚ QUỐC SUN GROUP',           tenNccKey:'CHKQT PHÚ QUỐC',   congTrinh:'CN CHK QUỐC TẾ PHÚ QUỐC - CSE' },
  ];

  // Group KVC-MTD KH Cũ gian by NCC
  const kvcCuGroupMap = new Map();
  KVC_MTD_CU_GIAN.forEach((g) => {
    const keyUpper = g.tenNccKey ? g.tenNccKey.toUpperCase() : "";
    // Use KH Cũ MST map first, fallback to KH Mới map, then full-name lookup
    const mstByKey = keyUpper ? (tenNccToMstCu.get(keyUpper) || tenNccToMst.get(keyUpper) || "") : "";
    const mst = mstByKey || lookupMst(g.tenNCC) || "";
    const tenNccDisplay = g.tenNCC || "(Không rõ NCC)";
    const key = tenNccDisplay;
    if (!kvcCuGroupMap.has(key)) {
      kvcCuGroupMap.set(key, { tenNCC: key, mst, tenNccKey: g.tenNccKey || "", gianList: [] });
    }
    if (mst && !kvcCuGroupMap.get(key).mst) kvcCuGroupMap.get(key).mst = mst;
    const allInvoicesCu = g.tenNccKey ? (tenNccKeyToInvoicesCu.get(g.tenNccKey) || []) : [];
    // KH Cũ dùng nhiều pattern: "Phí thuê...", "TIỀN THUÊ", "tiền thuê...", "thuê vị trí..."
    const candidateTTCu = allInvoicesCu.filter(r => {
      const dg = (r.dienGiai || '').toLowerCase();
      return dg.includes('thuê') || dg.includes('thue');
    });
    // Filter by tienThueKy if set (amount-based per-gian matching from contract data)
    let gianInvoicesCu = candidateTTCu;
    if (g.tienThueKy && g.tienThueKy > 0) {
      const tol = g.tienThueKy * 0.12; // 12% tolerance to handle VAT rounding
      gianInvoicesCu = candidateTTCu.filter(r => {
        const inv = parseInt(String(r.soTienTong || '0').replace(/\./g, '').replace(/,/g, '.').replace(/[^0-9]/g, ''), 10);
        return Math.abs(inv - g.tienThueKy) <= tol;
      });
    }
    kvcCuGroupMap.get(key).gianList.push({ ...g, mst, maCTFound: g.congTrinh || '', tienThueInvoices: gianInvoicesCu });
  });
  const kvcCuGroups = Array.from(kvcCuGroupMap.values())
    .sort((a, b) => a.tenNCC.localeCompare(b.tenNCC, 'vi'));

  res.render("phapdanh-unc-tien-thue", {
    title: "UNC Tiền Thuê",
    activeCompany: store.activeCompany || 'kh_cu',
    COMPANIES,
    groupsKhCu,
    groupsKhMoi,
    kvcMoiGroups,
    kvcCuGroups,
    userName: req.session && req.session.userName,
    success: req.query.success,
    error: req.query.error,
  }, function(err, html) {
    if (err) {
      console.error("[unc-tien-thue] RENDER ERROR:", err.message, err.stack);
      return res.status(500).send("UNC Tiền Thuê render lỗi: " + err.message + "<br><pre>" + err.stack + "</pre>");
    }
    res.send(html);
  });
  } catch(e) {
    console.error("[unc-tien-thue] ROUTE ERROR:", e.message, e.stack);
    res.status(500).send("UNC Tiền Thuê route lỗi: " + e.message + "<br><pre>" + e.stack + "</pre>");
  }
});

module.exports = router;
// Luyen, 2026-08-01: "từ cái hợp đồng thuê gian á nó sẽ có tên đối tác ký hợp
// đồng với mình từ cái tên đó bạn map với lại hóa đơn đầu vào" -- trang moi
// ---------- Hop Dong Thue Gian Tong ----------
// Nhan, 2026-08-13: "thêm 1 mục mới trong dropdown Pháp danh ... gọi là Hợp
// Đồng Thuê Gian Tổng ... có 4 sheet KVC MN, KVC MB, MTD MN, MTD MB" -- khac
// voi 2 trang Thue Gian Hang / Thue Gian Hang Mien Bac o tren (nguon la sheet
// Tinh trang thai mau cho 1 dong, so voi ngay he thong hien tai (todayStr,
// dang "YYYY-MM-DD"): "het_han" (do) neu ngayHetHanThue < hom nay; "het_han_
// thang_nay" (cam) neu ngayHetHanThue cung nam-thang voi hom nay VA >= hom
// nay; con lai (chua co ngay, hoac het han xa hon thang nay) khong to mau.
// Chi dung khi ngayHetHanThue da parse duoc ro rang (dang "dd/mm/yyyy -
// dd/mm/yyyy") -- gian chua ro thoi han (con lai phan lon o MTD MB) se KHONG
// bi to mau, dung y "khong doan" da noi trong yeu cau.
function computeTrangThaiThueTong(r, todayStr) {
    if (r.daChamDut) return "cham_dut";
    if (!r.ngayHetHanThue) return "";
    if (r.ngayHetHanThue < todayStr) return "het_han";
    if (r.ngayHetHanThue.slice(0, 7) === todayStr.slice(0, 7)) return "het_han_thang_nay";
    return "";
}

// Nhan, 2026-08-13 (lan 2): "lọc theo tháng nhá" -- them bo loc thang giong
// het cach lam voi trang "Thue Gian Hang" (monthOverlaps ben tren), nhung
// gian tong dung ten truong rieng ngayBatDauThue/ngayHetHanThue (khac
// ngayBatDauHD/ngayHetHanHD cua trang Thue Gian Hang) nen viet ham rieng.
// Gian CHUA parse duoc ngay (phan lon MTD MB, xem chu thich sheetKey/
// computeTrangThaiThueTong o tren) LUON duoc giu lai, khong bi loc mat.
function monthOverlapsThueTong(r, thang) {
    if (!thang) return true;
    if (!r.ngayBatDauThue && !r.ngayHetHanThue) return true;
    const monthStart = thang + "-01";
    const monthEnd = thang + "-31";
    const bd = r.ngayBatDauThue || "0000-00-00";
    const hh = r.ngayHetHanThue || "9999-99-99";
    return bd <= monthEnd && hh >= monthStart;
}

router.get("/phap-danh/hop-dong-thue-gian-tong", (req, res) => {
    const thangFilter = resolveThangFilter(req);
    const store = load();
    ensureShape(store);
    const activeCompany = getCompany(req);
    const sheetParam = HOP_DONG_THUE_TONG_SHEETS.some((s) => s.key === req.query.sheet)
      ? req.query.sheet
          : "kvc-mn";
    const sheetInfo = HOP_DONG_THUE_TONG_SHEETS.find((s) => s.key === sheetParam);
    const todayStr = new Date().toISOString().slice(0, 10);

    // Luyen, 2026-08-19: build lookup tienThueThang tu phap_danh_hop_dong_thue
    // (cac hop dong chi tiet da nhap / dong bo tu sheet HCM) de hien thi tien thue
    // trong bang hop dong tong ma khong can them cot vao GG Sheet.
    // Join theo maCongTrinh (normalize uppercase + trim).
    // Neu nhieu gian cung maCongTrinh → lay max (thuong la gian lon nhat / hop dong chinh).
    const normMa = (v) => String(v || "").toUpperCase().replace(/\s+/g, " ").trim();
    const tienThueByMa = {};
    for (const h of store.phap_danh_hop_dong_thue || []) {
      const ma = normMa(h.maCongTrinh);
      if (!ma || !h.tienThueThang) continue;
      const amt = Number(h.tienThueThang) || 0;
      if (amt <= 0) continue;
      // Sum: mot maCongTrinh co the co nhieu dong (nhieu gian o cung 1 diem)
      tienThueByMa[ma] = (tienThueByMa[ma] || 0) + amt;
    }

    let rows = store.phap_danh_hop_dong_thue_tong.filter(
          (r) => r.sheetKey === sheetInfo.storeKey && (r.congTy === activeCompany || r.congTy === "ca_2")
              );
if (thangFilter) rows = rows.filter((r) => monthOverlapsThueTong(r, thangFilter));
    rows.forEach((r) => {
          r.trangThaiMau = computeTrangThaiThueTong(r, todayStr);
          // Enrich tienThueThang tu hop dong chi tiet neu chua co
          if (!r.tienThueThang) {
            const ma = normMa(r.maCongTrinh);
            r.tienThueThang = tienThueByMa[ma] || 0;
          }
    });
  
    const counts = {};
    HOP_DONG_THUE_TONG_SHEETS.forEach((s) => {
          let sheetRows = store.phap_danh_hop_dong_thue_tong.filter(
                  (r) => r.sheetKey === s.storeKey && (r.congTy === activeCompany || r.congTy === "ca_2")
                        );
    if (thangFilter) sheetRows = sheetRows.filter((r) => monthOverlapsThueTong(r, thangFilter));
    counts[s.key] = sheetRows.length;
    });
  
    res.render("phapdanh-hopdong-thue-tong", {
          userName: req.session.userName,
          rows,
          sheets: HOP_DONG_THUE_TONG_SHEETS,
          sheetParam,
          thangFilter,
          counts,
          error: req.query.error || null,
          success: req.query.success || null,
    });
});

// Parser CSV THAT SU (khong chi split theo dong) -- gviz CSV co the co 1 o
// chua xuong dong that trong dau ngoac kep (gap 1 dong o sheet MTD MB), neu
// chi split("\n") don gian se cat nham 1 dong thanh 2, lam sai lech so dem.
function parseGvizCsv(text) {
    const rows = [];
    let row = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
          const c = text[i];
          if (inQuotes) {
                  if (c === '"') {
                            if (text[i + 1] === '"') {
                                        cur += '"';
                                        i++;
                            } else inQuotes = false;
                  } else cur += c;
          } else if (c === '"') inQuotes = true;
          else if (c === ",") {
                  row.push(cur);
                  cur = "";
          } else if (c === "\r") {
                  // bo qua
          } else if (c === "\n") {
                  row.push(cur);
                  cur = "";
                  rows.push(row);
                  row = [];
          } else cur += c;
    }
    if (cur.length > 0 || row.length > 0) {
          row.push(cur);
          rows.push(row);
    }
    return rows.filter((r) => r.length > 1 || (r[0] || "").trim() !== "");
}

function buildGvizHeaderIndex(headerRow) {
    const idx = {};
    headerRow.forEach((h, i) => {
          const key = (h || "").trim();
          if (key && !(key in idx)) idx[key] = i;
    });
    return idx;
}

function gvizCol(rowArr, idx, name) {
    const i = idx[name];
    return i === undefined ? "" : (rowArr[i] || "").trim();
}

// "KH mới"/"KH cũ"/"Cả 2" (nhu cot Phap nhan tren sheet) -> "kh_moi"/"kh_cu"/
// "ca_2"; rong hoac gia tri la khac -> "" (chua xac dinh, KHONG doan).
function congTyFromPhapNhan(v) {
    const n = (v || "").trim();
    if (n === "KH mới") return "kh_moi";
    if (n === "KH cũ") return "kh_cu";
    if (n.toLowerCase().includes("cả 2") || n.toLowerCase().includes("ca 2")) return "ca_2";
    return "";
}

// "dd/mm/yyyy - dd/mm/yyyy" -> {start, end} dang ISO "yyyy-mm-dd". Khong
// khop duoc dinh dang (thoi han ghi tu do, hoac de trong) -> "" ca 2, KHONG doan.
function parseThoiHanThueRange(raw) {
    const m = (raw || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s*-\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!m) return { start: "", end: "" };
    const p2 = (s) => String(s).padStart(2, "0");
    return {
          start: `${m[3]}-${p2(m[2])}-${p2(m[1])}`,
          end: `${m[6]}-${p2(m[5])}-${p2(m[4])}`,
    };
}

function stripLinkQuery(v) {
    return (v || "").split("?")[0];
}

// Nhan, 2026-08-13 (lan 3): moi tab (KVC MN/MB, MTD MN/MB) co SO CỘT KHAC
// NHAU (vd MTD co them "Bên giữ tiền"/"TK NỢ"/"TK CÓ", KVC MB khong co cot
// "Địa điểm"/"Ghi Chú") -- tra cot theo TEN HEADER (khong theo vi tri cot cu
// the) de doc dung du lieu moi sheet, khong bi lech cot.
async function fetchHopDongThueTongSheet(storeKey) {
    const gid = HOP_DONG_THUE_TONG_GIDS[storeKey];
    const url = `https://docs.google.com/spreadsheets/d/${HOP_DONG_THUE_TONG_SPREADSHEET_ID}/gviz/tq?tqx=out:csv&gid=${gid}`;
    const resp = await fetch(url);
    if (!resp.ok) {
          throw new Error(
                  `Không đọc được sheet ${storeKey} (mã lỗi ${resp.status}). Kiểm tra lại sheet đã chia sẻ "Bất kỳ ai có link đều xem được" chưa.`
                );
    }
    const text = await resp.text();
    const rows = parseGvizCsv(text);
    if (rows.length === 0) return [];
    const idx = buildGvizHeaderIndex(rows[0]);
    return rows
      .slice(1)
      .map((r) => {
              const thoiHanThueRaw = gvizCol(r, idx, "Thời hạn hợp đồng");
              const { start, end } = parseThoiHanThueRange(thoiHanThueRaw);
              return {
                        sheetKey: storeKey,
                        congTy: congTyFromPhapNhan(gvizCol(r, idx, "Pháp nhân")),
                        congTyRaw: gvizCol(r, idx, "Pháp nhân"),
                        khuVuc: gvizCol(r, idx, "Khu Vực"),
                        dichVu: gvizCol(r, idx, "Dịch vụ"),
                        tenNoiBo: gvizCol(r, idx, "Mã Điểm Nội Bộ"),
                        maCongTrinh: gvizCol(r, idx, "Mã Điểm Thuế"),
                        diaDiem: gvizCol(r, idx, "Địa điểm"),
                        tenKhachHang: gvizCol(r, idx, "Tên khách hàng"),
                        mstKhachHang: gvizCol(r, idx, "MST khách hàng"),
                        hinhThucThue: gvizCol(r, idx, "Hình Thức Hợp Tác"),
                        tienThueThang: (function() {
                          const raw = gvizCol(r, idx, "Tiền thuê/tháng") || gvizCol(r, idx, "Tien thue/thang") || "";
                          const n = Number(String(raw).replace(/[^\d]/g, ""));
                          return isNaN(n) ? 0 : n;
                        })(),
                        thoiHanThueRaw,
                        ngayBatDauThue: start,
                        ngayHetHanThue: end,
                        ghiChu: gvizCol(r, idx, "Ghi Chú"),
                        linkHopDong: stripLinkQuery(gvizCol(r, idx, "Link hợp đồng")),
                        trangThaiRaw: gvizCol(r, idx, "còn hạn không"),
              };
      })
      .filter((r) => r.tenNoiBo || r.diaDiem || r.tenKhachHang); // bo dong rong hoan toan
}

router.post("/phap-danh/hop-dong-thue-gian-tong/cap-nhat-tu-sheet", requireAdmin, async (req, res) => {
    const store = load();
    ensureShape(store);
    try {
          const perSheetCounts = {};
          let allNewRows = [];
          for (const s of HOP_DONG_THUE_TONG_SHEETS) {
                  const rows = await fetchHopDongThueTongSheet(s.storeKey);
                  perSheetCounts[s.key] = rows.length;
                  allNewRows = allNewRows.concat(rows);
          }

          // Luyen, 2026-08-24: thay vi replace toan bo, merge: giu du lieu cu,
          // chi bu vao cac field dang rong tu sheet moi. Dong moi (chua co trong
          // store) duoc them vao. Dong cu khong co trong sheet van duoc giu.
          // Key = sheetKey + "|" + tenNoiBo (Mã Điểm Nội Bộ).
          const FILL_FIELDS = [
            "congTy", "congTyRaw", "khuVuc", "dichVu", "maCongTrinh", "diaDiem",
            "tenKhachHang", "mstKhachHang", "hinhThucThue", "tienThueThang",
            "thoiHanThueRaw", "ngayBatDauThue", "ngayHetHanThue",
            "ghiChu", "linkHopDong", "trangThaiRaw",
          ];
          const existingMap = {};
          store.phap_danh_hop_dong_thue_tong.forEach((r) => {
            if (r.sheetKey && r.tenNoiBo) {
              existingMap[r.sheetKey + "|" + r.tenNoiBo] = r;
            }
          });

          let added = 0, filled = 0;
          allNewRows.forEach((newRow) => {
            const key = newRow.sheetKey + "|" + newRow.tenNoiBo;
            const existing = existingMap[key];
            if (existing) {
              // bu vao cac field dang rong trong ban ghi cu
              let changed = false;
              FILL_FIELDS.forEach((f) => {
                const newVal = newRow[f];
                const oldVal = existing[f];
                // "co du lieu" = khong rong va khac 0 (cho tienThueThang)
                const newHasData = newVal !== undefined && newVal !== null && newVal !== "" && newVal !== 0;
                const oldEmpty = oldVal === undefined || oldVal === null || oldVal === "" || oldVal === 0;
                if (newHasData && oldEmpty) { existing[f] = newVal; changed = true; }
              });
              if (changed) filled++;
            } else {
              // dong moi hoan toan
              newRow.id = nextId(store, "phap_danh_hop_dong_thue_tong_seq") || Date.now();
              store.phap_danh_hop_dong_thue_tong.push(newRow);
              existingMap[key] = newRow;
              added++;
            }
          });

          save(store);
          const msg =
                  `Đã đồng bộ từ Google Sheet: ${added} dòng mới, ${filled} dòng được bù thêm dữ liệu (` +
                  HOP_DONG_THUE_TONG_SHEETS.map((s) => `${s.label} ${perSheetCounts[s.key]}`).join(", ") +
                  `).`;
          res.redirect("/phap-danh/hop-dong-thue-gian-tong?success=" + encodeURIComponent(msg));
    } catch (e) {
          res.redirect("/phap-danh/hop-dong-thue-gian-tong?error=" + encodeURIComponent(e.message));
    }
});

// Nhan, 2026-08-25: cap nhat tung truong gian (tu modal chi tiet tren trang
// Hop Dong Thue Gian Tong) -- luu cac truong form vao ban ghi trong store.
router.post("/phap-danh/hop-dong-thue-gian-tong/:id/cap-nhat", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.phap_danh_hop_dong_thue_tong.find(x => String(x.id) === req.params.id);
  if (!r) return res.json({ success: false, error: "Không tìm thấy" });
  const fields = ['tenNoiBo','maCongTrinh','mstKhachHang','tenKhachHang','hinhThucThue',
    'tienThueThang','thoiHanBatDau','thoiHanKetThuc','thoiHanThueRaw','ghiChu','linkHopDong'];
  fields.forEach(f => { if (req.body[f] !== undefined) r[f] = (req.body[f]||'').trim(); });
  // Dong bo voi ten truong cu (ngayBatDauThue/ngayHetHanThue) de khong gay sai
  // le voi cac ham tinh trang thai mau / loc thang dang dung ten truong cu.
  if (req.body.thoiHanBatDau !== undefined) r.ngayBatDauThue = r.thoiHanBatDau;
  if (req.body.thoiHanKetThuc !== undefined) r.ngayHetHanThue = r.thoiHanKetThuc;
  const amt = Number(String(r.tienThueThang||'').replace(/[.,\s]/g,''));
  if (!isNaN(amt)) r.tienThueThang = amt;
  // Nhan, 2026-09-14: "chấm dứt hợp đồng ... tô nền mờ nhạt ... vẫn để" --
  // checkbox gui tu client dang chuoi 'true'/'false' (khong phai checkbox
  // form mac dinh) nen so sanh chuoi, khong dung truthy cua string 'false'.
  if (req.body.daChamDut !== undefined) r.daChamDut = req.body.daChamDut === 'true';
  save(store);
  res.json({ success: true });
});

// Nhan, 2026-09-14: "cho thêm 1 chỗ ô là Thêm Gian" -- them 1 gian moi (rong)
// ngay trong modal "Chi tiet gian" (dung LAI modal co san, chi khac cho ID
// rong = che do tao moi thay vi sua). Gian moi gan vao sheetKey dang xem
// (tab KVC MN/MB, MTD MN/MB dang chon) va cong ty dang chon (Cu/Moi) de hien
// ra ngay trong bang sau khi luu, khong can F5 chuyen tab/cong ty.
router.post("/phap-danh/hop-dong-thue-gian-tong/them", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const sheetInfo = HOP_DONG_THUE_TONG_SHEETS.find((s) => s.key === req.body.sheetKey);
  if (!sheetInfo) return res.json({ success: false, error: "Sheet không hợp lệ" });
  const activeCompany = getCompany(req);

  const r = {
    id: nextId(store, "phap_danh_hop_dong_thue_tong_seq") || Date.now(),
    sheetKey: sheetInfo.storeKey,
    congTy: activeCompany,
    congTyRaw: "",
    khuVuc: "",
    dichVu: "",
    tenNoiBo: "",
    maCongTrinh: "",
    diaDiem: "",
    tenKhachHang: "",
    mstKhachHang: "",
    hinhThucThue: "",
    tienThueThang: 0,
    thoiHanThueRaw: "",
    ngayBatDauThue: "",
    ngayHetHanThue: "",
    ghiChu: "",
    linkHopDong: "",
    trangThaiRaw: "",
    daChamDut: false,
  };
  const fields = ['tenNoiBo','maCongTrinh','mstKhachHang','tenKhachHang','hinhThucThue',
    'tienThueThang','thoiHanBatDau','thoiHanKetThuc','thoiHanThueRaw','ghiChu','linkHopDong'];
  fields.forEach(f => { if (req.body[f] !== undefined) r[f] = (req.body[f]||'').trim(); });
  if (req.body.thoiHanBatDau !== undefined) r.ngayBatDauThue = r.thoiHanBatDau;
  if (req.body.thoiHanKetThuc !== undefined) r.ngayHetHanThue = r.thoiHanKetThuc;
  const amt = Number(String(r.tienThueThang||'').replace(/[.,\s]/g,''));
  r.tienThueThang = isNaN(amt) ? 0 : amt;
  if (req.body.daChamDut !== undefined) r.daChamDut = req.body.daChamDut === 'true';

  store.phap_danh_hop_dong_thue_tong.push(r);
  save(store);
  res.json({ success: true, row: r });
});

// Nhan, 2026-09-14: xoa 1 gian (dung khi tao nham / test) -- cap sau "Them
// gian" cho doi xung voi cac trang khac (hoa-don-dau-ra.js co ca them/sua/
// xoa), admin-only giong cac route sua/them o tren.
router.post("/phap-danh/hop-dong-thue-gian-tong/:id/xoa", requireAdmin, (req, res) => {
  const store = load();
  ensureShape(store);
  const idx = store.phap_danh_hop_dong_thue_tong.findIndex(x => String(x.id) === req.params.id);
  if (idx === -1) return res.json({ success: false, error: "Không tìm thấy" });
  const removed = store.phap_danh_hop_dong_thue_tong.splice(idx, 1)[0];
  save(store);
  res.json({ success: true, removed });
});

// Nhan, 2026-08-25: doc noi dung hop dong PDF tu Google Drive (link co san
// trong truong linkHopDong cua ban ghi). Lay file qua Drive API bang access
// token hien co (dung chung voi Gmail OAuth), parse PDF bang pdf-parse, trich
// xuat cac dieu khoan chinh (vi tri, tien thue, thoi han, thanh toan).
function extractHopDongInfo(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  let dieuKhoanThanhToan = '', viTri = '', tienThue = '', thoiHan = '';
  const dkIdx = lines.findIndex(l => /thanh\s*to[áa]n|payment/i.test(l));
  if (dkIdx >= 0) dieuKhoanThanhToan = lines.slice(dkIdx, dkIdx+5).join('\n');
  const vtIdx = lines.findIndex(l => /v[ịi]\s*tr[íi]|di[eệ]n\s*t[íi]ch|location|premises/i.test(l));
  if (vtIdx >= 0) viTri = lines.slice(vtIdx, vtIdx+3).join('\n');
  const ttIdx = lines.findIndex(l => /ti[eề]n\s*thu[eê]|gi[áa]\s*thu[eê]|rent|fee/i.test(l));
  if (ttIdx >= 0) tienThue = lines.slice(ttIdx, ttIdx+3).join('\n');
  const thIdx = lines.findIndex(l => /th[oờ]i\s*h[aạ]n|term|duration|hi[eệ]u\s*l[uự]c/i.test(l));
  if (thIdx >= 0) thoiHan = lines.slice(thIdx, thIdx+3).join('\n');
  return { dieuKhoanThanhToan, viTri, tienThue, thoiHan };
}

router.get("/phap-danh/hop-dong-thue-gian-tong/:id/doc-hop-dong", async (req, res) => {
  const store = load();
  ensureShape(store);
  const r = store.phap_danh_hop_dong_thue_tong.find(x => String(x.id) === req.params.id);
  if (!r || !r.linkHopDong) return res.json({ success: false, error: "Không có link hợp đồng" });
  const driveMatch = r.linkHopDong.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  const openMatch = r.linkHopDong.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  const fileId = (driveMatch && driveMatch[1]) || (openMatch && openMatch[1]);
  if (!fileId) return res.json({ success: false, error: "Không nhận dạng được file ID từ link Drive" });
  try {
    const gmailApi = require('../utils/gmailApi');
    const accessToken = await gmailApi.getValidAccessToken(store);
    const dlResp = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    if (!dlResp.ok) throw new Error('Drive trả lỗi ' + dlResp.status + ': ' + await dlResp.text());
    const arrayBuffer = await dlResp.arrayBuffer();
    const pdfBuffer = Buffer.from(arrayBuffer);
    const pdfParse = require('pdf-parse');
    const data = await pdfParse(pdfBuffer);
    const result = extractHopDongInfo(data.text || '');
    res.json({ success: true, ...result, rawText: (data.text||'').substring(0, 3000) });
  } catch(e) {
    res.json({ success: false, error: e.message });
  }
});

// "Đối chiếu gian XHD, Tiền thuê" (routes/hoa-don-dau-vao.js) can doc lai
// danh sach hop dong thue gian (benChoThue/tienThueThang/...) -- xuat them
// cac ham nay (truoc gio chi co router duoc export) de dung LAI, khong doan
// lai/copy code (giong cach congno-ncc.js dang dung lai ensureShape/
// ensureDefaults/groupRowsByInvoice cua hoa-don-dau-vao.js).
module.exports.ensureShape = ensureShape;
module.exports.ensureThueDefaults = ensureThueDefaults;
module.exports.computeTrangThaiHD = computeTrangThaiHD;
// trigger deploy Fri Sep 25 09:24:24 +07 2026
