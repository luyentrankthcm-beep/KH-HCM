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
  if (thangFilter) rows = rows.filter((r) => monthOverlaps(r, thangFilter));
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
  // Hop dong nhap tay cu (khung rong ban dau) co the chua co congTy -- van
  // hien thi o ca 2 cong ty thay vi an mat, cho den khi Luyen bo sung.
  rows = rows.filter((r) => !r.congTy || r.congTy === activeCompany);
  rows.forEach((r) => {
    r.trangThaiNCC = computeTrangThaiNCC(r, todayStr);
  });
  rows.sort((a, b) => (a.ngayHetHan < b.ngayHetHan ? 1 : a.ngayHetHan > b.ngayHetHan ? -1 : a.tenNCC.localeCompare(b.tenNCC)));
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
    const { tenNCC, noiDung, hangHoaMua, phanLoaiHD, soHopDong, ngayKy, ngayHetHan, giaTriHopDong, linkHopDong, ghiChu } = req.body;
    if (!tenNCC || !tenNCC.trim()) throw new Error("Thiếu Tên NCC.");
    const amt = giaTriHopDong ? Number(String(giaTriHopDong).replace(/[^\d]/g, "")) : 0;
    store.phap_danh_hop_dong_ncc.push({
      id: nextId(store, "phap_danh_hop_dong_ncc_seq") || Date.now(),
      tenNCC: tenNCC.trim(),
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

module.exports = router;
// Luyen, 2026-08-01: "từ cái hợp đồng thuê gian á nó sẽ có tên đối tác ký hợp
// đồng với mình từ cái tên đó bạn map với lại hóa đơn đầu vào" -- trang moi
// "Đối chiếu gian XHD, Tiền thuê" (routes/hoa-don-dau-vao.js) can doc lai
// danh sach hop dong thue gian (benChoThue/tienThueThang/...) -- xuat them
// cac ham nay (truoc gio chi co router duoc export) de dung LAI, khong doan
// lai/copy code (giong cach congno-ncc.js dang dung lai ensureShape/
// ensureDefaults/groupRowsByInvoice cua hoa-don-dau-vao.js).
module.exports.ensureShape = ensureShape;
module.exports.ensureThueDefaults = ensureThueDefaults;
module.exports.computeTrangThaiHD = computeTrangThaiHD;
