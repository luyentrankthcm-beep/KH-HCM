const express = require("express");
const XLSX = require("xlsx");
const multer = require("multer");
const { load, save, nextId } = require("../store");
const { requireLogin, requireAdmin } = require("../middleware/auth");
const { getCompany } = require("../utils/companies");
const { parseHopDongHcmWorkbook, isNccRow } = require("../utils/hopDongHcmParser");

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

router.get("/phap-danh/hop-dong-thue-gian-hang", (req, res) => {
  const store = load();
  ensureShape(store);
  const activeCompany = getCompany(req);
  const loaiFilter = req.query.loai || "";
  const thangFilter = req.query.thang || "";
  const todayStr = new Date().toISOString().slice(0, 10);
  let rows = store.phap_danh_hop_dong_thue
    .map(ensureThueDefaults)
    .filter((r) => (r.congTy || "kh_cu") === activeCompany);
  if (loaiFilter) rows = rows.filter((r) => r.loaiHinh === loaiFilter);
  if (thangFilter) rows = rows.filter((r) => monthOverlaps(r, thangFilter));
  rows.forEach((r) => {
    r.trangThaiHD = computeTrangThaiHD(r, todayStr);
  });
  rows.sort((a, b) => (a.loaiHinh === b.loaiHinh ? a.gian.localeCompare(b.gian) : a.loaiHinh.localeCompare(b.loaiHinh)));
  const counts = {
    "Khu vui chơi": store.phap_danh_hop_dong_thue.filter((r) => (r.congTy || "kh_cu") === activeCompany && r.loaiHinh === "Khu vui chơi").length,
    "Máy tự động": store.phap_danh_hop_dong_thue.filter((r) => (r.congTy || "kh_cu") === activeCompany && r.loaiHinh === "Máy tự động").length,
  };
  res.render("phapdanh-hopdong-thue", {
    userName: req.session.userName,
    rows,
    loaiFilter,
    thangFilter,
    counts,
    totalForCompany: store.phap_danh_hop_dong_thue.filter((r) => (r.congTy || "kh_cu") === activeCompany).length,
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
    .filter((r) => (r.congTy || "kh_cu") === activeCompany);
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
    function groupKey(diaDiem, congTy) {
      return (diaDiem || "(không có địa điểm)") + "||" + (congTy || "");
    }
    const groups = new Map();
    nccRows.forEach((r) => {
      const key = groupKey(r.diaDiem, r.congTy);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });

    const existingByKey = new Map();
    store.phap_danh_hop_dong_ncc.forEach((r) => existingByKey.set(groupKey(r.tenNCC, r.congTy), r));

    // Xay dung toan bo cac truong tong hop tu 1 nhom groupRows (dung chung
    // cho ca NCC moi va NCC da co, vi gio ca 2 truong hop deu GHI DE giong het
    // nhau -- chi khac o cho tao moi object hay ghi vao object da co san).
    function buildFieldsFromRows(diaDiem, groupRows) {
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

    let addedNew = 0;
    let updatedExisting = 0;
    const addedNames = [];
    const updatedNames = [];
    const UPLOAD_TAG = "upload " + req.file.originalname + " " + new Date().toISOString().slice(0, 10);

    groups.forEach((groupRows, key) => {
      const diaDiem = groupRows[0].diaDiem || "(không có địa điểm)";
      const existing = existingByKey.get(key);
      const fields = buildFieldsFromRows(diaDiem, groupRows);
      if (!existing) {
        // NCC hoan toan moi
        store.phap_danh_hop_dong_ncc.push({
          id: nextId(store, "phap_danh_hop_dong_ncc_seq") || Date.now(),
          ...fields,
          ghiChu: `Import tự động từ ${UPLOAD_TAG} (Thuộc BP: Khác/Mua bán vocher), ${groupRows.length} dòng nguồn -- chị kiểm tra lại các trường tổng hợp, dữ liệu lấy tự động chưa qua rà soát tay như 48 hợp đồng ban đầu.`,
          createdAt: new Date().toISOString(),
          source: UPLOAD_TAG,
        });
        addedNew++;
        addedNames.push(diaDiem);
      } else {
        // NCC da co -- Luyen xac nhan 2026-07-21 muon GHI DE lai toan bo cac
        // truong (ke ca link/han/gia tri/chi tiet) theo dung sheet moi nhat
        // moi lan bam "Cap nhat hop dong", giu nguyen id/createdAt goc.
        Object.assign(existing, fields);
        existing.ghiChu = `Đã cập nhật lại từ ${UPLOAD_TAG} (Thuộc BP: Khác/Mua bán vocher), ${groupRows.length} dòng nguồn.`;
        existing.source = UPLOAD_TAG;
        updatedExisting++;
        updatedNames.push(diaDiem);
      }
    });

    save(store);
    let msg = `Đã quét ${nccRows.length} dòng NCC từ "${req.file.originalname}" (${groups.size} địa điểm). `;
    if (addedNew > 0) msg += `Thêm ${addedNew} NCC mới: ${addedNames.join(", ")}. `;
    if (updatedExisting > 0) msg += `Đã cập nhật lại ${updatedExisting} NCC đã có theo sheet mới nhất: ${updatedNames.join(", ")}. `;
    if (addedNew === 0 && updatedExisting === 0) msg += "Không có gì mới so với dữ liệu hiện tại.";
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

module.exports = router;
