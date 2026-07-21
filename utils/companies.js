// 2 cong ty (phap nhan) rieng cung dung 1 web nay: "KH Cu" (cong ty goc,
// TNHH DICH VU VA GIAI TRI K&H) va "KH Moi" (phap nhan moi them, TNHH GIAI
// TRI K&H -- xem sao ke BIDV7701/BIDV7702/VP58888, ten tai khoan tren sao ke
// la "CONG TY TNHH GIAI TRI K&H"/"K VA H"). Cung 1 danh muc gian/cong trinh
// (Momo/ZVP gian_mapping, cua_hang_mapping...) duoc dung CHUNG cho ca 2 cong
// ty vi la cung 1 he thong van hanh (gian/mat bang), chi khac ngan hang nhan
// tien -- giong cach cac kenh Momo/ZVP/VietQR cua KH Cu da chia se cac bang
// nay voi nhau tu truoc.
//
// Luu lua chon cong ty dang xem trong cookie-session (giong cach userId/
// userName da luu) -- moi nguoi dung (moi may) co the dang xem 1 cong ty
// khac nhau cung luc, khong anh huong nguoi khac.
const COMPANIES = {
  kh_cu: {
    key: "kh_cu",
    label: "KH Cũ",
    shortLabel: "Cũ",
    fullName: "CÔNG TY TNHH DỊCH VỤ VÀ GIẢI TRÍ K&H",
  },
  kh_moi: {
    key: "kh_moi",
    label: "KH Mới",
    shortLabel: "Mới",
    fullName: "CÔNG TY TNHH GIẢI TRÍ K&H",
  },
};

const COMPANY_KEYS = Object.keys(COMPANIES);
const DEFAULT_COMPANY = "kh_cu";

function getCompany(req) {
  const c = req && req.session && req.session.company;
  return COMPANIES[c] ? c : DEFAULT_COMPANY;
}

module.exports = { COMPANIES, COMPANY_KEYS, DEFAULT_COMPANY, getCompany };
