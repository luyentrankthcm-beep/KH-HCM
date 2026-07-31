// Chi Nhan, 2026-07-30: mapping TAI KHOAN NGAN HANG THUC -> cong ty (kh_cu/
// kh_moi) -- rut ra rieng file nay (truoc do dinh nghia trung lap ben trong
// routes/congno.js) de routes/baocao.js (bao cao Tờ khai thuế GTGT) dung
// LAI CHINH XAC, khong doan lai/co the lech nhau giua 2 noi.
const BANK_COMPANY = {
  BIDV123456: "kh_cu",
  ACB31268: "kh_cu",
  BIDV7704: "kh_cu",
  BIDV77020: "kh_cu",
  MB11521268: "kh_cu",
  BIDV7701: "kh_moi",
  VTB982: "kh_moi",
  BIDV7702: "kh_moi",
  BIDV77021: "kh_moi",
  MB02865168: "kh_moi",
  BIDV8613600999: "kh_moi",
};

function companyForBankLabel(bankLabel) {
  return BANK_COMPANY[bankLabel] || "kh_cu";
}

module.exports = { BANK_COMPANY, companyForBankLabel };
