// Luyen, 2026-07-21: "từ cái nội dung với 2 chi 2 ngân hàng này á bạn sẽ liên
// kết qua cái bên file chi phí á có thanh toán kh cũ và kh mới của gian nào á
// note lại chi ngày mấy ngân hàng nào bên chi phí cho tôi nhá cập nhật hàng
// ngày cho tôi luôn nhá" -- doc giao dich "Chi" (tien thue gian) tren cac tai
// khoan ngan hang da co san trong he thong, tu tach ten gian tu dien giai, doi
// chieu voi danh sach gian (phap_danh_hop_dong_thue) de biet dung cong ty +
// gian nao, roi tao dong Chi Phi tuong ung (an toan: CHI tao khi khop chac
// chan 1 gian duy nhat, con lai bao cao de Luyen tu xu ly, khong doan bua).
function removeDiacritics(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, (m) => (m === "đ" ? "d" : "D"));
}

const PAYER_RE = /^(?:CTY|C\.?ty|CONG TY|Cong ty)\.?\s*(?:TNHH\s*)?(?:DICH VU VA\s*)?(?:GIAI TRI\s*)?K\s*(?:VA|V\s*A|&)?\s*H\.?\s*/i;
const VERB_RE = /^(?:TT|tt|THANH TOAN|Thanh toan|thanh toan|NOP TIEN|Nop tien|nop tien)\s*/;
const CATEGORY_RE = /(tien thue ghe|tien thue)\s*(.*)$/i;

// Tach ten gian tho tu 1 dong dien giai giao dich ngan hang. Tra ve null neu
// khong nhan dien duoc mau "CTY K&H (tt) tien thue ..." ro rang -- CHU DINH
// bo qua nhung dong mo ho (tien hang, tien dien, luong, phi ngan hang, CTNB
// chuyen noi bo...) thay vi doan sai.
// Nhieu giao dich (chuyen khoan lien ngan hang) co tien to boilerplate cua
// ngan hang truoc noi dung thuc su, vd "BDR-TKThe :xxx| tai SCB VN. ND  CTY
// K VA H TT ..." -- bo tien to nay truoc khi tim PAYER_RE (von can nam o dau
// chuoi), neu khong se bi coi la "khong nhan dien duoc" oan.
const BANK_BOILERPLATE_RE = /^(?:BDR-TKThe\s*:[^|]*\|\s*tai\s+[^.]*\.\s*ND\s*)/i;

function extractGianRentText(descRaw) {
  let s = removeDiacritics(descRaw || "").trim().replace(/\s+/g, " ");
  s = s.replace(BANK_BOILERPLATE_RE, "").trim();
  const afterPayer = s.replace(PAYER_RE, "");
  if (afterPayer === s) return null; // khong phai giao dich do K&H tu chi (vd BDR- tien ngan hang khac, CTNB...)
  let rest = afterPayer.replace(VERB_RE, "");
  const catMatch = rest.match(CATEGORY_RE);
  if (!catMatch) return null;
  let gian = catMatch[2].trim();
  gian = gian.replace(/\s*theo[\s\S]*$/i, "");
  gian = gian.replace(/^\s*(?:mat bang|vi tri)?\s*\d{1,2}\.\d{2,4}\s*/i, "");
  gian = gian.replace(/\s*(?:t\d{1,2}\.\d{2,4}|thang\s*\d{1,2}[.\/]\d{2,4}(?:\s*-\s*\d{1,2}[.\/]\d{2,4})?)\s*$/i, "");
  gian = gian.replace(/\s*(?:tu\s+\d{1,2}[-.]\d{1,2}[\s\S]*)$/i, "");
  gian = gian.replace(/\s*-\s*(?:cong ty|cty|chi nhanh)[\s\S]*$/i, "");
  gian = gian.replace(/\s*hd\s*(?:so)?\s*[\w\-\/]+$/i, "");
  gian = gian.replace(/\s*ctlnhido\d+[\s\S]*$/i, "");
  gian = gian.replace(/[\s\-]+$/, "");
  gian = gian.trim();
  if (!gian || /^(nha|van phong|mat bang)$/i.test(gian)) return null;
  return gian;
}

// "posh"/"jp"/"aeon"/"mall" xuat hien qua nhieu gian nen bo qua (chi gay
// loang diem, khong giup phan biet). NGUOC LAI "tutu"/"fz"/"farm"/"pinball"
// la ten thuong hieu rieng, PHAI GIU LAI de phan biet cac gian CHUNG 1 mat
// bang (vd "tàu Lotte Gò Vấp" khac "FZ MN VR Lotte mart Gò Vấp") -- bo qua
// nham 2 tu nay tung lam 2 gian nay bi diem hoa lam 1, khong the phan biet.
const STOPWORDS = new Set(["mn", "ghe", "gian", "kvc", "p", "tang", "1", "2", "3", "posh", "jp", "aeon", "mall"]);

function tokenize(s) {
  return removeDiacritics(s)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((w) => w.length >= 2 && !STOPWORDS.has(w));
}

// Doi chieu ten gian tho (tu extractGianRentText) voi danh sach gian hien co
// (store.phap_danh_hop_dong_thue) -- so sanh ca 2 truong "gian" (ten diem
// xuat hoa don) va "tenDiemNoiBo" (ten noi bo). Diem = ty le token cua chuoi
// giao dich xuat hien trong ten gian. Chi coi la KHOP CHAC CHAN khi diem cao
// nhat >= 0.6 VA vuot han (>=0.15) diem cao thu nhi -- neu khong ro rang, tra
// ve null (khong doan) de Luyen tu xu ly.
function matchGianRecord(rentText, gianList) {
  const queryTokens = tokenize(rentText);
  if (queryTokens.length === 0) return null;
  const queryTokenSet = new Set(queryTokens);
  const scored = gianList.map((r) => {
    const targetText = `${r.gian || ""} ${r.tenDiemNoiBo || ""}`;
    const targetTokensArr = tokenize(targetText);
    const targetTokens = new Set(targetTokensArr);
    const hits = queryTokens.filter((t) => targetTokens.has(t)).length;
    // Giao dich ngan hang thuong co nhieu chu thua ("vuot doanh thu",
    // "theo hop dong"...) quanh ten gian that -- dung MIN(so token cua ben
    // NGAN hon) lam mau so, de 1 cai ten ngan (vd "sc vivo") xuat hien tron
    // ven trong 1 dong dai van tinh la khop hoan toan, thay vi bi pha loang
    // diem vi qua nhieu tu thua ben con lai.
    const denom = Math.min(queryTokenSet.size, targetTokens.size) || 1;
    const score = hits / denom;
    return { record: r, score };
  });
  scored.sort((a, b) => b.score - a.score);
  if (scored.length === 0) return null;
  const best = scored[0];
  const second = scored[1] || { score: 0 };
  if (best.score >= 0.6 && best.score - second.score >= 0.15) {
    return best.record;
  }
  return null;
}

// Luyen, 2026-07-23: "thêm cho tôi 1 cột tài khoản ... doanh thu chia sẻ gian
// thì đưa vô 1388 còn lại thì để 131 ... dựa vào hợp đồng thuê gian" -- xac
// dinh 1 hop dong trong "Hop Dong Thue Gian Hang" (store.phap_danh_hop_dong_thue)
// co phai kieu "doanh thu chia se" hay khong, dua vao 3 dau hieu:
//   1. doanhThuChiaSe = true -- Luyen TU TAY tich chon truc tiep tren trang
//      Hop Dong Thue Gian Hang (2026-07-23 lan 2: "có tàu bình tân với nhà ma
//      bình dương hay ghost bình dương cũng là doanh thu chia sẻ, cho thêm
//      thủ công cũng được" -- vi khong phai gian nao cung tu doan duoc tu van
//      ban hop dong, can co cach Luyen tu danh dau).
//   2. dieuKhoanThanhToan co ghi ro "chia se doanh thu" hoac ty le "% doanh thu"
//      (vd "POSH MN KIWOOZA QUẬN 2": "Chia se doanh thu: Khach 35% / Cong ty 65%").
//   3. hinhThucThuTien duoc dien (truong nay CHI dien khi la kieu "mall giu
//      tien roi cuoi thang tra ve qua tai khoan cho minh phan da tru tien thue
//      voi phi dich vu" -- xem chu thich trong phap-danh.js -- day cung la 1
//      dang doanh thu chia se, khac voi thu tien truc tiep qua VietQR/POS).
// Khong khop duoc gian nao / gian khong co dau hieu nao o tren -> mac dinh TK
// 131 (thu thuong), giu nguyen quy uoc TKCO da dung ben doi soat Momo.
function isDoanhThuChiaSeRecord(r) {
  if (!r) return false;
  if (r.doanhThuChiaSe === true) return true;
  const dtt = removeDiacritics(r.dieuKhoanThanhToan || "").toLowerCase();
  if (/chia se doanh thu|%\s*doanh thu/.test(dtt)) return true;
  if (r.hinhThucThuTien && String(r.hinhThucThuTien).trim()) return true;
  return false;
}

// Chuan hoa 1 chuoi ten gian de so sanh CHINH XAC (bo dau, thuong, chi giu
// chu+so, gop khoang trang thua) -- dung cho bang alias ben duoi, KHAC voi
// tokenize() (dung cho fuzzy match, giu tap hop tu roi rac).
function normalizeGianKey(s) {
  return removeDiacritics(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

// Luyen, 2026-07-23 (lan 2): "có thể viết tắc á" -- ten gian trong Chi Phi
// thuong la ma viet tat rieng cua Luyen (vd "TÀU BT" = Funzone/tàu AE Bình
// Tân, "GHOST AMBD" = Nhà ma AE Bình Dương) MA fuzzy match (matchGianRecord)
// KHONG nhan ra duoc vi khac han ve chu cai voi ten day du trong hop dong.
// Thay vi TU DOAN bang rieng viet tat -> ten day du (de sai, vd BD co the la
// "Bình Dương" hoac "Bà Rịa" tuy ngu canh), them 1 truong "aliasGian" de
// CHINH LUYEN tu dien cac ma viet tat cho tung hop dong (vd dien "TAU BT,
// TÀU BT" vao hop dong "Funzone (tàu+DIY )AE Bình Tân") -- bang tra nay dung
// KHOP CHINH XAC (khong doan), an toan hon fuzzy rat nhieu.
function buildGianAliasIndex(gianList) {
  const idx = new Map();
  (gianList || []).forEach((r) => {
    const aliases = String(r.aliasGian || "")
      .split(/[,;\n]/)
      .map((a) => a.trim())
      .filter(Boolean);
    aliases.forEach((a) => {
      const key = normalizeGianKey(a);
      if (key) idx.set(key, r);
    });
  });
  return idx;
}

// Nhieu dia diem gop chung 1 dong Chi Phi (vd "Tàu GV, TP, BD" -- noi 3 gian
// qua dau phay) -- KHONG THE gan chac chan 1 hop dong duy nhat cho ca cum,
// nen BO QUA (tra ve null, mac dinh TK 131) thay vi doan dai theo gian dau
// tien/manh nhat.
const MULTI_SITE_RE = /,|\+| va /i;

// Luyen, 2026-07-23 (lan 2): rat nhieu dong Chi Phi ghi kieu "TEN HANG HOA
// \n TÊN GIAN" (vd "KẸO ĐỒ CHƠI \nTÀU BT", "XÚC XÍCH CAO BỒI \nTÀU EST") --
// ten gian nam O DONG CUOI CUNG (sau dau xuong dong), chu khong phai truoc
// dau "-"/"_". Thu ca 2 kieu (truoc dau "-"/"_" VA dong cuoi cung sau "\n")
// lam ung vien tra alias, cung voi chuoi day du ban dau.
// CHI lay dong cuoi cung khi dung DUNG 2 dong (ten hang + 1 gian) -- tu 3
// dong tro len (vd "NƯỚC SUỐI\nDIY SORA\nTÀU TP\nTÀU BT", moi gian 1 dong
// rieng) la dau hieu 1 dong Chi Phi gop CHI PHI CHUNG cho NHIEU gian, lay dong
// cuoi se bo sot cac gian khac trong cung dong -- BO QUA ca dong cuoi (van con
// candidate chuoi day du/truoc dau "-" o tren, se khong khop alias don le nao
// vi la chuoi gop nhieu ten, an toan).
function buildGianTextCandidates(gianText) {
  const candidates = [gianText];
  const beforeDash = gianText.split(/[-_]/)[0];
  if (beforeDash !== gianText) candidates.push(beforeDash);
  const lines = gianText.split(/\n/).map((l) => l.trim()).filter(Boolean);
  if (lines.length === 2) candidates.push(lines[1]);
  return candidates;
}

// Tim hop dong thue gian khop voi 1 gian tho trong Chi Phi -- uu tien alias
// CHINH XAC (Luyen tu dien qua aliasGian, thu ca chuoi day du/phan truoc dau
// "-"/dong cuoi cung sau "\n" -- xem buildGianTextCandidates), cuoi cung moi
// fallback ve matchGianRecord (fuzzy token, nguong 0.6) cho ten gian day du/
// ro rang. Gian gop nhieu dia diem (dau phay/"+"/"và") thi BO QUA alias/fuzzy
// (tra ve null, mac dinh TK 131) vi khong the gan chac chan 1 hop dong duy
// nhat cho ca cum.
function findContractForGianText(gianText, gianList, aliasIndex) {
  if (!gianText || !gianText.trim()) return null;
  if (aliasIndex) {
    for (const candidate of buildGianTextCandidates(gianText)) {
      const key = normalizeGianKey(candidate);
      if (key && aliasIndex.has(key)) return aliasIndex.get(key);
    }
  }
  const normNoDiacritic = removeDiacritics(gianText).toLowerCase();
  if (MULTI_SITE_RE.test(" " + normNoDiacritic + " ")) return null;
  return matchGianRecord(gianText, gianList);
}

module.exports = {
  extractGianRentText,
  matchGianRecord,
  isDoanhThuChiaSeRecord,
  buildGianAliasIndex,
  findContractForGianText,
  normalizeGianKey,
  removeDiacritics,
};
