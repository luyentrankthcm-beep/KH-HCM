// Nhan, 2026-09-17: "hóa đơn 115 có nhiều hàng hóa lắm mà 1 cái là 1 dòng cho
// tôi nhá đvt số lượng và cả thành tiền và cả tổng cho tôi nhá ... còn địa chỉ
// số điện thoại đồ trên hóa đơn có mà sao bạn không lấy từ hóa đơn" -- doc PDF
// hoa don dien tu (Ban the hien) de trich xuat:
//  - Dia chi / Dien thoai cua NGUOI BAN (Ben A) -- lay tu chinh file hoa don,
//    KHONG lay tu Hop dong NCC (khac voi "Dai dien" luon lay tu Hop dong).
//  - Bang hang hoa day du (STT, Ten hang hoa, DVT, So luong, Don gia, Thanh
//    tien) + dong Tong cong, thay vi chi 1 dong "Noi dung" rut gon.
//
// pdf-parse tra text lien tuc, moi dong 1 field (KHONG co khoang trang giua
// STT/Ten/DVT/SL/DonGia/ThanhTien vi PDF export dinh dang bang khong co dau
// phan cach). Vi du 1 dong that: "4Đầu bắn tôn 8liCái1015.000150.000" (STT=4,
// Ten="Đầu bắn tôn 8li", DVT="Cái", SL=10, DonGia=15.000, ThanhTien=150.000).
// Chien luoc tin cay nhat: DUNG PHEP TOAN (Thanh tien = So luong x Don gia,
// dung nhu tieu de cot "6 = 4 x 5" tren hoa don) de tim diem cat 3 so o CUOI
// dong, thay vi doan boi khoang trang (khong co).

const UNIT_WORDS = [
  "cái", "cây", "viên", "bịch", "cuộn", "mét", "bộ", "kg", "hộp", "thùng",
  "lít", "chai", "tấm", "con", "chiếc", "đôi", "bao", "gói", "ổ", "sợi",
  "ống", "tuýp", "cuốn", "quyển", "tờ", "kiện", "lon", "thanh", "bó", "cặp",
  "m2", "m3", "kg/hộp", "m",
].sort((a, b) => b.length - a.length);

function removeDiacritics(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");
}

// Tim diem cat numBlob (chuoi so + dau cham) thanh 3 phan lien tiep SL|DonGia|ThanhTien
// sao cho SL * DonGia === ThanhTien (sau khi bo dau cham). Uu tien SL ngan nhat.
function splitByMath(numBlob) {
  const n = numBlob.length;
  let best = null;
  for (let i = 1; i < n - 1; i++) {
    for (let j = i + 1; j < n; j++) {
      const p1 = numBlob.slice(0, i);
      const p2 = numBlob.slice(i, j);
      const p3 = numBlob.slice(j);
      const n1 = Number(p1.replace(/\./g, ""));
      const n2 = Number(p2.replace(/\./g, ""));
      const n3 = Number(p3.replace(/\./g, ""));
      if (!n1 || !n2 || !n3) continue;
      if (n1 * n2 === n3) {
        if (!best || p1.length < best.p1.length) best = { n1, n2, n3 };
      }
    }
  }
  if (!best) return null;
  return {
    soLuong: String(best.n1),
    donGia: best.n2.toLocaleString("de-DE"),
    thanhTien: best.n3.toLocaleString("de-DE"),
  };
}

// Tach "Ten hang hoaDVT" dinh lien nhau: thu khop don vi tinh o CUOI chuoi
// (danh sach UNIT_WORDS, khong phan biet dau/hoa thuong). Khong khop duoc
// thi coi ca chuoi la ten hang, de trong DVT.
function splitTenDvt(prefix) {
  const p = (prefix || "").trim();
  const normP = removeDiacritics(p.toLowerCase());
  for (const u of UNIT_WORDS) {
    const normU = removeDiacritics(u.toLowerCase());
    if (normP.endsWith(normU) && normP.length > normU.length) {
      return { ten: p.slice(0, p.length - normU.length).trim(), dvt: p.slice(p.length - normU.length).trim() };
    }
  }
  return { ten: p, dvt: "" };
}

// Parse 1 dong hang hoa (da bo STT o dau). Tra ve null neu khong tach duoc so.
function parseItemBody(rest) {
  const numBlobMatch = rest.match(/([\d.]+)$/);
  if (!numBlobMatch) return { ten: rest.trim(), dvt: "", soLuong: "", donGia: "", thanhTien: "" };
  const numBlob = numBlobMatch[1];
  const prefix = rest.slice(0, rest.length - numBlob.length);
  const split = splitByMath(numBlob);
  const { ten, dvt } = splitTenDvt(prefix);
  if (!split) return { ten: ten || rest.trim(), dvt, soLuong: "", donGia: "", thanhTien: "" };
  return { ten, dvt, ...split };
}

// Lay gia tri dong NGAY SAU dong chua label (vd "Địa chỉ:" rồi giá trị ở dòng kế).
function grabAfterLabel(section, labelRegex) {
  const m = section.match(labelRegex);
  if (!m) return "";
  const rest = section.slice(m.index + m[0].length);
  const lines = rest.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  return lines[0] || "";
}

// Ham chinh: nhan text da pdf-parse, tra ve { diaChi, dienThoai, items, tongCong }
function extractHoaDonPdfInfo(rawText) {
  const text = rawText || "";
  const out = { diaChi: "", dienThoai: "", items: [], tongCong: "" };

  // Chi lay thong tin NGUOI BAN (truoc doan "Ten don vi"/"Ho ten nguoi mua
  // hang" cua Ben mua), tranh nham dia chi/dien thoai cua Ben mua (K&H).
  const buyerIdx = text.search(/Họ tên người mua hàng|Tên đơn vị/i);
  const sellerSection = buyerIdx > -1 ? text.slice(0, buyerIdx) : text;

  out.diaChi = grabAfterLabel(sellerSection, /Địa\s*chỉ\s*:?/i);
  const dienThoaiRaw = grabAfterLabel(sellerSection, /Điện\s*thoại\s*:?/i);
  out.dienThoai = /^[\d][\d\s.\-]{6,}$/.test(dienThoaiRaw) ? dienThoaiRaw : "";

  // Bang hang hoa: giua dong header "STT Tên hàng hóa..." va dong
  // "Cộng tiền hàng hóa, dịch vụ:".
  const headerMatch = text.match(/STT\s*Tên\s*hàng\s*hóa/i);
  const footerMatch = text.match(/Cộng\s*tiền\s*hàng\s*hóa[^:\n]*:?/i);
  if (headerMatch && footerMatch) {
    const headerEnd = headerMatch.index + headerMatch[0].length;
    const footerIdx = text.indexOf(footerMatch[0], headerEnd);
    if (footerIdx > headerEnd) {
      const tableBlock = text.slice(headerEnd, footerIdx);
      const rows = tableBlock.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      let expectStt = 1;
      rows.forEach((line) => {
        const m = line.match(/^(\d+)(.*)$/);
        if (!m) return;
        const stt = parseInt(m[1], 10);
        if (stt !== expectStt) return; // bo qua dong tieu de phu ("123456 = 4 x 5"...)
        const item = parseItemBody(m[2]);
        out.items.push(item);
        expectStt++;
      });

      const afterFooter = text.slice(footerIdx + footerMatch[0].length);
      const totalLine = afterFooter.split(/\r?\n/).map((l) => l.trim()).filter(Boolean)[0];
      if (totalLine && /^[\d.]+$/.test(totalLine)) out.tongCong = totalLine;
    }
  }

  return out;
}

module.exports = { extractHoaDonPdfInfo };
