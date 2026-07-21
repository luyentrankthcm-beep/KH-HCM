// Parser cho file MISA "So chi tiet cong no phai thu theo cong trinh" (TK 131),
// dang xuat tu MISA: 1 sheet duy nhat, moi "cong trinh" la 1 khoi gom dong tieu
// de "Ten cong trinh: X", dong "Ten khach hang: Y", roi cac dong giao dich
// (Ngay hach toan | Ngay chung tu | So chung tu | Dien giai | TK cong no |
// TK doi ung | No | Co | Du No | Du Co), ket bang 1 dong "Cong" chua tong
// No/Co va so du cuoi ky cua RIENG cong trinh do. File ket thuc bang 1 dong
// "Tong cong" cho toan bo so.
//
// CHI luu tom tat theo tung cong trinh (khong luu tung dong giao dich) de
// khong lam phinh store.json (file goc ~12.7k dong) -- du de doi soat
// "du no MISA" vs "chua thu theo ngan hang" o muc cong trinh, dung nhu
// Luyen yeu cau (2026-07-17).
const XLSX = require("xlsx");

function parseMisaCongNoXlsx(buffer) {
  const wbLite = XLSX.read(buffer, { type: "buffer", bookSheets: true });
  const sheetName = wbLite.SheetNames[0];
  const wb = XLSX.read(buffer, { type: "buffer", sheets: [sheetName] });
  const ws = wb.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const byCongTrinh = {};
  const order = [];
  let current = null;
  let tongCong = { tongNo: 0, tongCo: 0 };

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    const c0 = row[0];
    const c0s = typeof c0 === "string" ? c0.trim() : "";

    if (c0s.toLowerCase().startsWith("tên công trình:")) {
      const tenCongTrinh = c0s.replace(/^tên công trình:\s*/i, "").trim();
      if (!tenCongTrinh) continue;
      current = {
        tenCongTrinh,
        tenKhachHang: null,
        soDong: 0,
        tongNo: 0,
        tongCo: 0,
        soDuNo: 0,
        soDuCo: 0,
      };
      byCongTrinh[tenCongTrinh] = current;
      order.push(tenCongTrinh);
      continue;
    }
    if (c0s.toLowerCase().startsWith("tên khách hàng:")) {
      if (current) current.tenKhachHang = c0s.replace(/^tên khách hàng:\s*/i, "").trim();
      continue;
    }
    if (c0s === "Tổng cộng") {
      tongCong = { tongNo: Number(row[6]) || 0, tongCo: Number(row[7]) || 0 };
      continue;
    }
    const dienGiai = typeof row[3] === "string" ? row[3].trim() : row[3];
    if (dienGiai === "Cộng") {
      if (current) {
        current.tongNo = Number(row[6]) || 0;
        current.tongCo = Number(row[7]) || 0;
        current.soDuNo = Number(row[8]) || 0;
        current.soDuCo = Number(row[9]) || 0;
      }
      continue;
    }
    // Dong giao dich thong thuong: dem so dong cho cong trinh dang mo (de
    // hien thi "N giao dich" tren UI), khong can luu chi tiet tung dong.
    if (current && (row[2] || row[3])) {
      current.soDong += 1;
    }
  }

  return { sheetName, byCongTrinh, order, tongCong };
}

module.exports = { parseMisaCongNoXlsx };
