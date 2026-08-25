// Luyen, 2026-08-25: "cập nhật hóa đơn hàng ngày từ gg drive" -- list file
// tu folder Google Drive, dung chung OAuth token voi Gmail (store.gmail_oauth).
// Neu token cu chua co scope drive.readonly, Drive API tra 403 va message
// huong dan chi bam "Ket noi Gmail" lai (OAuth buildAuthUrl gio da request ca
// 2 scope cung luc trong gmailApi.js).

const gmailApi = require("./gmailApi");

// Folder IDs -- co the override qua env var de linh hoat hon
const DRIVE_HOADON_FOLDER_ID =
  process.env.DRIVE_HOADON_FOLDER_ID || "16Ybt8h75NcrvkcXDXb-eMTd0yRxtcNDS";
const DRIVE_PHI_CANG_FOLDER_ID =
  process.env.DRIVE_PHI_CANG_FOLDER_ID || "18r4LNki4LtdFj1ndk2x4BVQeS55QIuaX";

// List tat ca file (khong phai folder) trong 1 Drive folder, su dung
// Drive API v3. Tra ve mang [{id, name, mimeType, createdTime}].
// `store` duoc truyen vao de getValidAccessToken co the lam moi token va
// caller tu goi save(store) sau.
async function listFolderFiles(store, folderId, mimeFilter) {
  const accessToken = await gmailApi.getValidAccessToken(store);
  let q = `'${folderId}' in parents and trashed = false`;
  if (mimeFilter) q += ` and mimeType = '${mimeFilter}'`;
  const params = new URLSearchParams({
    q,
    fields: "files(id,name,mimeType,createdTime)",
    pageSize: "1000",
    orderBy: "name",
  });
  const resp = await fetch(
    "https://www.googleapis.com/drive/v3/files?" + params.toString(),
    { headers: { Authorization: "Bearer " + accessToken } }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    const msg = (err.error && err.error.message) ? err.error.message : resp.statusText;
    if (resp.status === 403 || resp.status === 401) {
      throw new Error(
        "Google Drive lỗi " + resp.status + ": " + msg +
        " — Cần bấm 'Kết nối Gmail' lại để cấp thêm quyền truy cập Drive."
      );
    }
    throw new Error("Google Drive lỗi " + resp.status + ": " + msg);
  }
  const data = await resp.json();
  return data.files || [];
}

// Chuyen anh (JPG/PNG buffer) thanh PDF 1 trang dung pdf-lib.
// Tra ve Buffer chua PDF.
async function imageToPdf(imageBuffer, mimeType) {
  const { PDFDocument } = require("pdf-lib");
  const pdfDoc = await PDFDocument.create();
  const mt = (mimeType || "").toLowerCase();
  let image;
  if (mt === "image/jpeg" || mt === "image/jpg") {
    image = await pdfDoc.embedJpg(imageBuffer);
  } else if (mt === "image/png") {
    image = await pdfDoc.embedPng(imageBuffer);
  } else {
    throw new Error("Chỉ hỗ trợ ảnh JPG hoặc PNG (mimeType: " + mimeType + ").");
  }
  const { width, height } = image.scale(1);
  const page = pdfDoc.addPage([width, height]);
  page.drawImage(image, { x: 0, y: 0, width, height });
  return Buffer.from(await pdfDoc.save());
}

// Upload 1 file len Google Drive (multipart upload), tra ve {id, name}.
// `fileBuffer` la Buffer, `mimeType` la MIME cua file can upload (vd "application/pdf").
// Can scope drive.file trong OAuth -- neu token cu chua co scope nay, bao loi 403 huong
// dan bam "Ket noi Gmail" lai.
async function uploadFileToDrive(store, folderId, fileName, fileBuffer, mimeType) {
  const accessToken = await gmailApi.getValidAccessToken(store);
  const boundary = "kh_bnd_" + Date.now();
  const metadata = JSON.stringify({ name: fileName, parents: [folderId] });
  const body = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${metadata}\r\n`),
    Buffer.from(`--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`),
    fileBuffer,
    Buffer.from(`\r\n--${boundary}--`),
  ]);
  const resp = await fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": `multipart/related; boundary=${boundary}`,
      },
      body,
    }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    const msg = (err.error && err.error.message) ? err.error.message : resp.statusText;
    if (resp.status === 403 || resp.status === 401) {
      throw new Error(
        "Google Drive upload lỗi " + resp.status + ": " + msg +
        " — Cần bấm 'Kết nối Gmail' lại để cấp quyền upload Drive."
      );
    }
    throw new Error("Google Drive upload lỗi " + resp.status + ": " + msg);
  }
  return await resp.json(); // { id, name }
}

// Doc PDF bien lai phi cang, trich xuat cac truong chinh.
// Tra ve { ngay, soHoaDon, loaiPhi, soTien, rawText }.
// Dung pdf-parse de lay text, sau do dung regex tim cac gia tri.
async function extractPhiCangInfo(pdfBuffer) {
  const pdfParse = require("pdf-parse");
  const data = await pdfParse(pdfBuffer);
  const text = data.text || "";
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const fullText = lines.join(" ");

  // --- So hoa don: "Số 0000192", "So: 0000192", "HD: 0000192", "No: 0000192" ---
  let soHoaDon = "";
  const hdMatch = fullText.match(/[Ss][ốo]\s*[Hh][oóđ][aà]\s*[Đđ][oơ][nà]?\s*[:\-]?\s*(\d+)/i)
    || fullText.match(/[Ss][ốo]\s*[:\-]?\s*(\d+)/i)
    || fullText.match(/[Hh][Dd]\s*[:\-]?\s*(\d+)/i)
    || fullText.match(/[Nn]o\.?\s*(\d{4,})/i);
  if (hdMatch) soHoaDon = hdMatch[1].replace(/^0+/, "").padStart(hdMatch[1].length, "0");

  // --- Ngay: dd/mm/yyyy hoac yyyy-mm-dd ---
  let ngay = "";
  const dateMatch = fullText.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dateMatch) {
    const d = dateMatch[1].padStart(2, "0");
    const m = dateMatch[2].padStart(2, "0");
    const y = dateMatch[3];
    ngay = `${y}-${m}-${d}`; // YYYY-MM-DD cho input[type=date]
  }

  // --- So tien: tim gan "Tong cong" / "Total" / "Thanh tien" truoc, fallback largest ---
  let soTien = "";
  // Uu tien: lay so tien sau "Tong cong", "Total", "Thanh tien", "So tien"
  const tongCongPatterns = [
    /[Tt][oô]?ng\s*c[oộ]ng[^0-9]{0,30}([\d.,]+)/i,
    /[Tt]otal[^0-9]{0,20}([\d.,]+)/i,
    /[Tt]h[àa]nh\s*ti[eề]n[^0-9]{0,20}([\d.,]+)/i,
    /[Ss][oố]\s*ti[eề]n[^0-9]{0,20}([\d.,]+)/i,
  ];
  let foundByKeyword = false;
  for (const pat of tongCongPatterns) {
    const m = fullText.match(pat);
    if (m) {
      const raw = parseInt(m[1].replace(/[.,]/g, ""));
      if (raw > 0 && raw < 100000000) {
        soTien = raw.toLocaleString("de-DE");
        foundByKeyword = true;
        break;
      }
    }
  }
  // Fallback: lay so co dau cham phan ngan (vd 42.100, 63.225) - chi nhan 3-8 chu so
  if (!foundByKeyword) {
    const dotThousands = [...fullText.matchAll(/\b(\d{1,3}(?:\.\d{3})+)\b/g)];
    let best = 0, bestStr = "";
    dotThousands.forEach(m => {
      const num = parseInt(m[1].replace(/\./g, ""));
      if (num > best && num < 10000000) { best = num; bestStr = m[1]; }
    });
    if (bestStr) soTien = bestStr;
  }

  // --- Loai phi: dong dau tien co chu "phi" hoac "le phi" hoac toan bo text ngan ---
  let loaiPhi = "";
  for (const line of lines) {
    if (/[Pp]h[ií]|[Ll][eệ]\s*[Pp]h[ií]/i.test(line) && line.length > 8 && line.length < 200) {
      loaiPhi = line.replace(/\s+/g, " ").trim();
      break;
    }
  }
  // Fallback: lay dong thu 2-4 neu chua co
  if (!loaiPhi && lines.length > 2) {
    loaiPhi = lines.slice(1, 4).join(" ").trim().substring(0, 150);
  }

  return { ngay, soHoaDon, loaiPhi, soTien, rawText: text.substring(0, 500) };
}

module.exports = { listFolderFiles, imageToPdf, uploadFileToDrive, extractPhiCangInfo, DRIVE_HOADON_FOLDER_ID, DRIVE_PHI_CANG_FOLDER_ID };
