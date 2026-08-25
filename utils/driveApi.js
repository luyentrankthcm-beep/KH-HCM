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

module.exports = { listFolderFiles, imageToPdf, uploadFileToDrive, DRIVE_HOADON_FOLDER_ID, DRIVE_PHI_CANG_FOLDER_ID };
