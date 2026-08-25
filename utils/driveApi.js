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

module.exports = { listFolderFiles, DRIVE_HOADON_FOLDER_ID, DRIVE_PHI_CANG_FOLDER_ID };
