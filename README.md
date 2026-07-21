# K&H Bank Tracker

Web app theo dõi ngân hàng cho **CÔNG TY TNHH DỊCH VỤ VÀ GIẢI TRÍ K&H**.

Tính năng:
- Đăng nhập (nhiều người trong team dùng chung, có đổi mật khẩu).
- Quản lý danh sách ngân hàng/tài khoản (tên, số TK, ngân hàng, số dư đầu kỳ).
- Nhập giao dịch từng cái, hoặc **dán hàng loạt** (copy từ Excel/sao kê rồi dán vào 1 ô).
- Trang tổng quan: tổng số dư tất cả ngân hàng, tổng thu/chi tháng này, biểu đồ thu chi 30 ngày, danh sách giao dịch gần đây.
- Lọc giao dịch theo ngân hàng / khoảng ngày, và **xuất ra file Excel**.

Công nghệ: Node.js + Express + EJS, dữ liệu lưu trong 1 file JSON (`data/store.json`) —
không cần cài đặt database riêng, không có bước build phức tạp, dễ deploy lên bất kỳ nơi nào
chạy được Node.

---

## 1. Chạy thử trên máy cá nhân

Cần cài [Node.js](https://nodejs.org) bản 18 trở lên.

```bash
cd bank-tracker
npm install
cp .env.example .env      # rồi mở file .env, đổi SESSION_SECRET và ADMIN_PASSWORD
npm start
```

Mở trình duyệt vào `http://localhost:3000`. Đăng nhập bằng tài khoản mặc định:

- Username: giá trị `ADMIN_USERNAME` trong `.env` (mặc định `admin`)
- Password: giá trị `ADMIN_PASSWORD` trong `.env` (mặc định `khadmin123`)

Tài khoản này chỉ được tạo **một lần duy nhất** lúc chạy lần đầu (khi chưa có `data/store.json`).
Sau khi đăng nhập, vào trang chủ để đổi mật khẩu ngay.

---

## 2. Đưa lên internet để cả team cùng dùng

App này là một Node.js server bình thường, có thể deploy ở bất kỳ dịch vụ hosting Node nào.
Hai lựa chọn dễ nhất và có gói miễn phí:

### Cách A — Render.com (khuyên dùng, có disk lưu dữ liệu lâu dài)

1. Đưa code này lên một GitHub repository (private là được).
2. Vào [render.com](https://render.com) → **New** → **Web Service** → chọn repo vừa tạo.
3. Cấu hình:
   - Build Command: `npm install`
   - Start Command: `npm start`
   - Instance Type: Free hoặc Starter đều được.
4. Vào tab **Environment** → thêm các biến:
   - `SESSION_SECRET` = một chuỗi ngẫu nhiên dài (không dùng chuỗi mặc định trong `.env.example`)
   - `ADMIN_USERNAME`, `ADMIN_PASSWORD` = tài khoản quản trị đầu tiên
   - `DATA_DIR` = `/var/data` (đường dẫn tới ổ đĩa bền vững, xem bước tiếp theo)
5. Vào tab **Disks** → thêm 1 Persistent Disk, mount vào `/var/data`. **Bước này quan trọng**:
   nếu không có disk, dữ liệu (`store.json`) sẽ mất mỗi khi Render khởi động lại server.
6. Deploy xong sẽ có 1 địa chỉ dạng `https://ten-app.onrender.com` — gửi link này cho cả team.

### Cách B — Railway.app

1. Đưa code lên GitHub, vào [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**.
2. Railway tự nhận `npm install` / `npm start`.
3. Vào tab **Variables** thêm `SESSION_SECRET`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`, và
   `DATA_DIR=/data`.
4. Vào tab **Volumes** → gắn 1 volume vào đường dẫn `/data` để dữ liệu không mất khi redeploy.
5. Railway cấp sẵn 1 domain public dạng `https://ten-app.up.railway.app`.

> Lưu ý chung cho cả 2 cách: nếu không gắn ổ đĩa bền vững (persistent disk/volume) vào đúng
> đường dẫn khai báo ở `DATA_DIR`, dữ liệu sẽ bị xoá mỗi lần server khởi động lại. Đây là bước
> hay bị bỏ sót nhất khi deploy app dùng file JSON/SQLite.

---

## 3. Sau khi deploy xong

1. Vào link online, đăng nhập bằng tài khoản admin đã cấu hình.
2. Đổi mật khẩu admin ngay trong trang **Tổng quan**.
3. Vào mục **Ngân hàng** → xoá "Ngan hang mau 1" (dữ liệu mẫu) → thêm các tài khoản ngân hàng
   thật của công ty (MTD40222, KVC40111, v.v.), điền số dư đầu kỳ đúng ngày bắt đầu theo dõi.
4. Vào mục **Giao dịch / Sao kê** → dán sao kê từ Excel để nhập dữ liệu.
5. Hiện tại chỉ có 1 tài khoản đăng nhập dùng chung cho cả team. Nếu cần mỗi người một tài khoản
   riêng, báo lại — có thể bổ sung trang quản lý user.

### Định dạng khi dán sao kê hàng loạt

Mỗi dòng gồm (cách nhau bằng Tab khi copy từ Excel, hoặc dấu phẩy):

```
Ngày (dd/mm/yyyy)    Diễn giải    Số tiền    Thu/Chi (không bắt buộc)
```

Ví dụ:
```
20/07/2026	Thu tien khach ABC	1.500.000	Thu
21/07/2026	Chi tra nha cung cap	850.000	Chi
```

Nếu bỏ trống cột Thu/Chi: số dương sẽ tự tính là Thu, số âm tự tính là Chi.

---

## 4. Cấu trúc thư mục

```
bank-tracker/
├── server.js              # điểm khởi động app
├── store.js                # lớp lưu trữ dữ liệu (file JSON)
├── middleware/auth.js      # kiểm tra đăng nhập
├── routes/                 # các route: auth, banks, transactions, dashboard
├── utils/parse.js          # xử lý dán sao kê + số tiền định dạng VN
├── views/                  # giao diện (EJS)
├── public/css/style.css    # giao diện
└── data/store.json         # dữ liệu (tự tạo khi chạy lần đầu, KHÔNG commit lên Git)
```

## 5. Bảo mật cần lưu ý trước khi dùng thật

- Đổi `SESSION_SECRET` trong `.env` sang chuỗi ngẫu nhiên, dài, giữ bí mật.
- Đổi mật khẩu admin mặc định ngay sau lần đăng nhập đầu tiên.
- Không commit file `.env` hoặc thư mục `data/` lên GitHub (đã cấu hình sẵn trong `.gitignore`).
- Vì mọi người trong team dùng chung 1 tài khoản, nên nhắc nhau không chia sẻ mật khẩu ra ngoài công ty.
