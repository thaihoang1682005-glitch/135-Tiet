# OUTSOURCE STUDIO — Hướng dẫn cài đặt

## Cấu trúc thư mục

```
135 Tiết Phúc Thái/
├─ website/                     ← đưa thư mục này lên GitHub Pages / Netlify
│  ├─ index.html                ← trang KHÁCH HÀNG (thuê làm video)
│  ├─ cong-tac-vien.html        ← trang CỘNG TÁC VIÊN (nhận việc freelance)
│  └─ assets/
│     ├─ config.js              ← ⭐ MỌI THỨ CẦN SỬA NẰM Ở ĐÂY (liên hệ, video, gói, link Sheet)
│     ├─ style.css, common.js   ← giao diện + code dùng chung
│     └─ logo.svg, logo-full.svg, favicon.*, apple-touch-icon.png, logo-512.png
├─ apps-script/Code.gs          ← code dán vào Google Sheet
├─ OUTSOURCE STUDIO - Quản lý.xlsx   ← file quản lý (mở bằng Google Sheets)
└─ HUONG-DAN-CAI-DAT.md
```

Luồng dữ liệu: **Khách hoặc CTV điền form trên web → Apps Script → ghi thẳng vào Google Sheet → gửi email báo cho Thái.**

---

## Bước 1. Tạo Google Sheet quản lý (đăng nhập bằng Gmail của Thái)

> ⚠️ **Quan trọng nhất:** file `.xlsx` mở trên Google Sheets (trên thanh tên file có nhãn **.XLSX**) **không chạy được Apps Script**, nên không thể nhận dữ liệu realtime. Phải chuyển sang Google Sheet thật như bước 2 dưới đây.

1. Mở `OUTSOURCE STUDIO - Quản lý.xlsx` trên Google Drive.
2. Chọn **Tệp → Lưu dưới dạng Google Trang tính**. Từ giờ chỉ dùng **bản mới** này (bản không còn nhãn .XLSX).
3. Trong bản mới: **Tiện ích mở rộng → Apps Script**. Xóa code mẫu, dán toàn bộ `apps-script/Code.gs`, rồi bấm **Lưu**.
4. Quay lại Sheet và **tải lại trang (F5)**. Thanh menu sẽ có thêm **🎬 Studio**.
5. Chọn **🎬 Studio → ⚙️ Khởi tạo / sửa chữa hệ thống**. Google sẽ hỏi cấp quyền: chọn **Nâng cao → Đi tới… → Cho phép**.
   - Script sẽ dựng bảng **Tổng quan** đầy đủ (công thức realtime + 4 biểu đồ), thêm công thức tự tính, màu cảnh báo và bộ lọc.
   - Nếu file còn các sheet của phiên bản cũ (Khách hàng, Dịch vụ & Gói, Chi phí khác…), script sẽ **hỏi trước khi xóa**.
   - Chạy lại bao nhiêu lần cũng được, **dữ liệu không bị xóa**.

> Muốn tạo Sheet trắng thay vì dùng file `.xlsx`: vào <https://sheets.new> rồi làm từ bước 3. Script sẽ tự tạo đủ các sheet.

### Các sheet sau khi cài

| Sheet | Ai điền | Nội dung |
|---|---|---|
| **Tổng quan** | Tự tính | Số yêu cầu hôm nay / 7 ngày / tháng này; đơn theo trạng thái, gói, loại dịch vụ, nguồn khách, đối tượng khách; doanh thu, công nợ, lợi nhuận; số khách và khách quay lại; CTV theo kỹ năng; biểu đồ 14 ngày và 12 tháng; cảnh báo trễ hạn; 5 đơn và 5 CTV mới nhất; kết quả khảo sát |
| **Đơn hàng** | Website + Thái | Khách gửi form → dòng mới xuất hiện ngay (tô xanh trong 24h). Thái cập nhật trạng thái, giá, CTV, hạn. Cột "Khách cũ?" tự nhận biết khách đặt lần 2, 3 |
| **Cộng tác viên** | Website + Thái | Người xin việc gửi form → dòng mới xuất hiện ngay. Tự tính số job, tỉ lệ đúng hạn, công nợ |
| **Khảo sát** | Google Form | Có sau khi chạy "Tạo form khảo sát"; kết quả tự hiện trong Tổng quan |

**Realtime hoạt động thế nào:** website gửi form → Web App ghi ngay 1 dòng vào sheet (vài giây) → mọi công thức ở Tổng quan tự tính lại → Thái nhận email báo. Không cần bấm làm mới.

## Bước 2. Kết nối website với Sheet

1. Trong Apps Script, chọn **Triển khai → Tùy chọn triển khai mới**. Bấm ⚙️ → **Ứng dụng web**, thiết lập:
   - Thực thi bằng: **Tôi**
   - Người có quyền truy cập: **Bất kỳ ai**
2. Sao chép **URL Web App**, dán vào `website/assets/config.js`:
   ```js
   SHEET_ENDPOINT: "https://script.google.com/macros/s/…/exec",
   ```
3. Mở URL đó trên trình duyệt. Thấy dòng "✅ Web App … đang hoạt động" là đã kết nối thành công.

> Mỗi lần sửa `Code.gs`: vào **Triển khai → Quản lý triển khai → ✏️ → Phiên bản mới**. Làm theo cách này thì URL giữ nguyên.

## Bước 3. Form khảo sát

1. Chọn **🎬 Studio → 📝 Tạo form khảo sát**. Script sẽ tạo Google Form và sheet **Khảo sát**. Điểm trung bình và xếp hạng điểm yếu tự hiện trong mục 📝 của **Tổng quan**.
2. Dán link form vào `SURVEY_URL` trong `config.js`. Chân trang web sẽ hiện mục "Góp ý / khảo sát".
3. **Gửi link form cho người thật trả lời.** Báo cáo chỉ có giá trị khi dữ liệu là thật.

## Bước 4. Video mẫu

- Video đang được **nhúng trực tiếp từ Google Drive** (thư mục *Profile Hoàng Phúc Thái*). Mình không tải về vì tổng dung lượng khoảng 1,9 GB, vượt giới hạn 100 MB/file của GitHub.
- Các file trên Drive **phải để chế độ "Bất kỳ ai có đường liên kết"** thì khách mới xem được.
- Thêm video mới: mở video trên Drive → **Chia sẻ → Sao chép đường liên kết** → lấy đoạn ID nằm giữa `/d/` và `/view` → thêm 1 dòng vào `VIDEOS` trong `config.js`.

## Bước 5. Đưa web lên mạng

- **GitHub Pages:** tạo repo, đẩy nội dung thư mục `website/` lên, rồi vào **Settings → Pages → Branch: main**.
- **Netlify Drop:** kéo thả thư mục `website` vào <https://app.netlify.com/drop>.
- Đo nguồn khách: khi đăng lên từng kênh, thêm `?src=tiktok` hoặc `?src=facebook` vào cuối link. Bảng Tổng quan sẽ thống kê theo nguồn.

---

## Vận hành hằng ngày

| Khi | Làm gì trong sheet **Đơn hàng** |
|---|---|
| Có email "🆕 Đơn mới" | Gọi hoặc nhắn Zalo khách → **Trạng thái**: *Đã liên hệ* / *Đã báo giá* |
| Khách đặt cọc | Điền **Giá chốt**, **Đã thu** → *Đã cọc* |
| Giao việc cho CTV | Chọn **CTV phụ trách**, **Thù lao CTV**, **Hạn CTV nộp** → *Đang sản xuất* |
| CTV nộp bài | Điền **Ngày CTV nộp** (cột *CTV đúng hạn?* tự chấm) |
| Gửi khách duyệt / sửa | *Chờ khách duyệt* / *Đang chỉnh sửa* |
| Bàn giao xong | Chọn *Hoàn thành* (tự điền **Ngày bàn giao**), cập nhật **Đã thu** |
| Trả tiền CTV | Điền **Đã trả CTV** (sheet Cộng tác viên tự trừ công nợ) |

- Ô hạn chuyển **vàng** khi còn ≤ 2 ngày, chuyển **đỏ** khi đã trễ.
- CTV mới đăng ký: giao bài test, đạt thì đổi trạng thái sang **Đang hoạt động**.
- Cột có tiêu đề **xanh đậm** là cột tự tính, đừng gõ vào.
