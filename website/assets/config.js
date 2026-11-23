/* ==========================================================================
   CẤU HÌNH CHUNG — chỉ cần sửa file này (dùng cho cả 2 trang)
   ========================================================================== */
window.CONFIG = {
  BRAND: "OUTSOURCE STUDIO",

  // URL Web App của Google Apps Script (xem HUONG-DAN-CAI-DAT.md). Để trống = chế độ demo.
  SHEET_ENDPOINT: "https://script.google.com/macros/s/AKfycbyN9sOj7h3aUKdZ4QpQwmMeThbZpA1uFZ6wz4fuCzhLiULBfLe33bZleTmKm7XP66Ao/exec",

  // Link Google Form khảo sát (tạo từ menu 🎬 Studio trong Google Sheet). Để trống = ẩn.
  SURVEY_URL: "",

  CONTACT: {
    phone:    { label: "Hotline",  value: "0399 924 348",                url: "tel:0399924348" },
    zalo:     { label: "Zalo",     value: "0399 924 348",                url: "https://zalo.me/0399924348" },
    facebook: { label: "Facebook", value: "Hoàng Phúc Thái",             url: "https://www.facebook.com/share/1AAid3RHud/?mibextid=wwXIfr" },
    email:    { label: "Email",    value: "thaihoang1682005@gmail.com",  url: "mailto:thaihoang1682005@gmail.com" },
  },

  DRIVE_PORTFOLIO: "https://drive.google.com/drive/folders/1nP1izNrPhnTihH3itx07e0NKO4eFE8aB",

  /* VIDEO MẪU — lấy trực tiếp từ Google Drive (file phải để "Bất kỳ ai có đường liên kết")
     cat:   "su-kien" | "quang-cao" | "ngan" | "phim"
     type:  "drive" | "youtube" | "tiktok" | "mp4"
     id/url: Drive dùng id file; loại khác dùng url
     ratio: "16:9" | "9:16" | "3:4"                                                   */
  VIDEO_CATS: { "su-kien": "Sự kiện", "quang-cao": "Quảng cáo", "ngan": "Video ngắn", "phim": "Phim & hậu trường" },
  VIDEOS: [
    { title: "Countdown Trấn Biên 2026",          cat: "su-kien",   type: "drive", id: "1YAfso6ISMmf-L6qT82i8vlK6JTbImwFK", ratio: "16:9", desc: "Highlight sự kiện" },
    { title: "Giải Pickleball Khu vực 7",          cat: "su-kien",   type: "drive", id: "11OXPTEHhC2K_C8e5pyCDyRIUamMy8bCL", ratio: "16:9", desc: "Recap giải đấu" },
    { title: "Diễn đàn doanh nghiệp Slovakia – VN", cat: "su-kien",   type: "drive", id: "18kBfqvkJNHW3zlhhQE-kjfrh7CzByrKJ", ratio: "16:9", desc: "Sự kiện doanh nghiệp · có ứng dụng AI" },
    { title: "Lễ công bố hợp tác số",              cat: "su-kien",   type: "drive", id: "1-Lapepi47B0Lla0HzGNF22giAyULmSzV", ratio: "16:9", desc: "Highlight · có ứng dụng AI" },
    { title: "Hotdeal Ibis Saigon South",          cat: "quang-cao", type: "drive", id: "19YywgKnF6ypl_0SdRpfowbjtOmr3AR5H", ratio: "9:16", desc: "Review khách sạn" },
    { title: "Quảng cáo thẩm mỹ – Bling Kim",      cat: "quang-cao", type: "drive", id: "1ohtuZVTzv18zxJ3sxGVLW_zl-tUO-ghV", ratio: "9:16", desc: "Video quảng cáo dọc" },
    { title: "Quảng cáo thẩm mỹ – Hành trình",     cat: "quang-cao", type: "drive", id: "1pbvsLt8p7ojT3Q9rG6V9MjEgmxoSUPWq", ratio: "3:4",  desc: "Ads cho fanpage" },
    { title: "AMY – Nội dung màn hình LED",        cat: "quang-cao", type: "drive", id: "1r57GU5r8Wmhvb4IJ65zylSP-Bx05Rsba", ratio: "16:9", desc: "Trình chiếu màn hình LED" },
    { title: "Kiến thức sức khỏe",                 cat: "ngan",      type: "drive", id: "1kUOqWgqvCwuI4bFH0m1KkYesu7mO4qPh", ratio: "9:16", desc: "Video kiến thức cho TikTok" },
    { title: "Phim ngắn",                          cat: "phim",      type: "drive", id: "1mBYLGpC6B6tjh9vnNJfk9KcDEpXu18mE", ratio: "9:16", desc: "Quay & dựng phim" },
    { title: "Hậu trường sản xuất",                cat: "phim",      type: "drive", id: "1XFtBOKqNXXA95nYjGAiaqSUG8nt__Fys", ratio: "9:16", desc: "Behind the scenes" },
  ],

  /* GÓI DỊCH VỤ — mã gói phải khớp sheet "Dịch vụ & Gói" */
  SERVICES: {
    "quay-edit": { name: "Quay + Edit", short: "Studio quay tận nơi và dựng hoàn chỉnh. Phù hợp khi bạn chưa có tư liệu." },
    "edit":      { name: "Chỉ Edit",    short: "Bạn gửi footage (quay bằng điện thoại cũng được), studio dựng thành video hoàn chỉnh." },
  },
  PACKAGES: [
    { code:"QE1", type:"quay-edit", name:"Khởi đầu",   price:1500000, for:"Shop, quán nhỏ làm video đầu tiên",
      videos:"1 video", length:"30–60s", time:"3–4 ngày", revisions:"2 lần",
      items:["Quay 2 giờ, 1 địa điểm","1 video dọc TikTok/Reels","Nhạc, phụ đề, tiêu đề"],
      detail:["Tư vấn ý tưởng trước khi quay","Chỉnh màu cơ bản","File MP4 Full HD"] },
    { code:"QE2", type:"quay-edit", name:"Kinh doanh", price:3000000, hot:true, for:"Doanh nghiệp nhỏ cần nội dung đều",
      videos:"2 video", length:"60–90s", time:"5–7 ngày", revisions:"3 lần",
      items:["Quay nửa ngày, 2 địa điểm","1 video dọc + 1 ngang","Kịch bản do studio viết"],
      detail:["Storyboard trước khi quay","Text động, hiệu ứng, chỉnh màu","Thêm 2 bản 15s chạy ads","Kèm thumbnail"] },
    { code:"QE3", type:"quay-edit", name:"Toàn diện",  price:5500000, for:"Khai trương, ra mắt, video thương hiệu",
      videos:"3 video", length:"1–3 phút", time:"7–10 ngày", revisions:"3 lần",
      items:["Quay 1 ngày + flycam","Kịch bản + voice-over","1 video chính + 2 bản ngắn"],
      detail:["Khảo sát địa điểm trước","Color grading, motion graphics","Sound design","Bàn giao file gốc khi cần"] },
    { code:"ED1", type:"edit", name:"Edit nhanh",      price:200000, for:"Video ngắn cần đăng ngay",
      videos:"1 video", length:"≤ 60s", time:"24–48 giờ", revisions:"1 lần",
      items:["Cắt ghép theo nhịp TikTok","Nhạc, phụ đề","Tiêu đề + lời kêu gọi cuối video"],
      detail:["Chọn cảnh đẹp từ footage","Cân bằng âm thanh","File MP4 Full HD dọc"] },
    { code:"ED2", type:"edit", name:"Edit chuẩn",      price:450000, hot:true, for:"Video bán hàng, review, giới thiệu",
      videos:"1 video", length:"≤ 3 phút", time:"2–3 ngày", revisions:"2 lần",
      items:["Hiệu ứng theo beat nhạc","Chỉnh màu đồng bộ","Text động, khử ồn"],
      detail:["Tư vấn bố cục câu chuyện","Thêm 1 bản 15s","Kèm thumbnail"] },
    { code:"ED3", type:"edit", name:"Edit chuyên sâu", price:900000, for:"Video thương hiệu nhiều hiệu ứng",
      videos:"1 dài / 3 ngắn", length:"≤ 5 phút", time:"3–5 ngày", revisions:"3 lần",
      items:["Motion graphics, logo intro","Color grading theo thương hiệu","Sound design"],
      detail:["Moodboard trước khi dựng","Xuất 9:16, 1:1, 16:9","Phụ đề song ngữ (tùy chọn)"] },
  ],
  ADDONS: [
    { name:"Flycam",          price:"+800k" },
    { name:"Voice-over",      price:"+300k" },
    { name:"Viết kịch bản",   price:"+300k" },
    { name:"Phụ đề song ngữ", price:"+150k" },
    { name:"Giao gấp 24h",    price:"+30%" },
  ],

  /* KỸ NĂNG CỘNG TÁC VIÊN */
  SKILLS: [
    "Edit video ngắn (TikTok/Reels)", "Edit video dài / YouTube", "Quay sự kiện", "Quay sản phẩm / quảng cáo",
    "Flycam", "Chụp ảnh", "Motion graphics", "Color grading", "Âm thanh / sound design",
    "Thiết kế thumbnail, banner", "Viết kịch bản / content", "Lồng tiếng / voice-over",
    "Phụ đề / dịch thuật", "Diễn xuất / KOC / mẫu", "Ánh sáng / hỗ trợ set quay",
  ],
};
