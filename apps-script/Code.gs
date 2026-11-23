/**
 * ============================================================================
 *  OUTSOURCE STUDIO — SHEET QUẢN LÝ (Google Apps Script)
 * ----------------------------------------------------------------------------
 *  Chỉ 3 sheet chính + 1 sheet khảo sát:
 *    • Tổng quan      — thống kê REALTIME bằng công thức (tự tính lại khi có dữ liệu mới)
 *    • Đơn hàng       — khách gửi form trên web → hiện ngay 1 dòng mới
 *    • Cộng tác viên  — người xin việc gửi form → hiện ngay 1 dòng mới
 *    • Khảo sát       — Google Form tự đổ câu trả lời về (tạo bằng menu)
 *  Chạy "Khởi tạo" bao nhiêu lần cũng được: KHÔNG xóa dữ liệu đang có.
 *  ⚠️ Chỉ chạy được trên Google Sheet THẬT (không phải file .xlsx đang mở bằng Sheets).
 * ============================================================================
 */

/***** CẤU HÌNH *****/
const CFG = {
  BRAND: 'OUTSOURCE STUDIO',
  NOTIFY_EMAIL: 'thaihoang1682005@gmail.com', // nhận email khi có đơn mới / CTV mới
  SEND_NOTIFY: true,
  TZ: 'Asia/Ho_Chi_Minh',
  MAXR: 2000
};

const SH = { TQ: 'Tổng quan', DH: 'Đơn hàng', NS: 'Cộng tác viên', KS: 'Khảo sát' };
const OBSOLETE = ['Khách hàng', 'Dịch vụ & Gói', 'Chi phí khác', 'Hướng dẫn', 'Báo cáo khảo sát', 'Liên hệ'];
const RENAMED = { 'Nhân sự CTV': SH.NS, 'Khảo sát (phản hồi)': SH.KS };

const STATUS_DH = ['Mới', 'Đã liên hệ', 'Đã báo giá', 'Đã cọc', 'Đang sản xuất', 'Chờ khách duyệt', 'Đang chỉnh sửa', 'Hoàn thành', 'Hủy'];
const STATUS_ACTIVE = ['Đã cọc', 'Đang sản xuất', 'Chờ khách duyệt', 'Đang chỉnh sửa'];
const STATUS_NS = ['Mới đăng ký', 'Đang làm bài test', 'Đang hoạt động', 'Tạm nghỉ', 'Ngưng hợp tác'];

const H = {
  DH: ['Mã đơn', 'Thời gian', 'Họ tên', 'SĐT/Zalo', 'Email', 'Đối tượng', 'Lĩnh vực', 'Loại dịch vụ', 'Gói', 'Dịch vụ thêm',
       'Nền tảng', 'Địa điểm quay', 'Hạn giao khách', 'Mô tả yêu cầu', 'Link tư liệu', 'Kênh liên hệ', 'Nguồn',
       'Trạng thái', 'Khách cũ?', 'Giá chốt', 'Đã thu', 'Còn phải thu',
       'CTV phụ trách', 'Thù lao CTV', 'Hạn CTV nộp', 'Ngày CTV nộp', 'CTV đúng hạn?', 'Đã trả CTV',
       'Ngày bàn giao', 'Giao khách đúng hạn?', 'Lợi nhuận gộp', 'Ghi chú'],
  NS: ['Mã CTV', 'Thời gian ĐK', 'Họ tên', 'SĐT/Zalo', 'Email', 'Khu vực', 'Kỹ năng', 'Kinh nghiệm', 'Phần mềm', 'Thiết bị',
       'Portfolio', 'Thời gian rảnh', 'Mức giá mong muốn', 'Ghi chú của CTV', 'Trạng thái', 'Đánh giá (1-5)',
       'Số job', 'Job đúng hạn', 'Tỉ lệ đúng hạn', 'Tổng thù lao', 'Đã trả', 'Còn nợ', 'Ghi chú']
};
const LAST_INPUT = { DH: 'Trạng thái', NS: 'Trạng thái' }; // web chỉ ghi tới cột này

/* Danh mục để thống kê — khớp website/assets/config.js */
const PACKAGES = [['QE1', 'Quay + Edit · Khởi đầu'], ['QE2', 'Quay + Edit · Kinh doanh'], ['QE3', 'Quay + Edit · Toàn diện'],
                  ['ED1', 'Edit nhanh'], ['ED2', 'Edit chuẩn'], ['ED3', 'Edit chuyên sâu']];
const SERVICES = ['Quay + Edit', 'Chỉ Edit'];
const DOITUONG = ['Cá nhân kinh doanh', 'Hộ kinh doanh', 'Doanh nghiệp nhỏ', 'Người bán hàng online', 'Cá nhân xây dựng thương hiệu', 'Sinh viên', 'Khác'];
const SKILLS = ['Edit video ngắn (TikTok/Reels)', 'Edit video dài / YouTube', 'Quay sự kiện', 'Quay sản phẩm / quảng cáo',
  'Flycam', 'Chụp ảnh', 'Motion graphics', 'Color grading', 'Âm thanh / sound design', 'Thiết kế thumbnail, banner',
  'Viết kịch bản / content', 'Lồng tiếng / voice-over', 'Phụ đề / dịch thuật', 'Diễn xuất / KOC / mẫu', 'Ánh sáng / hỗ trợ set quay'];

const C = { navy: '#0B1020', soft: '#16203F', ink: '#EEF2FF', blue: '#2F6BFF', teal: '#00D1B2', zebra: '#F5F7FF', calc: '#2B3A6B' };
const MONEY = '#,##0" đ"';
const OK = '✔ Đúng hạn', LATE = '✘ Trễ hạn';

/* ============================================================================
 *  MENU
 * ========================================================================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎬 Studio')
    .addItem('⚙️ 1. Khởi tạo / sửa chữa hệ thống', 'khoiTaoHeThong')
    .addItem('📝 2. Tạo form khảo sát', 'taoFormKhaoSat')
    .addSeparator()
    .addItem('🔗 Xem các link quan trọng', 'xemLink')
    .addSeparator()
    .addItem('🧪 Thêm dữ liệu DEMO', 'themDuLieuDemo')
    .addItem('🧹 Xóa dữ liệu DEMO', 'xoaDuLieuDemo')
    .addToUi();
}

/* ============================================================================
 *  KHỞI TẠO / SỬA CHỮA
 * ========================================================================== */
function khoiTaoHeThong() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  // en_GB: công thức dùng dấu phẩy, ngày dd/mm/yyyy. Không dùng vi_VN vì công thức sẽ lỗi #ERROR!
  ss.setSpreadsheetLocale('en_GB');
  SpreadsheetApp.flush();
  ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.MINUTE); // tính lại khi có thay đổi + mỗi phút

  // Đổi tên sheet của phiên bản cũ
  Object.keys(RENAMED).forEach(o => { const s = ss.getSheetByName(o); if (s && !ss.getSheetByName(RENAMED[o])) s.setName(RENAMED[o]); });

  // Bước 1: đảm bảo cấu trúc 2 sheet dữ liệu (phải có cả 2 trước khi đặt công thức tham chiếu chéo)
  const dh = ensureStructure_(ss, 'DH'), ns = ensureStructure_(ss, 'NS');
  // Bước 2: công thức, định dạng, danh sách chọn
  setupNS_(ns); setupDH_(dh);

  // Bước 3: dọn các sheet không còn dùng
  const old = OBSOLETE.map(n => ss.getSheetByName(n)).filter(Boolean)
    .concat(ss.getSheets().filter(s => /^(Sheet|Trang tính)\s?\d+$/.test(s.getName()) && s.getLastRow() === 0));
  if (old.length) {
    const r = ui.alert('Dọn sheet không cần thiết', 'Xóa các sheet sau?\n\n• ' + old.map(s => s.getName()).join('\n• ') +
      '\n\n(Thống kê của chúng đã được gộp vào "Tổng quan")', ui.ButtonSet.YES_NO);
    if (r === ui.Button.YES) old.forEach(s => { if (ss.getSheets().length > 1) ss.deleteSheet(s); });
  }

  taoTongQuan_(ss);
  [SH.TQ, SH.DH, SH.NS, SH.KS].forEach((n, i) => { const s = ss.getSheetByName(n); if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); } });
  ss.setActiveSheet(ss.getSheetByName(SH.TQ));
  ui.alert('✅ Đã sẵn sàng', 'Tiếp theo:\n1. Triển khai → Tùy chọn triển khai mới → Ứng dụng web (Bất kỳ ai).\n' +
    '2. Dán URL vào SHEET_ENDPOINT trong website/assets/config.js.\n3. 🎬 Studio → Tạo form khảo sát.', ui.ButtonSet.OK);
}

/** Tạo sheet nếu chưa có. Nếu cấu trúc cột khác phiên bản này: sheet trống → làm lại; có dữ liệu → giữ lại dưới tên "(cũ)" */
function ensureStructure_(ss, key) {
  let sh = ss.getSheetByName(SH[key]);
  const n = H[key].length;
  if (sh) {
    const head = sh.getRange(1, 1, 1, Math.max(n, sh.getLastColumn() || 1)).getDisplayValues()[0];
    const same = H[key].every((h, i) => head[i] === h);
    if (!same) {
      if (colValues_(sh, 1).length === 0) { ss.deleteSheet(sh); sh = null; }
      else { sh.setName(SH[key] + ' (cũ ' + Utilities.formatDate(new Date(), CFG.TZ, 'ddMM-HHmm') + ')'); sh = null; }
    }
  }
  if (!sh) {
    sh = ss.insertSheet(SH[key]);
    if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);
  }
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  if (sh.getMaxRows() < CFG.MAXR) sh.insertRowsAfter(sh.getMaxRows(), CFG.MAXR - sh.getMaxRows());
  sh.getRange(1, 1, 1, n).setValues([H[key]]);
  sh.getRange(1, 1, 1, n).setBackground(C.soft).setFontColor(C.ink).setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
  sh.setRowHeight(1, 40); sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.clearConditionalFormatRules();
  if (sh.getFilter()) sh.getFilter().remove();
  return sh;
}

/* --------------------------- Đơn hàng --------------------------- */
function setupDH_(sh) {
  const R = n => rngLocal_('DH', n);
  sh.setTabColor(C.blue);
  const sdt = R('SĐT/Zalo'), lan = `COUNTIFS(${sdt},${sdt},ROW(${sdt}),"<="&ROW(${sdt}))`;
  setArr_(sh, 'DH', 'Khách cũ?',
    `ARRAYFORMULA(IF(${sdt}="","",IF(${lan}>1,"🔁 Lần "&${lan},"Mới")))`);
  setArr_(sh, 'DH', 'Còn phải thu',
    `ARRAYFORMULA(IF(${R('Giá chốt')}="","",${R('Giá chốt')}-IF(${R('Đã thu')}="",0,${R('Đã thu')})))`);
  setArr_(sh, 'DH', 'CTV đúng hạn?',
    `ARRAYFORMULA(IF((${R('Hạn CTV nộp')}="")+(${R('Ngày CTV nộp')}=""),"",IF(${R('Ngày CTV nộp')}<=${R('Hạn CTV nộp')},"${OK}","${LATE}")))`);
  setArr_(sh, 'DH', 'Giao khách đúng hạn?',
    `ARRAYFORMULA(IF((${R('Hạn giao khách')}="")+(${R('Ngày bàn giao')}=""),"",IF(${R('Ngày bàn giao')}<=${R('Hạn giao khách')},"${OK}","${LATE}")))`);
  setArr_(sh, 'DH', 'Lợi nhuận gộp',
    `ARRAYFORMULA(IF(${R('Giá chốt')}="","",${R('Giá chốt')}-IF(${R('Thù lao CTV')}="",0,${R('Thù lao CTV')})))`);

  fmtCol_(sh, 'DH', 'Thời gian', 'dd/MM/yyyy HH:mm');
  ['Hạn giao khách', 'Hạn CTV nộp', 'Ngày CTV nộp', 'Ngày bàn giao'].forEach(n => fmtCol_(sh, 'DH', n, 'dd/MM/yyyy'));
  ['Giá chốt', 'Đã thu', 'Còn phải thu', 'Thù lao CTV', 'Đã trả CTV', 'Lợi nhuận gộp'].forEach(n => fmtCol_(sh, 'DH', n, MONEY));
  fmtCol_(sh, 'DH', 'SĐT/Zalo', '@');

  listVal_(sh, 'DH', 'Trạng thái', STATUS_DH);
  const ctv = col_('DH', 'CTV phụ trách');
  sh.getRange(`${ctv}2:${ctv}`).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActive().getRange(`'${SH.NS}'!A2:A`), true).setAllowInvalid(false).build());
  dateVal_(sh, 'DH', ['Hạn giao khách', 'Hạn CTV nộp', 'Ngày CTV nộp', 'Ngày bàn giao']);

  statusColors_(sh, 'DH', 'Trạng thái', {
    'Mới': ['#FFF4D6', '#8A5A00'], 'Đã liên hệ': ['#E6F0FF', '#1D4F9F'], 'Đã báo giá': ['#E6F0FF', '#1D4F9F'],
    'Đã cọc': ['#EDE4FF', '#5B2CB8'], 'Đang sản xuất': ['#DDE8FF', '#1A3FA8'], 'Chờ khách duyệt': ['#FFE9D6', '#A04A00'],
    'Đang chỉnh sửa': ['#FFE9D6', '#A04A00'], 'Hoàn thành': ['#D7F7EE', '#00695A'], 'Hủy': ['#F1F1F4', '#80808A']
  });
  ['CTV đúng hạn?', 'Giao khách đúng hạn?'].forEach(n => {
    textColor_(sh, 'DH', n, '✔', '#00695A', '#D7F7EE'); textColor_(sh, 'DH', n, '✘', '#B42318', '#FDE2E1');
  });
  textColor_(sh, 'DH', 'Khách cũ?', '🔁', '#5B2CB8', '#EDE4FF');
  deadlineColors_(sh, 'Hạn CTV nộp', 'Ngày CTV nộp');
  deadlineColors_(sh, 'Hạn giao khách', 'Ngày bàn giao');
  newRowHighlight_(sh, 'DH', 'Thời gian');

  shadeHeader_(sh, 'DH', ['Trạng thái', 'Giá chốt', 'Đã thu', 'Ngày bàn giao', 'Ghi chú'], '#1A3FA8');
  shadeHeader_(sh, 'DH', ['CTV phụ trách', 'Thù lao CTV', 'Hạn CTV nộp', 'Ngày CTV nộp', 'Đã trả CTV'], '#5B2CB8');
  noteHeader_(sh, 'DH', {
    'Mã đơn': 'Dòng mới trong 24h được tô xanh nhạt. Cột từ "Mã đơn" đến "Nguồn" do website tự điền.',
    'Hạn giao khách': 'Mặc định = ngày khách mong muốn. Sửa nếu đã thỏa thuận lại.',
    'Trạng thái': 'Chọn "Hoàn thành" → tự điền Ngày bàn giao.',
    'Khách cũ?': 'Tự nhận biết khách đặt lần 2, 3… theo số điện thoại.',
    'Giá chốt': 'Giá cuối cùng đã thống nhất với khách.',
    'Đã thu': 'Tổng tiền đã nhận (cọc + thanh toán).',
    'CTV phụ trách': 'Chọn mã CTV (sheet Cộng tác viên).',
    'Hạn CTV nộp': 'Vàng = còn ≤ 2 ngày, đỏ = đã trễ.',
    'Đã trả CTV': 'Số tiền đã chuyển cho CTV của đơn này.'
  });
  widths_(sh, 'DH', { 'Mã đơn': 125, 'Thời gian': 125, 'Họ tên': 160, 'Gói': 210, 'Mô tả yêu cầu': 260, 'Link tư liệu': 170, 'Trạng thái': 140, 'Khách cũ?': 90, 'Ghi chú': 220 });
  sh.getRange(1, 1, CFG.MAXR, H.DH.length).createFilter();
}

/* --------------------------- Cộng tác viên --------------------------- */
function setupNS_(sh) {
  const R = n => rngLocal_('NS', n);
  const ctv = rng_('DH', 'CTV phụ trách'), dh = rng_('DH', 'CTV đúng hạn?');
  sh.setTabColor('#A970FF');
  setArr_(sh, 'NS', 'Số job', `ARRAYFORMULA(IF(${R('Mã CTV')}="","",COUNTIF(${ctv},${R('Mã CTV')})))`);
  setArr_(sh, 'NS', 'Job đúng hạn', `MAP(${R('Mã CTV')},LAMBDA(id,IF(id="","",COUNTIFS(${ctv},id,${dh},"✔*"))))`);
  setArr_(sh, 'NS', 'Tỉ lệ đúng hạn', `MAP(${R('Mã CTV')},LAMBDA(id,IF(id="","",IFERROR(COUNTIFS(${ctv},id,${dh},"✔*")/COUNTIFS(${ctv},id,${dh},"?*"),""))))`);
  setArr_(sh, 'NS', 'Tổng thù lao', `ARRAYFORMULA(IF(${R('Mã CTV')}="","",SUMIF(${ctv},${R('Mã CTV')},${rng_('DH', 'Thù lao CTV')})))`);
  setArr_(sh, 'NS', 'Đã trả', `ARRAYFORMULA(IF(${R('Mã CTV')}="","",SUMIF(${ctv},${R('Mã CTV')},${rng_('DH', 'Đã trả CTV')})))`);
  setArr_(sh, 'NS', 'Còn nợ', `ARRAYFORMULA(IF(${R('Mã CTV')}="","",${R('Tổng thù lao')}-${R('Đã trả')}))`);
  fmtCol_(sh, 'NS', 'Thời gian ĐK', 'dd/MM/yyyy HH:mm');
  ['Tổng thù lao', 'Đã trả', 'Còn nợ'].forEach(n => fmtCol_(sh, 'NS', n, MONEY));
  fmtCol_(sh, 'NS', 'Tỉ lệ đúng hạn', '0%');
  fmtCol_(sh, 'NS', 'SĐT/Zalo', '@');
  listVal_(sh, 'NS', 'Trạng thái', STATUS_NS);
  listVal_(sh, 'NS', 'Đánh giá (1-5)', ['1', '2', '3', '4', '5']);
  statusColors_(sh, 'NS', 'Trạng thái', {
    'Mới đăng ký': ['#FFF4D6', '#8A5A00'], 'Đang làm bài test': ['#E6F0FF', '#1D4F9F'],
    'Đang hoạt động': ['#D7F7EE', '#00695A'], 'Tạm nghỉ': ['#F1F1F4', '#80808A'], 'Ngưng hợp tác': ['#FDE2E1', '#B42318']
  });
  const c = col_('NS', 'Còn nợ');
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
    .setFontColor('#B42318').setBold(true).setRanges([sh.getRange(`${c}2:${c}${CFG.MAXR}`)]).build());
  sh.setConditionalFormatRules(rules);
  newRowHighlight_(sh, 'NS', 'Thời gian ĐK');
  shadeHeader_(sh, 'NS', ['Trạng thái', 'Đánh giá (1-5)', 'Ghi chú'], '#1A3FA8');
  noteHeader_(sh, 'NS', { 'Trạng thái': 'Mới đăng ký → Đang làm bài test → Đang hoạt động. Chỉ giao job cho CTV "Đang hoạt động".' });
  widths_(sh, 'NS', { 'Họ tên': 170, 'Kỹ năng': 240, 'Portfolio': 200, 'Ghi chú của CTV': 220, 'Trạng thái': 150 });
  sh.getRange(1, 1, CFG.MAXR, H.NS.length).createFilter();
}

/* ============================================================================
 *  TỔNG QUAN — toàn bộ là công thức, tự cập nhật realtime
 * ========================================================================== */
const UI = { bg: '#0B1020', card: '#121A33', soft: '#16203F', line: '#24305A', ink: '#EEF2FF', muted: '#9AA6C9', dim: '#26325C',
             blue: '#5B9BFF', teal: '#00D1B2', violet: '#A970FF', orange: '#FF9F1C', pink: '#FF5C8A', red: '#FF6B6B', font: 'Be Vietnam Pro' };

function taoTongQuan_(ss) {
  let sh = ss.getSheetByName(SH.TQ);
  if (sh) {
    sh.getCharts().forEach(c => sh.removeChart(c));
    if (sh.getFilter()) sh.getFilter().remove();
    const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    all.breakApart(); all.clear(); all.clearDataValidations(); all.clearNote();
    sh.clearConditionalFormatRules();
  } else sh = ss.insertSheet(SH.TQ);
  const ROWS = 70, COLS = 62;
  if (sh.getMaxRows() > ROWS) sh.deleteRows(ROWS + 1, sh.getMaxRows() - ROWS);
  if (sh.getMaxRows() < ROWS) sh.insertRowsAfter(sh.getMaxRows(), ROWS - sh.getMaxRows());
  if (sh.getMaxColumns() < COLS) sh.insertColumnsAfter(sh.getMaxColumns(), COLS - sh.getMaxColumns());
  sh.setFrozenRows(0); sh.setFrozenColumns(0);
  sh.setTabColor(C.teal); sh.setHiddenGridlines(true);
  sh.getRange(1, 1, ROWS, sh.getMaxColumns()).setBackground(UI.bg).setFontColor(UI.muted).setFontFamily(UI.font).setFontSize(10).setVerticalAlignment('middle');

  sh.setColumnWidth(1, 22);
  for (let c = 2; c <= 13; c++) sh.setColumnWidth(c, 112);
  sh.setColumnWidth(14, 22);
  sh.setColumnWidth(15, 40); sh.setColumnWidth(16, 40);
  const heights = { 1: 10, 2: 40, 3: 22, 4: 10, 5: 26, 6: 44, 7: 24, 8: 10, 9: 26, 10: 44, 11: 24, 12: 12, 28: 12, 44: 12, 45: 32, 46: 26, 52: 12, 53: 32, 54: 26, 60: 12, 61: 32, 62: 26, 68: 16 };
  for (let r = 1; r <= ROWS; r++) sh.setRowHeight(r, heights[r] || ((r >= 13 && r <= 43) ? 22 : 26));

  const D = n => rng_('DH', n), N = n => rng_('NS', n);
  const st = D('Trạng thái'), ma = D('Mã đơn'), tg = D('Thời gian'), sdt = D('SĐT/Zalo'), goi = D('Gói'), thu = D('Đã thu');
  const KS = c => `INDIRECT("'${SH.KS}'!${c}2:${c}")`;
  const active = '{' + STATUS_ACTIVE.map(s => `"${s}"`).join(',') + '}';
  const nActive = `SUMPRODUCT(COUNTIF(${st},${active}))`, nDone = `COUNTIF(${st},"Hoàn thành")`, nNew = `COUNTIF(${st},"Mới")`;
  const rate = n => `IFERROR(COUNTIF(${D(n)},"✔*")/COUNTIF(${D(n)},"?*"),0)`;
  const pct = f => `TEXT(${f},"0%")`;
  const money = f => `TEXT(${f},"#,##0")&" đ"`;
  const thuTC = `SUMIFS(${thu},${st},"<>Hủy")`, traCTV = `SUM(${D('Đã trả CTV')})`;
  const noCTV = `SUMIFS(${D('Thù lao CTV')},${st},"<>Hủy")-SUMIFS(${D('Đã trả CTV')},${st},"<>Hủy")`;
  const khach = `IFERROR(ROWS(UNIQUE(FILTER(${sdt},${sdt}<>""))),0)`;
  const quayLai = `IFERROR(SUMPRODUCT(--(COUNTIF(${sdt},UNIQUE(FILTER(${sdt},${sdt}<>"")))>=2)),0)`;
  const thangNay = `">="&EOMONTH(TODAY(),-1)+1`;

  /* ---------- Đầu trang ---------- */
  sh.getRange('B2:H2').merge().setValue(CFG.BRAND).setFontSize(22).setFontWeight('bold').setFontColor(UI.ink);
  sh.getRange('B3:H3').merge().setValue('Bảng điều khiển vận hành').setFontSize(11).setFontColor(UI.muted);
  sh.getRange('I2:M2').merge().setFormula('="●  LIVE"').setFontSize(12).setFontWeight('bold').setFontColor(UI.teal).setHorizontalAlignment('right');
  sh.getRange('I3:M3').merge().setFormula(`="Cập nhật "&TEXT(NOW(),"hh:mm · dd/mm/yyyy")&"  ·  Đơn gần nhất: "&IF(COUNTA(${ma})=0,"chưa có",TEXT(MAX(${tg}),"hh:mm dd/mm"))`).setHorizontalAlignment('right');

  /* ---------- Thẻ số liệu ---------- */
  card_(sh, 5, 2, 'YÊU CẦU HÔM NAY', `=COUNTIFS(${tg},">="&TODAY())`, `="7 ngày qua: "&COUNTIFS(${tg},">="&TODAY()-6)`, UI.blue, '#,##0');
  card_(sh, 5, 4, 'CHƯA XỬ LÝ', `=${nNew}`, `="Tổng yêu cầu: "&COUNTA(${ma})`, UI.pink, '#,##0');
  card_(sh, 5, 6, 'ĐANG SẢN XUẤT', `=${nActive}`, `="Hoàn thành: "&${nDone}`, UI.violet, '#,##0');
  card_(sh, 5, 8, 'TỈ LỆ CHỐT ĐƠN', `=IFERROR((${nActive}+${nDone})/(COUNTA(${ma})-${nNew}),0)`, `="Đã hủy: "&COUNTIF(${st},"Hủy")`, UI.teal, '0%');
  card_(sh, 5, 10, 'CTV NỘP ĐÚNG HẠN', `=${rate('CTV đúng hạn?')}`, `="Giao khách đúng hạn: "&${pct(rate('Giao khách đúng hạn?'))}`, UI.orange, '0%');
  card_(sh, 5, 12, 'CỘNG TÁC VIÊN', `=COUNTA(${N('Mã CTV')})`, `="Hoạt động: "&COUNTIF(${N('Trạng thái')},"Đang hoạt động")&" · Chờ: "&(COUNTIF(${N('Trạng thái')},"Mới đăng ký")+COUNTIF(${N('Trạng thái')},"Đang làm bài test"))`, UI.violet, '#,##0');
  card_(sh, 9, 2, 'DOANH THU ĐÃ THU', `=${thuTC}`, `="Tháng này: "&${money(`SUMIFS(${thu},${tg},${thangNay},${st},"<>Hủy")`)}`, UI.teal, MONEY);
  card_(sh, 9, 4, 'LỢI NHUẬN THỰC', `=${thuTC}-${traCTV}`, '="Đã thu trừ đã trả CTV"', UI.teal, MONEY);
  card_(sh, 9, 6, 'KHÁCH CÒN NỢ', `=SUMIFS(${D('Còn phải thu')},${st},"<>Hủy")`, '="Cần thu thêm"', UI.orange, MONEY);
  card_(sh, 9, 8, 'CÒN NỢ CTV', `=${noCTV}`, `="Đã trả: "&${money(traCTV)}`, UI.pink, MONEY);
  card_(sh, 9, 10, 'SỐ KHÁCH HÀNG', `=${khach}`, `="Quay lại: "&${quayLai}`, UI.blue, '#,##0');
  card_(sh, 9, 12, 'GIÁ TRỊ TB / ĐƠN', `=IFERROR(AVERAGEIFS(${D('Giá chốt')},${st},"<>Hủy"),0)`, '="Theo giá đã chốt"', UI.blue, MONEY);

  /* ---------- Dữ liệu cho biểu đồ & bảng (cột Q trở đi) ---------- */
  sh.getRange('Q1').setValue('Vùng dữ liệu cho biểu đồ — tự tính, không sửa');
  const put = (c, rows) => sh.getRange(2, c, rows.length, rows[0].length).setValues(rows);
  put(17, [['Trạng thái', 'Số đơn']].concat(STATUS_DH.map((s, i) => [s, `=COUNTIF(${st},Q${3 + i})`])));
  const days = [['Ngày', 'Yêu cầu', 'CTV đăng ký']];
  for (let i = 13; i >= 0; i--) days.push([`=TEXT(TODAY()-${i},"dd/mm")`,
    `=COUNTIFS(${tg},">="&TODAY()-${i},${tg},"<"&TODAY()-${i - 1})`,
    `=COUNTIFS(${N('Thời gian ĐK')},">="&TODAY()-${i},${N('Thời gian ĐK')},"<"&TODAY()-${i - 1})`]);
  put(20, days);
  const months = [['Tháng', 'Doanh thu', 'Lợi nhuận gộp']];
  for (let m = 1; m <= 12; m++) {
    const cond = `${tg},">="&DATE(YEAR(TODAY()),${m},1),${tg},"<"&DATE(YEAR(TODAY()),${m + 1},1),${st},"<>Hủy"`;
    months.push(['T' + m, `=SUMIFS(${thu},${cond})`, `=SUMIFS(${D('Lợi nhuận gộp')},${cond})`]);
  }
  put(24, months);
  put(28, [['Gói', 'Số đơn']].concat(PACKAGES.map(p => [p[0] + ' · ' + p[1].replace('Quay + Edit · ', ''), `=COUNTIF(${goi},"${p[0]}*")`]),
    [['Cần tư vấn', `=COUNTIF(${goi},"Chưa chắc*")`]]));
  sh.getRange('AE3').setFormula(`=IFERROR(QUERY({${D('Nguồn')}},"select Col1, count(Col1) where Col1 <> '' group by Col1 order by count(Col1) desc limit 5 label count(Col1) ''",0),"Chưa có dữ liệu")`);
  put(34, [['Kỹ năng', 'Số người']].concat(SKILLS.map((s, i) => [s, `=COUNTIF(${N('Kỹ năng')},"*"&AH${3 + i}&"*")`])));
  sh.getRange('AK3').setFormula(`=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER(AH3:AI17,AI3:AI17>0),2,FALSE),5,2),"Chưa có dữ liệu")`);
  const due = (cell, hanName, doneName, whoName) => {
    const han = D(hanName), done = D(doneName);
    sh.getRange(cell).setFormula(`=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({${ma},${D(whoName)},${han},${st}},${han}<>"",${done}="",${han}<=TODAY()+2,${st}<>"Hoàn thành",${st}<>"Hủy"),3,TRUE),5,4),"✓ Không có")`);
  };
  due('AN3', 'Hạn CTV nộp', 'Ngày CTV nộp', 'CTV phụ trách');
  due('AS3', 'Hạn giao khách', 'Ngày bàn giao', 'Họ tên');
  sh.getRange('AX3').setFormula(`=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({${ma},${D('Họ tên')},${goi},${st},${tg}},${ma}<>""),5,FALSE),5,4),"Chưa có yêu cầu")`);
  sh.getRange('BC3').setFormula(`=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({${N('Mã CTV')},${N('Họ tên')},${N('Kỹ năng')},${N('Trạng thái')},${N('Thời gian ĐK')}},${N('Mã CTV')}<>""),5,FALSE),5,4),"Chưa có CTV")`);
  put(60, [['Khảo sát', 'Giá trị'],
    ['Số phản hồi', `=IFERROR(COUNTA(${KS('A')}),0)`],
    ['Mức độ hữu ích', `=IFERROR(TEXT(AVERAGE(${KS('F')}),"0.00")&" / 5","—")`],
    ['Tính khả thi', `=IFERROR(TEXT(AVERAGE(${KS('G')}),"0.00")&" / 5","—")`],
    ['Website dễ dùng', `=IFERROR(TEXT(AVERAGE(${KS('H')}),"0.00")&" / 5","—")`],
    ['Giá hợp lý', `=IFERROR(TEXT(AVERAGE(${KS('I')}),"0.00")&" / 5","—")`]]);
  sh.getRange(1, 17, 25, COLS - 16).setFontColor(UI.dim).setFontSize(8);

  /* ---------- Bảng ---------- */
  const DATE_F = 'dd/MM/yyyy';
  panel_(sh, 45, 2, 'CTV sắp / đã trễ hạn nộp', UI.red, [2, 2, 1, 1], ['Mã đơn', 'CTV', 'Hạn', 'Trạng thái'], 40, 5, { 2: DATE_F });
  panel_(sh, 45, 8, 'Sắp / đã trễ hạn giao khách', UI.red, [2, 2, 1, 1], ['Mã đơn', 'Khách', 'Hạn', 'Trạng thái'], 45, 5, { 2: DATE_F });
  panel_(sh, 53, 2, 'Yêu cầu mới nhất', UI.blue, [2, 1, 2, 1], ['Mã đơn', 'Khách', 'Gói', 'Trạng thái'], 50, 5, {});
  panel_(sh, 53, 8, 'Cộng tác viên mới nhất', UI.violet, [1, 2, 2, 1], ['Mã', 'Họ tên', 'Kỹ năng', 'Trạng thái'], 55, 5, {});
  panel_(sh, 61, 2, 'Nguồn khách', UI.teal, [3, 1], ['Kênh', 'Đơn'], 31, 5, {});
  panel_(sh, 61, 6, 'Kỹ năng nhiều nhất', UI.violet, [3, 1], ['Kỹ năng', 'Người'], 37, 5, {});
  panel_(sh, 61, 10, 'Kết quả khảo sát', UI.orange, [3, 1], ['Chỉ số', 'Giá trị'], 60, 5, {});

  /* ---------- Biểu đồ ---------- */
  const W = 664, HT = 326;
  const dark = (b, legend) => b.setOption('backgroundColor', UI.card)
    .setOption('titleTextStyle', { color: UI.ink, fontSize: 13, bold: true })
    .setOption('legend', { position: legend || 'top', textStyle: { color: UI.muted } })
    .setOption('hAxis', { textStyle: { color: UI.muted }, gridlines: { color: UI.line }, baselineColor: UI.line })
    .setOption('vAxis', { textStyle: { color: UI.muted }, gridlines: { color: UI.line }, baselineColor: UI.line, minValue: 0 })
    .setOption('width', W).setOption('height', HT);
  sh.insertChart(dark(sh.newChart().asAreaChart().addRange(sh.getRange('T2:V16')).setNumHeaders(1)
    .setOption('title', 'Yêu cầu & CTV đăng ký · 14 ngày').setOption('colors', [UI.blue, UI.violet]))
    .setPosition(13, 2, 4, 0).build());
  sh.insertChart(dark(sh.newChart().asPieChart().addRange(sh.getRange('Q2:R11')).setNumHeaders(1)
    .setOption('title', 'Đơn theo trạng thái').setOption('pieHole', 0.58).setOption('pieSliceBorderColor', UI.card)
    .setOption('colors', ['#FF9F1C', '#5B9BFF', '#3B7BFF', '#A970FF', '#7B4BE0', '#FF5C8A', '#FFB86B', '#00D1B2', '#6B779C']), 'right')
    .setPosition(13, 8, 4, 0).build());
  sh.insertChart(dark(sh.newChart().asColumnChart().addRange(sh.getRange('X2:Z14')).setNumHeaders(1)
    .setOption('title', 'Doanh thu & lợi nhuận theo tháng').setOption('colors', [UI.blue, UI.teal]))
    .setPosition(29, 2, 4, 0).build());
  sh.insertChart(dark(sh.newChart().asBarChart().addRange(sh.getRange('AB2:AC9')).setNumHeaders(1)
    .setOption('title', 'Số đơn theo gói').setOption('colors', [UI.teal]), 'none')
    .setPosition(29, 8, 4, 0).build());
}

/** Thẻ KPI: nhãn / số lớn / chú thích, rộng 2 cột */
function card_(sh, r, c, label, value, sub, accent, fmt) {
  const box = sh.getRange(r, c, 3, 2).setBackground(UI.card);
  sh.getRange(r, c, 1, 2).merge().setValue(label).setFontSize(9).setFontWeight('bold').setFontColor(UI.muted);
  sh.getRange(r + 1, c, 1, 2).merge().setFormula(value).setNumberFormat(fmt).setFontSize(20).setFontWeight('bold').setFontColor(accent).setHorizontalAlignment('left');
  sh.getRange(r + 2, c, 1, 2).merge().setFormula(sub).setFontSize(9).setFontColor(UI.muted);
  box.setBorder(null, true, true, true, false, false, UI.bg, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange(r, c, 1, 2).setBorder(true, null, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/** Bảng nhỏ có tiêu đề; mỗi ô đọc từ vùng dữ liệu (bắt đầu dòng 3 của cột srcCol) */
function panel_(sh, r, c, title, accent, spans, heads, srcCol, n, fmts) {
  const width = spans.reduce((a, b) => a + b, 0);
  sh.getRange(r, c, n + 2, width).setBackground(UI.card);
  sh.getRange(r, c, 1, width).merge().setValue(title).setFontSize(12).setFontWeight('bold').setFontColor(UI.ink);
  let x = c;
  spans.forEach((sp, j) => {
    const h = sh.getRange(r + 1, x, 1, sp); if (sp > 1) h.merge();
    h.setValue(heads[j]).setBackground(UI.soft).setFontColor(UI.muted).setFontWeight('bold').setFontSize(9);
    const L = colLetter_(srcCol + j);
    for (let i = 0; i < n; i++) {
      const cell = sh.getRange(r + 2 + i, x, 1, sp); if (sp > 1) cell.merge();
      cell.setFormula(`=IF(${L}${3 + i}="","",${L}${3 + i})`).setFontColor(UI.ink).setWrap(false);
      if (fmts[j]) cell.setNumberFormat(fmts[j]);
    }
    x += sp;
  });
  sh.getRange(r, c, n + 2, width).setBorder(null, true, true, true, false, false, UI.bg, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange(r, c, 1, width).setBorder(true, null, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/* ============================================================================
 *  NHẬN DỮ LIỆU TỪ WEBSITE (realtime)
 * ========================================================================== */
function doGet() {
  return ContentService.createTextOutput('✅ Web App của ' + CFG.BRAND + ' đang hoạt động.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = parseBody_(e);
    if (d.website) return json_({ ok: true }); // bẫy spam
    let res;
    if (d.loai === 'dich_vu') res = ghiDonHang_(d);
    else if (d.loai === 'ctv') res = ghiCTV_(d);
    else throw new Error('Loại form không hợp lệ: ' + d.loai);
    SpreadsheetApp.flush(); // ghi ngay để sheet & Tổng quan cập nhật tức thì
    return json_(Object.assign({ ok: true }, res));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function ghiDonHang_(d, opt) {
  opt = opt || {};
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.DH);
  if (!sh) throw new Error('Chưa khởi tạo hệ thống');
  const now = opt.time || new Date();
  const ids = colValues_(sh, 1);
  let ma = String(d.ma_don || '');
  if (!/^(DH\d{6}-[A-Z0-9]{3}|DEMO-\d+)$/.test(ma) || ids.indexOf(ma) >= 0) ma = maDon_(now, ids);
  const row = {
    'Mã đơn': ma, 'Thời gian': now, 'Họ tên': clean_(d.ho_ten), 'SĐT/Zalo': phone_(d.sdt), 'Email': clean_(d.email),
    'Đối tượng': clean_(d.doi_tuong), 'Lĩnh vực': clean_(d.linh_vuc), 'Loại dịch vụ': clean_(d.dich_vu), 'Gói': clean_(d.goi),
    'Dịch vụ thêm': clean_(d.them), 'Nền tảng': clean_(d.nen_tang), 'Địa điểm quay': clean_(d.khu_vuc),
    'Hạn giao khách': toDate_(d.deadline), 'Mô tả yêu cầu': clean_(d.mo_ta), 'Link tư liệu': clean_(d.link),
    'Kênh liên hệ': clean_(d.kenh), 'Nguồn': clean_(d.nguon), 'Trạng thái': 'Mới'
  };
  const r = writeRow_(sh, 'DH', row);
  if (!opt.silent) notify_(`🆕 Đơn mới ${ma} — ${row['Họ tên']}`,
    'Có yêu cầu dịch vụ mới từ website:\n\n' +
    Object.keys(row).filter(k => row[k] && k !== 'Thời gian').map(k => `• ${k}: ${row[k] instanceof Date ? fmtDate_(row[k]) : row[k]}`).join('\n') +
    `\n\nMở sheet: ${ss.getUrl()}#gid=${sh.getSheetId()}&range=A${r}`);
  return { ma_don: ma, row: r };
}

function ghiCTV_(d, opt) {
  opt = opt || {};
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SH.NS);
  if (!sh) throw new Error('Chưa khởi tạo hệ thống');
  const p = phone_(d.sdt);
  const phones = colValues_(sh, H.NS.indexOf('SĐT/Zalo') + 1).map(digits_);
  const dup = digits_(p) ? phones.indexOf(digits_(p)) : -1; // trùng SĐT → cập nhật hồ sơ cũ
  const ma = dup >= 0 ? sh.getRange(dup + 2, 1).getValue() : nextCode_(sh, 'CTV', 3);
  const row = {
    'Mã CTV': ma, 'Thời gian ĐK': opt.time || new Date(), 'Họ tên': clean_(d.ho_ten), 'SĐT/Zalo': p, 'Email': clean_(d.email),
    'Khu vực': clean_(d.khu_vuc), 'Kỹ năng': clean_(d.ky_nang), 'Kinh nghiệm': clean_(d.kinh_nghiem), 'Phần mềm': clean_(d.phan_mem),
    'Thiết bị': clean_(d.thiet_bi), 'Portfolio': clean_(d.portfolio), 'Thời gian rảnh': clean_(d.thoi_gian),
    'Mức giá mong muốn': clean_(d.muc_gia), 'Ghi chú của CTV': clean_(d.gioi_thieu),
    'Trạng thái': dup >= 0 ? sh.getRange(dup + 2, H.NS.indexOf('Trạng thái') + 1).getValue() : (d.trang_thai || 'Mới đăng ký')
  };
  const r = writeRow_(sh, 'NS', row, dup >= 0 ? dup + 2 : null);
  if (!opt.silent) notify_(`🧑‍💻 ${dup >= 0 ? 'CTV cập nhật hồ sơ' : 'CTV mới'} — ${row['Họ tên']} (${ma})`,
    `Kỹ năng: ${row['Kỹ năng']}\nKinh nghiệm: ${row['Kinh nghiệm']}\nKhu vực: ${row['Khu vực']}\nPortfolio: ${row['Portfolio']}\nSĐT: ${row['SĐT/Zalo']}\n\n${ss.getUrl()}#gid=${sh.getSheetId()}&range=A${r}`);
  return { ma_ctv: ma, row: r };
}

/** Khi Thái sửa tay */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SH.DH || e.range.getRow() < 2) return;
    if (e.range.getColumn() === H.DH.indexOf('Trạng thái') + 1 && e.value === 'Hoàn thành') {
      const cell = sh.getRange(e.range.getRow(), H.DH.indexOf('Ngày bàn giao') + 1);
      if (!cell.getValue()) cell.setValue(new Date(new Date().toDateString()));
    }
  } catch (err) { console.error(err); }
}

/* ============================================================================
 *  FORM KHẢO SÁT → sheet "Khảo sát" (thống kê nằm trong Tổng quan)
 *  Cột: A thời gian | B tên | C đối tượng | D lĩnh vực | E đã thuê | F hữu ích | G khả thi | H website | I giá
 *       J gói | K ngân sách | L điểm yếu | M góp ý | N SĐT
 * ========================================================================== */
const KS_DIEMYEU = ['Giá còn cao', 'Ít video mẫu, khó tin tưởng', 'Chưa có đánh giá từ khách cũ', 'Thời gian giao còn lâu',
  'Website khó dùng / thiếu thông tin', 'Chưa rõ ai trực tiếp làm video', 'Kênh liên hệ chưa tiện', 'Chưa có chính sách bảo hành / hoàn tiền'];

function taoFormKhaoSat() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getDocumentProperties();
  if (props.getProperty('FORM_ID')) {
    if (ui.alert('Đã có form khảo sát', 'Tạo thêm form MỚI?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return xemLink();
  }
  const ss = SpreadsheetApp.getActive();
  const form = FormApp.create('Khảo sát dịch vụ video quảng cáo – ' + CFG.BRAND);
  form.setDescription('Mất khoảng 2 phút. Góp ý của bạn giúp ' + CFG.BRAND + ' đánh giá mức độ hữu ích, tính khả thi và điểm cần cải thiện. Thông tin chỉ dùng cho mục đích nghiên cứu.')
    .setProgressBar(true).setShowLinkToRespondAgain(false).setConfirmationMessage('Cảm ơn bạn đã góp ý! 🙏');
  form.addTextItem().setTitle('Họ và tên (không bắt buộc)');
  form.addMultipleChoiceItem().setTitle('Bạn là').setChoiceValues(DOITUONG.filter(x => x !== 'Khác')).showOtherOption(true).setRequired(true);
  form.addTextItem().setTitle('Lĩnh vực / sản phẩm bạn kinh doanh hoặc quan tâm');
  form.addMultipleChoiceItem().setTitle('Bạn đã từng thuê làm video quảng cáo chưa?').setChoiceValues(['Đã từng thuê', 'Chưa, nhưng có ý định', 'Chưa và chưa có nhu cầu']).setRequired(true);
  form.addSectionHeaderItem().setTitle('Đánh giá dự án').setHelpText('Hãy xem website trước khi trả lời (1 = thấp nhất, 5 = cao nhất).');
  form.addScaleItem().setTitle('Mức độ HỮU ÍCH của dịch vụ với bạn').setBounds(1, 5).setLabels('Không hữu ích', 'Rất hữu ích').setRequired(true);
  form.addScaleItem().setTitle('Mức độ KHẢ THI: khả năng bạn sẽ thực sự sử dụng dịch vụ').setBounds(1, 5).setLabels('Chắc chắn không', 'Chắc chắn có').setRequired(true);
  form.addScaleItem().setTitle('Website có dễ xem, dễ chọn gói và gửi yêu cầu không?').setBounds(1, 5).setLabels('Rất khó', 'Rất dễ').setRequired(true);
  form.addScaleItem().setTitle('Mức giá các gói có hợp lý không?').setBounds(1, 5).setLabels('Quá đắt', 'Rất hợp lý').setRequired(true);
  form.addMultipleChoiceItem().setTitle('Gói bạn quan tâm nhất').setChoiceValues(['Edit nhanh (200k)', 'Edit chuẩn (450k)', 'Edit chuyên sâu (900k)', 'Quay + Edit Khởi đầu (1,5tr)', 'Quay + Edit Kinh doanh (3tr)', 'Quay + Edit Toàn diện (5,5tr)', 'Chưa quan tâm gói nào']).setRequired(true);
  form.addMultipleChoiceItem().setTitle('Ngân sách bạn sẵn sàng chi cho 1 video').setChoiceValues(['Dưới 500.000đ', '500.000đ – 1.500.000đ', '1.500.000đ – 3.000.000đ', '3.000.000đ – 6.000.000đ', 'Trên 6.000.000đ']).setRequired(true);
  form.addSectionHeaderItem().setTitle('Điểm yếu & góp ý');
  form.addCheckboxItem().setTitle('Dự án còn những ĐIỂM YẾU nào cần cải thiện? (chọn nhiều)').setChoiceValues(KS_DIEMYEU).showOtherOption(true).setRequired(true);
  form.addParagraphTextItem().setTitle('Góp ý chi tiết để dự án tốt hơn');
  form.addTextItem().setTitle('SĐT/Zalo nếu bạn muốn được tư vấn (không bắt buộc)');

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush(); Utilities.sleep(3000);
  const old = ss.getSheetByName(SH.KS);
  if (old) old.setName(SH.KS + ' cũ ' + Utilities.formatDate(new Date(), CFG.TZ, 'ddMM-HHmm'));
  const resp = ss.getSheets().find(s => { try { return (s.getFormUrl() || '').indexOf(form.getId()) >= 0; } catch (e) { return false; } });
  if (resp) {
    resp.setName(SH.KS); resp.setTabColor('#E0A100'); resp.setFrozenRows(1);
    resp.getRange(1, 1, 1, resp.getLastColumn()).setBackground(C.soft).setFontColor(C.ink).setFontWeight('bold').setWrap(true);
    ss.setActiveSheet(resp); ss.moveActiveSheet(4);
  }
  props.setProperty('FORM_ID', form.getId());
  props.setProperty('FORM_URL', form.getPublishedUrl());
  props.setProperty('FORM_EDIT', form.getEditUrl());
  xemLink();
}

function xemLink() {
  const p = PropertiesService.getDocumentProperties(), ss = SpreadsheetApp.getActive();
  let webapp = '';
  try { webapp = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial;font-size:13px;line-height:1.7">
      <b>Google Sheet:</b><br><a href="${ss.getUrl()}" target="_blank">${ss.getUrl()}</a><br><br>
      <b>Web App (dán vào SHEET_ENDPOINT):</b><br>${webapp ? `<code>${webapp}</code>` : '<i>Chưa triển khai. Triển khai → Tùy chọn triển khai mới → Ứng dụng web.</i>'}<br><br>
      <b>Form khảo sát (gửi cho người khảo sát):</b><br>${p.getProperty('FORM_URL') ? `<a href="${p.getProperty('FORM_URL')}" target="_blank">${p.getProperty('FORM_URL')}</a>` : '<i>Chưa tạo</i>'}<br><br>
      <b>Sửa form:</b><br>${p.getProperty('FORM_EDIT') ? `<a href="${p.getProperty('FORM_EDIT')}" target="_blank">Mở trình sửa form</a>` : '<i>—</i>'}
    </div>`).setWidth(560).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, '🔗 Các link quan trọng');
}

/* ============================================================================
 *  DỮ LIỆU DEMO — có chữ [DEMO], xóa bằng menu trước khi dùng thật
 * ========================================================================== */
function themDuLieuDemo() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SH.DH)) khoiTaoHeThong();
  const day = n => { const d = new Date(new Date().toDateString()); d.setDate(d.getDate() + n); d.setHours(9 + Math.abs(n) % 9); return d; };
  const o = t => ({ silent: true, time: t });
  const ctv = [
    ghiCTV_({ ho_ten: '[DEMO] CTV Editor A', sdt: '0900000001', ky_nang: 'Edit video ngắn (TikTok/Reels), Color grading', portfolio: 'https://example.com', trang_thai: 'Đang hoạt động' }, o(day(-40))).ma_ctv,
    ghiCTV_({ ho_ten: '[DEMO] CTV Quay phim B', sdt: '0900000002', ky_nang: 'Quay sự kiện, Flycam', portfolio: 'https://example.com', trang_thai: 'Đang hoạt động' }, o(day(-35))).ma_ctv,
    ghiCTV_({ ho_ten: '[DEMO] CTV Editor C', sdt: '0900000003', ky_nang: 'Motion graphics, Edit video ngắn (TikTok/Reels)', portfolio: 'https://example.com' }, o(day(-2))).ma_ctv
  ];
  const sh = ss.getSheetByName(SH.DH);
  const set = (row, obj) => Object.keys(obj).forEach(k => sh.getRange(row, H.DH.indexOf(k) + 1).setValue(obj[k]));
  const demo = [
    [{ ho_ten: '[DEMO] Quán cà phê X', sdt: '0911000001', doi_tuong: 'Hộ kinh doanh', dich_vu: 'Quay + Edit', goi: 'QE1 · Khởi đầu (1.500.000đ)', deadline: day(-23), nguon: 'tiktok' }, -30,
      { 'Trạng thái': 'Hoàn thành', 'Giá chốt': 1500000, 'Đã thu': 1500000, 'CTV phụ trách': ctv[1], 'Thù lao CTV': 600000, 'Hạn CTV nộp': day(-26), 'Ngày CTV nộp': day(-26), 'Đã trả CTV': 600000, 'Ngày bàn giao': day(-24) }],
    [{ ho_ten: '[DEMO] Shop mỹ phẩm Y', sdt: '0911000002', doi_tuong: 'Người bán hàng online', dich_vu: 'Chỉ Edit', goi: 'ED2 · Edit chuẩn (450.000đ)', deadline: day(-16), nguon: 'facebook' }, -20,
      { 'Trạng thái': 'Hoàn thành', 'Giá chốt': 450000, 'Đã thu': 450000, 'CTV phụ trách': ctv[0], 'Thù lao CTV': 230000, 'Hạn CTV nộp': day(-18), 'Ngày CTV nộp': day(-17), 'Đã trả CTV': 230000, 'Ngày bàn giao': day(-15) }],
    [{ ho_ten: '[DEMO] Shop mỹ phẩm Y', sdt: '0911000002', doi_tuong: 'Người bán hàng online', dich_vu: 'Chỉ Edit', goi: 'ED1 · Edit nhanh (200.000đ)', deadline: day(2), nguon: 'facebook' }, -3,
      { 'Trạng thái': 'Đang sản xuất', 'Giá chốt': 200000, 'Đã thu': 100000, 'CTV phụ trách': ctv[0], 'Thù lao CTV': 100000, 'Hạn CTV nộp': day(1) }],
    [{ ho_ten: '[DEMO] Công ty nội thất Z', sdt: '0911000003', doi_tuong: 'Doanh nghiệp nhỏ', dich_vu: 'Quay + Edit', goi: 'QE2 · Kinh doanh (3.000.000đ)', khu_vuc: 'Biên Hòa', deadline: day(8), nguon: 'Trực tiếp' }, -4,
      { 'Trạng thái': 'Đã cọc', 'Giá chốt': 3000000, 'Đã thu': 1500000, 'CTV phụ trách': ctv[1], 'Thù lao CTV': 1300000, 'Hạn CTV nộp': day(5) }],
    [{ ho_ten: '[DEMO] Tiệm bánh W', sdt: '0911000004', doi_tuong: 'Cá nhân kinh doanh', dich_vu: 'Chỉ Edit', goi: 'Chưa chắc, cần tư vấn', nguon: 'tiktok' }, 0, { 'Trạng thái': 'Mới' }]
  ];
  demo.forEach(([d, offs, upd]) => set(ghiDonHang_(Object.assign({ loai: 'dich_vu' }, d), o(day(offs))).row, upd));
  SpreadsheetApp.getUi().alert('Đã thêm dữ liệu DEMO. Mở "Tổng quan" để xem. Xóa bằng 🎬 Studio → Xóa dữ liệu DEMO.');
}

function xoaDuLieuDemo() {
  const ss = SpreadsheetApp.getActive();
  ['DH', 'NS'].forEach(k => {
    const sh = ss.getSheetByName(SH[k]); if (!sh) return;
    const last = colValues_(sh, 1).length; if (!last) return;
    const vals = sh.getRange(2, 1, last, 4).getDisplayValues();
    let del = 0;
    for (let i = vals.length - 1; i >= 0; i--) if (vals[i].join('|').indexOf('[DEMO]') >= 0) { sh.deleteRow(i + 2); del++; }
    if (del) sh.insertRowsAfter(Math.max(2, colValues_(sh, 1).length + 1), del); // giữ phạm vi công thức 2:MAXR
  });
  SpreadsheetApp.getUi().alert('Đã xóa toàn bộ dữ liệu DEMO.');
}

/* ============================================================================
 *  HELPERS
 * ========================================================================== */
function colLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function col_(key, name) { const i = H[key].indexOf(name); if (i < 0) throw new Error(`Không có cột "${name}" trong ${SH[key]}`); return colLetter_(i + 1); }
function rng_(key, name) { const c = col_(key, name); return `'${SH[key]}'!${c}2:${c}${CFG.MAXR}`; }
function rngLocal_(key, name) { const c = col_(key, name); return `${c}2:${c}${CFG.MAXR}`; }

/** Công thức mảng đặt ở ô tiêu đề → tự áp dụng cho mọi dòng mới (realtime) */
function setArr_(sh, key, name, formula) {
  const c = H[key].indexOf(name) + 1;
  sh.getRange(2, c, sh.getMaxRows() - 1, 1).clearContent().clearDataValidations();
  sh.getRange(1, c).setFormula(`={"${name}";${formula}}`).setBackground(C.calc).setNote('Cột tự tính — không gõ vào.');
}
function fmtCol_(sh, key, name, fmt) { const c = col_(key, name); sh.getRange(`${c}2:${c}`).setNumberFormat(fmt); }
function listVal_(sh, key, name, list) {
  const c = col_(key, name);
  sh.getRange(`${c}2:${c}`).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build());
}
function dateVal_(sh, key, names) {
  names.forEach(n => { const c = col_(key, n); sh.getRange(`${c}2:${c}`).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()); });
}
function addRule_(sh, rule) { const r = sh.getConditionalFormatRules(); r.push(rule); sh.setConditionalFormatRules(r); }
function statusColors_(sh, key, name, map) {
  const c = col_(key, name), rg = sh.getRange(`${c}2:${c}${CFG.MAXR}`);
  const rules = sh.getConditionalFormatRules();
  Object.keys(map).forEach(v => rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
    .setBackground(map[v][0]).setFontColor(map[v][1]).setBold(true).setRanges([rg]).build()));
  sh.setConditionalFormatRules(rules);
}
function textColor_(sh, key, name, txt, fg, bg) {
  const c = col_(key, name);
  addRule_(sh, SpreadsheetApp.newConditionalFormatRule().whenTextContains(txt).setFontColor(fg).setBackground(bg)
    .setRanges([sh.getRange(`${c}2:${c}${CFG.MAXR}`)]).build());
}
/** Tô xanh nhạt các dòng vừa vào trong 24h */
function newRowHighlight_(sh, key, timeName) {
  const t = col_(key, timeName);
  addRule_(sh, SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($${t}2<>"",NOW()-$${t}2<1)`)
    .setBackground('#E9FBF6').setRanges([sh.getRange(2, 1, CFG.MAXR - 1, H[key].indexOf(LAST_INPUT[key]))]).build());
}
function deadlineColors_(sh, hanName, doneName) {
  const han = col_('DH', hanName), done = col_('DH', doneName), tt = col_('DH', 'Trạng thái');
  const rg = sh.getRange(`${han}2:${han}${CFG.MAXR}`);
  const open = `$${han}2<>"",$${done}2="",$${tt}2<>"Hoàn thành",$${tt}2<>"Hủy"`;
  addRule_(sh, SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${open},$${han}2<TODAY())`)
    .setBackground('#FDE2E1').setFontColor('#B42318').setBold(true).setRanges([rg]).build());
  addRule_(sh, SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${open},$${han}2-TODAY()<=2)`)
    .setBackground('#FFF4D6').setFontColor('#8A5A00').setBold(true).setRanges([rg]).build());
}
function shadeHeader_(sh, key, names, color) { names.forEach(n => sh.getRange(1, H[key].indexOf(n) + 1).setBackground(color)); }
function noteHeader_(sh, key, notes) { Object.keys(notes).forEach(n => sh.getRange(1, H[key].indexOf(n) + 1).setNote(notes[n])); }
function widths_(sh, key, map) { H[key].forEach((n, i) => sh.setColumnWidth(i + 1, map[n] || 115)); }

function writeRow_(sh, key, obj, atRow) {
  const r = atRow || colValues_(sh, 1).length + 2;
  const n = H[key].indexOf(LAST_INPUT[key]) + 1;
  sh.getRange(r, 1, 1, n).setValues([H[key].slice(0, n).map(h => (obj[h] === undefined ? '' : obj[h]))]);
  return r;
}
function colValues_(sh, col) {
  const last = sh.getLastRow(); if (last < 2) return [];
  const v = sh.getRange(2, col, last - 1, 1).getValues().map(r => r[0]);
  let n = v.length; while (n > 0 && (v[n - 1] === '' || v[n - 1] === null)) n--;
  return v.slice(0, n);
}
function nextCode_(sh, prefix, pad) {
  const n = colValues_(sh, 1).reduce((m, v) => Math.max(m, parseInt(String(v).replace(/\D/g, ''), 10) || 0), 0) + 1;
  return prefix + String(n).padStart(pad, '0');
}
function maDon_(d, ids) {
  const A = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let ma;
  do { ma = 'DH' + Utilities.formatDate(d, CFG.TZ, 'yyMMdd') + '-' + [0, 0, 0].map(() => A[Math.floor(Math.random() * A.length)]).join(''); }
  while (ids.indexOf(ma) >= 0);
  return ma;
}
function clean_(v) {
  let s = String(v === undefined || v === null ? '' : v).trim().slice(0, 3000);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // chặn chèn công thức
  return s;
}
function toDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}
function phone_(v) { return String(v || '').replace(/[^\d+]/g, ''); }
function digits_(v) { let s = String(v || '').replace(/\D/g, ''); if (s.indexOf('84') === 0 && s.length === 11) s = '0' + s.slice(2); return s; }
function fmtDate_(d) { return Utilities.formatDate(d, CFG.TZ, 'dd/MM/yyyy'); }
function parseBody_(e) {
  if (e && e.postData && e.postData.contents) { try { return JSON.parse(e.postData.contents); } catch (x) {} }
  return (e && e.parameter) || {};
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function notify_(subject, body) {
  if (!CFG.SEND_NOTIFY) return;
  try {
    const to = CFG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    if (to) MailApp.sendEmail(to, '[' + CFG.BRAND + '] ' + subject, body);
  } catch (e) { console.error('Không gửi được email: ' + e); }
}
