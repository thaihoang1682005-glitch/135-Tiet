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
  try { ss.setSpreadsheetLocale('vi_VN'); } catch (e) {}
  ss.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);

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
  const sdt = col_('DH', 'SĐT/Zalo');
  setArr_(sh, 'DH', 'Khách cũ?',
    `MAP(${R('SĐT/Zalo')},ROW(${R('SĐT/Zalo')}),LAMBDA(p,r,IF(p="","",LET(n,COUNTIF(${sdt}$2:INDEX(${sdt}$2:${sdt}$${CFG.MAXR},r-1),p),IF(n>1,"🔁 Lần "&n,"Mới")))))`);
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
function taoTongQuan_(ss) {
  let sh = ss.getSheetByName(SH.TQ);
  if (sh) { sh.getCharts().forEach(c => sh.removeChart(c)); sh.clear(); sh.clearConditionalFormatRules(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); }
  else sh = ss.insertSheet(SH.TQ);
  sh.setTabColor(C.teal); sh.setHiddenGridlines(true);
  if (sh.getMaxColumns() < 22) sh.insertColumnsAfter(sh.getMaxColumns(), 22 - sh.getMaxColumns());
  if (sh.getMaxRows() < 90) sh.insertRowsAfter(sh.getMaxRows(), 90 - sh.getMaxRows());

  const D = n => rng_('DH', n), N = n => rng_('NS', n);
  const st = D('Trạng thái'), ma = D('Mã đơn'), tg = D('Thời gian'), sdt = D('SĐT/Zalo');
  const KS = c => `INDIRECT("'${SH.KS}'!${c}2:${c}")`;
  const active = '{' + STATUS_ACTIVE.map(s => `"${s}"`).join(',') + '}';
  const rate = n => `=IFERROR(COUNTIF(${D(n)},"✔*")/COUNTIF(${D(n)},"?*"),"—")`;

  sh.getRange('A1').setValue('🎬 ' + CFG.BRAND + ' — BẢNG ĐIỀU KHIỂN').setFontSize(18).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('A2').setFormula(`="🟢 Realtime · cập nhật "&TEXT(NOW(),"hh:mm dd/mm/yyyy")&"  ·  Đơn mới nhất: "&IFERROR(TEXT(MAX(${tg}),"hh:mm dd/mm"),"—")`).setFontColor('#00695A');

  /* ---------- KPI (A:B) ---------- */
  const kpi = [
    ['⚡ HÔM NAY & GẦN ĐÂY', ''],
    ['Yêu cầu hôm nay', `=COUNTIFS(${tg},">="&TODAY())`],
    ['Yêu cầu 7 ngày qua', `=COUNTIFS(${tg},">="&TODAY()-6)`],
    ['Yêu cầu tháng này', `=COUNTIFS(${tg},">="&EOMONTH(TODAY(),-1)+1)`],
    ['CTV đăng ký 7 ngày qua', `=COUNTIFS(${N('Thời gian ĐK')},">="&TODAY()-6)`],
    ['📦 ĐƠN HÀNG', ''],
    ['Tổng yêu cầu', `=COUNTA(${ma})`],
    ['Chưa xử lý (Mới)', `=COUNTIF(${st},"Mới")`],
    ['Đang sản xuất', `=SUM(COUNTIF(${st},${active}))`],
    ['Hoàn thành', `=COUNTIF(${st},"Hoàn thành")`],
    ['Hủy', `=COUNTIF(${st},"Hủy")`],
    ['Tỉ lệ chốt đơn', '=IFERROR((B12+B13)/(B10-B11),"—")'],
    ['CTV nộp đúng hạn', rate('CTV đúng hạn?')],
    ['Giao khách đúng hạn', rate('Giao khách đúng hạn?')],
    ['💰 TÀI CHÍNH', ''],
    ['Doanh thu đã thu', `=SUM(${D('Đã thu')})`],
    ['Khách còn nợ', `=SUM(${D('Còn phải thu')})`],
    ['Giá trị TB / đơn chốt', `=IFERROR(AVERAGE(${D('Giá chốt')}),0)`],
    ['Đã trả CTV', `=SUM(${D('Đã trả CTV')})`],
    ['Còn nợ CTV', `=SUM(${D('Thù lao CTV')})-SUM(${D('Đã trả CTV')})`],
    ['Lợi nhuận thực (tiền mặt)', '=B19-B22'],
    ['Lợi nhuận gộp (theo giá chốt)', `=SUMIFS(${D('Lợi nhuận gộp')},${st},"<>Hủy")`],
    ['👥 KHÁCH & CỘNG TÁC VIÊN', ''],
    ['Số khách (theo SĐT)', `=IFERROR(ROWS(UNIQUE(FILTER(${sdt},${sdt}<>""))),0)`],
    ['Khách quay lại (≥ 2 đơn)', `=IFERROR(SUMPRODUCT(--(COUNTIF(${sdt},UNIQUE(FILTER(${sdt},${sdt}<>"")))>=2)),0)`],
    ['CTV đã đăng ký', `=COUNTA(${N('Mã CTV')})`],
    ['CTV đang hoạt động', `=COUNTIF(${N('Trạng thái')},"Đang hoạt động")`],
    ['CTV chờ duyệt / test', `=COUNTIF(${N('Trạng thái')},"Mới đăng ký")+COUNTIF(${N('Trạng thái')},"Đang làm bài test")`]
  ];
  const k0 = 4;
  sh.getRange(k0, 1, kpi.length, 1).setValues(kpi.map(r => [r[0]]));
  kpi.forEach((r, i) => { if (r[1]) sh.getRange(k0 + i, 2).setFormula(r[1]); });
  sh.getRange(k0, 2, kpi.length, 1).setFontWeight('bold').setHorizontalAlignment('right').setNumberFormat('#,##0');
  // Dòng: 5 hôm nay · 10 tổng · 11 mới · 12 đang SX · 13 xong · 15-17 tỉ lệ · 19-25 tiền · 27-31 người
  ['B19', 'B20', 'B21', 'B22', 'B23', 'B24', 'B25'].forEach(a => sh.getRange(a).setNumberFormat(MONEY));
  ['B15', 'B16', 'B17'].forEach(a => sh.getRange(a).setNumberFormat('0%'));
  kpi.forEach((r, i) => {
    const rg = sh.getRange(k0 + i, 1, 1, 2);
    if (!r[1]) rg.merge().setBackground(C.soft).setFontColor(C.ink).setFontWeight('bold');
    else rg.setBackground(i % 2 ? '#FFFFFF' : C.zebra);
  });
  sh.getRange(k0, 1, kpi.length, 2).setBorder(true, true, true, true, null, null, '#D5DAEA', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B5').setFontColor('#1A3FA8').setFontSize(13);
  sh.getRange('B11').setFontColor('#B42318').setFontSize(13);
  sh.getRange('B24').setFontColor('#00695A').setFontSize(13);
  sh.getRange('B31').setFontColor('#5B2CB8');

  /* ---------- Khảo sát (A:B, dưới KPI) ---------- */
  const s0 = k0 + kpi.length + 2; // 34
  const ks = [
    ['📝 KHẢO SÁT (Google Form)', ''],
    ['Số phản hồi', `=IFERROR(COUNTA(${KS('A')}),0)`],
    ['Điểm HỮU ÍCH (TB /5)', `=IFERROR(AVERAGE(${KS('F')}),"—")`],
    ['Điểm KHẢ THI (TB /5)', `=IFERROR(AVERAGE(${KS('G')}),"—")`],
    ['Website dễ dùng (TB /5)', `=IFERROR(AVERAGE(${KS('H')}),"—")`],
    ['Giá hợp lý (TB /5)', `=IFERROR(AVERAGE(${KS('I')}),"—")`],
    ['% thấy hữu ích (≥ 4)', `=IFERROR(COUNTIF(${KS('F')},">=4")/B${s0 + 1},"—")`],
    ['% sẵn sàng sử dụng (≥ 4)', `=IFERROR(COUNTIF(${KS('G')},">=4")/B${s0 + 1},"—")`]
  ];
  sh.getRange(s0, 1, ks.length, 1).setValues(ks.map(r => [r[0]]));
  ks.forEach((r, i) => { if (r[1]) sh.getRange(s0 + i, 2).setFormula(r[1]); });
  sh.getRange(s0, 1, 1, 2).merge().setBackground('#E0A100').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(s0 + 1, 2, ks.length - 1, 1).setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(s0 + 2, 2, 4, 1).setNumberFormat('0.00');
  sh.getRange(s0 + 6, 2, 2, 1).setNumberFormat('0%');
  zebra_(sh, s0 + 1, 1, ks.length - 1, 2);
  const KS_DIEMYEU_ = KS_DIEMYEU;
  const y0 = s0 + ks.length + 1;
  sh.getRange(y0, 1, 1, 2).setValues([['Điểm yếu được chọn nhiều nhất', 'Lượt']]);
  KS_DIEMYEU_.forEach((x, i) => { sh.getRange(y0 + 1 + i, 1).setValue(x); sh.getRange(y0 + 1 + i, 2).setFormula(`=IFERROR(COUNTIF(${KS('L')},"*${x}*"),0)`); });
  tbl_(sh, y0, 1, KS_DIEMYEU_.length + 1, 2);

  /* ---------- Trạng thái (D:E) ---------- */
  hdr_(sh, 3, 4, ['Trạng thái đơn', 'Số đơn']);
  STATUS_DH.forEach((s, i) => { sh.getRange(4 + i, 4).setValue(s); sh.getRange(4 + i, 5).setFormula(`=COUNTIF(${st},D${4 + i})`); });
  tbl_(sh, 3, 4, STATUS_DH.length + 1, 2);

  /* ---------- Theo gói (G:I) ---------- */
  hdr_(sh, 3, 7, ['Gói', 'Số đơn', 'Doanh thu đã thu']);
  const goi = D('Gói'), thu = D('Đã thu');
  PACKAGES.concat([['Chưa chắc', 'Chưa chọn / cần tư vấn']]).forEach((p, i) => {
    const r = 4 + i;
    sh.getRange(r, 7).setValue(p[0] === 'Chưa chắc' ? p[1] : p[0] + ' · ' + p[1].replace(/^Quay \+ Edit · /, ''));
    sh.getRange(r, 8).setFormula(`=COUNTIF(${goi},"${p[0]}*")`);
    sh.getRange(r, 9).setFormula(`=SUMIF(${goi},"${p[0]}*",${thu})`).setNumberFormat(MONEY);
  });
  tbl_(sh, 3, 7, PACKAGES.length + 2, 3);

  /* ---------- Nguồn khách (K:L) ---------- */
  hdr_(sh, 3, 11, ['Nguồn khách (?src=)', 'Số đơn']);
  sh.getRange(4, 11).setFormula(`=IFERROR(QUERY({${D('Nguồn')}},"select Col1, count(Col1) where Col1 <> '' group by Col1 order by count(Col1) desc limit 8 label count(Col1) ''",0),"Chưa có")`);
  tbl_(sh, 3, 11, 9, 2);

  /* ---------- Loại dịch vụ (G:I) ---------- */
  hdr_(sh, 13, 7, ['Loại dịch vụ', 'Số đơn', 'Doanh thu đã thu']);
  SERVICES.forEach((s, i) => {
    sh.getRange(14 + i, 7).setValue(s);
    sh.getRange(14 + i, 8).setFormula(`=COUNTIF(${D('Loại dịch vụ')},G${14 + i})`);
    sh.getRange(14 + i, 9).setFormula(`=SUMIF(${D('Loại dịch vụ')},G${14 + i},${thu})`).setNumberFormat(MONEY);
  });
  tbl_(sh, 13, 7, SERVICES.length + 1, 3);

  /* ---------- Đối tượng khách (D:E) ---------- */
  hdr_(sh, 15, 4, ['Khách là', 'Số đơn']);
  DOITUONG.forEach((s, i) => { sh.getRange(16 + i, 4).setValue(s); sh.getRange(16 + i, 5).setFormula(`=COUNTIF(${D('Đối tượng')},D${16 + i})`); });
  tbl_(sh, 15, 4, DOITUONG.length + 1, 2);

  /* ---------- Kỹ năng CTV (K:L) ---------- */
  hdr_(sh, 14, 11, ['Kỹ năng CTV', 'Số người']);
  SKILLS.forEach((s, i) => { sh.getRange(15 + i, 11).setValue(s); sh.getRange(15 + i, 12).setFormula(`=COUNTIF(${N('Kỹ năng')},"*"&K${15 + i}&"*")`); });
  tbl_(sh, 14, 11, SKILLS.length + 1, 2);

  /* ---------- Theo tháng (D:I) ---------- */
  const m0 = 34;
  sh.getRange(m0, 4).setValue('📅 THEO THÁNG — năm:').setFontWeight('bold');
  sh.getRange(m0, 5).setFormula('=YEAR(TODAY())').setBackground('#FFF4D6').setFontWeight('bold').setNumberFormat('0').setNote('Gõ năm khác để xem.');
  hdr_(sh, m0 + 1, 4, ['Tháng', 'Yêu cầu', 'Hoàn thành', 'Doanh thu đã thu', 'Thù lao CTV', 'Lợi nhuận gộp']);
  for (let m = 1; m <= 12; m++) {
    const r = m0 + 1 + m;
    const cond = `${tg},">="&DATE($E$${m0},${m},1),${tg},"<"&EDATE(DATE($E$${m0},${m},1),1)`;
    sh.getRange(r, 4).setValue('Tháng ' + m);
    sh.getRange(r, 5).setFormula(`=COUNTIFS(${cond})`);
    sh.getRange(r, 6).setFormula(`=COUNTIFS(${cond},${st},"Hoàn thành")`);
    sh.getRange(r, 7).setFormula(`=SUMIFS(${thu},${cond})`);
    sh.getRange(r, 8).setFormula(`=SUMIFS(${D('Thù lao CTV')},${cond},${st},"<>Hủy")`);
    sh.getRange(r, 9).setFormula(`=G${r}-H${r}`);
  }
  const mt = m0 + 14;
  sh.getRange(mt, 4).setValue('Cả năm');
  ['E', 'F', 'G', 'H', 'I'].forEach(cl => sh.getRange(`${cl}${mt}`).setFormula(`=SUM(${cl}${m0 + 2}:${cl}${m0 + 13})`));
  sh.getRange(`G${m0 + 2}:I${mt}`).setNumberFormat(MONEY);
  tbl_(sh, m0 + 1, 4, 14, 6);
  sh.getRange(mt, 4, 1, 6).setFontWeight('bold').setBackground('#E6F0FF');

  /* ---------- 14 ngày gần nhất (K:M) ---------- */
  sh.getRange(m0, 11).setValue('📈 14 NGÀY GẦN NHẤT').setFontWeight('bold');
  hdr_(sh, m0 + 1, 11, ['Ngày', 'Yêu cầu', 'CTV đăng ký']);
  for (let i = 0; i < 14; i++) {
    const r = m0 + 2 + i;
    sh.getRange(r, 11).setFormula(`=TODAY()-${13 - i}`).setNumberFormat('dd/MM');
    sh.getRange(r, 12).setFormula(`=COUNTIFS(${tg},">="&K${r},${tg},"<"&K${r}+1)`);
    sh.getRange(r, 13).setFormula(`=COUNTIFS(${N('Thời gian ĐK')},">="&K${r},${N('Thời gian ĐK')},"<"&K${r}+1)`);
  }
  tbl_(sh, m0 + 1, 11, 15, 3);

  /* ---------- Cảnh báo hạn ---------- */
  const a0 = mt + 3; // 51
  const open = `${st}<>"Hoàn thành",${st}<>"Hủy"`;
  const alert = (row, col, title, hanName, doneName, whoName, whoLabel) => {
    const han = D(hanName), done = D(doneName);
    sh.getRange(row, col).setValue(title).setFontWeight('bold').setFontColor('#B42318');
    hdr_(sh, row + 1, col, ['Mã đơn', whoLabel, 'Hạn', 'Trạng thái']);
    sh.getRange(row + 2, col).setFormula(
      `=IFERROR(SORT(FILTER({${ma},${D(whoName)},${han},${st}},${han}<>"",${done}="",${han}<=TODAY()+2,${open}),3,TRUE),"✓ Không có")`);
    sh.getRange(row + 2, col + 2, 12, 1).setNumberFormat('dd/MM/yyyy');
    tbl_(sh, row + 1, col, 1, 4);
  };
  alert(a0, 4, '⚠️ CTV SẮP / ĐÃ TRỄ HẠN NỘP (≤ 2 ngày)', 'Hạn CTV nộp', 'Ngày CTV nộp', 'CTV phụ trách', 'CTV');
  alert(a0, 10, '⚠️ SẮP / ĐÃ TRỄ HẠN GIAO KHÁCH (≤ 2 ngày)', 'Hạn giao khách', 'Ngày bàn giao', 'Họ tên', 'Khách');

  /* ---------- Mới nhất ---------- */
  const l0 = a0 + 16; // 67
  sh.getRange(l0, 4).setValue('🆕 5 YÊU CẦU MỚI NHẤT').setFontWeight('bold');
  hdr_(sh, l0 + 1, 4, ['Mã đơn', 'Thời gian', 'Khách', 'Gói', 'Trạng thái']);
  sh.getRange(l0 + 2, 4).setFormula(
    `=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({${ma},${tg},${D('Họ tên')},${goi},${st}},${ma}<>""),2,FALSE),5,5),"Chưa có yêu cầu")`);
  sh.getRange(l0 + 2, 5, 5, 1).setNumberFormat('hh:mm dd/MM');
  tbl_(sh, l0 + 1, 4, 1, 5);
  sh.getRange(l0, 10).setValue('🧑‍💻 5 CTV ĐĂNG KÝ MỚI NHẤT').setFontWeight('bold');
  hdr_(sh, l0 + 1, 10, ['Mã CTV', 'Họ tên', 'Kỹ năng', 'Trạng thái']);
  sh.getRange(l0 + 2, 10).setFormula(
    `=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({${N('Mã CTV')},${N('Họ tên')},${N('Kỹ năng')},${N('Trạng thái')},${N('Thời gian ĐK')}},${N('Mã CTV')}<>""),5,FALSE),5,4),"Chưa có CTV")`);
  tbl_(sh, l0 + 1, 10, 1, 4);

  [240, 150, 16, 170, 120, 110, 190, 110, 140, 120, 240, 90, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setColumnWidth(14, 16);

  /* ---------- Biểu đồ (cột O) ---------- */
  const put = (b, row) => sh.insertChart(b.setPosition(row, 15, 0, 0).setOption('width', 480).setOption('height', 290).build());
  put(sh.newChart().asPieChart().addRange(sh.getRange(`D3:E${3 + STATUS_DH.length}`)).setNumHeaders(1)
    .setOption('title', 'Đơn theo trạng thái').setOption('pieHole', 0.45).setOption('legend', { position: 'right' }), 3);
  put(sh.newChart().asLineChart().addRange(sh.getRange(`K${m0 + 1}:M${m0 + 15}`)).setNumHeaders(1)
    .setOption('title', '14 ngày gần nhất').setOption('colors', [C.blue, '#A970FF']).setOption('legend', { position: 'top' }), 19);
  put(sh.newChart().asColumnChart().addRange(sh.getRange(`D${m0 + 1}:D${m0 + 13}`)).addRange(sh.getRange(`G${m0 + 1}:G${m0 + 13}`))
    .addRange(sh.getRange(`I${m0 + 1}:I${m0 + 13}`)).setNumHeaders(1)
    .setOption('title', 'Doanh thu & lợi nhuận theo tháng').setOption('colors', [C.blue, C.teal]).setOption('legend', { position: 'top' }), 35);
  put(sh.newChart().asBarChart().addRange(sh.getRange(`K14:L${14 + SKILLS.length}`)).setNumHeaders(1)
    .setOption('title', 'Kỹ năng của đội CTV').setOption('colors', ['#A970FF']).setOption('legend', { position: 'none' }), 51);
}

function hdr_(sh, r, c, vals) { sh.getRange(r, c, 1, vals.length).setValues([vals]); }
function zebra_(sh, r, c, nr, nc) { for (let i = 0; i < nr; i++) sh.getRange(r + i, c, 1, nc).setBackground(i % 2 ? '#FFFFFF' : C.zebra); }
function tbl_(sh, r, c, nr, nc) {
  sh.getRange(r, c, 1, nc).setBackground(C.soft).setFontColor(C.ink).setFontWeight('bold');
  sh.getRange(r, c, nr, nc).setBorder(true, true, true, true, true, true, '#D5DAEA', SpreadsheetApp.BorderStyle.SOLID);
  zebra_(sh, r + 1, c, nr - 1, nc);
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
