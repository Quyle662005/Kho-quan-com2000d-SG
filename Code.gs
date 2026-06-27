// ============================================================
//  Code.gs — Webhook cho hệ thống quản lý kho
//  Quán Cơm 2000đ Sài Gòn
// ============================================================

const SHEET_GIAO_DICH = 'SoGiaoDich';
const SHEET_TON_KHO   = 'TonKho';

// ------------------------------------------------------------
//  doPost — Nhận dữ liệu từ Web Form (GitHub Pages)
// ------------------------------------------------------------
function doPost(e) {
  try {
    // Parse JSON body
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    const loai         = (data.loaiGiaoDich  || '').trim();
    const nguoi        = (data.nguoiThucHien || '').trim();
    const sanPham      = (data.sanPham       || '').trim();
    const soLuong      = parseFloat(data.soLuong) || 0;
    const donVi        = (data.donVi         || '').trim();
    const ghiChu       = (data.ghiChu        || '').trim();

    // Validation cơ bản
    if (!loai || !nguoi || !sanPham || soLuong <= 0 || !donVi) {
      return jsonResponse({ status: 'error', message: 'Thiếu thông tin bắt buộc.' });
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_GIAO_DICH);

    if (!sheet) {
      return jsonResponse({ status: 'error', message: `Không tìm thấy sheet "${SHEET_GIAO_DICH}".` });
    }

    // Xác định trạng thái duyệt:
    //   Xuất kho  → TRUE  (tự động duyệt, trừ ngay vào kho)
    //   Nhập kho  → FALSE (chờ Admin tick duyệt)
    const duyet = (loai === 'Xuất kho');

    const timestamp = new Date();

    // Thêm dòng mới vào cuối sheet
    sheet.appendRow([
      timestamp,    // A — Thời gian
      loai,         // B — Loại giao dịch
      nguoi,        // C — Người thực hiện
      sanPham,      // D — Sản phẩm
      soLuong,      // E — Số lượng
      donVi,        // F — Đơn vị
      ghiChu,       // G — Ghi chú
      duyet,        // H — Trạng thái Duyệt (checkbox / boolean)
    ]);

    // Format cột Thời gian cho dòng vừa thêm
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    return jsonResponse({ status: 'success', message: 'Ghi nhận thành công.' });

  } catch (err) {
    console.error('doPost error:', err);
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ------------------------------------------------------------
//  doGet — Health-check / ping để kiểm tra deploy
// ------------------------------------------------------------
function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'Webhook đang hoạt động.' });
}

// ------------------------------------------------------------
//  Helper — trả về JSON với CORS headers
// ------------------------------------------------------------
function jsonResponse(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
//  setupSheets — Chạy MỘT LẦN để tạo header và định dạng
//  Menu: Extensions → Apps Script → Run → setupSheets
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Sheet 1: SoGiaoDich ──────────────────────────────────
  let gdSheet = ss.getSheetByName(SHEET_GIAO_DICH);
  if (!gdSheet) {
    gdSheet = ss.insertSheet(SHEET_GIAO_DICH);
  }

  // Header row
  const gdHeaders = [
    'Thời gian', 'Loại giao dịch', 'Người thực hiện',
    'Sản phẩm', 'Số lượng', 'Đơn vị', 'Ghi chú', 'Trạng thái Duyệt'
  ];
  const gdHeader = gdSheet.getRange(1, 1, 1, gdHeaders.length);
  gdHeader.setValues([gdHeaders]);
  gdHeader.setFontWeight('bold');
  gdHeader.setBackground('#C84B2F');
  gdHeader.setFontColor('#ffffff');
  gdSheet.setFrozenRows(1);

  // Column widths
  gdSheet.setColumnWidth(1, 160);  // Thời gian
  gdSheet.setColumnWidth(2, 110);  // Loại
  gdSheet.setColumnWidth(3, 160);  // Người
  gdSheet.setColumnWidth(4, 140);  // Sản phẩm
  gdSheet.setColumnWidth(5, 90);   // Số lượng
  gdSheet.setColumnWidth(6, 80);   // Đơn vị
  gdSheet.setColumnWidth(7, 200);  // Ghi chú
  gdSheet.setColumnWidth(8, 130);  // Duyệt

  // Đặt validation Checkbox cho cột H (từ hàng 2)
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  gdSheet.getRange('H2:H1000').setDataValidation(checkboxRule);

  // ── Sheet 2: TonKho ─────────────────────────────────────
  let tkSheet = ss.getSheetByName(SHEET_TON_KHO);
  if (!tkSheet) {
    tkSheet = ss.insertSheet(SHEET_TON_KHO);
  }

  const tkHeaders = [
    'Sản phẩm', 'Đơn vị', 'Tổng Nhập (đã duyệt)', 'Tổng Xuất', 'Tồn kho hiện tại'
  ];
  const tkHeader = tkSheet.getRange(1, 1, 1, tkHeaders.length);
  tkHeader.setValues([tkHeaders]);
  tkHeader.setFontWeight('bold');
  tkHeader.setBackground('#2E7D5A');
  tkHeader.setFontColor('#ffffff');
  tkSheet.setFrozenRows(1);

  // Danh sách sản phẩm mẫu (từ file Bảng mã)
  const products = [
    ['Gạo','kg'], ['Dầu ăn','lít'], ['Nước mắm','lít'], ['Nước tương','lít'],
    ['Hạt nêm','kg'], ['Bột ngọt','kg'], ['Muối','kg'], ['Đường','kg'],
    ['Tiêu','kg'], ['Tương ớt','lít'], ['Dầu hào','lít'],
    ['Thịt heo','kg'], ['Xúc xích','kg'], ['Trứng','cái'], ['Lạp xưởng','kg'],
    ['Cá hộp','hộp'], ['Mì gói','gói'], ['Nấm','kg'], ['Nấm mèo','kg'],
    ['Bún gạo','kg'], ['Rau củ','kg'], ['Trái cây','kg'],
    ['Nước rửa chén','kg'],
  ];

  products.forEach(([ten, dvt], i) => {
    const row = i + 2;
    tkSheet.getRange(row, 1).setValue(ten);
    tkSheet.getRange(row, 2).setValue(dvt);

    // Tổng Nhập — chỉ tính dòng đã duyệt (H=TRUE) và loại = "Nhập kho"
    tkSheet.getRange(row, 3).setFormula(
      `=IFERROR(SUMPRODUCT((SoGiaoDich!D$2:D$5000=A${row})*(SoGiaoDich!B$2:B$5000="Nhập kho")*(SoGiaoDich!H$2:H$5000=TRUE)*SoGiaoDich!E$2:E$5000),0)`
    );

    // Tổng Xuất — chỉ tính dòng đã duyệt (H=TRUE) và loại = "Xuất kho"
    tkSheet.getRange(row, 4).setFormula(
      `=IFERROR(SUMPRODUCT((SoGiaoDich!D$2:D$5000=A${row})*(SoGiaoDich!B$2:B$5000="Xuất kho")*(SoGiaoDich!H$2:H$5000=TRUE)*SoGiaoDich!E$2:E$5000),0)`
    );

    // Tồn kho = Nhập - Xuất
    tkSheet.getRange(row, 5).setFormula(`=C${row}-D${row}`);
  });

  // Conditional formatting: tồn kho < 0 → đỏ cảnh báo
  const tonKhoRange = tkSheet.getRange('E2:E100');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#FDECEA')
    .setFontColor('#C84B2F')
    .setRanges([tonKhoRange])
    .build();
  tkSheet.setConditionalFormatRules([rule]);

  // Column widths
  tkSheet.setColumnWidth(1, 160);
  tkSheet.setColumnWidth(2, 80);
  tkSheet.setColumnWidth(3, 160);
  tkSheet.setColumnWidth(4, 110);
  tkSheet.setColumnWidth(5, 150);

  SpreadsheetApp.getUi().alert('✅ Thiết lập hoàn tất! Sheet SoGiaoDich và TonKho đã sẵn sàng.');
}


// ============================================================
//  ADMIN API — đọc dữ liệu & cập nhật trạng thái duyệt
// ============================================================

function doGet(e) {
  const action = e.parameter.action || 'ping';

  if (action === 'ping') {
    return jsonResponse({ status: 'ok', message: 'Webhook đang hoạt động.' });
  }

  if (action === 'getAll') {
    return getAllData();
  }

  return jsonResponse({ status: 'error', message: 'Unknown action.' });
}

function getAllData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_GIAO_DICH);
  if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found.' });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ status: 'ok', transactions: [] });

  const data  = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const rows  = data
    .filter(r => r[0] !== '')   // bỏ dòng trống
    .map((r, i) => ({
      id:       i + 2,          // row number in sheet (for update)
      time:     r[0] ? Utilities.formatDate(new Date(r[0]), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm') : '',
      type:     r[1] || '',
      person:   r[2] || '',
      product:  r[3] || '',
      qty:      parseFloat(r[4]) || 0,
      unit:     r[5] || '',
      note:     r[6] || '',
      approved: r[7] === true,
    }));

  return jsonResponse({ status: 'ok', transactions: rows });
}

// Hàm doPost mở rộng để xử lý cả "approve" action
// (thêm vào đầu hàm doPost cũ, trước phần parse giao dịch)
function doPostAdmin(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    // Action: duyệt / bỏ duyệt
    if (data.action === 'approve') {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_GIAO_DICH);
      if (!sheet) return jsonResponse({ status:'error', message:'Sheet not found.' });

      const rowNum = parseInt(data.id);
      if (isNaN(rowNum) || rowNum < 2) return jsonResponse({ status:'error', message:'Invalid row id.' });

      sheet.getRange(rowNum, 8).setValue(data.approved === true);
      return jsonResponse({ status:'success', message:'Đã cập nhật trạng thái.' });
    }

    // Delegate to original doPost logic (nhận giao dịch mới)
    return doPost(e);

  } catch(err) {
    return jsonResponse({ status:'error', message: err.toString() });
  }
}