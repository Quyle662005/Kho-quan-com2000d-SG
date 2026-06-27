/**
 * ════════════════════════════════════════════════════════════════
 * APPS SCRIPT — Kho Quán Cơm 2000đ  (Code.gs)
 * Xử lý toàn bộ API cho cả form nhân viên & trang admin
 *
 * Sheets cần có trong Google Spreadsheet:
 *   "ỦNG HỘ HIỆN VẬT"  — nhập kho đã duyệt
 *   "XUẤT KHO"          — xuất kho (ghi thẳng)
 *   "CHỜ DUYỆT"         — nhập kho chờ admin phê duyệt (tự tạo)
 *
 * DEPLOY (làm 1 lần):
 *   Extensions → Apps Script → dán code → Save
 *   Deploy → New deployment → Web app
 *   Execute as: Me | Who has access: Anyone
 *   → Copy URL → dán vào SHEET_URL trong cả 2 file HTML
 * ════════════════════════════════════════════════════════════════
 */

// ── CẤU HÌNH ────────────────────────────────────────────────────
const SH_NHAP    = "ỦNG HỘ HIỆN VẬT";
const SH_XUAT    = "XUẤT KHO";
const SH_PENDING = "CHỜ DUYỆT";

// Email admin được phép truy cập (thêm email vào đây)
const ADMIN_EMAILS = [
  "leq662005gmail.com",   // ← thay bằng email của bạn
  // "another_admin@gmail.com",
];

// ── ROUTER ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    switch(action) {
      case "submit_nhap":    return submitNhap(body);
      case "submit_xuat":    return submitXuat(body);
      case "get_pending":    return getPending(body);
      case "approve":        return approve(body);
      case "reject":         return reject(body);
      case "get_dashboard":  return getDashboard(body);
      case "get_history":    return getHistory(body);
      default:
        return res({status:"error", message:"action không hợp lệ: " + action});
    }
  } catch(err) {
    return res({status:"error", message: err.toString()});
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === "ping") return res({status:"ok", message:"API đang hoạt động 🟢"});
  return res({status:"ok", message:"Kho API 🟢"});
}


// ════════════════════════════════════════════════════════════════
//  KIỂM TRA QUYỀN ADMIN
// ════════════════════════════════════════════════════════════════
function checkAdmin(token) {
  // token là Google ID token từ frontend
  // Verify bằng Google tokeninfo endpoint
  try {
    const url  = `https://oauth2.googleapis.com/tokeninfo?id_token=${token}`;
    const resp = UrlFetchApp.fetch(url);
    const info = JSON.parse(resp.getContentText());

    if (info.error) return {ok: false, reason: "Token không hợp lệ"};

    const email = info.email;
    if (!ADMIN_EMAILS.includes(email)) {
      return {ok: false, reason: `Email ${email} không có quyền admin`};
    }
    return {ok: true, email};
  } catch(e) {
    return {ok: false, reason: "Không xác thực được token: " + e.toString()};
  }
}


// ════════════════════════════════════════════════════════════════
//  NHÂN VIÊN: NỘP PHIẾU NHẬP → GHI VÀO "CHỜ DUYỆT"
// ════════════════════════════════════════════════════════════════
function submitNhap(body) {
  ensurePendingSheet();
  const sheet = getSheet(SH_PENDING);
  const items = body.items || [];
  const ts    = new Date();
  // Tạo ID phiếu duy nhất: PND + timestamp milliseconds
  const phieuId = "PND" + ts.getTime();

  items.forEach(item => {
    sheet.appendRow([
      phieuId,             // A: Mã phiếu (để nhóm các dòng cùng phiếu)
      ts,                  // B: Thời gian nộp
      formatDateVN(body.ngay), // C: Ngày ghi nhận
      body.nguon,          // D: Người ủng hộ
      body.nguoi,          // E: Người ghi nhận
      body.diaChi,         // F: Địa chỉ
      body.sdt,            // G: SĐT
      (item.ma||'').toLowerCase(), // H: Mã NVL
      item.ten,            // I: Tên hàng
      item.qty,            // J: Số lượng
      item.dvt,            // K: ĐVT
      "CHỜ DUYỆT",        // L: Trạng thái
    ]);
  });

  return res({
    status: "ok",
    message: `Đã gửi ${items.length} mặt hàng. Chờ admin phê duyệt!`,
    phieuId,
    rows: items.length,
  });
}


// ════════════════════════════════════════════════════════════════
//  NHÂN VIÊN: XUẤT KHO → GHI THẲNG VÀO "XUẤT KHO"
// ════════════════════════════════════════════════════════════════
function submitXuat(body) {
  const sheet = getSheet(SH_XUAT);
  if (!sheet) return res({status:"error", message:`Không tìm thấy sheet "${SH_XUAT}"`});

  const items = body.items || [];
  const ts    = parseDate(body.ngay);

  items.forEach(item => {
    sheet.appendRow([
      null,                // A: trống
      ts,                  // B: Thời gian
      body.dienGiai,       // C: Diễn giải
      body.diaChi,         // D: Địa chỉ
      body.sdt,            // E: SĐT
      null,                // F: Địa chỉ 2
      (item.ma||'').toLowerCase(), // G: Mã
      item.ten,            // H: Tên
      item.qty,            // I: Số lượng
      item.dvt,            // J: ĐVT
    ]);
  });

  fmtNewRows(sheet, items.length, 9);
  return res({status:"ok", message:`Đã xuất kho ${items.length} mặt hàng`, rows: items.length});
}


// ════════════════════════════════════════════════════════════════
//  ADMIN: LẤY DANH SÁCH CHỜ DUYỆT
// ════════════════════════════════════════════════════════════════
function getPending(body) {
  const auth = checkAdmin(body.token);
  if (!auth.ok) return res({status:"error", message: auth.reason});

  ensurePendingSheet();
  const sheet = getSheet(SH_PENDING);
  const data  = sheet.getDataRange().getValues();

  // Nhóm các dòng theo phieuId
  const phieuMap = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const [phieuId, thoiGianNop, ngay, nguon, nguoi, diaChi, sdt, maNVL, ten, qty, dvt, ttrang] = row;
    if (!phieuId) continue;

    if (!phieuMap[phieuId]) {
      phieuMap[phieuId] = {
        phieuId,
        thoiGianNop: thoiGianNop ? new Date(thoiGianNop).toLocaleString("vi-VN") : "",
        ngay: ngay || "",
        nguon: nguon || "",
        nguoi: nguoi || "",
        diaChi: diaChi || "",
        sdt: sdt || "",
        trangThai: ttrang || "CHỜ DUYỆT",
        items: [],
        rowStart: i + 1, // 1-indexed
        rowCount: 0,
      };
    }
    phieuMap[phieuId].items.push({maNVL, ten, qty, dvt, rowIdx: i + 1});
    phieuMap[phieuId].rowCount++;
  }

  // Chỉ trả về phiếu còn CHỜ DUYỆT
  const pending = Object.values(phieuMap).filter(p => p.trangThai === "CHỜ DUYỆT");

  return res({status:"ok", data: pending, total: pending.length});
}


// ════════════════════════════════════════════════════════════════
//  ADMIN: PHÊ DUYỆT → CHUYỂN SANG "ỦNG HỘ HIỆN VẬT"
// ════════════════════════════════════════════════════════════════
function approve(body) {
  const auth = checkAdmin(body.token);
  if (!auth.ok) return res({status:"error", message: auth.reason});

  const { phieuId } = body;
  if (!phieuId) return res({status:"error", message:"Thiếu phieuId"});

  const pendingSheet = getSheet(SH_PENDING);
  const nhapSheet    = getSheet(SH_NHAP);
  if (!nhapSheet) return res({status:"error", message:`Không tìm thấy sheet "${SH_NHAP}"`});

  const data     = pendingSheet.getDataRange().getValues();
  const tsApprove = new Date();
  let   count    = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] !== phieuId) continue;
    if (row[11] !== "CHỜ DUYỆT") continue;

    // Ghi vào sheet chính
    nhapSheet.appendRow([
      null,           // A: trống
      parseDate(row[2]), // B: Ngày ghi nhận gốc
      row[3],         // C: Người ủng hộ
      row[5],         // D: Địa chỉ
      row[6],         // E: SĐT
      null,           // F: Địa chỉ 2
      row[7],         // G: Mã NVL
      row[8],         // H: Tên
      row[9],         // I: Số lượng
      row[10],        // J: ĐVT
    ]);

    // Cập nhật trạng thái trong CHỜ DUYỆT
    pendingSheet.getRange(i + 1, 12).setValue("ĐÃ DUYỆT");
    pendingSheet.getRange(i + 1, 13).setValue(auth.email);
    pendingSheet.getRange(i + 1, 14).setValue(tsApprove);
    count++;
  }

  fmtNewRows(nhapSheet, count, 9);

  return res({
    status: "ok",
    message: `Đã duyệt phiếu ${phieuId} (${count} mặt hàng)`,
    approved: count,
  });
}


// ════════════════════════════════════════════════════════════════
//  ADMIN: TỪ CHỐI PHIẾU
// ════════════════════════════════════════════════════════════════
function reject(body) {
  const auth = checkAdmin(body.token);
  if (!auth.ok) return res({status:"error", message: auth.reason});

  const { phieuId, lyDo } = body;
  if (!phieuId) return res({status:"error", message:"Thiếu phieuId"});

  const pendingSheet = getSheet(SH_PENDING);
  const data         = pendingSheet.getDataRange().getValues();
  const ts           = new Date();
  let   count        = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] !== phieuId || row[11] !== "CHỜ DUYỆT") continue;

    pendingSheet.getRange(i + 1, 12).setValue("TỪ CHỐI");
    pendingSheet.getRange(i + 1, 13).setValue(auth.email);
    pendingSheet.getRange(i + 1, 14).setValue(ts);
    pendingSheet.getRange(i + 1, 15).setValue(lyDo || "");
    count++;
  }

  return res({
    status: "ok",
    message: `Đã từ chối phiếu ${phieuId}`,
    rejected: count,
  });
}


// ════════════════════════════════════════════════════════════════
//  ADMIN: DASHBOARD — THỐNG KÊ TỔNG QUAN
// ════════════════════════════════════════════════════════════════
function getDashboard(body) {
  const auth = checkAdmin(body.token);
  if (!auth.ok) return res({status:"error", message: auth.reason});

  const nhapData    = getSheetValues(SH_NHAP,    5); // data từ dòng 5
  const xuatData    = getSheetValues(SH_XUAT,    5);
  const pendingData = getSheetValues(SH_PENDING, 2); // header dòng 1

  // ── THỐNG KÊ NHẬP ──
  // Nhóm theo tên hàng, tính tổng số lượng
  const nhapByItem  = {};
  nhapData.forEach(r => {
    const ten = r[7] || ""; // cột H: Tên
    const qty = parseFloat(r[8]) || 0; // cột I: SL
    const dvt = r[9] || "";
    if (!ten) return;
    if (!nhapByItem[ten]) nhapByItem[ten] = {ten, dvt, tongNhap: 0};
    nhapByItem[ten].tongNhap += qty;
  });

  // ── THỐNG KÊ XUẤT ──
  const xuatByItem = {};
  xuatData.forEach(r => {
    const ten = r[7] || "";
    const qty = parseFloat(r[8]) || 0;
    const dvt = r[9] || "";
    if (!ten) return;
    if (!xuatByItem[ten]) xuatByItem[ten] = {ten, dvt, tongXuat: 0};
    xuatByItem[ten].tongXuat += qty;
  });

  // ── TỒN KHO = NHẬP - XUẤT ──
  const allItems = new Set([...Object.keys(nhapByItem), ...Object.keys(xuatByItem)]);
  const tonKho   = [];
  allItems.forEach(ten => {
    const nhap = nhapByItem[ten]?.tongNhap || 0;
    const xuat = xuatByItem[ten]?.tongXuat || 0;
    const dvt  = nhapByItem[ten]?.dvt || xuatByItem[ten]?.dvt || "";
    tonKho.push({ten, dvt, tongNhap: nhap, tongXuat: xuat, ton: nhap - xuat});
  });
  tonKho.sort((a,b) => a.ten.localeCompare(b.ten, 'vi'));

  // ── CHỜ DUYỆT ──
  const chouDuyet = pendingData.filter(r => r[11] === "CHỜ DUYỆT");
  const pendingPhieu = {};
  chouDuyet.forEach(r => {
    const id = r[0];
    if (!pendingPhieu[id]) pendingPhieu[id] = 0;
    pendingPhieu[id]++;
  });

  // ── GIAO DỊCH GẦN ĐÂY ──
  const recent = [];

  // 10 dòng nhập gần nhất
  nhapData.slice(-10).reverse().forEach(r => {
    if (!r[7]) return;
    recent.push({
      loai: "nhap", thoiGian: r[1] ? formatDateStr(r[1]) : "",
      nguon: r[2] || "", ten: r[7], qty: r[8], dvt: r[9],
    });
  });

  // 10 dòng xuất gần nhất
  xuatData.slice(-10).reverse().forEach(r => {
    if (!r[7]) return;
    recent.push({
      loai: "xuat", thoiGian: r[1] ? formatDateStr(r[1]) : "",
      dienGiai: r[2] || "", ten: r[7], qty: r[8], dvt: r[9],
    });
  });

  // Sắp xếp giao dịch gần nhất lên đầu
  recent.sort((a,b) => b.thoiGian.localeCompare(a.thoiGian));

  return res({
    status: "ok",
    data: {
      tongNhapPhieu:   nhapData.filter(r => r[7]).length,
      tongXuatPhieu:   xuatData.filter(r => r[7]).length,
      pendingCount:    Object.keys(pendingPhieu).length,
      pendingItemCount: chouDuyet.length,
      tonKho,
      recent: recent.slice(0, 20),
    }
  });
}


// ════════════════════════════════════════════════════════════════
//  ADMIN: LỊCH SỬ CHI TIẾT (NHẬP hoặc XUẤT)
// ════════════════════════════════════════════════════════════════
function getHistory(body) {
  const auth = checkAdmin(body.token);
  if (!auth.ok) return res({status:"error", message: auth.reason});

  const { loai, page = 1, pageSize = 30 } = body;
  const sheetName = loai === "nhap" ? SH_NHAP : SH_XUAT;
  const allData   = getSheetValues(sheetName, 5).filter(r => r[7]);

  // Phân trang từ mới nhất
  const reversed = [...allData].reverse();
  const total    = reversed.length;
  const start    = (page - 1) * pageSize;
  const rows     = reversed.slice(start, start + pageSize).map(r => ({
    thoiGian: r[1] ? formatDateStr(r[1]) : "",
    col3:     r[2] || "",   // nguon (nhập) hoặc dienGiai (xuất)
    diaChi:   r[3] || "",
    sdt:      r[4] || "",
    maNVL:    r[6] || "",
    ten:      r[7] || "",
    qty:      r[8] || 0,
    dvt:      r[9] || "",
  }));

  return res({status:"ok", data:{rows, total, page, pageSize, totalPages: Math.ceil(total/pageSize)}});
}


// ════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════
function ensurePendingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SH_PENDING)) {
    const sh = ss.insertSheet(SH_PENDING);
    sh.appendRow([
      "Mã phiếu","Thời gian nộp","Ngày","Người ủng hộ","Người ghi nhận",
      "Địa chỉ","SĐT","Mã NVL","Tên hàng","Số lượng","ĐVT",
      "Trạng thái","Admin xử lý","Thời gian xử lý","Lý do từ chối"
    ]);
    // Style header
    sh.getRange(1,1,1,15).setBackground("#1a3050").setFontColor("#fff").setFontWeight("bold");
    sh.setFrozenRows(1);
  }
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// Lấy values từ dòng startRow trở xuống, bỏ qua dòng header/tiêu đề
function getSheetValues(sheetName, startRow) {
  const sh = getSheet(sheetName);
  if (!sh || sh.getLastRow() < startRow) return [];
  return sh.getRange(startRow, 1, sh.getLastRow() - startRow + 1, sh.getLastColumn()).getValues();
}

function fmtNewRows(sheet, count, qtyCol) {
  if (count <= 0) return;
  const last  = sheet.getLastRow();
  const first = last - count + 1;
  sheet.getRange(first, qtyCol, count, 1).setNumberFormat('#,##0.##');
  sheet.getRange(first, 2, count, 1).setHorizontalAlignment('center');
  sheet.getRange(first, 10, count, 1).setHorizontalAlignment('center');
}

function parseDate(dateStr) {
  if (!dateStr) return new Date();
  if (dateStr instanceof Date) return dateStr;
  const [y,m,d] = String(dateStr).split('-').map(Number);
  return new Date(y, m-1, d);
}

function formatDateVN(dateStr) {
  if (!dateStr) return "";
  const d = parseDate(dateStr);
  return `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`;
}

function formatDateStr(val) {
  try {
    const d = new Date(val);
    return d.toLocaleDateString("vi-VN") + " " + d.toLocaleTimeString("vi-VN", {hour:"2-digit",minute:"2-digit"});
  } catch(e) { return String(val); }
}

function res(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}