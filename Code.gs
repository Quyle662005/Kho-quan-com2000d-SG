function doGet(e) {
  const cache = CacheService.getScriptCache();
  // Nếu URL có ?refresh=1 -> bỏ qua cache, quét dữ liệu mới ngay lập tức
  const boQuaCache = e && e.parameter && e.parameter.refresh === '1';

  if (!boQuaCache) {
    const cached = cache.get('cuumang_data');
    if (cached) {
      return ContentService.createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Cache rỗng, vừa hết hạn, hoặc người dùng chủ động yêu cầu làm mới -> quét trực tiếp
  const result = quetDuLieu_();
  const jsonStr = JSON.stringify(result);

  try {
    cache.put('cuumang_data', jsonStr, 600); // cache 10 phút
  } catch (e) {
    // Nếu JSON quá 100KB (giới hạn CacheService) thì bỏ qua cache, vẫn trả kết quả bình thường
  }

  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

// Hàm chạy nền tự động làm mới cache trước khi hết hạn (gắn Trigger chạy mỗi 5 phút)
function lamMoiCache() {
  const result = quetDuLieu_();
  const jsonStr = JSON.stringify(result);
  try {
    CacheService.getScriptCache().put('cuumang_data', jsonStr, 600);
  } catch (e) {
    console.log('Dữ liệu quá lớn để cache: ' + e.message);
  }
}

// Quét toàn bộ workbook, trả về object dữ liệu (không tự trả JSON ở đây)
function quetDuLieu_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const result = {
    "_meta": {
      "updated": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
    }
  };

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Bỏ qua sheet hướng dẫn hoặc các sheet không cần thiết
    if (sheetName.toUpperCase().includes("HƯỚNG DẪN") || sheetName.toUpperCase().includes("TEMPLATE")) return;

    const range = sheet.getDataRange();
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();

    // 1. Tìm dòng tiêu đề chuẩn chứa "MÃ HC" hoặc "MÃ HOÀN CẢNH"
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(15, values.length); i++) {
      const rowStr = values[i].join("").toUpperCase();
      if (rowStr.includes("MÃ HC") || rowStr.includes("MÃ HOÀN CẢNH")) {
        headerRowIdx = i;
        break;
      }
    }

    if (headerRowIdx === -1) return; // Bỏ qua nếu không tìm thấy tiêu đề

    // Xử lý tiêu đề (xóa khoảng trắng thừa, gộp dòng nếu có Alt+Enter)
    const headers = values[headerRowIdx].map(h => h.toString().toUpperCase().replace(/\n/g, ' ').trim());

    // 2. Map vị trí các cột để lấy dữ liệu chính xác
    const colMap = { months: {} };
    headers.forEach((h, idx) => {
      if (h.includes("MÃ HC") || h.includes("MÃ HOÀN CẢNH")) colMap.ma = idx;
      else if (h === "HOÀN CẢNH") colMap.ten = idx;
      else if (h === "KHU VỰC") colMap.khuVuc = idx;
      else if (h.includes("MẠNH THƯỜNG QUÂN") || h.includes("MẠNH THƯỜNG QUÂN")) colMap.mtq = idx;
      else if (h.includes("MỨC HỖ TRỢ")) colMap.muc = idx;
      else if (h === "GHI CHÚ") colMap.ghiChu = idx;
      else if (/^T([1-9]|1[0-2])$/.test(h)) colMap.months[h] = idx; // Tìm T1 -> T12
    });

    const sheetData = [];
    let currentHC = null; // Biến lưu hoàn cảnh hiện tại để gộp dòng

    // 3. Duyệt qua các dòng dữ liệu bên dưới tiêu đề
    for (let i = headerRowIdx + 1; i < values.length; i++) {
      const row = values[i];
      const bgRow = backgrounds[i];

      // Bỏ qua nếu dòng hoàn toàn trống
      if (!row.some(cell => cell !== "")) continue;

      const maRaw = colMap.ma !== undefined ? row[colMap.ma] : "";
      const tenRaw = colMap.ten !== undefined ? row[colMap.ten] : "";

      // NẾU CÓ MÃ HOẶC TÊN HOÀN CẢNH -> Tạo object Hoàn Cảnh mới
      if (maRaw !== "" || tenRaw !== "") {
        // Lấy link bài viết gắn vào ô tên hoàn cảnh (dòng sheet = i+1, cột = colMap.ten+1)
        const link = colMap.ten !== undefined
          ? layLinkBaiViet_(sheet, i + 1, colMap.ten + 1)
          : "";

        currentHC = {
          ma: maRaw.toString().trim() || "N/A",
          ten: tenRaw.toString().trim() || "Chưa cập nhật",
          link: link,
          khu_vuc: colMap.khuVuc !== undefined ? row[colMap.khuVuc].toString().trim() : "",
          ghi_chu: colMap.ghiChu !== undefined ? row[colMap.ghiChu].toString().trim() : "",
          mtq_list: []
        };
        sheetData.push(currentHC);
      }

      // NẾU CÓ DỮ LIỆU MẠNH THƯỜNG QUÂN -> Đẩy vào mảng mtq_list của Hoàn Cảnh hiện tại
      if (currentHC) {
        const mtqName = colMap.mtq !== undefined ? row[colMap.mtq].toString().trim() : "";

        // Bỏ qua dòng "NYP" (dòng xác nhận nhắc bài "ok", không phải Mạnh Thường Quân thật)
        if (mtqName !== "" && mtqName.toUpperCase() !== "NYP") {
          const mtq = {
            ten: mtqName,
            muc: colMap.muc !== undefined ? parseFloat(row[colMap.muc]) || 0 : 0,
            thang: {}
          };

          // Quét qua các cột tháng T1 -> T12
          for (const [mName, mIdx] of Object.entries(colMap.months)) {
            const cellVal = row[mIdx];
            const cellBg = bgRow[mIdx].toLowerCase();

            // Đánh giá: Có giá trị hoặc ô được tô màu nền (khác trắng/đen) thì tính là được hỗ trợ
            const hasValue = cellVal !== "";
            const isColored = cellBg !== "#ffffff" && cellBg !== "#000000";

            if (hasValue) {
              const value = String(cellVal).trim().toLowerCase();

              // Là số tiền
              if (!isNaN(parseFloat(value)) && parseFloat(value) > 0) {
                mtq.thang[mName] = parseFloat(value);
              }
              // Là chữ ok
              else if (value === "ok") {
                mtq.thang[mName] = "ok";
              }
              // Khác
              else {
                mtq.thang[mName] = "";
              }
            }
          }

          currentHC.mtq_list.push(mtq);
        }
      }
    }

    // Chỉ đưa vào kết quả nếu sheet đó có dữ liệu
    if (sheetData.length > 0) {
      result[sheetName] = sheetData;
    }
  });

  return result;
}

/**
 * Lấy URL link gắn vào 1 ô: ưu tiên link "chèn liên kết" (rich text),
 * nếu không có thì thử công thức =HYPERLINK("url","text")
 */
function layLinkBaiViet_(sheet, row, col) {
  try {
    const cell = sheet.getRange(row, col);
    const rich = cell.getRichTextValue();
    if (rich) {
      const url = rich.getLinkUrl();
      if (url) return url;
    }
  } catch (e) { /* bỏ qua */ }

  try {
    const congThuc = sheet.getRange(row, col).getFormula();
    if (congThuc) {
      const match = congThuc.match(/HYPERLINK\(\s*"([^"]+)"/i);
      if (match) return match[1];
    }
  } catch (e) { /* bỏ qua */ }

  return "";
}