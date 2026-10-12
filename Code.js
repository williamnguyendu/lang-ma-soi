const SHEET_ID = '1-H_mbJpwwLQvzLsjuF2ZG1lQaMlV0eXm0nIkICHphg8';

// Xử lý request GET để lấy thông tin sự kiện
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('LichLangMaSoi');
    const data = sheet.getDataRange().getValues();
    data.shift(); // Bỏ qua dòng tiêu đề
    
    let latestEvent = null;
    let maxDate = 0;

    // Lọc các sự kiện "Mở đăng ký" và lấy sự kiện có NgayToChuc lớn nhất
    data.forEach(row => {
      const trangThai = row[5]; 
      if (trangThai === 'Mở đăng ký') {
        const dateTimestamp = new Date(row[1]).getTime();
        if (dateTimestamp > maxDate) {
          maxDate = dateTimestamp;
          latestEvent = {
            maSuKien: row[0],
            ngayToChuc: dateTimestamp, // Trả về timestamp để Front-end dễ so sánh và format
            diaDiem: row[2],
            dangKyTu: new Date(row[3]).getTime(),
            dangKyDen: new Date(row[4]).getTime()
          };
        }
      }
    });

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: latestEvent
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Xử lý request POST để lưu thông tin đăng ký
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('DangKyLangMaSoi');
    
    // Lấy data từ request
    const hoVaTen = e.parameter.hoVaTen;
    const soDienThoai = e.parameter.soDienThoai;
    
    // Format thời gian đăng ký (Giờ Việt Nam GMT+7)
    const thoiGianDangKy = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    // Ghi vào sheet
    sheet.appendRow([thoiGianDangKy, hoVaTen, soDienThoai]);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Đăng ký thành công"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}