var SPREADSHEET_ID = "1pp75BhhqmqeHLf1NY9FdCRo4IlfE8V-VqZA6dzEiClo";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Cảng Sài Gòn | Kết nối | Chia sẻ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- 1. TRA CỨU KHÁCH / CÔNG VỤ (Form_Responses & NhatKyRaVao) ---
function traCuuCCCD(cccd, forForm) {
  try {
    var searchKey = String(cccd || '').replace(/['\s]+/g, '').trim();

    if (!searchKey) return { found: false, error: "Chưa nhập số CCCD!" };
    if (!/^\d{12}$/.test(searchKey)) return { found: false, error: "Số CCCD phải có đúng 12 chữ số!" };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetForm = ss.getSheetByName("Form_Responses");
    var sheetLog = ss.getSheetByName("NhatKyRaVao");
    
    if (!sheetForm) return { found: false, error: "Không tìm thấy tab 'Form_Responses' trong Google Sheet!" };

    var data = sheetForm.getDataRange().getValues();
    var workerInfo = null;

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var cccdInSheet = String(row[5] || '').replace(/['\s]+/g, '').trim();

      if (cccdInSheet !== "" && cccdInSheet === searchKey) {
        var ngayKhaiBao = new Date();
        if (row[0] instanceof Date && !isNaN(row[0].getTime())) {
          ngayKhaiBao = row[0];
        } else if (row[0]) {
          var parsedDate = new Date(row[0]);
          if (!isNaN(parsedDate.getTime())) ngayKhaiBao = parsedDate;
        }

        var rawThoiHan = row[7]; 
        var ngayHetHan = null;
        var thoiHanStr = "Chưa cập nhật";
        var khongNhapThoiHan = false;

        if (rawThoiHan) {
          if (rawThoiHan instanceof Date && !isNaN(rawThoiHan.getTime())) {
            ngayHetHan = new Date(rawThoiHan.getTime());
          } else {
            var parts = String(rawThoiHan).split('/');
            if (parts.length === 3) {
              ngayHetHan = new Date(parts[2], parts[1] - 1, parts[0]);
            } else {
              var d = new Date(rawThoiHan);
              if (!isNaN(d.getTime())) ngayHetHan = d;
            }
          }
        }

        if (ngayHetHan && !isNaN(ngayHetHan.getTime())) {
          thoiHanStr = Utilities.formatDate(ngayHetHan, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          khongNhapThoiHan = true;
          ngayHetHan = new Date(ngayKhaiBao.getTime());
          ngayHetHan.setMonth(ngayHetHan.getMonth() + 6);
        }

        ngayHetHan.setHours(23, 59, 59, 999);
        var homNay = new Date();

        if (!forForm && homNay > ngayHetHan) {
          var thongBaoLoi = khongNhapThoiHan 
            ? "Bạn CHƯA CẬP NHẬT thời hạn thẻ an toàn và đã quá 6 tháng kể từ ngày khai báo ban đầu. Vui lòng bấm khai báo lại!" 
            : "Thẻ an toàn của bạn đã HẾT HẠN (Hạn đến ngày: " + thoiHanStr + "). Vui lòng khai báo/cập nhật lại!";
          return { found: false, error: thongBaoLoi };
        }

        var diffTime = ngayHetHan.getTime() - homNay.getTime();
        var diffDays = Math.ceil(diffTime / (1000 * 3600 * 24));

        var thoiHanInputDate = "";
        if (ngayHetHan && !khongNhapThoiHan) {
          thoiHanInputDate = Utilities.formatDate(ngayHetHan, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }

        workerInfo = {
          found: true,
          hoTen: row[1] || "",               
          donVi: row[2] || "",               
          chucVu: row[3] || "",              
          lyDo: row[4] || "",                
          cccd: cccdInSheet,                 
          soDienThoai: row[6] || "",         
          thoiHanTheAnToan: thoiHanStr,
          thoiHanInputDate: thoiHanInputDate,
          canhBaoSapHetHan: (diffDays <= 30 && diffDays >= 0),
          soNgayConLai: diffDays,
          khongNhapThoiHan: khongNhapThoiHan,
          dangTrongCang: false
        };
        break;
      }
    }

    if (!workerInfo) return { found: false };

    if (sheetLog) {
      var logData = sheetLog.getDataRange().getValues();
      for (var j = logData.length - 1; j >= 1; j--) {
        var logCCCD = String(logData[j][4] || '').replace(/['\s]+/g, '').trim();
        if (logCCCD === searchKey) {
          var trangThaiGanNhat = String(logData[j][9] || '').trim();
          if (trangThaiGanNhat === "VÀO CỔNG") {
            workerInfo.dangTrongCang = true;
          }
          break;
        }
      }
    }

    return workerInfo;
  } catch (err) {
    return { found: false, error: "Lỗi Server: " + err.toString() };
  }
}

function ghiNhatKyVaoCong(info) {
  try {
    var searchKey = String(info.cccd || '').replace(/['\s]+/g, '').trim();
    if (!/^\d{12}$/.test(searchKey)) return "LỖI: Số CCCD phải có đúng 12 chữ số!";

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetLog = ss.getSheetByName("NhatKyRaVao");
    if (!sheetLog) return "LỖI: Không tìm thấy tab 'NhatKyRaVao'!";

    var logData = sheetLog.getDataRange().getValues();

    for (var j = logData.length - 1; j >= 1; j--) {
      var logCCCD = String(logData[j][4] || '').replace(/['\s]+/g, '').trim();
      if (logCCCD === searchKey) {
        var trangThai = String(logData[j][9] || '').trim();
        if (trangThai === "VÀO CỔNG") {
          return "LỖI: Bạn chưa xác nhận RA CỔNG lượt trước! Phải RA CỔNG trước khi xác nhận VÀO CỔNG lượt mới.";
        }
        break;
      }
    }

    var now = new Date();
    var ngayVao = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var gioVao = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

    var nextRow = logData.length + 1;
    for (var k = 1; k < logData.length; k++) {
      if (String(logData[k][0] || '').trim() === "") {
        nextRow = k + 1;
        break;
      }
    }

    sheetLog.getRange(nextRow, 1, 1, 10).setValues([[
      info.hoTen,                               
      info.donVi,                               
      info.chucVu,                              
      info.lyDo,                                
      "'" + info.cccd,                          
      "'" + info.soDienThoai,                   
      info.thoiHanTheAnToan || "Chưa cập nhật",  
      ngayVao + " " + gioVao,
      "",
      "VÀO CỔNG"
    ]]);

    return "XÁC NHẬN VÀO CẢNG THÀNH CÔNG!\nThời gian: " + gioVao + " - " + ngayVao;
  } catch (err) {
    return "LỖI: " + err.toString();
  }
}

function ghiNhatKyRaCong(info) {
  try {
    var searchKey = String(info.cccd || '').replace(/['\s]+/g, '').trim();
    if (!/^\d{12}$/.test(searchKey)) return "LỖI: Số CCCD phải có đúng 12 chữ số!";

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetLog = ss.getSheetByName("NhatKyRaVao");
    if (!sheetLog) return "LỖI: Không tìm thấy tab 'NhatKyRaVao'!";

    var logData = sheetLog.getDataRange().getValues();
    var targetRow = -1;

    for (var j = logData.length - 1; j >= 1; j--) {
      var logCCCD = String(logData[j][4] || '').replace(/['\s]+/g, '').trim();
      if (logCCCD === searchKey) {
        var trangThai = String(logData[j][9] || '').trim();
        if (trangThai === "VÀO CỔNG") {
          targetRow = j + 1;
          break;
        }
      }
    }

    if (targetRow === -1) {
      return "LỖI: Không tìm thấy dữ liệu lượt VÀO hoặc RA của bạn!";
    }

    var now = new Date();
    var ngayRa = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var gioRa = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");

    sheetLog.getRange(targetRow, 9).setValue(ngayRa + " " + gioRa);
    sheetLog.getRange(targetRow, 10).setValue("RA CỔNG");

    return "XÁC NHẬN RA CỔNG THÀNH CÔNG!\nThời gian: " + gioRa + " - " + ngayRa;
  } catch (err) {
    return "LỖI: " + err.toString();
  }
}

function luuKhaiBaoMoi(data) {
  try {
    var inputCCCD = String(data.cccd || '').replace(/['\s]+/g, '').trim();
    var inputPhone = String(data.soDienThoai || '').replace(/['\s]+/g, '').trim();

    if (!/^\d{12}$/.test(inputCCCD)) {
      return { success: false, error: "Số CCCD phải có đúng 12 chữ số!" };
    }
    if (!/^\d{10}$/.test(inputPhone)) {
      return { success: false, error: "Số điện thoại phải có đúng 10 chữ số!" };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetForm = ss.getSheetByName("Form_Responses");

    if (!sheetForm) return { success: false, error: "Không tìm thấy tab 'Form_Responses'!" };

    var now = new Date();
    var timeStamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    var rawThoiHan = data.thoiHanThe || data.thoiHanTheAnToan || "";
    var thoiHanFormatted = "";

    if (data.lyDo !== "VENDOR") {
      thoiHanFormatted = "";
    } else if (rawThoiHan) {
      if (rawThoiHan.indexOf('-') !== -1) {
        var parts = rawThoiHan.split('-');
        if (parts.length === 3) thoiHanFormatted = parts[2] + "/" + parts[1] + "/" + parts[0];
      } else if (rawThoiHan.indexOf('/') !== -1) {
        thoiHanFormatted = rawThoiHan;
      }
    }

    var allData = sheetForm.getDataRange().getValues();
    var targetRow = -1;
    var trangThaiKhaiBao = "KHAI BÁO MỚI";

    for (var i = 1; i < allData.length; i++) {
      var cccdInSheet = String(allData[i][5] || '').replace(/['\s]+/g, '').trim();
      if (cccdInSheet !== "" && cccdInSheet === inputCCCD) {
        targetRow = i + 1;
        trangThaiKhaiBao = "CẬP NHẬT";

        if (data.lyDo === "VENDOR") {
          var oldThoiHanRaw = allData[i][7];
          var oldNgayHetHan = null;

          if (oldThoiHanRaw) {
            if (oldThoiHanRaw instanceof Date && !isNaN(oldThoiHanRaw.getTime())) {
              oldNgayHetHan = new Date(oldThoiHanRaw.getTime());
            } else {
              var parts = String(oldThoiHanRaw).split('/');
              if (parts.length === 3) {
                oldNgayHetHan = new Date(parts[2], parts[1] - 1, parts[0]);
              } else {
                var d = new Date(oldThoiHanRaw);
                if (!isNaN(d.getTime())) oldNgayHetHan = d;
              }
            }
          }

          if (oldNgayHetHan && !isNaN(oldNgayHetHan.getTime())) {
            oldNgayHetHan.setHours(23, 59, 59, 999);
            var homNay = new Date();

            if (homNay <= oldNgayHetHan) {
              thoiHanFormatted = Utilities.formatDate(oldNgayHetHan, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
          }
        }
        break;
      }
    }

    if (targetRow === -1) targetRow = allData.length + 1;

    sheetForm.getRange(targetRow, 1, 1, 9).setValues([[
      timeStamp, data.hoTen, data.donVi, data.chucVu, data.lyDo, 
      "'" + inputCCCD, "'" + inputPhone, thoiHanFormatted, trangThaiKhaiBao        
    ]]);

    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// --- 2. CÁC HÀM XỬ LÝ DÀNH RIÊNG CHO THUYỀN VIÊN (RaVaoThuyenVien) ---

// Tra cứu thông tin thuyền viên lần gần nhất từ nhật ký RaVaoThuyenVien
function traCuuThuyenVien(cccd) {
  try {
    var searchKey = String(cccd || '').replace(/['\s]+/g, '').trim();
    if (!searchKey) return { found: false, error: "Chưa nhập số giấy tờ!" };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetLog = ss.getSheetByName("RaVaoThuyenVien");

    if (!sheetLog) return { found: false, error: "Không tìm thấy tab 'RaVaoThuyenVien'!" };

    var logData = sheetLog.getDataRange().getValues();
    var crewInfo = null;

    // Tìm dòng mới nhất của thuyền viên theo CCCD/Passport
    for (var j = logData.length - 1; j >= 1; j--) {
      var logCCCD = String(logData[j][3] || '').replace(/['\s]+/g, '').trim();
      if (logCCCD === searchKey) {
        var thoiGianVao = String(logData[j][4] || '').trim();
        var thoiGianRa = String(logData[j][5] || '').trim();

        crewInfo = {
          found: true,
          hoTen: logData[j][0] || "",
          tenTau: logData[j][1] || "",
          quocTich: logData[j][2] || "",
          cccd: logCCCD,
          dangTrongCang: (thoiGianVao !== "" && thoiGianRa === "")
        };
        break;
      }
    }

    if (!crewInfo) return { found: false };
    return crewInfo;
  } catch (err) {
    return { found: false, error: "Lỗi Server: " + err.toString() };
  }
}

// Ghi nhận thuyền viên VÀO CỔNG trực tiếp
function ghiNhatKyThuyenVienVaoCong(info) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetLog = ss.getSheetByName("RaVaoThuyenVien");
    if (!sheetLog) return "LỖI: Không tìm thấy tab 'RaVaoThuyenVien'!";

    var searchKey = String(info.cccd || '').replace(/['\s]+/g, '').trim();
    var logData = sheetLog.getDataRange().getValues();

    // Kiểm tra xem thuyền viên này có lượt VÀO mà CHƯA RA hay không
    for (var j = logData.length - 1; j >= 1; j--) {
      var logCCCD = String(logData[j][3] || '').replace(/['\s]+/g, '').trim();
      if (logCCCD === searchKey) {
        var thoiGianVao = String(logData[j][4] || '').trim();
        var thoiGianRa = String(logData[j][5] || '').trim();
        if (thoiGianVao !== "" && thoiGianRa === "") {
          return "LỖI: Thuyền viên đang ở trong cảng! Phải xác nhận RA CỔNG trước khi vào lại.";
        }
        break;
      }
    }

    var now = new Date();
    var ngayGioVao = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    var nextRow = logData.length + 1;
    for (var k = 1; k < logData.length; k++) {
      if (String(logData[k][0] || '').trim() === "") {
        nextRow = k + 1;
        break;
      }
    }

    // Ghi dòng mới vào RaVaoThuyenVien: [Họ tên, Tên tàu, Quốc tịch, Số giấy tờ, Thời gian vào, Thời gian ra]
    sheetLog.getRange(nextRow, 1, 1, 6).setValues([[
      info.hoTen,
      info.tenTau,
      info.quocTich,
      "'" + searchKey,
      ngayGioVao,
      ""
    ]]);

    return "XÁC NHẬN VÀO CẢNG THÀNH CÔNG!\nThời gian: " + ngayGioVao;
  } catch (err) {
    return "LỖI: " + err.toString();
  }
}

// Ghi nhận thuyền viên RA CỔNG
function ghiNhatKyThuyenVienRaCong(info) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetLog = ss.getSheetByName("RaVaoThuyenVien");
    if (!sheetLog) return "LỖI: Không tìm thấy tab 'RaVaoThuyenVien'!";

    var searchKey = String(info.cccd || '').replace(/['\s]+/g, '').trim();
    var logData = sheetLog.getDataRange().getValues();
    var targetRow = -1;

    for (var j = logData.length - 1; j >= 1; j--) {
      var logCCCD = String(logData[j][3] || '').replace(/['\s]+/g, '').trim();
      if (logCCCD === searchKey) {
        var thoiGianVao = String(logData[j][4] || '').trim();
        var thoiGianRa = String(logData[j][5] || '').trim();
        if (thoiGianVao !== "" && thoiGianRa === "") {
          targetRow = j + 1;
          break;
        }
      }
    }

    if (targetRow === -1) {
      return "LỖI: Không tìm thấy lượt VÀO CỔNG đang hoạt động của thuyền viên này!";
    }

    var now = new Date();
    var ngayGioRa = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    // Cập nhật lại Họ tên, Tên tàu, Quốc tịch nếu có chỉnh sửa lúc bấm Ra Cổng
    sheetLog.getRange(targetRow, 1, 1, 3).setValues([[info.hoTen, info.tenTau, info.quocTich]]);
    sheetLog.getRange(targetRow, 6).setValue(ngayGioRa);

    return "XÁC NHẬN RA CỔNG THÀNH CÔNG!\nThời gian: " + ngayGioRa;
  } catch (err) {
    return "LỖI: " + err.toString();
  }
}