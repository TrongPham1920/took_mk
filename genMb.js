const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const port = 3000;

// setup multer
const upload = multer({ dest: "uploads/" });

// tạo folder xuất file nếu chưa có
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

// ĐỊNH NGHĨA THỨ TỰ CỘT MONG MUỐN (ĐÃ KHÔI PHỤC "Giá bán")
const columnOrder = ["stb", "stb gốc", "Giá bán", "Giá khuyến mãi", "DẠNG SỐ"];

// --- HÀM TIỆN ÍCH CHUẨN HÓA TÊN CỘT ---
const normalizeKeys = (data) => {
  return data.map((row) => {
    const newRow = {};
    for (const key in row) {
      if (Object.prototype.hasOwnProperty.call(row, key)) {
        // Chuyển key (tên cột) sang chữ thường và loại bỏ khoảng trắng thừa
        const normalizedKey = key.toString().trim().toLowerCase();
        newRow[normalizedKey] = row[key];
      }
    }
    return newRow;
  });
};
// ----------------------------------------

app.post("/split-excel", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).send("Chưa upload file");

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    let data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // 1. CHUẨN HÓA TÊN CỘT: chuyển tất cả tên cột về chữ thường
    data = normalizeKeys(data);

    const grouped = {};
    data.forEach((row) => {
      // Vì tên cột đã được chuẩn hóa thành chữ thường, ta sử dụng key chữ thường:
      const originalStb = row["stb"] ? row["stb"].toString() : "";

      // 1a. Chuẩn hóa stb (đã dùng key chữ thường)
      let stb_clean = row["stb"]
        ? row["stb"].toString().replace(/\D/g, "")
        : "";

      // 1b. Chuẩn hóa stb gốc
      let stb_goc_clean = originalStb ? originalStb.replace(/\D/g, "") : "";

      // 1c. LẤY GIÁ BÁN TỪ KEY CHỮ THƯỜNG ĐÃ CHUẨN HÓA
      // Đảm bảo bắt được cả "giá bán lẻ" hoặc "giá bán" từ file gốc
      let gia_ban_clean = row["giá bán lẻ"] || row["giá bán"] || "";

      // 1d. Chuẩn hóa DẠNG SỐ (đã dùng key chữ thường)
      let dang_so = row["dạng số"]
        ? row["dạng số"].toString().toUpperCase()
        : "UNKNOWN";

      // 2. TẠO ĐỐI TƯỢNG SẠCH
      const cleanRow = {
        // Tên cột đầu ra vẫn giữ chữ hoa/chữ thường theo yêu cầu (columnOrder)
        stb: stb_clean,
        "stb gốc": stb_goc_clean,
        "Giá bán": gia_ban_clean, // Sử dụng giá trị đã lấy từ key chữ thường
        "Giá khuyến mãi": 0,
        "DẠNG SỐ": dang_so,
      };

      // 3. Gom nhóm
      if (!grouped[dang_so]) grouped[dang_so] = [];
      grouped[dang_so].push(cleanRow);
    });

    // tạo workbook mới
    const newWorkbook = XLSX.utils.book_new();
    for (const [type, rows] of Object.entries(grouped)) {
      // Sử dụng 'header' với tên cột gốc (không chuẩn hóa) để định dạng file xuất ra
      const ws = XLSX.utils.json_to_sheet(rows, { header: columnOrder });
      XLSX.utils.book_append_sheet(newWorkbook, ws, type.substring(0, 31));
    }

    // ghi file ra folder output
    const outFile = path.join(outputDir, `output_${Date.now()}.xlsx`);
    XLSX.writeFile(newWorkbook, outFile);

    // xóa file tạm upload
    fs.unlinkSync(req.file.path);

    res.send(`Xử lý xong! File xuất ra: ${outFile}`);
  } catch (err) {
    console.error(err);
    res.status(500).send("Xử lý file lỗi");
  }
});

app.listen(port, () => {
  console.log(`Server chạy tại http://localhost:${port}`);
});
