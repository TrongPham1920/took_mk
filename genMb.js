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

// THỨ TỰ CỘT ĐẦU RA
const columnOrder = ["stb", "stb gốc", "Giá bán", "Giá khuyến mãi", "DẠNG SỐ"];

// --- HÀM TIỆN ÍCH CHUẨN HÓA TÊN CỘT ---
const normalizeKeys = (data) => {
  return data.map((row) => {
    const newRow = {};
    for (const key in row) {
      const normalizedKey = key.toString().trim().toLowerCase();
      newRow[normalizedKey] = row[key];
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

    // 1. Chuẩn hóa key về chữ thường
    data = normalizeKeys(data);

    const grouped = {};

    data.forEach((row) => {
      let stb = row["stb"] ? row["stb"].toString().replace(/\D/g, "") : "";
      let stbGoc = row["stb"] ? row["stb"].toString().replace(/\D/g, "") : "";

      let giaBan = row["giá bán lẻ"] || row["giá bán"] || "";

      let dangSo = row["dạng số"]
        ? row["dạng số"].toString().toUpperCase()
        : "UNKNOWN";

      const cleanRow = {
        stb,
        "stb gốc": stbGoc,
        "Giá bán": giaBan,
        "Giá khuyến mãi": 0,
        "DẠNG SỐ": dangSo,
      };

      if (!grouped[dangSo]) grouped[dangSo] = [];
      grouped[dangSo].push(cleanRow);
    });

    // --- XUẤT MỖI DẠNG SỐ THÀNH 1 FILE RIÊNG ---
    let outFiles = [];

    for (const [type, rows] of Object.entries(grouped)) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(rows, { header: columnOrder });

      XLSX.utils.book_append_sheet(wb, ws, type.substring(0, 31));

      // Tên file an toàn
      const cleanName = type.replace(/[^a-zA-Z0-9]/g, "_");

      const filePath = path.join(outputDir, `${cleanName}_${Date.now()}.xlsx`);

      XLSX.writeFile(wb, filePath);
      outFiles.push(filePath);
    }

    // Xóa file upload tạm
    fs.unlinkSync(req.file.path);

    res.send(`Xong! Đã tạo ${outFiles.length} file:\n\n` + outFiles.join("\n"));
  } catch (err) {
    console.error(err);
    res.status(500).send("Xử lý file lỗi");
  }
});

app.listen(port, () => {
  console.log(`Server chạy tại http://localhost:${port}`);
});
