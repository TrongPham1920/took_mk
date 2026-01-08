const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const port = 3000;
const upload = multer({ dest: "uploads/" });

const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

// chuẩn hóa key
const normalizeRow = (row) => {
  const obj = {};
  for (const k in row) {
    obj[k.trim().toLowerCase()] = row[k];
  }
  return obj;
};

app.post("/split-excel", upload.single("file"), (req, res) => {
  if (!req.file) return res.status(400).send("Chưa upload file");

  try {
    const wb = XLSX.readFile(req.file.path);
    const originalSheetName = wb.SheetNames[0];
    const originalSheet = wb.Sheets[originalSheetName];

    let data = XLSX.utils.sheet_to_json(originalSheet);
    data = data.map(normalizeRow);

    // 👉 TẠO DỮ LIỆU CHO SHEET nam92
    const nam92Data = data.map((row) => ({
      "Mã Vận Đơn": row["order id"] || "",
      "ID Theo Dõi": row["tracking id"] || "",
      "Ngày Đặt Hàng": row["created time"] || "",
      "Địa Chỉ": "",
      "Sản Phẩm": row["seller sku"] || row["sku id"] || "",
      "Số Lượng Trước Hủy": row["quantity"] || 0,
      "Số Lượng Sau Hủy": 0,
      "Số Lượng Cuối Cùng": row["quantity"] || 0,
      Sàn: "TIKTOK",
      Shop: "Sim Hải Đăng",
      "Doanh Thu": row["order amount"] || 0,
      "Ngày Quyết Toán": "",
      "Tình Trạng": row["order status"] || "",
    }));

    // 👉 TẠO SHEET nam92
    const nam92Sheet = XLSX.utils.json_to_sheet(nam92Data);

    // 👉 GẮN SHEET nam92 VÀO FILE GỐC
    XLSX.utils.book_append_sheet(wb, nam92Sheet, "nam92");

    // 👉 GHI FILE
    const outputFile = path.join(outputDir, `output_${Date.now()}.xlsx`);
    XLSX.writeFile(wb, outputFile);

    fs.unlinkSync(req.file.path);

    res.send(`✅ Xử lý xong!\nFile xuất ra:\n${outputFile}`);
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Lỗi xử lý file");
  }
});

app.listen(port, () => {
  console.log(`🚀 Server chạy tại http://localhost:${port}`);
});
