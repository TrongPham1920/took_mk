const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");

const genMKT = require("./handlers/genMKTV1.js");

const app = express();
const port = 3000;

const uploadsDir = path.join(__dirname, "uploads");
const outputDir = path.join(__dirname, "output");

if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

app.use(express.static("public"));

// Cấu hình Multer để nhận 2 file từ giao diện
const upload = multer({ dest: "uploads/" }).fields([
  { name: "file1", maxCount: 1 }, // Đây sẽ là File Tổng Đơn
  { name: "file2", maxCount: 1 }, // Đây sẽ là File Doanh Thu
]);

// Route chính để xử lý gộp file
app.post("/merge-excel", upload, async (req, res) => {
  try {
    if (!req.files?.file1 || !req.files?.file2) {
      return res
        .status(400)
        .send(
          "❌ Lỗi: Bạn cần chọn đầy đủ cả File Tổng Đơn và File Doanh Thu."
        );
    }

    const outputFile = await genMKT(req.files.file1[0], req.files.file2[0]);

    res.download(outputFile, "bao_cao_tong_hop_doi_soat.xlsx", (err) => {
      if (err) {
        console.error("Lỗi khi gửi file:", err);
      }

      if (fs.existsSync(outputFile)) fs.unlinkSync(outputFile);
    });
  } catch (err) {
    console.error("Lỗi Server:", err);
    res.status(500).send(`❌ Lỗi xử lý: ${err.message}`);
  } finally {
    if (req.files) {
      const files = Object.values(req.files).flat();
      files.forEach((f) => {
        if (fs.existsSync(f.path)) {
          fs.unlinkSync(f.path);
        }
      });
    }
  }
});

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/indexWebV1.html"));
});

app.listen(port, () => {
  console.log(`--------------------------------------------------`);
  console.log(`🚀 Server đang chạy tại: http://localhost:${port}`);
  console.log(`📂 Thư mục tạm: ${uploadsDir}`);
  console.log(`📂 Thư mục kết quả: ${outputDir}`);
  console.log(`--------------------------------------------------`);
});
