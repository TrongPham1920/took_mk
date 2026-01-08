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

const upload = multer({ dest: "uploads/" }).fields([
  { name: "file1", maxCount: 1 },
  { name: "file2", maxCount: 1 },
]);

app.post("/merge-excel", upload, async (req, res) => {
  try {
    if (!req.files?.file1 || !req.files?.file2) {
      return res.status(400).send("❌ Cần upload đủ 2 file");
    }

    const outputFile = await genMKT(req.files.file1[0], req.files.file2[0]);

    res.download(outputFile);
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Lỗi xử lý file");
  } finally {
    Object.values(req.files || {})
      .flat()
      .forEach((f) => {
        if (fs.existsSync(f.path)) fs.unlinkSync(f.path);
      });
  }
});

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/indexWebV1.html"));
});

app.listen(port, () => {
  console.log(`🚀 Server gộp báo cáo chạy tại http://localhost:${port}`);
});
