const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const archiver = require("archiver");

const genMk = require("./handlers/genMk.js");
const genMb = require("./handlers/genMb.js");

const app = express();
const port = 3000;

/* ================== STATIC ================== */
app.use(express.static("public"));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/index.html"));
});

/* ================== UPLOAD ================== */
const upload = multer({ dest: "uploads/" });

const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");

/* ================== ROUTER ================== */
app.post("/split-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).send("❌ Chưa upload file");

  const mode = req.body.mode;

  if (!mode) {
    fs.unlinkSync(req.file.path);
    return res.status(400).send("❌ Thiếu mode xử lý");
  }

  try {
    if (mode === "beauty") {
      return await genMk(req, res);
    }

    if (mode === "mmb") {
      return await genMb(req, res);
    }

    throw new Error("Mode không hợp lệ");
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Lỗi xử lý file");
  } finally {
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
  }
});

/* ================== START ================== */
app.listen(port, () => {
  console.log(`🚀 Server chạy tại http://localhost:${port}`);
});
