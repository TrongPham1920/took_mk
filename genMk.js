const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const archiver = require("archiver");

const app = express();
const port = process.env.PORT || 3000;

app.use(express.static("public"));

const upload = multer({ dest: "uploads/" });
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

/* ================== UTILS ================== */

// Dùng dấu phẩy làm dấu phân cách chuẩn CSV
const DELIMITER = ",";

const escapeCSV = (val) => {
  if (val === undefined || val === null) return "";
  let str = val.toString();
  // Nếu dữ liệu có dấu phẩy hoặc nháy kép, bọc lại bằng nháy kép để không lỗi cột
  if (str.includes(DELIMITER) || str.includes('"') || str.includes("\n")) {
    str = '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
};

const normalizeKey = (key) =>
  key
    ? key
        .toString()
        .replace(/đ/g, "d")
        .replace(/Đ/g, "D")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim()
        .replace(/[^a-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "")
    : "";

const parseMoney = (val) => {
  if (!val) return 0;
  if (typeof val === "number") return val;
  const clean = val.toString().replace(/[,.\s]/g, "");
  return isNaN(Number(clean)) ? 0 : Number(clean);
};

const toSnakeCaseNoAccent = (str) =>
  str
    ? str
        .replace(/đ/g, "d")
        .replace(/Đ/g, "D")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "")
    : "khong_co_dang_so";

const DANG_SO_ALIAS_MAP = {
  "tien-don-3": "so-tien-don-3",
  "phat-loc": "so-phat-loc",
  "loc-phat": "so-loc-phat",
  "loc-loc": "so-loc-loc",
  "phat-phat": "so-phat-phat",
  "tien-don-5": "so-tien-don-5",
  "so-tien-giua": "tien-giua",
};

/* ================== CHÍNH: XỬ LÝ STREAM ================== */

app.post("/split-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).send("❌ Chưa upload file");

  const timestamp = Date.now();
  const sessionDir = path.join(outputDir, `session_${timestamp}`);
  fs.mkdirSync(sessionDir);

  const summary = {};
  const writers = new Map();
  let headers = [];

  try {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
      req.file.path,
      {
        entries: "emit",
        sharedStrings: "cache",
        worksheets: "emit",
      }
    );

    for await (const worksheet of workbookReader) {
      for await (const row of worksheet) {
        if (row.number === 1) {
          headers = row.values.map((v) => normalizeKey(v));
          continue;
        }

        const rawRow = {};
        headers.forEach((h, idx) => {
          rawRow[h] = row.values[idx];
        });

        const rawDangSo = rawRow["dang_so"];
        let dangSo =
          rawDangSo && rawDangSo.toString().toUpperCase() !== "#N/A"
            ? toSnakeCaseNoAccent(rawDangSo.toString())
            : "khong_co_dang_so";

        if (DANG_SO_ALIAS_MAP[dangSo]) dangSo = DANG_SO_ALIAS_MAP[dangSo];

        if (!writers.has(dangSo)) {
          const csvPath = path.join(sessionDir, `${dangSo}.csv`);
          const writeStream = fs.createWriteStream(csvPath);

          // MẸO QUAN TRỌNG: Thêm "sep=," để ép Excel tách cột
          writeStream.write(`sep=${DELIMITER}\n`);
          const head = [
            "phone_number",
            "telco",
            "tier",
            "purchase_price",
            "price",
            "distributor_price",
            "plan",
            "serial",
            "variations",
          ];
          writeStream.write(head.map(escapeCSV).join(DELIMITER) + "\n");

          writers.set(dangSo, writeStream);
        }

        const stb = rawRow["stb"]
          ? rawRow["stb"].toString().replace(/\D/g, "")
          : "";

        const columns = [
          stb,
          "GMB",
          "NORMAL",
          parseMoney(rawRow["gia_nhap"]),
          parseMoney(rawRow["gia_tho"]),
          parseMoney(rawRow["gia_ban_le"]),
          "",
          "",
          dangSo,
        ];

        writers
          .get(dangSo)
          .write(columns.map(escapeCSV).join(DELIMITER) + "\n");
        summary[dangSo] = (summary[dangSo] || 0) + 1;
      }
    }

    // Đóng các stream
    for (const [name, stream] of writers) {
      stream.end();
    }

    // Ghi file tổng hợp với lệnh sep=,
    const summaryPath = path.join(sessionDir, `00_tong_hop.csv`);
    let summaryContent = `sep=${DELIMITER}\ndang_so${DELIMITER}so_luong\n`;
    Object.entries(summary).forEach(([k, v]) => {
      summaryContent += `${escapeCSV(k)}${DELIMITER}${escapeCSV(v)}\n`;
    });
    fs.writeFileSync(summaryPath, summaryContent);

    /* ===== ZIP & SEND ===== */
    const zipPath = path.join(outputDir, `ket_qua_${timestamp}.zip`);
    const outputZip = fs.createWriteStream(zipPath);
    const archive = archiver("zip", { zlib: { level: 5 } });

    res.on("finish", () => {
      setTimeout(() => {
        try {
          if (fs.existsSync(sessionDir))
            fs.rmSync(sessionDir, { recursive: true, force: true });
          if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
          if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
        } catch (e) {}
      }, 5000);
    });

    archive.pipe(outputZip);
    archive.directory(sessionDir, false);
    await archive.finalize();

    res.download(zipPath);
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Lỗi xử lý dữ liệu.");
  }
});

app.listen(port, () => console.log(`🚀 Server running on port ${port}`));
