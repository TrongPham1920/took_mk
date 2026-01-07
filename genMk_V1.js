const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const archiver = require("archiver");

const app = express();
const port = 3000;

app.use(express.static("public"));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/index.html"));
});

const upload = multer({ dest: "uploads/" });

const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

/* ================== UTIL ================== */

// normalize header excel → snake_case không dấu
const normalizeKey = (key) =>
  key
    .toString()
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");

// normalize giá tiền
const parseMoney = (val) => {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") return val;

  if (typeof val === "string") {
    const clean = val
      .trim()
      .replace(/,/g, "")
      .replace(/\./g, "")
      .replace(/\s/g, "");
    const num = Number(clean);
    return isNaN(num) ? 0 : num;
  }
  return 0;
};

// normalize toàn bộ key trong row
const normalizeKeys = (data) =>
  data.map((row) => {
    const newRow = {};
    for (const key in row) {
      newRow[normalizeKey(key)] = row[key];
    }
    return newRow;
  });

// normalize value dạng số
const toSnakeCaseNoAccent = (str) =>
  str
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

// alias dạng số
const DANG_SO_ALIAS_MAP = {
  "tien-don-3": "so-tien-don-3",
  "phat-loc": "so-phat-loc",
  "loc-phat": "so-loc-phat",
  "loc-loc": "so-loc-loc",
  "phat-phat": "so-phat-phat",
  "tien-don-5": "so-tien-don-5",
  "so-tien-giua": "tien-giua",
};

// zip file util
const zipFiles = (files, zipPath) =>
  new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver("zip", { zlib: { level: 9 } });

    output.on("close", resolve);
    archive.on("error", reject);

    archive.pipe(output);
    files.forEach((file) => archive.file(file, { name: path.basename(file) }));
    archive.finalize();
  });

/* ================== API ================== */

app.post("/split-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).send("❌ Chưa upload file");

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    let data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    if (!data.length) throw new Error("File Excel rỗng");

    const totalInputRows = data.length;
    data = normalizeKeys(data);

    const grouped = {};
    const summary = {};
    let totalProcessed = 0;

    /* ===== XỬ LÝ DATA ===== */
    data.forEach((row) => {
      const stb = row["stb"] ? row["stb"].toString().replace(/\D/g, "") : "";

      const rawDangSo = row["dang_so"];
      let dangSo =
        rawDangSo &&
        rawDangSo.toString().trim() !== "" &&
        rawDangSo.toString().toUpperCase() !== "#N/A"
          ? toSnakeCaseNoAccent(rawDangSo.toString())
          : "khong_co_dang_so";

      if (DANG_SO_ALIAS_MAP[dangSo]) {
        dangSo = DANG_SO_ALIAS_MAP[dangSo];
      }

      const cleanRow = {
        phone_number: stb,
        telco: "GMB",
        tier: "NORMAL",
        purchase_price: parseMoney(row["gia_nhap"]),
        price: parseMoney(row["gia_tho"]),
        distributor_price: parseMoney(row["gia_ban_le"]),
        plan: "",
        serial: "",
        variations: dangSo,
      };

      if (!grouped[dangSo]) grouped[dangSo] = [];
      grouped[dangSo].push(cleanRow);

      summary[dangSo] = (summary[dangSo] || 0) + 1;
      totalProcessed++;
    });

    const outFiles = [];

    /* ===== FILE THEO DẠNG SỐ ===== */
    for (const [type, rows] of Object.entries(grouped)) {
      const wb = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "DATA");

      XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.json_to_sheet([
          { thong_tin: "DẠNG SỐ", gia_tri: type },
          { thong_tin: "SỐ LƯỢNG", gia_tri: rows.length },
          { thong_tin: "TỔNG SỐ ĐÃ XỬ LÝ", gia_tri: totalProcessed },
          { thong_tin: "SỐ DÒNG FILE UPLOAD", gia_tri: totalInputRows },
        ]),
        "TONG_HOP"
      );

      const filePath = path.join(outputDir, `${type}_${Date.now()}.xlsx`);
      XLSX.writeFile(wb, filePath);
      outFiles.push(filePath);
    }

    /* ===== FILE TỔNG ===== */
    const summaryWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      summaryWb,
      XLSX.utils.json_to_sheet(
        Object.entries(summary).map(([k, v]) => ({
          dang_so: k,
          so_luong: v,
        }))
      ),
      "TONG_HOP"
    );

    const summaryFile = path.join(outputDir, `tong_hop_${Date.now()}.xlsx`);
    XLSX.writeFile(summaryWb, summaryFile);
    outFiles.push(summaryFile);

    /* ===== ZIP ===== */
    const zipPath = path.join(outputDir, `ket_qua_${Date.now()}.zip`);
    await zipFiles(outFiles, zipPath);

    fs.unlinkSync(req.file.path);

    res.download(zipPath);
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Lỗi xử lý file Excel");
  }
});

app.listen(port, () => {
  console.log(`🚀 Server chạy tại http://localhost:${port}`);
});
