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

// ===== UTIL =====
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

const normalizeKeys = (data) =>
  data.map((row) => {
    const newRow = {};
    for (const key in row) {
      newRow[key.toString().trim().toLowerCase()] = row[key];
    }
    return newRow;
  });

// ===== API =====
app.post("/split-excel", upload.single("file"), (req, res) => {
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

    // ===== XỬ LÝ =====
    data.forEach((row) => {
      const stb = row["stb"] ? row["stb"].toString().replace(/\D/g, "") : "";

      const rawDangSo = row["dạng số"];
      const dangSo =
        rawDangSo &&
        rawDangSo.toString().trim() !== "" &&
        rawDangSo.toString().toUpperCase() !== "#N/A"
          ? toSnakeCaseNoAccent(rawDangSo.toString())
          : "khong_co_dang_so";

      const cleanRow = {
        phone_number: stb,
        telco: "GMB",
        tier: "NORMAL",
        distributor_price: row["giá bán lẻ"] || 0,
        price: row["giá thợ"] || 0,
        purchase_price: row["giá nhập"] || 0,
        plan: "",
        serial: "",
        variations: dangSo,
      };

      if (!grouped[dangSo]) grouped[dangSo] = [];
      grouped[dangSo].push(cleanRow);

      summary[dangSo] = (summary[dangSo] || 0) + 1;
      totalProcessed++;
    });

    console.log("=== BẮT ĐẦU XUẤT FILE ===");

    const outFiles = [];

    // ===== XUẤT FILE CON =====
    for (const [type, rows] of Object.entries(grouped)) {
      console.log(`${type}: ${rows.length} số`);

      const wb = XLSX.utils.book_new();

      const wsData = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, wsData, "DATA");

      const wsSummary = XLSX.utils.json_to_sheet([
        { thong_tin: "DẠNG SỐ", gia_tri: type },
        { thong_tin: "SỐ LƯỢNG", gia_tri: rows.length },
        { thong_tin: "TỔNG SỐ ĐÃ XỬ LÝ", gia_tri: totalProcessed },
        { thong_tin: "SỐ DÒNG FILE UPLOAD", gia_tri: totalInputRows },
      ]);
      XLSX.utils.book_append_sheet(wb, wsSummary, "TONG_HOP");

      const filePath = path.join(outputDir, `${type}_${Date.now()}.xlsx`);
      XLSX.writeFile(wb, filePath);
      outFiles.push(filePath);
    }

    // ===== FILE TỔNG RIÊNG =====
    const summaryRows = [];

    for (const [dangSo, count] of Object.entries(summary)) {
      summaryRows.push({
        dang_so: dangSo,
        so_luong: count,
      });
    }

    summaryRows.push({
      dang_so: "TONG_SO_DA_XU_LY",
      so_luong: totalProcessed,
    });

    summaryRows.push({
      dang_so: "SO_DONG_FILE_UPLOAD",
      so_luong: totalInputRows,
    });

    const summaryWb = XLSX.utils.book_new();
    const wsTong = XLSX.utils.json_to_sheet(summaryRows, {
      header: ["dang_so", "so_luong"],
    });
    XLSX.utils.book_append_sheet(summaryWb, wsTong, "TONG_HOP");

    const summaryFilePath = path.join(
      outputDir,
      `tong_hop_dang_so_${Date.now()}.xlsx`
    );

    XLSX.writeFile(summaryWb, summaryFilePath);
    outFiles.push(summaryFilePath);

    console.log("=== KẾT THÚC ===");
    console.log(`Tổng số đã xử lý: ${totalProcessed}`);
    console.log(`Số dòng file upload: ${totalInputRows}`);

    fs.unlinkSync(req.file.path);

    res.send(
      `✅ XỬ LÝ THÀNH CÔNG\n\n` +
        `Số dòng file upload: ${totalInputRows}\n` +
        `Tổng số đã xử lý: ${totalProcessed}\n\n` +
        `File đã tạo:\n${outFiles.join("\n")}`
    );
  } catch (err) {
    console.error(err);
    res.status(500).send("❌ Lỗi xử lý file Excel");
  }
});

app.listen(port, () => {
  console.log(`🚀 Server chạy tại http://localhost:${port}`);
});
