const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

/**
 * UTILS CHUẨN HÓA
 */
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

const getCellValue = (v) => {
  if (v && typeof v === "object" && "result" in v) return v.result;
  return v;
};

const parseMoney = (val) => {
  if (val === null || val === undefined) return 0;

  if (typeof val === "object" && val.result !== undefined) {
    return Number(val.result) || 0;
  }

  if (typeof val === "number") return val;

  const clean = val
    .toString()
    .trim()
    .replace(/[,.\s]/g, "");

  const num = Number(clean);
  return isNaN(num) ? 0 : num;
};

const formatPhone = (val) => {
  if (!val) return "";

  let phone = String(val);
  phone = phone.replace(/\D/g, "");

  if (phone.length === 9) {
    phone = "0" + phone;
  }

  if (phone.startsWith("84")) {
    phone = "0" + phone.slice(2);
  }

  return phone;
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

/**
 * HANDLER DESKTOP
 */
module.exports = async function genMb(inputPath) {
  const timestamp = Date.now();

  const baseDir = path.dirname(inputPath);
  const sessionDir = path.join(baseDir, `Ket_Qua_MMB_${timestamp}`);

  if (!fs.existsSync(sessionDir)) fs.mkdirSync(sessionDir);

  const workbooks = new Map();
  const summary = {};
  let headers = [];

  try {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(inputPath, {
      entries: "emit",
      sharedStrings: "cache",
      worksheets: "emit",
    });

    for await (const worksheet of workbookReader) {
      for await (const row of worksheet) {
        if (row.number === 1) {
          headers = row.values.map((v) => normalizeKey(v));
          continue;
        }

        const rawRow = {};
        headers.forEach((h, idx) => {
          if (h) rawRow[h] = row.values[idx];
        });

        // ===== XỬ LÝ DATA =====

        const stb = formatPhone(getCellValue(rawRow["stb"]));

        const rawDangSo = getCellValue(rawRow["dang_so"]);

        let dangSo =
          rawDangSo && rawDangSo.toString().toUpperCase() !== "#N/A"
            ? toSnakeCaseNoAccent(rawDangSo.toString())
            : "khong_co_dang_so";

        if (DANG_SO_ALIAS_MAP[dangSo]) {
          dangSo = DANG_SO_ALIAS_MAP[dangSo];
        }

        // ===== TẠO FILE THEO DẠNG SỐ =====

        if (!workbooks.has(dangSo)) {
          const cleanName = dangSo.replace(/[^a-zA-Z0-9]/g, "_");

          const filePath = path.join(sessionDir, `${cleanName}.xlsx`);

          const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({
            filename: filePath,
          });

          const sheet = workbookWriter.addWorksheet("DATA");

          sheet.columns = [
            { header: "phone_number", key: "phone_number", width: 15 },
            { header: "telco", key: "telco", width: 10 },
            { header: "tier", key: "tier", width: 10 },
            {
              header: "distributor_price",
              key: "distributor_price",
              width: 15,
            },
            { header: "price", key: "price", width: 15 },
            { header: "purchase_price", key: "purchase_price", width: 15 },
            { header: "plan", key: "plan", width: 10 },
            { header: "serial", key: "serial", width: 10 },
            { header: "variations", key: "variations", width: 20 },
          ];

          workbooks.set(dangSo, { writer: workbookWriter, sheet });
        }

        const { sheet } = workbooks.get(dangSo);

        sheet
          .addRow({
            phone_number: stb,
            telco: "GMB",
            tier: "NORMAL",
            distributor_price: parseMoney(getCellValue(rawRow["gia_ban_le"])),
            price: parseMoney(getCellValue(rawRow["gia_tho"])),
            purchase_price: parseMoney(getCellValue(rawRow["gia_nhap"])),
            plan: "",
            serial: "",
            variations: dangSo,
          })
          .commit();

        summary[dangSo] = (summary[dangSo] || 0) + 1;
      }
    }

    // ===== COMMIT FILE =====

    for (const [name, entry] of workbooks) {
      await entry.writer.commit();
    }

    // ===== FILE TỔNG HỢP =====

    const summaryFile = path.join(sessionDir, "00_tong_hop.xlsx");

    const sWriter = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: summaryFile,
    });

    const sSheet = sWriter.addWorksheet("TONG HOP");

    sSheet.columns = [
      { header: "dang_so", key: "k" },
      { header: "so_luong", key: "v" },
    ];

    Object.entries(summary).forEach(([k, v]) =>
      sSheet.addRow({ k, v }).commit(),
    );

    await sWriter.commit();

    console.log("✔ Xuất file thành công:", sessionDir);

    return sessionDir;
  } catch (err) {
    console.error("❌ Lỗi genMb:", err);
    throw err;
  }
};
