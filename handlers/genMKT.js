const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const os = require("os");
const readFileAuto = require("../utils/readFileAuto");

function normalizeHeader(h) {
  return String(h || "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[\/]/g, "_");
}

function normalizeValues(values) {
  if (!Array.isArray(values)) return [];
  return values[0] == null ? values.slice(1) : values;
}

function normalizeSheetName(name) {
  if (!name) return "";
  let n = name.toLowerCase();
  n = n.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  n = n.replace(/[\s\\\/\?\*\[\]:]/g, "_");
  return n.substring(0, 31);
}

function getProductName(sku) {
  const s = String(sku || "")
    .toUpperCase()
    .trim();
  if (!s) return null;

  if (s.includes("VL350") || s.includes("E350")) return "MDT350";
  if (s.includes("VL255") || s.includes("E255")) return "MDT255";
  if (s.includes("12A79S")) return "12A79S";
  if (s.includes("A69")) return "A69";
  if (s.includes("A79")) return "A79";
  if (s.includes("A89")) return "A89";
  if (s.includes("A99")) return "A99";
  if (s.includes("A119")) return "A119";
  if (s.startsWith("HVN")) return "SIM DU LỊCH - HI VIỆT NAM";
  if (s.includes("6MDT") || s.includes("6H")) return "6MDT150";
  if (s.includes("12MDT") || s.includes("12H")) return "12MDT150";

  return null;
}

module.exports = async (fileObj1, fileObj2) => {
  const timestamp = Date.now();
  const tempDir = os.tmpdir();
  const outputFile = path.join(tempDir, `report_${timestamp}.xlsx`);
  const workbook = new ExcelJS.Workbook();

  // ================== 1. KHỞI TẠO SHEET TỔNG ĐẦU TIÊN ==================
  const summary = workbook.addWorksheet("TONG_HOP_DOI_SOAT");
  summary.columns = [
    { header: "Order ID", key: "order_id", width: 25 },
    { header: "Product Name", key: "product_name", width: 25 },
    { header: "Order Status", key: "status", width: 20 },
    { header: "Seller SKU", key: "sku", width: 25 },
    { header: "Quantity", key: "qty", width: 10 },
    { header: "Sku Quantity of return", key: "qty_return", width: 20 },
    { header: "Cancel Reason", key: "cancel_reason", width: 25 },
    { header: "Tracking ID", key: "tracking_id", width: 25 },
    { header: "Package ID", key: "package_id", width: 25 },
    { header: "Order created time", key: "created_time", width: 20 },
    { header: "Order settled time", key: "settled_time", width: 20 },
    { header: "Total settlement amount", key: "settlement_amount", width: 22 },
    { header: "Total Revenue", key: "revenue", width: 22 },
  ];
  summary.getRow(1).font = { bold: true };
  summary.getColumn("settlement_amount").numFmt = "#,##0";
  summary.getColumn("revenue").numFmt = "#,##0";

  // ================== 2. ĐỌC DỮ LIỆU VÀ TẠO SHEET PHỤ (Sẽ nằm sau sheet tổng) ==================
  const res1 = await readFileAuto(fileObj1);
  const res2 = await readFileAuto(fileObj2);
  if (!res1 || !res2) throw new Error("Could not read data from files.");

  const processAndSaveSheet = (result, fileObj) => {
    const fileName = path.parse(fileObj.originalname).name;
    const sheetName = normalizeSheetName(fileName);
    const ws = workbook.addWorksheet(sheetName);
    result.data.forEach((row) => {
      const values = normalizeValues(row.values);
      if (values.length > 0) ws.addRow(values);
    });
  };

  processAndSaveSheet(res1, fileObj1);
  processAndSaveSheet(res2, fileObj2);

  // ================== 3. NHẬN DIỆN FILE ==================
  const h1 = (res1.data[0]?.values || []).map(normalizeHeader);
  const isRes1DoanhThu =
    h1.includes("related_order_id") || h1.includes("total_settlement_amount");

  let dataTongDon = isRes1DoanhThu ? res2 : res1;
  let dataDoanhThu = isRes1DoanhThu ? res1 : res2;

  const orderMap = new Map();

  // ================== 4. XỬ LÝ FILE TỔNG ĐƠN ==================
  const headersTD = normalizeValues(dataTongDon.data[0]?.values || []).map(
    normalizeHeader
  );

  for (let i = 1; i < dataTongDon.data.length; i++) {
    const rowValues = normalizeValues(dataTongDon.data[i].values || []);
    if (rowValues.length === 0) continue;

    const record = {};
    headersTD.forEach((h, idx) => {
      if (h) record[h] = rowValues[idx];
    });

    const sku = String(record.seller_sku || "").trim();
    const productName = getProductName(sku);
    const orderId = String(record.order_id || "").trim();

    if (productName && orderId) {
      if (orderMap.has(orderId)) {
        const existing = orderMap.get(orderId);
        existing.qty += Number(record.quantity || 0);
        existing.qty_return += Number(record.sku_quantity_of_return || 0);
      } else {
        orderMap.set(orderId, {
          order_id: orderId,
          product_name: productName,
          status: record.order_status || "",
          sku: sku,
          qty: Number(record.quantity || 0),
          qty_return: Number(record.sku_quantity_of_return || 0),
          cancel_reason: record.cancel_reason || "",
          tracking_id: record.tracking_id || "",
          package_id: record.package_id || "",
          created_time: "",
          settled_time: "",
          settlement_amount: 0,
          revenue: 0,
        });
      }
    }
  }

  // ================== 5. XỬ LÝ FILE DOANH THU ==================
  const headersDT = normalizeValues(dataDoanhThu.data[0]?.values || []).map(
    normalizeHeader
  );

  for (let i = 1; i < dataDoanhThu.data.length; i++) {
    const rowValues = normalizeValues(dataDoanhThu.data[i].values || []);
    if (rowValues.length === 0) continue;

    const record = {};
    headersDT.forEach((h, idx) => {
      if (h) record[h] = rowValues[idx];
    });

    const relatedId = String(
      record.related_order_id || record.order_id || ""
    ).trim();

    if (orderMap.has(relatedId)) {
      const existing = orderMap.get(relatedId);
      if (record.order_created_time)
        existing.created_time = record.order_created_time;
      if (record.order_settled_time)
        existing.settled_time = record.order_settled_time;

      existing.settlement_amount += Number(record.total_settlement_amount || 0);
      existing.revenue += Number(record.total_revenue || 0);
    }
  }

  // ================== 6. GHI DỮ LIỆU VÀO SHEET TỔNG ==================
  orderMap.forEach((item) => summary.addRow(item));

  await workbook.xlsx.writeFile(outputFile);
  return outputFile;
};
