const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const os = require("os"); // Thêm thư viện hệ điều hành
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

function getProductName(sku) {
  const s = String(sku || "")
    .toUpperCase()
    .trim();
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

  // SỬA TẠI ĐÂY: Lưu vào thư mục Temp của hệ điều hành
  const tempDir = os.tmpdir();
  const outputFile = path.join(tempDir, `bao_cao_tong_hop_${timestamp}.xlsx`);

  const workbook = new ExcelJS.Workbook();

  // ================== 1. KHỞI TẠO SHEET TỔNG ==================
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

  const orderMap = new Map();

  const processAndSaveSheet = async (fileObj) => {
    const result = await readFileAuto(fileObj);
    if (!result || !result.data || result.data.length === 0) return null;
    const fileName = path.parse(fileObj.originalname).name;
    const sheetName = `${fileName}`
      .substring(0, 31)
      .replace(/[\\\/\?\*\[\]:]/g, "");
    const ws = workbook.addWorksheet(sheetName);
    result.data.forEach((row) => {
      const values = normalizeValues(row.values);
      if (values.length > 0) ws.addRow(values);
    });
    return result;
  };

  const res1 = await processAndSaveSheet(fileObj1);
  const res2 = await processAndSaveSheet(fileObj2);
  if (!res1 || !res2) throw new Error("Could not read data from files.");

  // ================== 2. NHẬN DIỆN FILE ==================
  let dataTongDon = (res1.data[0]?.values || [])
    .map(normalizeHeader)
    .includes("seller_sku")
    ? res1
    : res2;
  let dataDoanhThu = dataTongDon === res1 ? res2 : res1;

  // ================== 3. XỬ LÝ FILE TỔNG ĐƠN ==================
  const headersTD = normalizeValues(dataTongDon.data[0]?.values || []).map(
    normalizeHeader
  );

  for (let i = 2; i < dataTongDon.data.length; i++) {
    const rowValues = normalizeValues(dataTongDon.data[i].values || []);
    if (rowValues.length === 0) continue;
    const record = {};
    headersTD.forEach((h, idx) => {
      if (h) record[h] = rowValues[idx];
    });

    const sku = String(record.seller_sku || "").trim();
    const productName = getProductName(sku);

    if (productName !== null) {
      const orderId = String(record.order_id || "").trim();
      if (!orderId) continue;
      const mapKey = `${orderId}_${sku}`;
      orderMap.set(mapKey, {
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

  // ================== 4. XỬ LÝ FILE DOANH THU ==================
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

    const relatedId = String(record.related_order_id || "").trim();
    const skuDT = String(record.seller_sku || record.sku_id || "").trim();
    const mapKey = `${relatedId}_${skuDT}`;

    if (orderMap.has(mapKey)) {
      const item = orderMap.get(mapKey);
      item.created_time = record.order_created_time || "";
      item.settled_time = record.order_settled_time || "";
      item.settlement_amount = Number(record.total_settlement_amount || 0);
      item.revenue = Number(record.total_revenue || 0);
    }
  }

  // ================== 5. GHI DỮ LIỆU TỔNG HỢP ==================
  orderMap.forEach((item) => summary.addRow(item));

  // Ghi file vào đường dẫn Temp
  await workbook.xlsx.writeFile(outputFile);

  // Trả về đường dẫn tuyệt đối cho main.js
  return outputFile;
};
