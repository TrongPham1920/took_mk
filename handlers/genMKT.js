const ExcelJS = require("exceljs");
const path = require("path");
const readFileAuto = require("../utils/readFileAuto");

// Hàm chuẩn hóa tiêu đề để dễ so sánh
function normalizeHeader(h) {
  return String(h || "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[\/]/g, "_");
}

module.exports = async (fileObj1, fileObj2) => {
  const timestamp = Date.now();
  const outputFile = path.join(
    __dirname,
    "../output",
    `merged_${timestamp}.xlsx`
  );
  const outputWorkbook = new ExcelJS.Workbook();

  // 1. Đọc dữ liệu từ cả 2 file
  const res1 = await readFileAuto(fileObj1);
  const res2 = await readFileAuto(fileObj2);

  if (!res1 || !res2) throw new Error("Không thể đọc dữ liệu từ file upload.");

  let dataTongDon = null;
  let dataDoanhThu = null;

  // 2. LOGIC NHẬN DIỆN FILE THÔNG MINH
  // Kiểm tra file nào có cột 'seller_sku' -> đó là file Tổng đơn
  // Kiểm tra file nào có cột 'related_order_id' -> đó là file Doanh thu
  const h1 = (res1.data[0]?.values || []).map(normalizeHeader);
  const h2 = (res2.data[0]?.values || []).map(normalizeHeader);

  if (h1.includes("seller_sku") || h1.includes("order_id")) {
    dataTongDon = res1;
    dataDoanhThu = res2;
  } else {
    dataTongDon = res2;
    dataDoanhThu = res1;
  }

  // 3. Cấu hình Sheet TỔNG HỢP
  const summary = outputWorkbook.addWorksheet("TONG_HOP");
  summary.columns = [
    { header: "Order ID", key: "order_id", width: 25 },
    { header: "Order Status", key: "status", width: 20 },
    { header: "Seller SKU", key: "sku", width: 25 },
    { header: "Quantity", key: "qty", width: 10 },
    { header: "Sku Quantity of return", key: "qty_return", width: 15 },
    { header: "Cancel Reason", key: "cancel_reason", width: 25 },
    { header: "Tracking ID", key: "tracking_id", width: 25 },
    { header: "Package ID", key: "package_id", width: 25 },
    { header: "Order created time", key: "created_time", width: 20 },
    { header: "Order settled time", key: "settled_time", width: 20 },
    { header: "Total settlement amount", key: "settlement_amount", width: 15 },
    { header: "Total Revenue", key: "revenue", width: 15 },
  ];

  const orderMap = new Map();

  // --- BƯỚC 4: XỬ LÝ FILE TỔNG ĐƠN ---
  const headersTD = (dataTongDon.data[0]?.values || []).map(normalizeHeader);
  for (let i = 1; i < dataTongDon.data.length; i++) {
    const row = dataTongDon.data[i].values;
    if (!row) continue;

    const record = {};
    headersTD.forEach((h, idx) => {
      if (h) record[h] = row[idx];
    });

    const orderId = String(record.order_id || "").trim();
    if (!orderId) continue;

    orderMap.set(orderId, {
      order_id: orderId,
      status: record.order_status || "",
      sku: record.seller_sku || "",
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

  // --- BƯỚC 5: XỬ LÝ FILE DOANH THU ---
  const headersDT = (dataDoanhThu.data[0]?.values || []).map(normalizeHeader);
  for (let i = 1; i < dataDoanhThu.data.length; i++) {
    const row = dataDoanhThu.data[i].values;
    if (!row) continue;

    const record = {};
    headersDT.forEach((h, idx) => {
      if (h) record[h] = row[idx];
    });

    const relatedOrderId = String(record.related_order_id || "").trim();

    if (orderMap.has(relatedOrderId)) {
      const existing = orderMap.get(relatedOrderId);
      existing.created_time = record.order_created_time || "";
      existing.settled_time = record.order_settled_time || "";
      existing.settlement_amount = Number(record.total_settlement_amount || 0);
      existing.revenue = Number(record.total_revenue || 0);
    }
  }

  // --- BƯỚC 6: ĐỔ DỮ LIỆU VÀ ĐỊNH DẠNG ---
  orderMap.forEach((item) => summary.addRow(item));

  // Làm đẹp Header
  summary.getRow(1).font = { bold: true };
  summary.getRow(1).alignment = { vertical: "middle", horizontal: "center" };

  // Định dạng số cho cột tiền
  summary.getColumn("settlement_amount").numFmt = "#,##0";
  summary.getColumn("revenue").numFmt = "#,##0";

  await outputWorkbook.xlsx.writeFile(outputFile);
  return outputFile;
};
