import forge from "node-forge";
import CryptoJS from "crypto-js";

/* ===================== CONFIG INPUT ===================== */
const INPUT = {
  // Thông tin cá nhân mới
  sex: "1", // Nữ
  name: "DANH THƯƠNG",
  id: "094090010207",
  issue_date: "10-07-2021",
  issue_by: "CỤC TRƯỞNG CỤC CẢNH SÁT QUẢN LÝ HÀNH CHÍNH VỀ TRẬT TỰ XÃ HỘI",
  birthday: "12-09-1990",
  city: "94",
  district: "943",
  ward: "31528",
  address: "Ấp Kinh Giữa 1 Kế Thành, Kế Sách, Sóc Trăng",

  // SIM
  phone: "779599557",
  serialSim: "8406250412223466",
  // YYYY-MM-DD
  expiry: "2030-09-12",

  // Session & eKYC
  customerId: "VNS5160001466293",
  contractNo: "u6ZUBZnZ2KmIJcTujrNByA",

  id_ekyc: "18a8a78a-8c2d-4f9a-a9dc-f6026d8784f4",
  sessionToken: "8DC8E89EB4E44F34B74CEE68F5-1768986929357",

  // Decree 13 Accept (Mảng 6 phần tử)
  decree13Accept: ["DK1", "DK2", "DK3", "DK4", "DK5", "DK6"],

  // Public Key mới
  publicKey:
    "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAmxk8muAKwr0Dccah4OLLaKD7KJKFBojccPVZNFSzU5uZHawi+e6RoBLyjeFXvZJ+QLeL/VN3AGVNBMkpS1ZW9NPg0T0SBfNDHLXKF0epaOK52EMeTaPVOaQgk0dGFBrSzbiRX5SzsBKWhAEVNC4wTrbIEe4oEFet2iL3sIaJnzXnUSt2q4oVlFWVg57QTRL7OVyn+wWT2aHHAotzH7IYX0pUId5JUkeEkSGlV0yEc5v+o57hEUz33x2nGTekw1TJF8aw4+/mh5ZtOIAb/qK04eHNQrhHK+XIA87eEy6yQvQLK+ziFI+UfaSmoWzEjlMPc8Jla+666rGRd/fDJ6xy9wIDAQAB".trim(),

  document: "1",
};

/* ===================== UTILS ===================== */

function formatToStrictDMY(dateStr, type) {
  if (!dateStr) return "";
  let d, m, y;

  if (type === "YYYY-MM-DD") {
    [y, m, d] = dateStr.split("-");
  } else if (type === "DD-MM-YYYY") {
    [d, m, y] = dateStr.split("-");
  } else {
    throw new Error("Invalid date format type");
  }

  return `${d.padStart(2, "0")}/${m.padStart(2, "0")}/${y}`;
}

function getTodayDMY() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
}

/* ===================== ENCRYPT CORE ===================== */
function encryptVnsky(payload, publicKeyBase64) {
  const aesKey = forge.random.getBytesSync(32);
  const iv = forge.random.getBytesSync(12);

  const cipher = forge.cipher.createCipher("AES-GCM", aesKey);
  cipher.start({ iv, tagLength: 128 });
  cipher.update(forge.util.createBuffer(payload, "utf8"));
  cipher.finish();

  const encryptedData = forge.util.encode64(
    cipher.output.getBytes() + cipher.mode.tag.getBytes(),
  );

  const publicKeyPem = `-----BEGIN PUBLIC KEY-----\n${publicKeyBase64}\n-----END PUBLIC KEY-----`;
  const publicKey = forge.pki.publicKeyFromPem(publicKeyPem);

  const encryptedAESKey = forge.util.encode64(
    publicKey.encrypt(aesKey, "RSAES-PKCS1-V1_5"),
  );

  return {
    encryptedData,
    encryptedAESKey,
    iv: forge.util.encode64(iv),
  };
}

/* ===================== BUILD PAYLOAD ===================== */
const payloadObj = {
  request: {
    strSex: INPUT.sex,
    strSubName: INPUT.name,
    strIdNo: INPUT.id,
    strIdIssueDate: formatToStrictDMY(INPUT.issue_date, "DD-MM-YYYY"),
    strIdIssuePlace: INPUT.issue_by,
    strBirthday: formatToStrictDMY(INPUT.birthday, "DD-MM-YYYY"),
    strProvince: INPUT.city,
    strDistrict: INPUT.district,
    strPrecinct: INPUT.ward,
    strHome: INPUT.address,
    strAddress: INPUT.address,
    strContractNo: INPUT.contractNo,
    strIsdn: INPUT.phone,
    strSerial: INPUT.serialSim,
  },
  idExpiryDate: formatToStrictDMY(INPUT.expiry, "YYYY-MM-DD"),
  idType: INPUT.document,
  idEkyc: INPUT.id_ekyc,
  customerCode: INPUT.customerId,
  contractDate: getTodayDMY(),
  decree13Accept: INPUT.decree13Accept?.toString(), // 🔥 CHUẨN
  sessionToken: INPUT.sessionToken,
  signature: CryptoJS.MD5(INPUT.id_ekyc + INPUT.sessionToken).toString(), // 🔥 CHUẨN
};

/* ===================== EXECUTE ===================== */
const payloadString = JSON.stringify(payloadObj, null, 2);
const encrypted = encryptVnsky(payloadString, INPUT.publicKey);

console.log("--- JSON DỮ LIỆU TRƯỚC MÃ HÓA ---");
console.log(payloadString);
console.log("---------------------------------");
console.log("👉 DÁN VÀO POSTMAN (KEY 'data'):");
console.log(JSON.stringify(encrypted));
