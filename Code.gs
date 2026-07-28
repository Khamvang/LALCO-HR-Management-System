function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('LALCO HR Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function formatDateText(rawDate) {
  if (rawDate instanceof Date) {
    var year = rawDate.getFullYear();
    var month = ("0" + (rawDate.getMonth() + 1)).slice(-2);
    var day = ("0" + rawDate.getDate()).slice(-2);
    return year + "-" + month + "-" + day;
  }
  return rawDate || "";
}

function formatDateTimeText(date) {
  if (!(date instanceof Date)) date = new Date(date);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);
  var hours = ("0" + date.getHours()).slice(-2);
  var minutes = ("0" + date.getMinutes()).slice(-2);
  var seconds = ("0" + date.getSeconds()).slice(-2);
  return year + "-" + month + "-" + day + " " + hours + ":" + minutes + ":" + seconds;
}

function asText(val) {
  if (val === "" || val === null || val === undefined) return "";
  return "'" + val.toString().trim();
}

function uploadToDrive(base64Data, folderName, fileName) {
  if (!base64Data || base64Data.indexOf('base64,') === -1) return "";
  var splitData = base64Data.split('base64,');
  var contentType = splitData[0].split(';')[0].replace('data:', '');
  var byteCharacters = Utilities.base64Decode(splitData[1]);
  var blob = Utilities.newBlob(byteCharacters, contentType, fileName);
  
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function processForm(fd) {
  var today = new Date();
  today.setHours(23, 59, 59, 999); // allow the whole current day
  var issueFieldsToCheck = [
    { key: 'idIssue', label: 'ວັນອອກບັດ (ID Issue Date)' },
    { key: 'famIssue', label: 'ວັນອອກສຳມະໂນຄົວ (Family Book Issue Date)' },
    { key: 'passIssue', label: 'ວັນທີ່ອອກພາດສະປອດ (Passport Issue Date)' }
  ];
  for (var k = 0; k < issueFieldsToCheck.length; k++) {
    var rawVal = fd[issueFieldsToCheck[k].key];
    if (rawVal) {
      var checkDate = new Date(rawVal);
      if (!isNaN(checkDate.getTime()) && checkDate.getTime() > today.getTime()) {
        return "⚠️ " + issueFieldsToCheck[k].label + " ຕ້ອງນ້ອຍກວ່າ ຫຼື ເທົ່າກັບມື້ນີ້! ກະລຸນາແກ້ໄຂແລ້ວບັນທຶກໃໝ່.";
      }
    }
  }

  // ກຸ່ມເອກະສານຢັ້ງຢືນຕົວຕົນ — ຖ້າໃສ່ຫົວຂໍ້ໃດໜຶ່ງໃນກຸ່ມ ຕ້ອງໃສ່ໃຫ້ຄົບທຸກຫົວຂໍ້ໃນກຸ່ມນັ້ນ
  // ຢ່າງໜ້ອຍຕ້ອງມີ 1 ກຸ່ມທີ່ຄົບຖ້ວນ ຈຶ່ງບັນທຶກໄດ້
  function isFieldFilled(v) {
    return v !== undefined && v !== null && v.toString().trim() !== "";
  }
  function checkDocGroup(fieldKeys) {
    var filledCount = 0;
    for (var i = 0; i < fieldKeys.length; i++) {
      if (isFieldFilled(fd[fieldKeys[i]])) filledCount++;
    }
    if (filledCount === 0) return "empty";
    if (filledCount === fieldKeys.length) return "complete";
    return "partial";
  }

  var idCardStatus = checkDocGroup(['idCard', 'idName', 'idIssue', 'idExp']);
  var famBookStatus = checkDocGroup(['famBook', 'famIssue']);
  var passportStatus = checkDocGroup(['passport', 'passName', 'passIssue', 'passExp']);

  if (idCardStatus === "partial") {
    return "⚠️ ຂໍ້ມູນ ID Card (ບັດປະຈຳຕົວ) ບໍ່ຄົບ! ກະລຸນາປ້ອນ ID Card Number, Name on ID Card, ID Issue Date, ແລະ ID Expiration Date ໃຫ້ຄົບທຸກຢ່າງ.";
  }
  if (famBookStatus === "partial") {
    return "⚠️ ຂໍ້ມູນ Family Book (ສຳມະໂນຄົວ) ບໍ່ຄົບ! ກະລຸນາປ້ອນ Family Book ID ແລະ Family Book Issue Date ໃຫ້ຄົບທຸກຢ່າງ.";
  }
  if (passportStatus === "partial") {
    return "⚠️ ຂໍ້ມູນ Passport (ໜັງສືຜ່ານແດນ) ບໍ່ຄົບ! ກະລຸນາປ້ອນ Passport Number, Name on Passport, Passport Issue Date, ແລະ Passport Expiration Date ໃຫ້ຄົບທຸກຢ່າງ.";
  }
  if (idCardStatus !== "complete" && famBookStatus !== "complete" && passportStatus !== "complete") {
    return "⚠️ ກະລຸນາປ້ອນຂໍ້ມູນຢັ້ງຢືນຕົວຕົນຢ່າງໜ້ອຍ 1 ຢ່າງໃຫ້ຄົບຖ້ວນ: ID Card, Family Book, ຫຼື Passport! ຈຶ່ງສາມາດບັນທຶກໄດ້.";
  }

  var PHONE_LENGTH_RULES = {
    "+856": [8, 10], "+66": [9, 9], "+84": [9, 10], "+855": [8, 9], "+95": [7, 10],
    "+86": [11, 11], "+82": [9, 10], "+81": [9, 10], "+65": [8, 8], "+60": [9, 10],
    "+63": [10, 10], "+91": [10, 10], "+1": [10, 10], "+44": [10, 10], "+61": [9, 9]
  };
  if (fd.phone) {
    var phoneVal = fd.phone.toString().trim();
    var phoneMatch = phoneVal.match(/^(\+\d{1,4})(\d+)$/);
    if (!phoneMatch) {
      return "⚠️ ຮູບແບບເບີໂທບໍ່ຖືກຕ້ອງ! ກະລຸນາເລືອກລະຫັດປະເທດ ແລະ ປ້ອນເບີໂທເປັນຕົວເລກເທົ່ານັ້ນ.";
    }
    var phoneCode = phoneMatch[1];
    var phoneDigits = phoneMatch[2];
    if (phoneCode === "+856") {
      var isMobile = /^20\d{8}$/.test(phoneDigits);
      var isLandline = /^30\d{7}$/.test(phoneDigits);
      var isOffice = /^\d{8}$/.test(phoneDigits);
      if (!isMobile && !isLandline && !isOffice) {
        return "⚠️ ຮູບແບບເບີໂທລາວບໍ່ຖືກຕ້ອງ! ຕ້ອງເປັນ: ມືຖື (20+8ໂຕເລກ), ຕັ້ງໂຕະ (30+7ໂຕເລກ), ຫຼື ຫ້ອງການ/ອື່ນໆ (8ໂຕເລກ).";
      }
    } else {
      var phoneRule = PHONE_LENGTH_RULES[phoneCode];
      if (phoneRule && (phoneDigits.length < phoneRule[0] || phoneDigits.length > phoneRule[1])) {
        return "⚠️ ເບີໂທບໍ່ຄົບ ຫຼື ຍາວເກີນໄປ ສຳລັບລະຫັດ " + phoneCode + "! ກະລຸນາກວດເບີໂທຄືນ.";
      }
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Staff_List");
  
  if (!sheet) {
    sheet = ss.insertSheet("Staff_List");
    sheet.appendRow([
      "ID", "Nickname", "First Name Lao", "Last Name Lao", "First Name En", "Last Name En",
      "Phone", "Branch", "Department", "Unit", "Position", "Hire Date", "Basic Salary", "DOB",
      "Gender", "Nationality", "Marital Status", 
      "Birth Province", "Birth District", "Birth Village", "Current Province", "Current District", "Current Village",
      "Education Level", 
      "School", "Major", "GPA", "Driver License", "ID Card", "ID Card Name", "ID Issue Date", "ID Expire Date",
      "Family Book", "Family Book Date", "Passport No", "Passport Name", "Passport Issue Date", "Passport Expire", 
      "Photo URL", "CV File URL", "Contract File URL", "Created At", "Updated At", "Application Source", "Referrer Name",
      "Status", "Resign Date", "Resign Reason", "Resign Letter URL",
      "Bank Account Name", "Bank Account Number", "Currency", "Bank Book Photo URL"
    ]);
  }
  
  var data = sheet.getDataRange().getValues();
  var id = fd.staffId;
  var foundRow = -1;
  var createdAt = formatDateTimeText(new Date());
  var updatedAt = formatDateTimeText(new Date());

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() == id) {
      foundRow = i + 1;
      if(data[i][41]) {
        createdAt = data[i][41] instanceof Date ? formatDateTimeText(data[i][41]) : data[i][41];
      }
      break;
    }
  }

  var finalPhotoUrl = fd.oldPhotoUrl; 
  if (fd.photoBase64 && fd.photoBase64.indexOf('data:image') !== -1) {
    var splitBase = fd.photoBase64.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var bytes = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(bytes, type, id + "_profile.jpg");
    var imgFolder = DriveApp.getFoldersByName("LALCO_HR_Profiles").hasNext() ? DriveApp.getFoldersByName("LALCO_HR_Profiles").next() : DriveApp.createFolder("LALCO_HR_Profiles");
    imgFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var file = imgFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    finalPhotoUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800";
  }

  var finalCvUrl = fd.cvBase64 && fd.cvFileName ? uploadToDrive(fd.cvBase64, "LALCO_HR_CVs", id + "_CV_" + fd.cvFileName) : fd.oldCvUrl;
  var finalContractUrl = fd.contractBase64 && fd.contractFileName ? uploadToDrive(fd.contractBase64, "LALCO_HR_Contracts", id + "_Contract_" + fd.contractFileName) : fd.oldContractUrl;
  var finalResignLetterUrl = fd.resignLetterBase64 && fd.resignLetterFileName ? uploadToDrive(fd.resignLetterBase64, "LALCO_HR_ResignLetters", id + "_ResignLetter_" + fd.resignLetterFileName) : fd.oldResignLetterUrl;
  var finalBankBookUrl = fd.bankBookBase64 && fd.bankBookFileName ? uploadToDrive(fd.bankBookBase64, "LALCO_HR_BankBooks", id + "_BankBook_" + fd.bankBookFileName) : fd.oldBankBookUrl;

  var rowData = [
    asText(id), fd.nickname, fd.firstNameLao, fd.lastNameLao, fd.firstNameEn, fd.lastNameEn,
    asText(fd.phone), fd.branch, fd.department, fd.unit, fd.position, fd.hireDate, fd.salary, fd.dob,
    fd.gender, fd.nationality, fd.maritalStatus,
    fd.pBirth, fd.dBirth, fd.vBirth, fd.pCurr, fd.dCurr, fd.vCurr,
    fd.eduLevel,
    fd.school, fd.major, fd.gpa, asText(fd.driverId),
    asText(fd.idCard), fd.idName, fd.idIssue, fd.idExp,
    asText(fd.famBook), fd.famIssue, asText(fd.passport), fd.passName, fd.passIssue, fd.passExp, 
    finalPhotoUrl, finalCvUrl, finalContractUrl,
    asText(createdAt), asText(updatedAt), fd.source, fd.referrer,
    fd.statusSelect, fd.resignDate, fd.resignReason, finalResignLetterUrl,
    fd.bankAccountName, asText(fd.bankAccountNumber), fd.currency, finalBankBookUrl
  ];

  if (foundRow > -1) {
    sheet.getRange(foundRow, 1, 1, rowData.length).setValues([rowData]);
    return "✅ ອັບເດດຂໍ້ມູນພະນັກງານລະຫັດ " + id + " ສຳເລັດ!";
  } else {
    sheet.appendRow(rowData);
    return "🎉 ບັນທຶກພະນັກງານໃໝ່ສຳເລັດ!";
  }
}

function getStaffById(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff_List");
  if (!sheet) throw new Error("ບໍ່ພົບໜ້າ Sheet 'Staff_List'"); 
  
  var data = sheet.getDataRange().getValues();
  var searchId = id.toString().trim(); 
  
  for (var i = 1; i < data.length; i++) {
    var rowId = data[i][0] ? data[i][0].toString().trim() : ""; 
    if (rowId === searchId) {
      return {
        id: data[i][0], nickname: data[i][1], firstNameLao: data[i][2], lastNameLao: data[i][3], firstNameEn: data[i][4], lastNameEn: data[i][5],
        phone: data[i][6], branch: data[i][7], department: data[i][8], unit: data[i][9], position: data[i][10], hireDate: formatDateText(data[i][11]), salary: data[i][12], dob: formatDateText(data[i][13]),
        gender: data[i][14], nationality: data[i][15], maritalStatus: data[i][16],
        pBirth: data[i][17], dBirth: data[i][18], vBirth: data[i][19], pCurr: data[i][20], dCurr: data[i][21], vCurr: data[i][22],
        eduLevel: data[i][23], school: data[i][24], major: data[i][25], gpa: data[i][26], driverId: data[i][27],
        idCard: data[i][28], idName: data[i][29], idIssue: formatDateText(data[i][30]), idExp: formatDateText(data[i][31]),
        famBook: data[i][32], famIssue: formatDateText(data[i][33]), passport: data[i][34], passName: data[i][35], passIssue: formatDateText(data[i][36]), passExp: formatDateText(data[i][37]),
        photoUrl: data[i][38], cvUrl: data[i][39], contractUrl: data[i][40], source: data[i][43], referrer: data[i][44],
        status: data[i][45] || "Active", resignDate: formatDateText(data[i][46]), resignReason: data[i][47], resignLetterUrl: data[i][48],
        bankAccountName: data[i][49], bankAccountNumber: data[i][50], currency: data[i][51], bankBookUrl: data[i][52] 
      };
    }
  }
  return null; 
}

function getAllStaff() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff_List");
  return sheet.getDataRange().getDisplayValues();
}

function getStaffListForDropdown() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff_List");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues(); 
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var id = data[i][0].toString().trim();
    var nickname = data[i][1].toString().trim();
    var nameLao = data[i][2].toString().trim();
    if (id !== "") {
       var display = id;
       if (nickname !== "") display += " - " + nickname;
       else if (nameLao !== "") display += " - " + nameLao;
       list.push(display);
    }
  }
  return list;
}

function getUnitListForDropdown() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff_List");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 10, lastRow - 1, 1).getDisplayValues(); 
  var unitsSet = {};
  for (var i = 0; i < data.length; i++) {
    var unit = data[i][0].toString().trim();
    if (unit !== "") unitsSet[unit] = true;
  }
  return Object.keys(unitsSet).sort();
}

function getEducationDataForDatalist() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff_List");
  if (!sheet) return { schools: [], majors: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { schools: [], majors: [] };
  var data = sheet.getRange(2, 25, lastRow - 1, 2).getDisplayValues(); 
  var schoolsSet = {};
  var majorsSet = {};
  for (var i = 0; i < data.length; i++) {
    var school = data[i][0].toString().trim();
    var major = data[i][1].toString().trim();
    if (school !== "") schoolsSet[school] = true;
    if (major !== "") majorsSet[major] = true;
  }
  return { schools: Object.keys(schoolsSet).sort(), majors: Object.keys(majorsSet).sort() };
}

// 📌 ຟັງຊັນໃໝ່ສຳລັບດຶງຂໍ້ມູນ ສາຂາ, ພະແນກ ແລະ ຕຳແໜ່ງ
function getWorkDataForDatalist() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Staff_List");
  if (!sheet) return { branches: [], departments: [], positions: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { branches: [], departments: [], positions: [] };
  
  // ດຶງຂໍ້ມູນຖັນທີ 8 (Branch) ເຖິງຖັນທີ 11 (Position)
  var data = sheet.getRange(2, 8, lastRow - 1, 4).getDisplayValues(); 
  var branchSet = {};
  var deptSet = {};
  var posSet = {};
  
  for (var i = 0; i < data.length; i++) {
    var branch = data[i][0].toString().trim(); // ຖັນ Branch
    var dept = data[i][1].toString().trim();   // ຖັນ Department
    var pos = data[i][3].toString().trim();    // ຖັນ Position (Index 3 ໃນ Range)
    
    if (branch !== "") branchSet[branch] = true;
    if (dept !== "") deptSet[dept] = true;
    if (pos !== "") posSet[pos] = true;
  }
  
  return { 
    branches: Object.keys(branchSet).sort(), 
    departments: Object.keys(deptSet).sort(), 
    positions: Object.keys(posSet).sort() 
  };
}

function getProvinces() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Address_Data");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var provinces = {};
  for (var i = 0; i < data.length; i++) {
    var p = data[i][0].toString().trim();
    if (p !== "") provinces[p] = true;
  }
  return Object.keys(provinces).sort();
}

function getDistricts(province) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Address_Data");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var districts = {};
  for (var i = 0; i < data.length; i++) {
    var p = data[i][0].toString().trim();
    var d = data[i][1].toString().trim();
    if (p === province && d !== "") districts[d] = true;
  }
  return Object.keys(districts).sort();
}

function getVillages(province, district) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Address_Data");
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var villages = {};
  for (var i = 0; i < data.length; i++) {
    var p = data[i][0].toString().trim();
    var d = data[i][1].toString().trim();
    var v = data[i][2].toString().trim();
    if (p === province && d === district && v !== "") villages[v] = true;
  }
  return Object.keys(villages).sort();
}
