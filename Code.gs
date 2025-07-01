// #####################################################
// #                 Code.gs (Backend)                 #
// #####################################################

function doPost(e) {
  let response = { success: false, message: 'Invalid request' };
  let payload;
  try {
    if (e?.postData?.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      response.message = 'No data received.';
      Logger.log('Warning: No postData.contents found.');
      return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload?.action) {
      if (payload.action === 'login') {
        response = handleLogin(payload.mobile, payload.password);
      } else {
        if (!payload.spreadsheetId) {
          response = { success: false, message: 'Spreadsheet ID is required for this action.' };
        } else {
          const verification = verifyTokenAndPermission(payload.spreadsheetId, payload.token);
          if (verification.success) {
            sendNotification(payload.spreadsheetId, "School App Activity", `An action ('${payload.action}') was successfully authorized and is being processed.`);
            
            switch (payload.action) {
              case 'getSheetNames':
                response = getSheetNames(payload.spreadsheetId);
                break;
              case 'getSheetData':
                response = getSheetData(payload.spreadsheetId, payload.sheetName);
                break;
              case 'getSheetDataAsJson':
                response = getSheetDataAsJson(payload.spreadsheetId, payload.sheetName);
                break;
              case 'getStudents':
                response = getStudentsDataWithClassMapping(payload.spreadsheetId);
                break;
              case 'getStudentFullDetails':
                response = getStudentFullDetails(payload.spreadsheetId, payload.studentId);
                break;
              case 'editStudent':
                response = editStudentData(payload.spreadsheetId, payload.studentData);
                break;
              case 'deleteStudent':
                response = deleteStudent(payload.spreadsheetId, payload.studentId);
                break;
              case 'updateSheetRow':
                response = updateSheetRow(payload.spreadsheetId, payload.sheetName, payload.rowIndex, payload.rowData);
                break;
              case 'deleteSheetRow':
                response = deleteSheetRow(payload.spreadsheetId, payload.sheetName, payload.rowIndex);
                break;
              case 'saveResult':
                const subjects = typeof payload.subjects === 'string' ? JSON.parse(payload.subjects) : payload.subjects || [];
                const marks = typeof payload.marks === 'string' ? JSON.parse(payload.marks) : payload.marks || {};
                response = saveStudentResult(
                  payload.spreadsheetId,
                  payload.resultName,
                  payload.className,
                  subjects,
                  payload.studentId,
                  marks
                );
                break;
              case 'addBulkFees':
                response = addBulkFees(payload.spreadsheetId, payload.classId, payload.feeTypeId, payload.monthYear, payload.dueDate, payload.academicYear, payload.amount);
                break;
              case 'updateStudentFeeStatus':
                response = updateStudentFeeStatus(payload.spreadsheetId, payload.feeRecordId, payload.newStatus, payload.paidDate);
                break;
              case 'getStaffList':
                response = getStaffList(payload.spreadsheetId);
                break;
              case 'getStaffDetails':
                response = getStaffDetails(payload.spreadsheetId, payload.staffId);
                break;
              case 'addStaffSalaryPayment':
                response = addStaffSalaryPayment(payload.spreadsheetId, payload.staffId, payload.paymentDate, payload.amount, payload.monthYear, payload.notes);
                break;
              case 'getDashboardOverview':
                response = getDashboardOverview(payload.spreadsheetId);
                break;
              case 'completedata_fatch':
                response = getCompleteDataFetch(payload.spreadsheetId);
                break;
              default:
                response = { success: false, message: 'Unknown action specified' };
                break;
            }
          } else {
            response = verification;
          }
        }
      }
    } else {
      response.message = 'Action not specified in payload.';
    }

  } catch (error) {
    Logger.log(`Error in doPost: ${error.toString()}\nPayload: ${JSON.stringify(payload)}\nStack: ${error.stack}`);
    response = { success: false, error: 'Server error occurred: ' + error.message };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

const MAIN_SPREADSHEET_ID = '1PjNIMBpDWqU_Vj8SHnCG39mvAqjZ1S51lcLxK5Apzf8';
const LOGIN_SHEET_NAME = 'Schools';
const PERMISSIONS_SHEET_NAME = 'Permissions';

const STUDENTS_SHEET_NAME = 'Students';
const STAFFS_SHEET_NAME = 'Staffs';
const CLASSES_SHEET_NAME = 'Classes';
const SUBJECTS_SHEET_NAME = 'Subjects';
const CLASS_SUBJECTS_SHEET_NAME = 'ClassSubjects';
const STUDENTS_FEES_SHEET_NAME = 'StudentsFees';
const FEE_TYPES_SHEET_NAME = 'FeeTypes';
const STAFF_SALARY_PAYMENTS_SHEET_NAME = 'StaffSalaryPayments';
const RESULTS_SHEET_NAME_FROM_CODE = "results1129";
const ATTENDANCE_SHEET_NAME = 'Attendance';
const EXPENSES_SHEET_NAME = 'Expenses';

const STUDENT_HEADERS = ['StudentID', 'RollNumber', 'Name', 'Mobile', 'Gmail', 'Password', 'FatherName', 'MotherName', 'Class', 'Address', 'PhotoURL', 'Aadhar', 'Gender', 'RegistrationDate'];
const STAFF_HEADERS = ['StaffID', 'Name', 'Mobile', 'Gmail', 'Password', 'JoiningDate', 'PhotoURL', 'SalaryAmount', 'TotalPaid', 'TotalDues', 'IsActive'];
const CLASS_HEADERS = ['ClassID', 'ClassName', 'Section', 'ClassTeacherStaffID'];
const STUDENTS_FEES_HEADERS = ['FeeRecordID', 'StudentID', 'FeeTypeID', 'Amount', 'DueDate', 'PaidDate', 'Status', 'AcademicYear', 'Notes'];
const FEE_TYPE_HEADERS = ['FeeTypeID', 'FeeTypeName', 'DefaultAmount', 'Frequency'];
const STAFF_SALARY_PAYMENT_HEADERS = ['PaymentID', 'StaffID', 'PaymentDate', 'Amount', 'MonthYear', 'Notes'];
const ATTENDANCE_HEADERS = ['AttendanceID', 'Date', 'ClassID', 'PresentStudentIDs', 'AbsentStudentIDs'];
const EXPENSE_HEADERS = ['ExpenseID', 'Date', 'Category', 'Description', 'Amount'];

function getSheetSafely(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
  return sheet;
}

function getHeaderMap(sheet, normalize = true) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    const key = normalize ? String(header).trim().replace(/\s+/g, '') : String(header).trim();
    if (key) map[key] = index;
  });
  return map;
}

function findRowIndexByValue(sheet, colName, valueToFind, headerMap) {
  const colIndex = headerMap[colName.replace(/\s+/g, '')];
  if (colIndex === undefined) throw new Error(`Column "${colName}" not found in sheet "${sheet.getName()}".`);

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] != null && String(data[i][colIndex]).trim() === String(valueToFind).trim()) {
      return i + 1;
    }
  }
  return -1;
}

function generateUniqueId() {
  return new Date().getTime().toString(36) + Math.random().toString(36).substring(2, 7);
}

function getCurrentAcademicYear(currentDate = new Date()) {
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth() + 1;
  if (month >= 4) {
    return `${year}-${String(year + 1).slice(-2)}`;
  } else {
    return `${year - 1}-${String(year).slice(-2)}`;
  }
}

function verifyTokenAndPermission(spreadsheetId, receivedToken) {
  const isPermissionGranted = checkDataTransferPermission(spreadsheetId);
  if (!isPermissionGranted) {
    sendNotification(spreadsheetId, "SECURITY ALERT: Data Access Denied", `A request was blocked due to missing data transfer permissions.`);
    return {
      success: false,
      message: 'Data transfer permission denied by the school administrator. Please enable it in the "Permissions" sheet.'
    };
  }
  
  if (!receivedToken) {
      sendNotification(spreadsheetId, "SECURITY ALERT: Missing Token", `An access attempt was blocked due to a missing authentication token.`);
      return { success: false, message: 'Authentication token is missing. Please log in again.' };
  }
  
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const permissionsSheet = ss.getSheetByName(PERMISSIONS_SHEET_NAME);
    
    if (!permissionsSheet) {
      return { success: false, message: `Configuration error: "${PERMISSIONS_SHEET_NAME}" sheet not found.` };
    }
    
    const storedToken = permissionsSheet.getRange('C2').getValue().toString().trim();
    
    if (storedToken && storedToken === receivedToken) {
      return { success: true };
    } else {
      sendNotification(spreadsheetId, "SECURITY ALERT: Invalid Token", `An access attempt was blocked due to an invalid or expired token.`);
      return { success: false, message: 'Invalid or expired session. Please log in again.' };
    }

  } catch (error) {
    Logger.log(`Error during token verification for spreadsheet ID ${spreadsheetId}: ${error.toString()}`);
    return { success: false, message: 'Server error during token verification.' };
  }
}

function checkDataTransferPermission(spreadsheetId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const permissionsSheet = ss.getSheetByName(PERMISSIONS_SHEET_NAME);

    if (!permissionsSheet) {
      Logger.log(`Permission check failed for ${spreadsheetId}: Sheet "${PERMISSIONS_SHEET_NAME}" not found.`);
      return false;
    }

    const data = permissionsSheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log(`Permission check failed for ${spreadsheetId}: No data found in "${PERMISSIONS_SHEET_NAME}" sheet.`);
      return false;
    }

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const permissionColIndex = headers.indexOf('data transfer permission');

    if (permissionColIndex === -1) {
      Logger.log(`Permission check failed for ${spreadsheetId}: Header "data transfer permission" not found in "${PERMISSIONS_SHEET_NAME}".`);
      return false;
    }

    const permissionValue = String(data[1][permissionColIndex] || '').trim().toLowerCase();
    const isAllowed = permissionValue === 'true';
    
    if (!isAllowed) {
        Logger.log(`Data transfer denied for ${spreadsheetId}. Value found: "${permissionValue}".`);
    }
    return isAllowed;

  } catch (error) {
    Logger.log(`Error during permission check for spreadsheet ID ${spreadsheetId}: ${error.toString()}`);
    return false;
  }
}

function sendNotification(spreadsheetId, title, message) {
  Logger.log(`NOTIFICATION TRIGGERED for Spreadsheet ID: ${spreadsheetId}`);
  Logger.log(`Title: ${title}`);
  Logger.log(`Message: ${message}`);
}

function handleLogin(mobile, password) {
  if (!mobile || !password) {
    return { success: false, message: 'Mobile number and password are required.' };
  }
  try {
    const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(LOGIN_SHEET_NAME);
    if (!sheet) throw new Error(`Login sheet "${LOGIN_SHEET_NAME}" not found.`);

    const values = sheet.getDataRange().getValues();
    const mobileColIndex = 2; const passwordColIndex = 4;
    const schoolSheetIdColIndex = 8; const premiumStatusColIndex = 11;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const storedMobile = String(row[mobileColIndex] || '').trim();
      const storedPassword = String(row[passwordColIndex] || '').trim();
      const premiumStatus = String(row[premiumStatusColIndex] || '').trim().toLowerCase();
      const schoolSheetId = String(row[schoolSheetIdColIndex] || '').trim();

      if (storedMobile === mobile && storedPassword === password) {
        if (premiumStatus === 'activate') {
          if (schoolSheetId) {
              const token = Utilities.getUuid();
              try {
                  const schoolSs = SpreadsheetApp.openById(schoolSheetId);
                  const permissionsSheet = getSheetSafely(schoolSs, PERMISSIONS_SHEET_NAME);
                  
                  if(permissionsSheet.getRange('C1').getValue() === ''){
                      permissionsSheet.getRange('C1').setValue('AuthToken').setFontWeight('bold');
                  }
                  permissionsSheet.getRange('C2').setValue(token);
                  SpreadsheetApp.flush();

                  Logger.log(`Login successful for mobile: ${mobile}, School ID: ${schoolSheetId}. Token generated.`);
                  return { success: true, spreadsheetId: schoolSheetId, token: token };

              } catch(e) {
                  Logger.log(`Error saving token during login for ${mobile}: ${e.toString()}`);
                  return { success: false, error: 'Login successful, but failed to set up secure session. ' + e.message };
              }
          } else {
            Logger.log(`Login failed for ${mobile}: Premium but School ID missing.`);
            throw new Error('Configuration error: School Spreadsheet ID not found for your account.');
          }
        } else {
          Logger.log(`Login failed for ${mobile}: Not premium (Status: ${premiumStatus}).`);
          return { success: false, reason: 'not_premium', message: 'Access denied. Premium subscription required.' };
        }
      }
    }
    Logger.log(`Login failed for ${mobile}: Invalid credentials.`);
    return { success: false, message: 'Invalid mobile number or password.' };
  } catch (error) {
    Logger.log(`Error during login validation: ${error.toString()}`);
    return { success: false, error: 'An error occurred during login: ' + error.message };
  }
}

function getSheetNames(spreadsheetId) {
  if (!spreadsheetId) return { success: false, error: 'Spreadsheet ID is required.' };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    Logger.log(`Fetched ${sheetNames.length} sheet names from ID: ${spreadsheetId}`);
    return { success: true, data: sheetNames };
  } catch (error) {
    Logger.log(`Error fetching sheet names from ${spreadsheetId}: ${error.toString()}`);
    return { success: false, error: 'Error fetching sheet names: ' + error.message };
  }
}

function getSheetData(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) return { success: false, error: 'Spreadsheet ID and Sheet Name are required.' };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getSheetSafely(ss, sheetName);

    const dataRange = sheet.getDataRange();
    const values = dataRange.getDisplayValues();

    if (values.length === 0) {
      Logger.log(`Sheet "${sheetName}" (ID: ${spreadsheetId}) is empty.`);
      return { success: true, data: { headers: [], rows: [] } };
    }

    const headers = values[0].map(h => String(h).trim());
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      rows.push({ rowIndex: i + 1, values: values[i] });
    }

    Logger.log(`Fetched ${rows.length} rows from sheet "${sheetName}", ID: ${spreadsheetId}`);
    return { success: true, data: { headers: headers, rows: rows } };
  } catch (error) {
    Logger.log(`Error fetching data from sheet "${sheetName}", ID ${spreadsheetId}: ${error.toString()}`);
    return { success: false, error: `Error fetching data from "${sheetName}": ${error.message}` };
  }
}

function getSheetDataAsJson(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) return { success: false, error: 'Spreadsheet ID and Sheet Name are required.' };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getSheetSafely(ss, sheetName);

    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) {
      Logger.log(`Sheet "${sheetName}" (ID: ${spreadsheetId}) has no data or only headers.`);
      return { success: true, data: [] };
    }

    const headers = values[0].map(h => String(h).trim());
    const jsonData = [];
    for (let i = 1; i < values.length; i++) {
      const rowObject = {};
      let hasDataInRow = false;
      headers.forEach((header, colIndex) => {
        if (header) {
          rowObject[header] = values[i][colIndex];
          if (values[i][colIndex] !== null && values[i][colIndex] !== '') hasDataInRow = true;
        }
      });
      if (hasDataInRow) jsonData.push(rowObject);
    }

    Logger.log(`Fetched ${jsonData.length} objects from sheet "${sheetName}", ID: ${spreadsheetId} for JSON export.`);
    return { success: true, data: jsonData };
  } catch (error) {
    Logger.log(`Error fetching JSON data from sheet "${sheetName}", ID ${spreadsheetId}: ${error.toString()}`);
    return { success: false, error: `Error fetching JSON data from "${sheetName}": ${error.message}` };
  }
}

function getClassMapping(spreadsheet) {
  const mapping = {};
  try {
    const classesSheet = spreadsheet.getSheetByName(CLASSES_SHEET_NAME);
    if (!classesSheet) {
      Logger.log(`Warning: Sheet "${CLASSES_SHEET_NAME}" not found in spreadsheet ID ${spreadsheet.getId()}. Cannot map Class IDs.`);
      return mapping;
    }
    const values = classesSheet.getDataRange().getValues();
    const headerMap = getHeaderMap(classesSheet);
    const idCol = headerMap['ClassID'];
    const nameCol = headerMap['ClassName'.replace(/\s+/g, '')];
    const sectionCol = headerMap['Section'.replace(/\s+/g, '')];

    if (idCol === undefined) {
      Logger.log(`Warning: "ClassID" column not found in "${CLASSES_SHEET_NAME}".`);
      return mapping;
    }

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const classId = String(row[idCol] || '').trim();
      const className = nameCol !== undefined ? String(row[nameCol] || '').trim() : '';
      const section = sectionCol !== undefined ? String(row[sectionCol] || '').trim() : '';

      if (classId) {
        const namePart = className || classId;
        const sectionPart = section;
        const displayName = sectionPart ? `${namePart} - ${sectionPart}` : namePart;
        mapping[classId] = { id: classId, name: namePart, section: sectionPart, displayName: displayName };
      }
    }
    Logger.log(`Created class mapping with ${Object.keys(mapping).length} entries from spreadsheet ID ${spreadsheet.getId()}.`);
  } catch (error) {
    Logger.log(`Error reading Classes sheet in spreadsheet ID ${spreadsheet.getId()}: ${error}`);
  }
  return mapping;
}

function getStudentsDataWithClassMapping(spreadsheetId) {
  if (!spreadsheetId) return { success: false, error: 'School Spreadsheet ID is required.' };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const classMapping = getClassMapping(ss);

    const studentsSheet = getSheetSafely(ss, STUDENTS_SHEET_NAME);
    const values = studentsSheet.getDataRange().getValues();

    if (values.length <= 1) {
      Logger.log(`No student data found in "${STUDENTS_SHEET_NAME}" (ID: ${spreadsheetId}).`);
      return { success: true, data: { students: [], classes: classMapping } };
    }

    const headers = values[0].map(h => String(h).trim().replace(/\s+/g, ''));
    const studentData = [];
    const headerIndices = {};
    STUDENT_HEADERS.forEach(expectedHeader => {
      const normalizedHeader = expectedHeader.replace(/\s+/g, '');
      const index = headers.indexOf(normalizedHeader);
      if (index !== -1) headerIndices[expectedHeader] = index;
      else Logger.log(`Warning: Header "${expectedHeader}" not found in sheet "${STUDENTS_SHEET_NAME}".`);
    });

    if (headerIndices['StudentID'] === undefined) {
      throw new Error(`Essential header "StudentID" not found in sheet "${STUDENTS_SHEET_NAME}".`);
    }

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const student = {};
      let hasData = false;
      let studentIdValue = row[headerIndices['StudentID']];

      for (const headerKey in headerIndices) {
        const colIndex = headerIndices[headerKey];
        student[headerKey] = (row[colIndex] !== undefined && row[colIndex] !== null) ? String(row[colIndex]) : '';
        if (student[headerKey] !== '') hasData = true;
      }

      if (hasData && studentIdValue != null && String(studentIdValue).trim() !== '') {
        const classId = student['Class'] ? student['Class'].trim() : null;
        const mappedClassInfo = classId ? classMapping[classId] : null;
        student['ClassNameWithSection'] = mappedClassInfo ? mappedClassInfo.displayName : (classId || 'N/A');
        studentData.push(student);
      } else if (hasData) {
        Logger.log(`Skipping row ${i + 1} in ${STUDENTS_SHEET_NAME} (ID: ${spreadsheetId}) due to missing or empty StudentID.`);
      }
    }
    Logger.log(`Fetched ${studentData.length} valid students with class mapping from spreadsheet ID: ${spreadsheetId}`);
    return { success: true, data: { students: studentData, classes: classMapping } };
  } catch (error) {
    Logger.log(`Error fetching students with mapping (ID: ${spreadsheetId}): ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Error fetching students: ${error.message}` };
  }
}

function getStudentFullDetails(spreadsheetId, studentId) {
  if (!spreadsheetId || !studentId) return { success: false, error: 'Spreadsheet ID and Student ID are required.' };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const studentDetails = {};

    const studentsSheet = getSheetSafely(ss, STUDENTS_SHEET_NAME);
    const studentHeaderMap = getHeaderMap(studentsSheet, true);
    const studentRowIdx = findRowIndexByValue(studentsSheet, 'StudentID', studentId, studentHeaderMap);
    if (studentRowIdx === -1) return { success: false, error: `Student with ID "${studentId}" not found.` };

    const studentRowValues = studentsSheet.getRange(studentRowIdx, 1, 1, studentsSheet.getLastColumn()).getValues()[0];
    const studentRawHeaders = studentsSheet.getRange(1, 1, 1, studentsSheet.getLastColumn()).getValues()[0];
    studentRawHeaders.forEach((header, idx) => {
      const key = String(header).trim();
      if (key && key.toLowerCase() !== 'password') studentDetails[key] = studentRowValues[idx];
    });

    const classMapping = getClassMapping(ss);
    const classIdFromStudent = studentDetails['Class'] ? String(studentDetails['Class']).trim() : null;
    if (classIdFromStudent && classMapping[classIdFromStudent]) {
      studentDetails['ClassNameWithSection'] = classMapping[classIdFromStudent].displayName;
    } else {
      studentDetails['ClassNameWithSection'] = classIdFromStudent || 'N/A';
    }


    const fees = { due: [], paid: [], totalDue: 0, totalPaid: 0, byMonth: {} };
    try {
      const feesSheet = ss.getSheetByName(STUDENTS_FEES_SHEET_NAME);
      if (feesSheet) {
        const feeValues = feesSheet.getDataRange().getValues();
        const feeHeaderMap = getHeaderMap(feesSheet);
        const feeStudentIDCol = feeHeaderMap['StudentID'];
        const feeAmountCol = feeHeaderMap['Amount'];
        const feeStatusCol = feeHeaderMap['Status'];
        const feePaidDateCol = feeHeaderMap['PaidDate'];

        if (feeStudentIDCol !== undefined && feeAmountCol !== undefined && feeStatusCol !== undefined) {
          for (let i = 1; i < feeValues.length; i++) {
            if (String(feeValues[i][feeStudentIDCol]).trim() === studentId) {
              const record = {};
              Object.keys(feeHeaderMap).forEach(key => {
                record[key.replace(/\s+/g, '')] = feeValues[i][feeHeaderMap[key]];
              });
              feesSheet.getRange(1, 1, 1, feesSheet.getLastColumn()).getValues()[0].forEach((header, idx) => {
                record[header] = feeValues[i][idx];
              });

              const amount = parseFloat(record.Amount) || 0;
              if (String(record.Status).toLowerCase() === 'paid') {
                fees.paid.push(record);
                fees.totalPaid += amount;
                if (record.PaidDate && feePaidDateCol !== undefined) {
                  try {
                    const paidDateObj = new Date(record.PaidDate);
                    const monthYear = Utilities.formatDate(paidDateObj, Session.getScriptTimeZone(), "MMM yyyy");
                    fees.byMonth[monthYear] = (fees.byMonth[monthYear] || 0) + amount;
                  } catch (dateErr) { Logger.log("Fee Paid Date parse error for " + record.PaidDate) }
                }
              } else if (['due', 'partial'].includes(String(record.Status).toLowerCase())) {
                fees.due.push(record);
                fees.totalDue += amount;
              }
            }
          }
        } else { Logger.log(`Fee sheet "${STUDENTS_FEES_SHEET_NAME}" missing required columns (StudentID, Amount, Status).`); }
      } else { Logger.log(`Fee sheet "${STUDENTS_FEES_SHEET_NAME}" not found.`); }
    } catch (feeError) { Logger.log(`Error processing fees for student ${studentId}: ${feeError.toString()}`); }
    studentDetails.fees = fees;

    const attendance = { records: [], byMonth: {}, totalPresent: 0, totalAbsent: 0, totalSchoolDaysForStudent: 0 };
    try {
      const attSheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
      const studentsSheetForCount = getSheetSafely(ss, STUDENTS_SHEET_NAME);
      const studentHeaderMapForCount = getHeaderMap(studentsSheetForCount);
      const studentClassCol = studentHeaderMapForCount['Class'];
      const allStudentsInClass = {};

      if (attSheet && studentDetails.Class && studentClassCol !== undefined) {
        const studentClassId = String(studentDetails.Class).trim();
        const attValues = attSheet.getDataRange().getValues();
        const attHeaderMap = getHeaderMap(attSheet);
        const attDateCol = attHeaderMap['Date'];
        const attClassIDCol = attHeaderMap['ClassID'];
        const attPresentCol = attHeaderMap['PresentStudentIDs'];
        const attAbsentCol = attHeaderMap['AbsentStudentIDs'];

        if (attDateCol !== undefined && attClassIDCol !== undefined && (attPresentCol !== undefined || attAbsentCol !== undefined)) {
          for (let i = 1; i < attValues.length; i++) {
            if (String(attValues[i][attClassIDCol]).trim() === studentClassId) {
              attendance.totalSchoolDaysForStudent++;
              const dateStr = attValues[i][attDateCol];
              const presentIDs = attPresentCol !== undefined ? String(attValues[i][attPresentCol] || '').split(',').map(id => id.trim()) : [];
              const absentIDs = attAbsentCol !== undefined ? String(attValues[i][attAbsentCol] || '').split(',').map(id => id.trim()) : [];
              let status = "N/A";

              if (presentIDs.includes(studentId)) {
                status = "Present";
                attendance.totalPresent++;
              } else if (absentIDs.includes(studentId)) {
                status = "Absent";
                attendance.totalAbsent++;
              }

              attendance.records.push({ Date: dateStr, Status: status });
              try {
                const dateObj = new Date(dateStr);
                const monthYear = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMM yyyy");
                if (!attendance.byMonth[monthYear]) attendance.byMonth[monthYear] = { present: 0, absent: 0, schoolDays: 0 };
                attendance.byMonth[monthYear].schoolDays++;
                if (status === "Present") attendance.byMonth[monthYear].present++;
                else if (status === "Absent") attendance.byMonth[monthYear].absent++;
              } catch (dateErr) { Logger.log("Att Date parse error for " + dateStr) }
            }
          }
        } else { Logger.log(`Attendance sheet "${ATTENDANCE_SHEET_NAME}" missing required columns.`); }
      } else { Logger.log(`Attendance sheet or student class details missing for student ${studentId}`); }
    } catch (attError) { Logger.log(`Error processing attendance for student ${studentId}: ${attError.toString()}`); }
    studentDetails.attendance = attendance;

    return { success: true, data: studentDetails };
  } catch (error) {
    Logger.log(`Error fetching full student details for ID ${studentId}: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Error fetching details: ${error.message}` };
  }
}

function editStudentData(spreadsheetId, studentData) {
  if (!spreadsheetId || !studentData || !studentData.StudentID || String(studentData.StudentID).trim() === '') {
    return { success: false, error: 'Spreadsheet ID and valid Student Data (including non-empty StudentID) are required.' };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getSheetSafely(ss, STUDENTS_SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0].map(h => String(h).trim().replace(/\s+/g, ''));

    const studentIdColIndex = headers.indexOf('StudentID');
    if (studentIdColIndex === -1) throw new Error('"StudentID" column not found in sheet.');

    let dataRowIndex = -1;
    const studentIdToFind = String(studentData.StudentID).trim();
    for (let i = 1; i < values.length; i++) {
      if (values[i][studentIdColIndex] != null && String(values[i][studentIdColIndex]).trim() === studentIdToFind) {
        dataRowIndex = i; break;
      }
    }
    if (dataRowIndex === -1) return { success: false, error: `Student with ID "${studentData.StudentID}" not found.` };
    const sheetRowIndex = dataRowIndex + 1;

    const headerIndexMap = {};
    headers.forEach((header, index) => { headerIndexMap[header] = index; });

    const currentValues = sheet.getRange(sheetRowIndex, 1, 1, headers.length).getValues()[0];
    const newValues = [...currentValues];
    let updated = false;

    for (const key in studentData) {
      const normalizedKey = key.replace(/\s+/g, '');
      if (key !== 'StudentID' && headerIndexMap[normalizedKey] !== undefined) {
        const colIndex0Based = headerIndexMap[normalizedKey];
        const newValue = studentData[key];

        if (key === 'Password' && (newValue === undefined || newValue === '')) {
          continue;
        }

        if (String(currentValues[colIndex0Based] ?? '') !== String(newValue ?? '')) {
          newValues[colIndex0Based] = newValue;
          updated = true;
        }
      } else if (key !== 'StudentID' && headerIndexMap[normalizedKey] === undefined) {
        Logger.log(`Warning: Header "${key}" from edit data not found in sheet "${STUDENTS_SHEET_NAME}". Field skipped.`);
      }
    }

    if (updated) {
      sheet.getRange(sheetRowIndex, 1, 1, newValues.length).setValues([newValues]);
      SpreadsheetApp.flush();
      Logger.log(`Successfully updated student ID ${studentData.StudentID} in ${STUDENTS_SHEET_NAME}, Row: ${sheetRowIndex}`);
    } else {
      Logger.log(`No changes detected for student ID ${studentData.StudentID} in ${STUDENTS_SHEET_NAME}. Update skipped.`);
    }
    return { success: true, message: updated ? 'Student updated successfully.' : 'No changes applied.' };
  } catch (error) {
    Logger.log(`Error editing student ${studentData.StudentID} in ${spreadsheetId}: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: 'Error updating student data: ' + error.message };
  }
}

function deleteStudent(spreadsheetId, studentId) {
  if (!spreadsheetId || !studentId) {
    return { success: false, error: 'Spreadsheet ID and Student ID are required.' };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getSheetSafely(ss, STUDENTS_SHEET_NAME);
    const headerMap = getHeaderMap(sheet);
    const studentIdCol = headerMap['StudentID'];

    if (studentIdCol === undefined) {
      throw new Error(`"StudentID" column not found in "${STUDENTS_SHEET_NAME}".`);
    }

    const values = sheet.getDataRange().getValues();
    let rowIndexToDelete = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][studentIdCol] != null && String(values[i][studentIdCol]).trim() === String(studentId).trim()) {
        rowIndexToDelete = i + 1;
        break;
      }
    }

    if (rowIndexToDelete !== -1) {
      sheet.deleteRow(rowIndexToDelete);
      SpreadsheetApp.flush();
      Logger.log(`Deleted student with ID "${studentId}" from row ${rowIndexToDelete} in "${STUDENTS_SHEET_NAME}".`);
      return { success: true, message: 'Student deleted successfully.' };
    } else {
      Logger.log(`Student with ID "${studentId}" not found for deletion.`);
      return { success: false, error: `Student with ID "${studentId}" not found.` };
    }
  } catch (error) {
    Logger.log(`Error deleting student ${studentId}: ${error.toString()}`);
    return { success: false, error: `Error deleting student: ${error.message}` };
  }
}

function updateSheetRow(spreadsheetId, sheetName, rowIndex, rowData) {
  if (!spreadsheetId || !sheetName || !rowIndex || !rowData || typeof rowData !== 'object') {
    return { success: false, error: 'Missing required data for updating row.' };
  }
  const targetRowIndex = Number(rowIndex);
  if (isNaN(targetRowIndex) || targetRowIndex < 2) {
    return { success: false, error: 'Invalid row index provided. Must be 2 or greater.' };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getSheetSafely(ss, sheetName);
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) throw new Error(`Sheet "${sheetName}" appears to be empty or has no headers.`);
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    if (targetRowIndex > sheet.getLastRow()) {
      throw new Error(`Row index ${targetRowIndex} is out of bounds for sheet "${sheetName}".`);
    }

    const currentValues = sheet.getRange(targetRowIndex, 1, 1, lastCol).getValues()[0];
    const newValues = [...currentValues];
    let updated = false;

    headers.forEach((header, index) => {
      if (header && rowData.hasOwnProperty(header)) {
        const newValue = rowData[header];
        if (String(currentValues[index] ?? '') !== String(newValue ?? '')) {
          newValues[index] = newValue;
          updated = true;
        }
      }
    });

    if (updated) {
      sheet.getRange(targetRowIndex, 1, 1, newValues.length).setValues([newValues]);
      SpreadsheetApp.flush();
      Logger.log(`Updated row ${targetRowIndex} in sheet "${sheetName}", ID: ${spreadsheetId}`);
      return { success: true, message: 'Row updated successfully.' };
    } else {
      Logger.log(`No changes detected for row ${targetRowIndex} in sheet "${sheetName}". Update skipped.`);
      return { success: true, message: 'No changes applied.' };
    }
  } catch (error) {
    Logger.log(`Error updating row ${rowIndex} in sheet "${sheetName}", ID ${spreadsheetId}: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Error updating row: ${error.message}` };
  }
}

function deleteSheetRow(spreadsheetId, sheetName, rowIndex) {
  if (!spreadsheetId || !sheetName || !rowIndex) {
    return { success: false, error: 'Spreadsheet ID, Sheet Name, and Row Index are required.' };
  }
  const targetRowIndex = Number(rowIndex);
  if (isNaN(targetRowIndex) || targetRowIndex < 2) {
    return { success: false, error: 'Invalid row index provided. Must be 2 or greater.' };
  }

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getSheetSafely(ss, sheetName);

    if (targetRowIndex > sheet.getLastRow()) {
      return { success: false, error: `Row index ${targetRowIndex} is out of bounds for sheet "${sheetName}".` };
    }
    sheet.deleteRow(targetRowIndex);
    SpreadsheetApp.flush();
    Logger.log(`Deleted row ${targetRowIndex} from sheet "${sheetName}".`);
    return { success: true, message: 'Row deleted successfully.' };
  } catch (error) {
    Logger.log(`Error deleting row ${rowIndex} from sheet "${sheetName}": ${error.toString()}`);
    return { success: false, error: `Error deleting row: ${error.message}` };
  }
}

function saveStudentResult(spreadsheetId, resultName, classId, subjects, studentId, marks) {
  if (!spreadsheetId || !resultName || !classId || !subjects || !Array.isArray(subjects) || subjects.length === 0 || !studentId || !marks || typeof marks !== 'object') {
    Logger.log('Missing or invalid required data for saveStudentResult.', { spreadsheetId, resultName, classId, subjects, studentId, marks });
    return { success: false, error: 'Missing or invalid required data.' };
  }
  if (String(studentId).trim() === '') {
    return { success: false, error: 'StudentID cannot be empty.' };
  }
  const targetSheetName = RESULTS_SHEET_NAME_FROM_CODE;
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const studentsSheet = getSheetSafely(ss, STUDENTS_SHEET_NAME);
    const studentValues = studentsSheet.getDataRange().getValues();
    if (studentValues.length <= 1) throw new Error(`Sheet "${STUDENTS_SHEET_NAME}" contains no student data.`);

    const studentHeadersRaw = studentValues[0].map(h => String(h).trim());
    const studentHeadersNormalized = studentHeadersRaw.map(h => h.replace(/\s+/g, ''));
    const studentIdColIdx = studentHeadersNormalized.indexOf('StudentID');
    if (studentIdColIdx === -1) throw new Error(`"StudentID" column not found in "${STUDENTS_SHEET_NAME}".`);

    let studentDetails = null;
    const studentIdToFind = String(studentId).trim();
    for (let i = 1; i < studentValues.length; i++) {
      if (studentValues[i][studentIdColIdx] != null && String(studentValues[i][studentIdColIdx]).trim() === studentIdToFind) {
        studentDetails = {};
        studentHeadersRaw.forEach((header, idx) => {
          if (header && header.toLowerCase() !== 'password') {
            studentDetails[header] = studentValues[i][idx];
          }
        });
        break;
      }
    }
    if (!studentDetails) throw new Error(`Student with ID "${studentId}" not found.`);

    const classMapping = getClassMapping(ss);
    const mappedClass = classMapping[classId] || { displayName: classId };

    let resultsSheet = ss.getSheetByName(targetSheetName);
    const baseResultHeaders = ['ResultName', 'ClassName', 'Timestamp'];
    const studentDetailHeadersForResult = studentHeadersRaw.filter(h => h && h.toLowerCase() !== 'password');
    const dynamicMarkHeaders = subjects.map(subj => `${subj.name}_Marks`);
    const requiredHeaders = [...baseResultHeaders, ...studentDetailHeadersForResult, ...dynamicMarkHeaders];

    let headerMap = {};
    let currentSheetHeaders = [];

    if (!resultsSheet) {
      resultsSheet = ss.insertSheet(targetSheetName);
      resultsSheet.appendRow(requiredHeaders);
      resultsSheet.setFrozenRows(1);
      resultsSheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight('bold').setBackground('#e0e0e0');
      currentSheetHeaders = requiredHeaders;
    } else {
      currentSheetHeaders = resultsSheet.getRange(1, 1, 1, resultsSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const missingHeaders = requiredHeaders.filter(reqHeader => !currentSheetHeaders.includes(reqHeader));
      if (missingHeaders.length > 0) {
        const firstNewCol = resultsSheet.getLastColumn() + 1;
        resultsSheet.getRange(1, firstNewCol, 1, missingHeaders.length).setValues([missingHeaders]);
        resultsSheet.getRange(1, firstNewCol, 1, missingHeaders.length).setFontWeight('bold').setBackground('#e0e0e0');
        currentSheetHeaders = resultsSheet.getRange(1, 1, 1, resultsSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      }
    }
    currentSheetHeaders.forEach((h, i) => { headerMap[h] = i; });

    const res_studentIdColIdx = headerMap['StudentID'];
    const res_resultNameColIdx = headerMap['ResultName'];
    let existingDataRowIndex = -1;

    if (res_studentIdColIdx !== undefined && res_resultNameColIdx !== undefined) {
      const sheetData = resultsSheet.getDataRange().getValues();
      for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        if (row[res_studentIdColIdx] != null && String(row[res_studentIdColIdx]).trim() === String(studentId).trim() &&
          row[res_resultNameColIdx] != null && String(row[res_resultNameColIdx]).trim() === String(resultName).trim()) {
          existingDataRowIndex = i;
          break;
        }
      }
    } else {
      Logger.log(`Warning: StudentID or ResultName column not found in results sheet "${targetSheetName}".`);
    }
    const rowDataArray = new Array(currentSheetHeaders.length).fill('');
    if (headerMap['ResultName'] !== undefined) rowDataArray[headerMap['ResultName']] = resultName;
    if (headerMap['ClassName'] !== undefined) rowDataArray[headerMap['ClassName']] = mappedClass.displayName;
    if (headerMap['Timestamp'] !== undefined) rowDataArray[headerMap['Timestamp']] = new Date();

    studentDetailHeadersForResult.forEach(header => {
      if (headerMap[header] !== undefined) {
        rowDataArray[headerMap[header]] = studentDetails[header] !== undefined ? studentDetails[header] : '';
      }
    });
    subjects.forEach(subj => {
      const headerName = `${subj.name}_Marks`;
      if (headerMap[headerName] !== undefined) {
        rowDataArray[headerMap[headerName]] = marks[subj.name] !== undefined ? marks[subj.name] : '';
      } else {
        Logger.log(`Warning: Header "${headerName}" was expected but not found in results sheet "${targetSheetName}".`);
      }
    });

    if (existingDataRowIndex > 0) {
      const sheetRowIndex = existingDataRowIndex + 1;
      resultsSheet.getRange(sheetRowIndex, 1, 1, rowDataArray.length).setValues([rowDataArray]);
      Logger.log(`Updated result for StudentID ${studentId}, ResultName ${resultName} in row ${sheetRowIndex} of sheet "${targetSheetName}"`);
    } else {
      resultsSheet.appendRow(rowDataArray);
      Logger.log(`Appended new result for StudentID ${studentId}, ResultName ${resultName} to sheet "${targetSheetName}"`);
    }
    SpreadsheetApp.flush();
    return { success: true, message: existingDataRowIndex > 0 ? 'Result updated successfully.' : 'Result saved successfully.' };
  } catch (error) {
    Logger.log(`Error saving result for student ${studentId} in ${spreadsheetId}: ${error.toString()}\nStack: ${error.stack}`);
    let userErrorMessage = 'Error saving result data: ' + error.message;
    if (error.message.includes("You do not have permission")) userErrorMessage = "Permission error: Script cannot write to the target sheet.";
    else if (error.message.includes("not found")) userErrorMessage = "Configuration error: A required sheet could not be found.";
    return { success: false, error: userErrorMessage };
  }
}

function addBulkFees(spreadsheetId, classId, feeTypeId, monthYear, dueDateStr, academicYearStr, customAmount) {
  if (!spreadsheetId || !classId || !feeTypeId || !monthYear || !dueDateStr) {
    return { success: false, error: "Spreadsheet ID, Class ID, Fee Type ID, Month/Year, and Due Date are required." };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const studentsSheet = getSheetSafely(ss, STUDENTS_SHEET_NAME);
    const feeTypesSheet = getSheetSafely(ss, FEE_TYPES_SHEET_NAME);
    const studentsFeesSheet = getSheetSafely(ss, STUDENTS_FEES_SHEET_NAME);

    const feeTypeHeaderMap = getHeaderMap(feeTypesSheet);
    const ftIdCol = feeTypeHeaderMap['FeeTypeID'];
    const ftAmountCol = feeTypeHeaderMap['DefaultAmount'];
    if (ftIdCol === undefined || ftAmountCol === undefined) throw new Error("FeeTypes sheet missing FeeTypeID or DefaultAmount columns.");

    const feeTypeValues = feeTypesSheet.getDataRange().getValues();
    let feeAmount = null;
    for (let i = 1; i < feeTypeValues.length; i++) {
      if (String(feeTypeValues[i][ftIdCol]).trim() === String(feeTypeId).trim()) {
        feeAmount = parseFloat(customAmount ?? feeTypeValues[i][ftAmountCol]);
        break;
      }
    }
    if (feeAmount === null || isNaN(feeAmount)) return { success: false, error: `Fee Type ID "${feeTypeId}" not found or has invalid default amount.` };

    const studentHeaderMap = getHeaderMap(studentsSheet);
    const sIdCol = studentHeaderMap['StudentID'];
    const sClassCol = studentHeaderMap['Class'];
    if (sIdCol === undefined || sClassCol === undefined) throw new Error("Students sheet missing StudentID or Class columns.");

    const studentValues = studentsSheet.getDataRange().getValues();
    const studentsInClass = [];
    for (let i = 1; i < studentValues.length; i++) {
      if (String(studentValues[i][sClassCol]).trim() === String(classId).trim() && studentValues[i][sIdCol]) {
        studentsInClass.push(String(studentValues[i][sIdCol]).trim());
      }
    }

    if (studentsInClass.length === 0) return { success: false, message: "No students found in the selected class." };

    const academicYear = academicYearStr || getCurrentAcademicYear(new Date(dueDateStr));
    const newFeeRecords = [];
    const now = new Date();

    studentsInClass.forEach(studentId => {
      const feeRecord = [
        `FR-${generateUniqueId()}`,
        studentId,
        feeTypeId,
        feeAmount,
        new Date(dueDateStr),
        null,
        'Due',
        academicYear,
        `Bulk added for ${monthYear}`
      ];
      newFeeRecords.push(feeRecord);
    });

    if (newFeeRecords.length > 0) {
      if (studentsFeesSheet.getLastRow() === 0) {
        studentsFeesSheet.appendRow(STUDENTS_FEES_HEADERS);
      }
      studentsFeesSheet.getRange(studentsFeesSheet.getLastRow() + 1, 1, newFeeRecords.length, newFeeRecords[0].length).setValues(newFeeRecords);
      SpreadsheetApp.flush();
    }

    Logger.log(`Added ${newFeeRecords.length} fee records for Class ID ${classId}, Fee Type ${feeTypeId}.`);
    return { success: true, message: `Successfully added fees for ${newFeeRecords.length} students.` };
  } catch (error) {
    Logger.log(`Error in addBulkFees: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Failed to add bulk fees: ${error.message}` };
  }
}

function updateStudentFeeStatus(spreadsheetId, feeRecordId, newStatus, paidDateStr) {
  if (!spreadsheetId || !feeRecordId || !newStatus) {
    return { success: false, error: "Spreadsheet ID, Fee Record ID, and New Status are required." };
  }
  if (newStatus.toLowerCase() === 'paid' && !paidDateStr) {
    return { success: false, error: "Paid Date is required when status is 'Paid'." };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const feesSheet = getSheetSafely(ss, STUDENTS_FEES_SHEET_NAME);
    const headerMap = getHeaderMap(feesSheet);
    const frIdCol = headerMap['FeeRecordID'];
    const statusCol = headerMap['Status'];
    const pdCol = headerMap['PaidDate'];

    if (frIdCol === undefined || statusCol === undefined || pdCol === undefined) {
      throw new Error(`Required columns (FeeRecordID, Status, PaidDate) not found in "${STUDENTS_FEES_SHEET_NAME}".`);
    }
    const feeValues = feesSheet.getDataRange().getValues();
    let rowIndexToUpdate = -1;
    for (let i = 1; i < feeValues.length; i++) {
      if (String(feeValues[i][frIdCol]).trim() === String(feeRecordId).trim()) {
        rowIndexToUpdate = i + 1;
        break;
      }
    }
    if (rowIndexToUpdate === -1) return { success: false, error: `Fee Record ID "${feeRecordId}" not found.` };

    feesSheet.getRange(rowIndexToUpdate, statusCol + 1).setValue(newStatus);
    if (newStatus.toLowerCase() === 'paid') {
      feesSheet.getRange(rowIndexToUpdate, pdCol + 1).setValue(new Date(paidDateStr));
    } else {
      feesSheet.getRange(rowIndexToUpdate, pdCol + 1).setValue(null);
    }
    SpreadsheetApp.flush();
    Logger.log(`Updated fee status for Fee Record ID ${feeRecordId} to ${newStatus}.`);
    return { success: true, message: "Fee status updated successfully." };

  } catch (error) {
    Logger.log(`Error updating fee status for ${feeRecordId}: ${error.toString()}`);
    return { success: false, error: `Error updating fee status: ${error.message}` };
  }
}

function getStaffList(spreadsheetId) {
  if (!spreadsheetId) return { success: false, error: "Spreadsheet ID is required." };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const staffSheet = getSheetSafely(ss, STAFFS_SHEET_NAME);
    const values = staffSheet.getDataRange().getValues();
    if (values.length <= 1) return { success: true, data: [] };

    const headers = values[0].map(h => String(h).trim());
    const staffList = [];
    for (let i = 1; i < values.length; i++) {
      const staffMember = {};
      let hasData = false;
      headers.forEach((header, index) => {
        if (header.toLowerCase() !== 'password') {
          staffMember[header.replace(/\s+/g, '')] = values[i][index];
          staffMember[header] = values[i][index];
          if (values[i][index]) hasData = true;
        }
      });
      if (hasData && staffMember.StaffID) staffList.push(staffMember);
    }
    return { success: true, data: staffList };
  } catch (error) {
    Logger.log(`Error fetching staff list: ${error.toString()}`);
    return { success: false, error: `Error fetching staff list: ${error.message}` };
  }
}

function getStaffDetails(spreadsheetId, staffId) {
  if (!spreadsheetId || !staffId) return { success: false, error: "Spreadsheet ID and Staff ID are required." };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const staffDetails = {};

    const staffsSheet = getSheetSafely(ss, STAFFS_SHEET_NAME);
    const staffHeaderMap = getHeaderMap(staffsSheet);
    const staffRowIdx = findRowIndexByValue(staffsSheet, 'StaffID', staffId, staffHeaderMap);

    if (staffRowIdx === -1) return { success: false, error: `Staff with ID "${staffId}" not found.` };

    const staffRowValues = staffsSheet.getRange(staffRowIdx, 1, 1, staffsSheet.getLastColumn()).getValues()[0];
    const staffRawHeaders = staffsSheet.getRange(1, 1, 1, staffsSheet.getLastColumn()).getValues()[0];
    staffRawHeaders.forEach((header, idx) => {
      const key = String(header).trim();
      if (key.toLowerCase() !== 'password') staffDetails[key] = staffRowValues[idx];
    });

    const payments = { records: [], totalPaid: 0 };
    try {
      const paymentsSheet = ss.getSheetByName(STAFF_SALARY_PAYMENTS_SHEET_NAME);
      if (paymentsSheet) {
        const paymentValues = paymentsSheet.getDataRange().getValues();
        const paymentHeaderMap = getHeaderMap(paymentsSheet);
        const pStaffIdCol = paymentHeaderMap['StaffID'];
        const pAmountCol = paymentHeaderMap['Amount'];

        if (pStaffIdCol !== undefined && pAmountCol !== undefined) {
          for (let i = 1; i < paymentValues.length; i++) {
            if (String(paymentValues[i][pStaffIdCol]).trim() === staffId) {
              const record = {};
              Object.keys(paymentHeaderMap).forEach(key => {
                record[key] = paymentValues[i][paymentHeaderMap[key]];
              });
              paymentsSheet.getRange(1, 1, 1, paymentsSheet.getLastColumn()).getValues()[0].forEach((header, idx) => {
                record[header] = paymentValues[i][idx];
              });
              payments.records.push(record);
              payments.totalPaid += parseFloat(record.Amount) || 0;
            }
          }
        }
      }
    } catch (paymentError) { Logger.log("Error processing staff payments: " + paymentError.toString()) }
    staffDetails.payments = payments;

    return { success: true, data: staffDetails };
  } catch (error) {
    Logger.log(`Error fetching staff details for ID ${staffId}: ${error.toString()}`);
    return { success: false, error: `Error fetching staff details: ${error.message}` };
  }
}

function addStaffSalaryPayment(spreadsheetId, staffId, paymentDateStr, amount, monthYear, notes) {
  if (!spreadsheetId || !staffId || !paymentDateStr || !amount || !monthYear) {
    return { success: false, error: "Spreadsheet ID, Staff ID, Payment Date, Amount, and Month/Year are required." };
  }
  const paymentAmount = parseFloat(amount);
  if (isNaN(paymentAmount) || paymentAmount <= 0) return { success: false, error: "Invalid payment amount." };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const paymentsSheet = getSheetSafely(ss, STAFF_SALARY_PAYMENTS_SHEET_NAME);
    const staffsSheet = getSheetSafely(ss, STAFFS_SHEET_NAME);

    const paymentRecord = [
      `PAY-${generateUniqueId()}`,
      staffId,
      new Date(paymentDateStr),
      paymentAmount,
      monthYear,
      notes || ''
    ];
    if (paymentsSheet.getLastRow() === 0) {
      paymentsSheet.appendRow(STAFF_SALARY_PAYMENT_HEADERS);
    }
    paymentsSheet.appendRow(paymentRecord);

    const staffHeaderMap = getHeaderMap(staffsSheet);
    const staffRowIdx = findRowIndexByValue(staffsSheet, 'StaffID', staffId, staffHeaderMap);
    if (staffRowIdx !== -1) {
      const totalPaidCol = staffHeaderMap['TotalPaid'];
      const salaryAmountCol = staffHeaderMap['SalaryAmount'];
      const totalDuesCol = staffHeaderMap['TotalDues'];

      if (totalPaidCol !== undefined) {
        const currentTotalPaid = parseFloat(staffsSheet.getRange(staffRowIdx, totalPaidCol + 1).getValue()) || 0;
        staffsSheet.getRange(staffRowIdx, totalPaidCol + 1).setValue(currentTotalPaid + paymentAmount);

        if (salaryAmountCol !== undefined && totalDuesCol !== undefined) {
          const salary = parseFloat(staffsSheet.getRange(staffRowIdx, salaryAmountCol + 1).getValue()) || 0;
          const currentDues = parseFloat(staffsSheet.getRange(staffRowIdx, totalDuesCol + 1).getValue()) || 0;
          if (currentDues > 0) {
            staffsSheet.getRange(staffRowIdx, totalDuesCol + 1).setValue(Math.max(0, currentDues - paymentAmount));
          }

        }
      }
    }
    SpreadsheetApp.flush();
    Logger.log(`Added salary payment for Staff ID ${staffId}.`);
    return { success: true, message: "Salary payment added successfully." };

  } catch (error) {
    Logger.log(`Error adding staff salary payment: ${error.toString()}`);
    return { success: false, error: `Error adding salary payment: ${error.message}` };
  }
}

function getCompleteDataFetch(spreadsheetId) {
  if (!spreadsheetId) {
    return { success: false, error: 'Spreadsheet ID is required.' };
  }
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    const allData = {};

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      try {
        const values = sheet.getDataRange().getValues();
        if (values.length > 1) {
          const headers = values[0].map(h => String(h).trim());
          const jsonData = [];
          for (let i = 1; i < values.length; i++) {
            const rowObject = {};
            let hasDataInRow = false;
            headers.forEach((header, colIndex) => {
              if (header) {
                rowObject[header] = values[i][colIndex];
                if (values[i][colIndex] !== null && values[i][colIndex] !== '') {
                  hasDataInRow = true;
                }
              }
            });
            if (hasDataInRow) {
              jsonData.push(rowObject);
            }
          }
          allData[sheetName] = jsonData;
        } else {
          allData[sheetName] = []; // Sheet is empty or has only headers
        }
      } catch (sheetError) {
        Logger.log(`Could not process sheet "${sheetName}" in spreadsheet ${spreadsheetId}: ${sheetError.toString()}`);
        allData[sheetName] = { error: `Failed to process sheet: ${sheetError.message}` };
      }
    });

    Logger.log(`Fetched complete data from ${Object.keys(allData).length} sheets from ID: ${spreadsheetId}`);
    return { success: true, data: allData };
  } catch (error) {
    Logger.log(`Error fetching complete data from ${spreadsheetId}: ${error.toString()}`);
    return { success: false, error: 'Error fetching complete data: ' + error.message };
  }
}

function getDashboardOverview(spreadsheetId) {
  if (!spreadsheetId) return { success: false, error: "Spreadsheet ID is required." };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const overview = {};
    const classMapping = getClassMapping(ss);

    const studentsSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
    if (studentsSheet) {
      const studentValues = studentsSheet.getDataRange().getValues();
      const studentHeaderMap = getHeaderMap(studentsSheet);
      const sIdCol = studentHeaderMap['StudentID'];
      const genderCol = studentHeaderMap['Gender'];
      const classCol = studentHeaderMap['Class'];

      let totalStudents = 0, totalGirls = 0, totalBoys = 0;
      const studentsPerClass = {};
      const genderSplitPerClass = {};

      if (sIdCol !== undefined) {
        for (let i = 1; i < studentValues.length; i++) {
          if (studentValues[i][sIdCol]) {
            totalStudents++;
            const gender = genderCol !== undefined ? String(studentValues[i][genderCol]).toLowerCase() : "";
            if (gender === 'female') totalGirls++;
            else if (gender === 'male') totalBoys++;

            if (classCol !== undefined) {
              const classId = String(studentValues[i][classCol]).trim();
              if (classId) {
                const className = classMapping[classId] ? classMapping[classId].displayName : classId;
                studentsPerClass[className] = (studentsPerClass[className] || 0) + 1;
                if (!genderSplitPerClass[className]) genderSplitPerClass[className] = { boys: 0, girls: 0, other: 0 };
                if (gender === 'female') genderSplitPerClass[className].girls++;
                else if (gender === 'male') genderSplitPerClass[className].boys++;
                else genderSplitPerClass[className].other++;
              }
            }
          }
        }
      }
      overview.totalStudents = totalStudents;
      overview.totalGirls = totalGirls;
      overview.totalBoys = totalBoys;
      overview.studentsPerClass = studentsPerClass;
      overview.genderSplitPerClass = genderSplitPerClass;
    }

    const staffSheet = ss.getSheetByName(STAFFS_SHEET_NAME);
    if (staffSheet) {
      const staffValues = staffSheet.getDataRange().getValues();
      const staffHeaderMap = getHeaderMap(staffSheet);
      const staffIdCol = staffHeaderMap['StaffID'];
      const isActiveCol = staffHeaderMap['IsActive'];
      let totalStaff = 0;
      if (staffIdCol !== undefined) {
        for (let i = 1; i < staffValues.length; i++) {
          if (staffValues[i][staffIdCol] && (isActiveCol === undefined || String(staffValues[i][isActiveCol]).toLowerCase() === 'true')) {
            totalStaff++;
          }
        }
      }
      overview.totalStaff = totalStaff;
    }

    overview.totalClasses = Object.keys(classMapping).length;

    const feesSheet = ss.getSheetByName(STUDENTS_FEES_SHEET_NAME);
    if (feesSheet) {
      const feeValues = feesSheet.getDataRange().getValues();
      const feeHeaderMap = getHeaderMap(feesSheet);
      const fStudentIdCol = feeHeaderMap['StudentID'];
      const fAmountCol = feeHeaderMap['Amount'];
      const fStatusCol = feeHeaderMap['Status'];
      const fPaidDateCol = feeHeaderMap['PaidDate'];

      let totalPaidMoney = 0, totalDuesMoney = 0;
      const feesCollectedPerClassThisMonth = {};
      const duesPerClass = {};
      const currentMonthYearStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");

      const studentToClassMap = {};
      if (studentsSheet && fStudentIdCol !== undefined && feeHeaderMap['StudentID'] !== undefined) {
        const studentValCache = studentsSheet.getDataRange().getValues();
        const studHMapCache = getHeaderMap(studentsSheet);
        const sIDColCache = studHMapCache['StudentID'];
        const classColCache = studHMapCache['Class'];
        if (sIDColCache !== undefined && classColCache !== undefined) {
          for (let i = 1; i < studentValCache.length; i++) {
            if (studentValCache[i][sIDColCache]) {
              studentToClassMap[String(studentValCache[i][sIDColCache]).trim()] = String(studentValCache[i][classColCache]).trim();
            }
          }
        }
      }

      if (fAmountCol !== undefined && fStatusCol !== undefined) {
        for (let i = 1; i < feeValues.length; i++) {
          const amount = parseFloat(feeValues[i][fAmountCol]) || 0;
          const status = String(feeValues[i][fStatusCol]).toLowerCase();
          const studentIdForFee = fStudentIdCol !== undefined ? String(feeValues[i][fStudentIdCol]).trim() : null;
          const classIdForFee = studentIdForFee ? studentToClassMap[studentIdForFee] : null;
          const classNameForFee = classIdForFee && classMapping[classIdForFee] ? classMapping[classIdForFee].displayName : (classIdForFee || "UnknownClass");

          if (status === 'paid') {
            totalPaidMoney += amount;
            if (fPaidDateCol !== undefined && feeValues[i][fPaidDateCol]) {
              try {
                const paidDate = new Date(feeValues[i][fPaidDateCol]);
                if (Utilities.formatDate(paidDate, Session.getScriptTimeZone(), "yyyy-MM") === currentMonthYearStr) {
                  feesCollectedPerClassThisMonth[classNameForFee] = (feesCollectedPerClassThisMonth[classNameForFee] || 0) + amount;
                }
              } catch (e) { }
            }
          } else if (status === 'due' || status === 'partial') {
            totalDuesMoney += amount;
            duesPerClass[classNameForFee] = (duesPerClass[classNameForFee] || 0) + amount;
          }
        }
      }
      overview.totalPaidMoney = totalPaidMoney;
      overview.totalDuesMoney = totalDuesMoney;
      overview.feesCollectedPerClassThisMonth = feesCollectedPerClassThisMonth;
      overview.duesPerClass = duesPerClass;
    }

    const expensesSheet = ss.getSheetByName(EXPENSES_SHEET_NAME);
    if (expensesSheet) {
      const expenseValues = expensesSheet.getDataRange().getValues();
      const expHeaderMap = getHeaderMap(expensesSheet);
      const eAmountCol = expHeaderMap['Amount'];
      let totalExpenses = 0;
      if (eAmountCol !== undefined) {
        for (let i = 1; i < expenseValues.length; i++) {
          totalExpenses += parseFloat(expenseValues[i][eAmountCol]) || 0;
        }
      }
      overview.totalExpenses = totalExpenses;
    }

    const attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    const attendancePerClass = {};
    if (attendanceSheet && studentsSheet) {
      const attValues = attendanceSheet.getDataRange().getValues();
      const attHeaderMap = getHeaderMap(attendanceSheet);
      const attDateCol = attHeaderMap['Date'];
      const attClassIDCol = attHeaderMap['ClassID'];
      const attPresentIDsCol = attHeaderMap['PresentStudentIDs'];

      const classStudentCounts = {};
      const studVals = studentsSheet.getDataRange().getValues();
      const studHMap = getHeaderMap(studentsSheet);
      const studClassCol = studHMap['Class'];
      const studIdCol = studHMap['StudentID'];
      if (studClassCol !== undefined && studIdCol !== undefined) {
        for (let i = 1; i < studVals.length; i++) {
          if (studVals[i][studIdCol]) {
            const cId = String(studVals[i][studClassCol]).trim();
            if (cId) classStudentCounts[cId] = (classStudentCounts[cId] || 0) + 1;
          }
        }
      }

      if (attDateCol !== undefined && attClassIDCol !== undefined && attPresentIDsCol !== undefined) {
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        const classAttendanceAgg = {};

        for (let i = 1; i < attValues.length; i++) {
          try {
            const attDate = new Date(attValues[i][attDateCol]);
            if (attDate < thirtyDaysAgo) continue;
            const classId = String(attValues[i][attClassIDCol]).trim();
            if (!classId || !classMapping[classId]) continue;
            if (!classAttendanceAgg[classId]) classAttendanceAgg[classId] = { totalPresentStudents: 0, totalSchoolDaysForClass: 0 };
            const presentStudentIds = String(attValues[i][attPresentIDsCol] || '').split(',').filter(id => id.trim() !== "");
            classAttendanceAgg[classId].totalPresentStudents += presentStudentIds.length;
            classAttendanceAgg[classId].totalSchoolDaysForClass += 1;
          } catch (e) { }
        }

        for (const classId in classAttendanceAgg) {
          const aggData = classAttendanceAgg[classId];
          const totalEnrolledInClass = classStudentCounts[classId] || 0;
          const className = classMapping[classId].displayName;
          if (aggData.totalSchoolDaysForClass > 0 && totalEnrolledInClass > 0) {
            const avgPresentStudentsPerDay = aggData.totalPresentStudents / aggData.totalSchoolDaysForClass;
            const avgAttendancePercent = (avgPresentStudentsPerDay / totalEnrolledInClass) * 100;
            attendancePerClass[className] = {
              averageAttendancePercent: parseFloat(avgAttendancePercent.toFixed(2)),
            };
          } else {
            attendancePerClass[className] = { averageAttendancePercent: 0 };
          }
        }
      }
      overview.attendancePerClass = attendancePerClass;
    }

    return { success: true, data: overview };
  } catch (error) {
    Logger.log(`Error in getDashboardOverview: ${error.toString()}\nStack: ${error.stack}`);
    return { success: false, error: `Error fetching dashboard overview: ${error.message}` };
  }
}