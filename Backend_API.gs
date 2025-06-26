//==================================================================================
// UTILITY & VALIDATION
//==================================================================================

function _validateToken(token) {
    if (!token) return null;
    const userCache = CacheService.getUserCache();
    const sessionData = userCache.get(token);
    return sessionData ? JSON.parse(sessionData) : null;
}

function _notifyHRUsers(spreadsheet, message, linkRequestId) {
    const empSheet = spreadsheet.getSheetByName("Employees");
    const empData = empSheet.getDataRange().getValues();
    const headers = empData[0];
    const roleIndex = headers.indexOf("Role");
    const idIndex = headers.indexOf("EmployeeID");

    const hrUsers = empData.filter(row => row[roleIndex] === "HR");
    hrUsers.forEach(hr => {
        const hrId = hr[idIndex];
        _createNotification(spreadsheet, hrId, message, linkRequestId);
    });
}

//==================================================================================
// MAIN API FUNCTIONS
//==================================================================================

function getDashboardData(token) {
    const user = _validateToken(token);
    if (!user) return { error: "Invalid session" };
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        
        const empSheet = ss.getSheetByName("Employees");
        const allEmpDataWithHeaders = empSheet.getDataRange().getValues();
        const empHeaders = allEmpDataWithHeaders[0];
        const allEmpData = allEmpDataWithHeaders.slice(1);

        const employeeMap = new Map(allEmpData.map(e => [e[empHeaders.indexOf("EmployeeID")], `${e[empHeaders.indexOf("Title")]}${e[empHeaders.indexOf("FirstName")]} ${e[empHeaders.indexOf("LastName")]}`]));
        const typeSheet = ss.getSheetByName("LeaveTypes");
        const typeData = typeSheet.getRange("A2:B" + typeSheet.getLastRow()).getValues();
        const leaveTypeMap = new Map(typeData.map(t => [t[0], t[1]]));
        
        const leaveBalance = _getLeaveBalance(ss, user.employeeId);

        let dashboardData = {
            leaveBalance: leaveBalance,
            usedLeaveSummary: [],
            calendarEvents: [],
            managerTasks: []
        };

        const allLeaveRequests = ss.getSheetByName("LeaveRequests").getLastRow() > 1 ? ss.getSheetByName("LeaveRequests").getDataRange().getValues().slice(1) : [];
        const allOTRequests = ss.getSheetByName("OTRequests").getLastRow() > 1 ? ss.getSheetByName("OTRequests").getDataRange().getValues().slice(1) : [];
        
        const isManagerOrHR = user.role === "HR" || user.role === "Manager" || user.role === "Supervisor";
        let subordinateIds = [];
        if (isManagerOrHR && user.role !== "HR") {
            subordinateIds = allEmpData
                .filter(emp => emp[empHeaders.indexOf("ManagerID")] === user.employeeId)
                .map(emp => emp[empHeaders.indexOf("EmployeeID")]);
        }
        
        if (user.role === "HR") {
            const pendingHRLeave = allLeaveRequests.filter(req => req[7] === 'Pending HR').map(req => _formatTaskCard(req, employeeMap, leaveTypeMap, 'leave'));
            const pendingHROT = allOTRequests.filter(req => req[7] === 'Pending HR').map(req => _formatTaskCard(req, employeeMap, null, 'ot'));
            dashboardData.managerTasks = [...pendingHRLeave, ...pendingHROT];
        } else if (user.role === "Manager" || user.role === "Supervisor") {
            const pendingManagerLeave = allLeaveRequests.filter(req => subordinateIds.includes(req[1]) && req[7] === 'Pending Manager').map(req => _formatTaskCard(req, employeeMap, leaveTypeMap, 'leave'));
            const pendingManagerOT = allOTRequests.filter(req => subordinateIds.includes(req[1]) && req[7] === 'Pending Manager').map(req => _formatTaskCard(req, employeeMap, null, 'ot'));
            dashboardData.managerTasks = [...pendingManagerLeave, ...pendingManagerOT];
        }

        const myApprovedLeaveRequests = allLeaveRequests.filter(req => req[1] === user.employeeId && req[7] === 'Approved');
        const summary = {};
        myApprovedLeaveRequests.forEach(req => {
            const typeName = leaveTypeMap.get(req[2]) || req[2];
            const days = Number(req[5]);
            if (!summary[typeName]) summary[typeName] = 0;
            summary[typeName] += days;
        });
        dashboardData.usedLeaveSummary = Object.entries(summary).map(([name, days]) => ({ typeName: name, totalDays: days }));

        let relevantLeaveRequestsForCalendar = [];
        if (user.role === "HR") {
            relevantLeaveRequestsForCalendar = allLeaveRequests;
        } else if (isManagerOrHR) {
             const teamIds = [...subordinateIds, user.employeeId];
             relevantLeaveRequestsForCalendar = allLeaveRequests.filter(req => teamIds.includes(req[1]));
        } else {
            relevantLeaveRequestsForCalendar = allLeaveRequests.filter(req => req[1] === user.employeeId);
        }

        relevantLeaveRequestsForCalendar.filter(req => req[7] === 'Approved').forEach(req => {
            let currentDate = new Date(req[3]);
            const endDate = new Date(req[4]);
            const employeeName = employeeMap.get(req[1]) || 'ไม่พบชื่อ';
            const leaveType = leaveTypeMap.get(req[2]) || 'ไม่ระบุ';
            while (currentDate <= endDate) {
                dashboardData.calendarEvents.push({
                    date: currentDate.toISOString().split('T')[0],
                    employeeName: employeeName,
                    leaveType: leaveType
                });
                currentDate.setDate(currentDate.getDate() + 1);
            }
        });

        return dashboardData;
    } catch (e) {
        Logger.log(`getDashboardData Error: ${e.toString()}`);
        return { error: `เกิดข้อผิดพลาดในการโหลดข้อมูล: ${e.message}` };
    }
}

function _formatTaskCard(req, employeeMap, leaveTypeMap, type) {
    if (type === 'leave') {
        return {
            type: 'leave',
            requestId: req[0], 
            employeeName: employeeMap.get(req[1]) || 'Unknown', 
            leaveType: leaveTypeMap.get(req[2]) || req[2],
            startDate: new Date(req[3]).toLocaleDateString('th-TH', { day: '2-digit', month: 'short', year: 'numeric'}),
            endDate: new Date(req[4]).toLocaleDateString('th-TH', { day: '2-digit', month: 'short', year: 'numeric'}),
            totalDays: req[5], 
        };
    } else {
        return {
            type: 'ot',
            requestId: req[0],
            employeeName: employeeMap.get(req[1]) || 'Unknown',
            leaveType: `OT (${req[5]} ชม.)`,
            startDate: new Date(req[2]).toLocaleDateString('th-TH', { day: '2-digit', month: 'short', year: 'numeric'}),
            endDate: `${req[3]} - ${req[4]}`,
            totalDays: req[5],
        };
    }
}

function getHistory(token) {
    const user = _validateToken(token);
    if (!user) return { success: false, leaveHistory: [], otHistory: [] };
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const statusMap = {'Approved':'อนุมัติแล้ว', 'Rejected':'ปฏิเสธ', 'Pending Manager':'รอหัวหน้าอนุมัติ', 'Pending HR': 'รอ HR อนุมัติ'};

        const typeSheet = ss.getSheetByName("LeaveTypes");
        const typeData = typeSheet.getDataRange().getValues();
        const leaveTypeMap = new Map(typeData.slice(1).map(row => [row[0], row[1]]));
        const leaveSheet = ss.getSheetByName("LeaveRequests");
        let userLeaveHistory = [];
        if (leaveSheet.getLastRow() > 1) {
            const leaveData = leaveSheet.getRange(2, 1, leaveSheet.getLastRow() - 1, leaveSheet.getLastColumn()).getValues();
            for(const row of leaveData){
                if(row[1] === user.employeeId) {
                    userLeaveHistory.push({
                        requestDate: new Date(row[8]).toLocaleDateString('th-TH'),
                        type: leaveTypeMap.get(row[2]) || row[2],
                        dateRange: `${new Date(row[3]).toLocaleDateString('th-TH')} - ${new Date(row[4]).toLocaleDateString('th-TH')}`,
                        total: `${row[5]} วัน`,
                        status: statusMap[row[7]] || row[7],
                        statusClass: row[7]
                    });
                }
            }
        }

        const otSheet = ss.getSheetByName("OTRequests");
        let userOTHistory = [];
        if (otSheet.getLastRow() > 1) {
            // ** CHANGE START: Use getDisplayValues to get time as string **
            const otDataRange = otSheet.getRange(2, 1, otSheet.getLastRow() - 1, otSheet.getLastColumn());
            const otValues = otDataRange.getValues();
            const otDisplayValues = otDataRange.getDisplayValues(); // Gets "11:29:00" instead of a date object

            for(let i = 0; i < otValues.length; i++){
                 const row = otValues[i];
                 const displayRow = otDisplayValues[i];

                 if(row[1] === user.employeeId) {
                    const otDate = new Date(row[2]);
                    const startTimeString = displayRow[3];
                    const endTimeString = displayRow[4];

                    userOTHistory.push({
                        requestDate: otDate.toLocaleDateString('th-TH'), 
                        type: 'OT',
                        dateRange: `${startTimeString} - ${endTimeString}`, // This will now be "11:29:00 - 12:30:00"
                        total: `${Number(row[5]).toFixed(2)} ชม.`,
                        status: statusMap[row[7]] || row[7],
                        statusClass: row[7]
                    });
                }
            }
            // ** CHANGE END **
        }
        
        return { success: true, leaveHistory: userLeaveHistory.reverse(), otHistory: userOTHistory.reverse() };
    } catch(e) {
        Logger.log("Get History Error: " + e.toString());
        return { success: false, message: e.message, leaveHistory: [], otHistory: [] };
    }
}

// ===== OT APPROVAL WORKFLOW FUNCTIONS =====
function approveOTRequest(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Pending HR";
    const messageForEmployee = `คำขอ OT (ID: ${requestId}) ได้รับการอนุมัติจากหัวหน้าแล้ว รอฝ่ายบุคคลตรวจสอบ`;
    return _updateOTStatus(requestId, nextStatus, messageForEmployee, true);
}

function rejectOTRequest(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Rejected";
    const messageForEmployee = `คำขอ OT (ID: ${requestId}) ถูกปฏิเสธโดยหัวหน้างาน`;
    return _updateOTStatus(requestId, nextStatus, messageForEmployee, false);
}

function finalizeOTApproval(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Approved";
    const messageForEmployee = `คำขอ OT ของคุณ (ID: ${requestId}) ได้รับการอนุมัติเรียบร้อยแล้ว`;
    return _updateOTStatus(requestId, nextStatus, messageForEmployee, false);
}

function finalizeOTRejection(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Rejected";
    const messageForEmployee = `คำขอ OT ของคุณ (ID: ${requestId}) ถูกปฏิเสธโดยฝ่ายบุคคล`;
    return _updateOTStatus(requestId, nextStatus, messageForEmployee, false);
}

function _updateOTStatus(requestId, newStatus, messageForEmployee, notifyHR) {
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const otSheet = ss.getSheetByName("OTRequests");
        
        const textFinder = otSheet.getRange("A:A").createTextFinder(requestId);
        const foundCell = textFinder.findNext();
        if (!foundCell) return { success: false, message: "ไม่พบคำขอ OT (ID: " + requestId + ")" };
        
        const targetRowIndex = foundCell.getRow();
        const requestData = otSheet.getRange(targetRowIndex, 1, 1, otSheet.getLastColumn()).getValues()[0];
        const employeeId = requestData[1];
        const otDate = new Date(requestData[2]).toLocaleDateString('th-TH');
        
        otSheet.getRange(targetRowIndex, 8).setValue(newStatus);
        otSheet.getRange(targetRowIndex, 10).setValue(new Date().toISOString());

        _createNotification(ss, employeeId, messageForEmployee, requestId);

        if (notifyHR) {
             const empSheet = ss.getSheetByName("Employees");
             const empData = empSheet.getDataRange().getValues();
             const empHeaders = empData[0];
             const idIndex = empHeaders.indexOf("EmployeeID");
             const titleIndex = empHeaders.indexOf("Title");
             const firstNameIndex = empHeaders.indexOf("FirstName");
             const lastNameIndex = empHeaders.indexOf("LastName");
             
             const employeeNameRow = empData.find(row => row[idIndex] === employeeId);
             const employeeName = employeeNameRow ? `${employeeNameRow[titleIndex]}${employeeNameRow[firstNameIndex]} ${employeeNameRow[lastNameIndex]}` : employeeId;

             const messageForHR = `มีคำขอ OT วันที่ ${otDate} จากคุณ ${employeeName} รอการตรวจสอบ`;
             _notifyHRUsers(ss, messageForHR, requestId);
             
             const lineMessage = `คำขอ OT ของคุณ ${employeeName} ได้รับการอนุมัติจากหัวหน้าแล้ว ขณะนี้คำขอถูกส่งให้ฝ่ายบุคคล (HR) ตรวจสอบ`;
             _sendLineNotify(lineMessage);
        }
        
        return { success: true, message: `ดำเนินการสำเร็จ! อัปเดตสถานะเป็น ${newStatus}` };
    } catch(e) {
        return { success: false, message: "เกิดข้อผิดพลาดในการอัปเดตข้อมูล: " + e.message };
    }
}


//==================================================================================
// OTHER FUNCTIONS
//==================================================================================

function getLeaveTypes(token) {
    const user = _validateToken(token);
    if (!user) return { success: false, types: [] };
    
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const typeSheet = ss.getSheetByName("LeaveTypes");
        if (typeSheet.getLastRow() < 2) return { success: true, types: [] };
        
        const typeData = typeSheet.getRange(2, 1, typeSheet.getLastRow() - 1, typeSheet.getLastColumn()).getValues();
        const activeTypes = typeData
            .filter(row => row[8] === true)
            .map(row => ({
                id: row[0],
                name: row[1],
                category: row[2],
                requireDoc: row[6]
            }));
            
        return { success: true, types: activeTypes };
    } catch (e) {
        Logger.log(`getLeaveTypes Error: ${e.message}`);
        return { success: false, types: [] };
    }
}

function getHRManagementData(token) {
    const user = _validateToken(token);
    if (!user || user.role !== 'HR') return { success: false, message: 'ไม่มีสิทธิ์เข้าถึง' };
    
    const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
    const empSheet = ss.getSheetByName("Employees");
    const empData = empSheet.getDataRange().getValues();
    const headers = empData.shift();
    const deptIndex = headers.indexOf("Department");
    const roleIndex = headers.indexOf("Role");
    const idIndex = headers.indexOf("EmployeeID");
    const titleIndex = headers.indexOf("Title");
    const firstNameIndex = headers.indexOf("FirstName");
    const lastNameIndex = headers.indexOf("LastName");

    const empList = empData.map(row => {
        let empObject = {};
        headers.forEach((header, index) => {
            if (header !== "PasswordHash") { empObject[header] = row[index]; }
        });
        empObject.FullName = `${row[titleIndex]}${row[firstNameIndex]} ${row[lastNameIndex]}`;
        return empObject;
    });

    const departments = [...new Set(empData.map(row => row[deptIndex]).filter(Boolean))];
    const potentialManagers = empData
        .filter(row => row[roleIndex] === 'Manager' || row[roleIndex] === 'Supervisor')
        .map(row => ({ 
            id: row[idIndex], 
            name: `${row[titleIndex]}${row[firstNameIndex]} ${row[lastNameIndex]}`, 
            department: row[deptIndex] 
        }));

    return { success: true, employees: empList, departments: departments, managers: potentialManagers };
}

function addNewEmployee(token, newEmployeeData) {
    const user = _validateToken(token);
    if (!user || user.role !== 'HR') return { success: false, message: 'ไม่มีสิทธิ์เข้าถึง' };
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const empSheet = ss.getSheetByName("Employees");
        
        const lastRow = empSheet.getLastRow();
        const existingUsers = lastRow > 1 ? empSheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
        if (existingUsers.some(row => row[0] === newEmployeeData.EmployeeID)) return { success: false, message: "รหัสพนักงานนี้มีอยู่แล้วในระบบ" };
        if (existingUsers.some(row => row[1] === newEmployeeData.Username)) return { success: false, message: "Username นี้มีอยู่แล้วในระบบ" };
        
        const defaultPassword = newEmployeeData.EmployeeID + "@pass";
        const passwordHash = _hashPasswordForNewUser(defaultPassword);
        
        const fullName = `${newEmployeeData.Title}${newEmployeeData.FirstName} ${newEmployeeData.LastName}`;
        const newRow = [
            newEmployeeData.EmployeeID, newEmployeeData.Username, passwordHash,
            newEmployeeData.Title, newEmployeeData.FirstName, newEmployeeData.LastName, newEmployeeData.Nickname,
            newEmployeeData.Role, newEmployeeData.Department, newEmployeeData.ManagerID,
            newEmployeeData.StartDate, newEmployeeData.EmploymentStatus,
            newEmployeeData.Quota_LT001, newEmployeeData.Quota_LT002, newEmployeeData.Quota_LT003
        ];
        empSheet.appendRow(newRow);
        
        return { success: true, message: `เพิ่มพนักงาน ${fullName} สำเร็จ\nรหัสผ่านเริ่มต้นคือ: ${defaultPassword}` };
    } catch (e) {
        return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
    }
}

function _hashPasswordForNewUser(password) {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
    return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function _getLeaveBalance(spreadsheet, employeeId) {
    const empSheet = spreadsheet.getSheetByName("Employees");
    const leaveReqSheet = spreadsheet.getSheetByName("LeaveRequests");
    
    const empData = empSheet.getDataRange().getValues();
    const headers = empData[0];
    const empHeaders = {
        employeeID: headers.indexOf("EmployeeID"),
        sickQuota: headers.indexOf("Quota_LT001"),
        businessQuota: headers.indexOf("Quota_LT002"),
        vacationQuota: headers.indexOf("Quota_LT003")
    };

    const employeeRow = empData.find(row => row[empHeaders.employeeID] === employeeId);
    if (!employeeRow) return { LT001: {}, LT002: {}, LT003: {} };
    
    const leaveData = leaveReqSheet.getLastRow() > 1 ? leaveReqSheet.getRange(2, 1, leaveReqSheet.getLastRow() - 1, leaveReqSheet.getLastColumn()).getValues() : [];
    
    const usedSick = leaveData.filter(r => r[1] === employeeId && r[2] === "LT001" && r[7] === "Approved").reduce((sum, r) => sum + Number(r[5]), 0);
    const usedBusiness = leaveData.filter(r => r[1] === employeeId && r[2] === "LT002" && r[7] === "Approved").reduce((sum, r) => sum + Number(r[5]), 0);
    const usedVacation = leaveData.filter(r => r[1] === employeeId && r[2] === "LT003" && r[7] === "Approved").reduce((sum, r) => sum + Number(r[5]), 0);
  
    return {
      LT001: { quota: Number(employeeRow[empHeaders.sickQuota] || 0), used: usedSick },
      LT002: { quota: Number(employeeRow[empHeaders.businessQuota] || 0), used: usedBusiness },
      LT003: { quota: Number(employeeRow[empHeaders.vacationQuota] || 0), used: usedVacation }
    };
}

function submitLeaveRequest(token, leaveData) {
  const user = _validateToken(token);
  if (!user) return { success: false, message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่" };

  try {
    const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
    
    const startDate = new Date(leaveData.startDate);
    const endDate = new Date(leaveData.endDate);
    const totalDays = Math.round((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
    if (totalDays <= 0) return { success: false, message: "จำนวนวันลาไม่ถูกต้อง" };
    
    if (["LT001", "LT002", "LT003"].includes(leaveData.leaveTypeID)) {
        const balance = _getLeaveBalance(ss, user.employeeId);
        const remaining = balance[leaveData.leaveTypeID].quota - balance[leaveData.leaveTypeID].used;
        if(totalDays > remaining) {
          return { success: false, message: `วันลาประเภทนี้คงเหลือไม่พอ (เหลือ ${remaining} วัน)` };
        }
    }
    
    const empSheet = ss.getSheetByName("Employees");
    const empValues = empSheet.getDataRange().getValues();
    const empHeaders = empValues[0];
    const userRow = empValues.find(row => row[empHeaders.indexOf("EmployeeID")] === user.employeeId);
    const managerId = userRow[empHeaders.indexOf("ManagerID")];
    
    const nextStatus = managerId ? "Pending Manager" : "Pending HR";

    const leaveReqSheet = ss.getSheetByName("LeaveRequests");
    const newRequestId = "LR" + Utilities.getUuid().substring(0, 5).toUpperCase();
    leaveReqSheet.appendRow([ newRequestId, user.employeeId, leaveData.leaveTypeID, startDate.toISOString(), endDate.toISOString(), totalDays, leaveData.reason, nextStatus, new Date().toISOString(), new Date().toISOString(), "" ]);
    
    const typeSheet = ss.getSheetByName("LeaveTypes");
    const typeData = typeSheet.getDataRange().getValues();
    const typeMap = new Map(typeData.map(row => [row[0], row[1]]));
    const leaveTypeName = typeMap.get(leaveData.leaveTypeID) || leaveData.leaveTypeID;
    
    if (managerId) {
      _createNotification(ss, managerId, `มีคำขออนุมัติ (${leaveTypeName}) จากคุณ ${user.fullName}`, newRequestId);
      const managerRow = empValues.find(row => row[empHeaders.indexOf("EmployeeID")] === managerId);
      const managerName = managerRow ? `${managerRow[empHeaders.indexOf("Title")]}${managerRow[empHeaders.indexOf("FirstName")]} ${managerRow[empHeaders.indexOf("LastName")]}` : 'N/A';
      const lineMessage = `\n== ใบลาใหม่รออนุมัติ ==\nผู้ลา: ${user.fullName}\nประเภทลา: ${leaveTypeName}\nวันที่ลา: ${startDate.toLocaleDateString('th-TH')} - ${endDate.toLocaleDateString('th-TH')}\nจำนวนวันลา: ${totalDays} วัน\nรบกวน คุณ${managerName} ตรวจสอบใน App ด้วยครับ`;
      _sendLineNotify(lineMessage.trim());
    } else {
      const messageForHR = `มีคำขอลา (${leaveTypeName}) จากคุณ ${user.fullName} (ไม่มีหัวหน้า) รอการตรวจสอบ`;
      _notifyHRUsers(ss, messageForHR, newRequestId);
      const lineMessage = `\n== ใบลาใหม่รอ HR ตรวจสอบ ==\nผู้ลา: ${user.fullName} (ไม่มีหัวหน้า)\nประเภทลา: ${leaveTypeName}\nจำนวนวันลา: ${totalDays} วัน\nคำขอถูกส่งให้ HR โดยตรง`;
      _sendLineNotify(lineMessage.trim());
    }

    return { success: true, message: "ยื่นใบลาสำเร็จ! รอการอนุมัติ" };
  } catch (e) {
    Logger.log(`submitLeaveRequest Error: ${e.toString()}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  }
}

function submitLeaveRequestWithAttachment(token, leaveData, fileObject) {
  const user = _validateToken(token);
  if (!user) return { success: false, message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่" };

  try {
    const attachmentUrl = _uploadFileToDrive(fileObject);
    if (!attachmentUrl) {
      return { success: false, message: "เกิดข้อผิดพลาดในการอัปโหลดไฟล์แนบ" };
    }

    const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
    
    const startDate = new Date(leaveData.startDate);
    const endDate = new Date(leaveData.endDate);
    const totalDays = Math.round((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
    if (totalDays <= 0) return { success: false, message: "จำนวนวันลาไม่ถูกต้อง" };
    
    if (["LT001", "LT002", "LT003"].includes(leaveData.leaveTypeID)) {
        const balance = _getLeaveBalance(ss, user.employeeId);
        const remaining = balance[leaveData.leaveTypeID].quota - balance[leaveData.leaveTypeID].used;
        if(totalDays > remaining) {
          return { success: false, message: `วันลาประเภทนี้คงเหลือไม่พอ (เหลือ ${remaining} วัน)` };
        }
    }
    
    const empSheet = ss.getSheetByName("Employees");
    const empValues = empSheet.getDataRange().getValues();
    const empHeaders = empValues[0];
    const userRow = empValues.find(row => row[empHeaders.indexOf("EmployeeID")] === user.employeeId);
    const managerId = userRow[empHeaders.indexOf("ManagerID")];

    const nextStatus = managerId ? "Pending Manager" : "Pending HR";

    const leaveReqSheet = ss.getSheetByName("LeaveRequests");
    const newRequestId = "LR" + Utilities.getUuid().substring(0, 5).toUpperCase();
    leaveReqSheet.appendRow([ newRequestId, user.employeeId, leaveData.leaveTypeID, startDate.toISOString(), endDate.toISOString(), totalDays, leaveData.reason, nextStatus, new Date().toISOString(), new Date().toISOString(), attachmentUrl ]);
    
    const typeSheet = ss.getSheetByName("LeaveTypes");
    const typeData = typeSheet.getDataRange().getValues();
    const typeMap = new Map(typeData.map(row => [row[0], row[1]]));
    const leaveTypeName = typeMap.get(leaveData.leaveTypeID) || leaveData.leaveTypeID;
    
    if (managerId) {
      _createNotification(ss, managerId, `มีคำขออนุมัติ (${leaveTypeName}) จากคุณ ${user.fullName}`, newRequestId);
      const managerRow = empValues.find(row => row[empHeaders.indexOf("EmployeeID")] === managerId);
      const managerName = managerRow ? `${managerRow[empHeaders.indexOf("Title")]}${managerRow[empHeaders.indexOf("FirstName")]} ${managerRow[empHeaders.indexOf("LastName")]}` : 'N/A';
      const lineMessage = `\n== ใบลาใหม่รออนุมัติ ==\nผู้ลา: ${user.fullName}\nประเภทลา: ${leaveTypeName}\nวันที่ลา: ${startDate.toLocaleDateString('th-TH')} - ${endDate.toLocaleDateString('th-TH')}\nจำนวนวันลา: ${totalDays} วัน\n**มีไฟล์แนบ**\nรบกวน คุณ${managerName} ตรวจสอบใน App ด้วยครับ`;
      _sendLineNotify(lineMessage.trim());
    } else {
      const messageForHR = `มีคำขอลา (${leaveTypeName}) จากคุณ ${user.fullName} (ไม่มีหัวหน้า) รอการตรวจสอบ`;
      _notifyHRUsers(ss, messageForHR, newRequestId);
      const lineMessage = `\n== ใบลาใหม่รอ HR ตรวจสอบ ==\nผู้ลา: ${user.fullName} (ไม่มีหัวหน้า)\nประเภทลา: ${leaveTypeName}\nจำนวนวันลา: ${totalDays} วัน\n**มีไฟล์แนบ**\nคำขอถูกส่งให้ HR โดยตรง`;
      _sendLineNotify(lineMessage.trim());
    }

    return { success: true, message: "ยื่นใบลาพร้อมไฟล์แนบสำเร็จ! รอการอนุมัติ" };
  } catch (e) {
    Logger.log(`submitLeaveRequestWithAttachment Error: ${e.toString()}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  }
}

function submitOTRequest(token, otData) {
  const user = _validateToken(token);
  if (!user) return { success: false, message: "Session หมดอายุ กรุณาเข้าสู่ระบบใหม่" };

  try {
    if (!otData.otDate || !otData.startTime || !otData.endTime || !otData.reason) {
      return { success: false, message: "กรุณากรอกข้อมูล OT ให้ครบถ้วน" };
    }

    const start = new Date(`${otData.otDate}T${otData.startTime}:00`);
    const end = new Date(`${otData.otDate}T${otData.endTime}:00`);
    if (end <= start) { return { success: false, message: "เวลาสิ้นสุดต้องอยู่หลังเวลาเริ่มต้น" }; }
    const totalHours = (end.getTime() - start.getTime()) / (1000 * 60 * 60);

    const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
    const empSheet = ss.getSheetByName("Employees");
    const empValues = empSheet.getDataRange().getValues();
    const empHeaders = empValues[0];
    const userRow = empValues.find(row => row[empHeaders.indexOf("EmployeeID")] === user.employeeId);
    const managerId = userRow[empHeaders.indexOf("ManagerID")];
    
    const nextStatus = managerId ? "Pending Manager" : "Pending HR";
    
    const otReqSheet = ss.getSheetByName("OTRequests");
    const newRequestId = "OT" + Utilities.getUuid().substring(0, 5).toUpperCase();
    const now = new Date().toISOString();

    otReqSheet.appendRow([ newRequestId, user.employeeId, otData.otDate, otData.startTime, otData.endTime, totalHours.toFixed(2), otData.reason, nextStatus, now, now ]);
    
    const otDateFormatted = new Date(otData.otDate).toLocaleDateString('th-TH', { dateStyle: 'long' });
    
    if (managerId) {
      const message = `มีคำขออนุมัติ OT วันที่ ${otDateFormatted} จากคุณ ${user.fullName}`;
      _createNotification(ss, managerId, message, newRequestId);
      _sendLineNotify(message);
    } else {
      const message = `มีคำขอ OT จากคุณ ${user.fullName} (ไม่มีหัวหน้า) รอการตรวจสอบ`;
      _notifyHRUsers(ss, message, newRequestId);
      _sendLineNotify(message);
    }
    
    return { success: true, message: "บันทึกข้อมูล OT สำเร็จ! รอการอนุมัติ" };
  } catch (e) {
    Logger.log(`submitOTRequest Error: ${e.toString()}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  }
}

function approveLeaveRequest(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Pending HR";
    const messageForEmployee = `ใบลา (ID: ${requestId}) ได้รับการอนุมัติจากหัวหน้าแล้ว รอฝ่ายบุคคลตรวจสอบ`;
    return _updateLeaveStatus(requestId, nextStatus, messageForEmployee, true);
}

function rejectLeaveRequest(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Rejected";
    const messageForEmployee = `ใบลา (ID: ${requestId}) ถูกปฏิเสธโดยหัวหน้างาน`;
    return _updateLeaveStatus(requestId, nextStatus, messageForEmployee, false);
}

function finalizeApproval(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Approved";
    const messageForEmployee = `ใบลาของคุณ (ID: ${requestId}) ได้รับการอนุมัติเรียบร้อยแล้ว`;
    return _updateLeaveStatus(requestId, nextStatus, messageForEmployee, false);
}

function finalizeRejection(token, requestId) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    const nextStatus = "Rejected";
    const messageForEmployee = `ใบลาของคุณ (ID: ${requestId}) ถูกปฏิเสธโดยฝ่ายบุคคล`;
    return _updateLeaveStatus(requestId, nextStatus, messageForEmployee, false);
}

function _updateLeaveStatus(requestId, newStatus, messageForEmployee, notifyHR, approver) {
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const leaveSheet = ss.getSheetByName("LeaveRequests");
        
        const textFinder = leaveSheet.getRange("A:A").createTextFinder(requestId);
        const foundCell = textFinder.findNext();
        if (!foundCell) return { success: false, message: "ไม่พบใบลาที่ต้องการ (ID: " + requestId + ")" };
        
        const targetRowIndex = foundCell.getRow();
        const requestData = leaveSheet.getRange(targetRowIndex, 1, 1, leaveSheet.getLastColumn()).getValues()[0];
        const employeeId = requestData[1]; // The person who requested leave
        const leaveTypeId = requestData[2];
        
        leaveSheet.getRange(targetRowIndex, 8).setValue(newStatus);
        leaveSheet.getRange(targetRowIndex, 10).setValue(new Date().toISOString());

        // This notification to the employee is always correct
        _createNotification(ss, employeeId, messageForEmployee, requestId);

        if (notifyHR) {
             const empSheet = ss.getSheetByName("Employees");
             const empData = empSheet.getDataRange().getValues();
             const headers = empData[0];
             const idIndex = headers.indexOf("EmployeeID");
             const titleIndex = headers.indexOf("Title");
             const firstNameIndex = headers.indexOf("FirstName");
             const lastNameIndex = headers.indexOf("LastName");
             
             // Find the name of the employee who requested leave
             const employeeRow = empData.find(row => row[idIndex] === employeeId);
             const employeeName = employeeRow ? `${employeeRow[titleIndex]}${employeeRow[firstNameIndex]} ${employeeRow[lastNameIndex]}` : employeeId;

             const typeSheet = ss.getSheetByName("LeaveTypes");
             const typeData = typeSheet.getDataRange().getValues();
             const typeMap = new Map(typeData.map(row => [row[0], row[1]]));
             const leaveTypeName = typeMap.get(leaveTypeId) || leaveTypeId;

             // Create in-app notification for HR
             const messageForHR = `มีคำขอลา (${leaveTypeName}) จากคุณ ${employeeName} รอการตรวจสอบ`;
             _notifyHRUsers(ss, messageForHR, requestId);

             // ** BUG FIX START: Correctly identify the manager and send LINE notification **
             try {
              // The approver's name is simply the full name from the 'approver' object passed into this function.
              const managerName = approver.fullName;
              
              const lineMessage = `คุณ ${managerName} ได้อนุมัติใบลาของคุณ ${employeeName} แล้ว ขณะนี้คำขอถูกส่งให้ฝ่ายบุคคล (HR) ตรวจสอบในลำดับถัดไป`;
              _sendLineNotify(lineMessage);
             } catch(lineError) {
              // This will log if there's an error sending the notification itself.
              Logger.log(`Failed to send LINE notification for manager approval on ${requestId}: ${lineError.message}`);
             }
             // ** BUG FIX END **
        }
        
        return { success: true, message: `ดำเนินการสำเร็จ! อัปเดตสถานะเป็น ${newStatus}` };
    } catch(e) {
        return { success: false, message: "เกิดข้อผิดพลาดในการอัปเดตข้อมูล: " + e.message };
    }
}


function _createNotification(spreadsheet, targetUserId, message, linkRequestId) {
    const notiSheet = spreadsheet.getSheetByName("Notifications");
    const newNotiId = "NOTI" + Utilities.getUuid().substring(0, 5).toUpperCase();
    const expiryDate = new Date();
    expiryDate.setDate(expiryDate.getDate() + 30);
    const now = new Date().toISOString();
    notiSheet.appendRow([ newNotiId, targetUserId, message, linkRequestId, "Unread", now, expiryDate.toISOString() ]);
}

function getAllNotifications(token) {
    const user = _validateToken(token);
    if (!user) return []; 
    const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
    const notiSheet = ss.getSheetByName("Notifications");
    if(notiSheet.getLastRow() < 2) return [];
    const notiData = notiSheet.getRange(2, 1, notiSheet.getLastRow() - 1, notiSheet.getLastColumn()).getValues();
    const userNotifications = [];
    const now = new Date();
    for (const row of notiData) {
        if (row[1] === user.employeeId && new Date(row[6]) > now) {
            userNotifications.push({
                notificationId: row[0], message: row[2], linkToRequestId: row[3], status: row[4],
                createdDate: new Date(row[5])
            });
        }
    }
    userNotifications.sort((a,b) => b.createdDate.getTime() - a.createdDate.getTime());
    return userNotifications.map(n => ({
      ...n,
      createdDate: n.createdDate.toLocaleString('th-TH', { dateStyle: 'medium', timeStyle: 'short' })
    }));
}

function markNotificationsAsRead(token, notificationIds) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Invalid session" };
    if (!notificationIds || notificationIds.length === 0) return { success: true };
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const notiSheet = ss.getSheetByName("Notifications");
        const dataRange = notiSheet.getRange("A:E");
        const data = dataRange.getValues();
        for (let i = 1; i < data.length; i++) {
            if (notificationIds.includes(data[i][0])) {
                notiSheet.getRange(i + 1, 5).setValue("Read");
            }
        }
        return { success: true };
    } catch(e) {
        return { success: false, message: e.message };
    }
}

function updateEmployeeData(token, employeeUpdate) {
    const user = _validateToken(token);
    if (!user || user.role !== 'HR') return { success: false, message: 'ไม่มีสิทธิ์เข้าถึง' };
    
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const empSheet = ss.getSheetByName("Employees");
        const empData = empSheet.getDataRange().getValues();
        const headers = empData[0];
        const targetRowIndex = empData.findIndex(row => row[headers.indexOf("EmployeeID")] === employeeUpdate.EmployeeID);
        
        if (targetRowIndex !== -1) {
            const rowNumber = targetRowIndex + 1;
            headers.forEach((header, index) => {
                if (employeeUpdate.hasOwnProperty(header) && header !== "EmployeeID") {
                    empSheet.getRange(rowNumber, index + 1).setValue(employeeUpdate[header]);
                }
            });
            const fullName = `${employeeUpdate.Title}${employeeUpdate.FirstName} ${employeeUpdate.LastName}`;
            return { success: true, message: `อัปเดตข้อมูลคุณ ${fullName} สำเร็จ` };
        } else {
            return { success: false, message: 'ไม่พบพนักงานที่ต้องการแก้ไข' };
        }
    } catch(e) {
        return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
    }
}

function changeUserPassword(token, oldPassword, newPassword) {
    const user = _validateToken(token);
    if (!user) return { success: false, message: "Session ไม่ถูกต้อง" };
    
    try {
        const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
        const empSheet = ss.getSheetByName("Employees");
        const empData = empSheet.getDataRange().getValues();
        const headers = empData.shift();
        const idIndex = headers.indexOf("EmployeeID");
        const hashIndex = headers.indexOf("PasswordHash");

        const userRowIndex = empData.findIndex(row => row[idIndex] === user.employeeId);

        if (userRowIndex === -1) {
            return { success: false, message: "ไม่พบข้อมูลผู้ใช้ในระบบ" };
        }

        const storedHash = empData[userRowIndex][hashIndex];
        const oldPasswordHash = _hashPasswordForNewUser(oldPassword);

        if (storedHash !== oldPasswordHash) {
            return { success: false, message: "รหัสผ่านปัจจุบันไม่ถูกต้อง" };
        }
        
        const newPasswordHash = _hashPasswordForNewUser(newPassword);
        empSheet.getRange(userRowIndex + 2, hashIndex + 1).setValue(newPasswordHash);
        
        return { success: true, message: "เปลี่ยนรหัสผ่านสำเร็จ!" };

    } catch (e) {
        Logger.log("Change Password Error: " + e.toString());
        return { success: false, message: "เกิดข้อผิดพลาดในการเปลี่ยนรหัสผ่าน" };
    }
}

function cleanupOldNotifications() {
    const ss = SpreadsheetApp.openById("1wI1mBucukSkHOfsxrh5iDDbqTRrISgzhAbpcqZV7MyQ");
    const notiSheet = ss.getSheetByName("Notifications");
    if (notiSheet.getLastRow() < 2) return;
    const data = notiSheet.getDataRange().getValues();
    const rowsToDelete = [];
    const now = new Date();
    for (let i = data.length - 1; i > 0; i--) {
        const status = data[i][4];
        const expiryDate = new Date(data[i][6]);
        const createdDate = new Date(data[i][5]);
        const isExpired = expiryDate < now;
        const isReadAndOld = (status === 'Read' && (now.getTime() - createdDate.getTime()) > (30 * 24 * 60 * 60 * 1000));
        if (isExpired || isReadAndOld) { rowsToDelete.push(i + 1); }
    }
    rowsToDelete.reverse().forEach(rowIndex => notiSheet.deleteRow(rowIndex));
}

function _sendLineNotify(message) {
  const token = "RtiBfQaVenMHb5JE/+jieZuFndO0aAOY2w4I+J1YJW8msAcGhsM/MAn1tEijoumeRinv47HGm4AHs+wLigNtdVQkX+82sbTSm1sDRzmWsicoyks/NuXZrDNKb3MrI82Mdoi734E8Z40kx6TNq/9nMwdB04t89/1O/w1cDnyilFU=";
  const groupId = "Cf0e827b2e64c73aa1250800b4c2e0ae9";
  
  if (!token || !groupId) {
    Logger.log("Line token or Group ID is not set.");
    return;
  }
  
  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    "to": groupId,
    "messages": [{ "type": "text", "text": message }]
  };

  const options = {
    "method": "post",
    "headers": { "Content-Type": "application/json", "Authorization": "Bearer " + token },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("Error sending Line notification: " + e.toString());
  }
}

const ATTACHMENT_FOLDER_ID = "1q1hhSTcV5W-6CEOuX6HPGIwDq0CKYwIY";

function _uploadFileToDrive(fileObject) {
  try {
    const folder = DriveApp.getFolderById(ATTACHMENT_FOLDER_ID);
    const decoded = Utilities.base64Decode(fileObject.base64Data, Utilities.Charset.UTF_8);
    const blob = Utilities.newBlob(decoded, fileObject.mimeType, fileObject.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log(`File Upload Error: ${e.toString()}`);
    return null;
  }
}
