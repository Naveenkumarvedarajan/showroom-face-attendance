const CONFIG = {
  SHOP_NAME: "Sri Lakshmi Groups",
  SPREADSHEET_ID: "1U4Lel-7RrmDgJgW4I9oTrmOnvUqrL_WNzTeynUAH05g", // Optional: paste a NEW Google Spreadsheet ID. Blank creates one on first run.
  ADMIN_PIN: "8182", // Change before deployment. This is a basic gate, not enterprise authentication.
  TIMEZONE: "Asia/Kolkata",
  SHEETS: {
    STAFF: "Staff",
    ATTENDANCE: "Attendance",
    PAYROLL: "Payroll",
    SETTINGS: "Settings"
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle(CONFIG.SHOP_NAME + " | Staff Attendance")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function setupApp() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, CONFIG.SHEETS.STAFF,
    ["EmployeeID","Name","MonthlySalary","Active","FaceDescriptorJSON","CreatedAt"]);
  ensureSheet_(ss, CONFIG.SHEETS.ATTENDANCE,
    ["Timestamp","Date","EmployeeID","Name","Action","Source","Notes"]);
  ensureSheet_(ss, CONFIG.SHEETS.PAYROLL,
    ["Month","EmployeeID","Name","MonthlySalary","PresentDays","AbsentDays","PaidLeaveDays","OvertimeAmount","Advance","OtherDeductions","EstimatedPay","ApprovedPay","PaymentStatus","UpdatedAt"]);
  ensureSheet_(ss, CONFIG.SHEETS.SETTINGS, ["Key","Value"]);
  return {ok:true, spreadsheetUrl:ss.getUrl(), message:"Setup complete. Keep this spreadsheet private."};
}

function adminLogin(pin) {
  if (String(pin || "") !== String(CONFIG.ADMIN_PIN)) return {ok:false, message:"Incorrect admin PIN."};
  const ss = getSpreadsheet_();
  return {ok:true, shopName:CONFIG.SHOP_NAME, spreadsheetUrl:ss.getUrl(), staff:listStaff_()};
}

function registerStaff(pin, employeeId, name, monthlySalary, descriptor) {
  assertAdmin_(pin);
  employeeId = clean_(employeeId, 40);
  name = clean_(name, 100);
  monthlySalary = Number(monthlySalary);
  if (!employeeId || !name || !Number.isFinite(monthlySalary) || monthlySalary < 0) {
    throw new Error("Enter a staff ID, name and valid monthly salary.");
  }
  if (!descriptor || !Array.isArray(descriptor) || descriptor.length !== 128) {
    throw new Error("Face enrollment failed. Scan one clear face again.");
  }
  const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEETS.STAFF);
  const rows = sh.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === employeeId.toLowerCase()) {
      throw new Error("That Employee ID already exists.");
    }
  }
  sh.appendRow([employeeId, name, monthlySalary, true, JSON.stringify(descriptor), new Date()]);
  return {ok:true, message:"Staff member registered."};
}

function listStaff() {
  return listStaff_().map(s => ({employeeId:s.employeeId, name:s.name, monthlySalary:s.monthlySalary, active:s.active}));
}

function listStaff_() {
  const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEETS.STAFF);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues().map(r => ({
    employeeId:String(r[0]), name:String(r[1]), monthlySalary:Number(r[2])||0,
    active:r[3] === true || String(r[3]).toLowerCase() === "true",
    descriptor: parseDescriptor_(r[4])
  })).filter(s => s.employeeId && s.active);
}

function verifyAndMarkAttendance(employeeId, action, descriptor, threshold) {
  employeeId = clean_(employeeId, 40);
  action = String(action || "").toUpperCase();
  if (!["IN","OUT"].includes(action)) throw new Error("Choose Check In or Check Out.");
  if (!descriptor || !Array.isArray(descriptor) || descriptor.length !== 128) throw new Error("No valid face scan received.");
  const staff = listStaff_().find(s => s.employeeId.toLowerCase() === employeeId.toLowerCase());
  if (!staff) throw new Error("Staff ID not found or inactive.");
  const distance = euclideanDistance_(staff.descriptor, descriptor);
  const maxDistance = Math.min(0.65, Math.max(0.35, Number(threshold) || 0.48));
  if (!Number.isFinite(distance) || distance > maxDistance) {
    return {ok:false, message:"Face did not match. Try again in good lighting or ask an admin for manual verification."};
  }

  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const now = new Date();
  const today = Utilities.formatDate(now, CONFIG.TIMEZONE, "yyyy-MM-dd");
  const rows = sh.getDataRange().getValues();
  let lastAction = "";
  for (let i=rows.length-1; i>=1; i--) {
    if (String(rows[i][1]) === today && String(rows[i][2]).toLowerCase() === employeeId.toLowerCase()) {
      lastAction = String(rows[i][4]).toUpperCase();
      break;
    }
  }
  if (action === "IN" && lastAction === "IN") return {ok:false, message:"You are already checked in today."};
  if (action === "OUT" && lastAction !== "IN") return {ok:false, message:"No matching check-in found for today."};

  sh.appendRow([now, today, staff.employeeId, staff.name, action, "Face match", "Distance="+distance.toFixed(3)]);
  return {ok:true, message:staff.name + " checked " + (action === "IN" ? "in" : "out") + " successfully.", time:Utilities.formatDate(now, CONFIG.TIMEZONE, "dd MMM yyyy, hh:mm:ss a")};
}

function buildMonthlyPayroll(pin, month, paidLeaveDays, overtimeById, advanceById, deductionsById) {
  assertAdmin_(pin);
  if (!/^\d{4}-\d{2}$/.test(String(month))) throw new Error("Month must be YYYY-MM.");
  paidLeaveDays = Math.max(0, Number(paidLeaveDays)||0);
  overtimeById = overtimeById || {};
  advanceById = advanceById || {};
  deductionsById = deductionsById || {};
  const ss = getSpreadsheet_();
  const staff = listStaff_();
  const att = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE).getDataRange().getValues();
  const sh = ss.getSheetByName(CONFIG.SHEETS.PAYROLL);
  const out = [];
  staff.forEach(s => {
    const days = {};
    let firstIn = {};
    let lastOut = {};
    for (let i=1;i<att.length;i++) {
      const date = String(att[i][1] || "");
      const id = String(att[i][2] || "");
      const action = String(att[i][4] || "").toUpperCase();
      if (date.slice(0,7) !== month || id.toLowerCase() !== s.employeeId.toLowerCase()) continue;
      days[date] = days[date] || {in:false,out:false};
      if (action === "IN") days[date].in = true;
      if (action === "OUT") days[date].out = true;
    }
    const presentDays = Object.values(days).filter(d=>d.in).length;
    const daysInMonth = new Date(Number(month.slice(0,4)), Number(month.slice(5,7)), 0).getDate();
    const absentDays = Math.max(0, daysInMonth - presentDays - paidLeaveDays);
    const dailyRate = s.monthlySalary / daysInMonth;
    const overtime = Math.max(0, Number(overtimeById[s.employeeId])||0);
    const advance = Math.max(0, Number(advanceById[s.employeeId])||0);
    const deductions = Math.max(0, Number(deductionsById[s.employeeId])||0);
    // Simple estimate only. Configure company policy and statutory rules before paying staff.
    const estimated = Math.max(0, (dailyRate * Math.min(daysInMonth, presentDays + paidLeaveDays)) + overtime - advance - deductions);
    out.push([month,s.employeeId,s.name,s.monthlySalary,presentDays,absentDays,paidLeaveDays,overtime,advance,deductions,Math.round(estimated*100)/100,"","Pending",new Date()]);
  });
  // Replace rows for this month, preserving other months.
  const all = sh.getDataRange().getValues();
  const headers = all.length ? all[0] : [];
  const kept = all.slice(1).filter(r=>String(r[0]) !== month);
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (kept.length) sh.getRange(2,1,kept.length,headers.length).setValues(kept);
  if (out.length) sh.getRange(sh.getLastRow()+1,1,out.length,out[0].length).setValues(out);
  return {ok:true, rows:out.map(r=>({employeeId:r[1],name:r[2],salary:r[3],presentDays:r[4],absentDays:r[5],estimatedPay:r[10]})), message:"Payroll estimates generated. Review and approve before payment."};
}

function getAttendanceReport(pin, month) {
  assertAdmin_(pin);
  const sh = getSpreadsheet_().getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(r=>String(r[1]).slice(0,7)===month)
    .map(r=>({timestamp:Utilities.formatDate(new Date(r[0]),CONFIG.TIMEZONE,"dd MMM yyyy, hh:mm:ss a"),date:String(r[1]),employeeId:String(r[2]),name:String(r[3]),action:String(r[4]),source:String(r[5])}));
}

function assertAdmin_(pin) {
  if (String(pin || "") !== String(CONFIG.ADMIN_PIN)) throw new Error("Admin PIN required or incorrect.");
}
function parseDescriptor_(v) {
  try { const x=JSON.parse(String(v||"[]")); return Array.isArray(x) && x.length===128 ? x.map(Number) : []; }
  catch(e) { return []; }
}
function euclideanDistance_(a,b) {
  if (!Array.isArray(a) || !Array.isArray(b) || a.length!==128 || b.length!==128) return Infinity;
  let sum=0;
  for(let i=0;i<128;i++){ const d=Number(a[i])-Number(b[i]); sum+=d*d; }
  return Math.sqrt(sum);
}
function clean_(v,max) { return String(v||"").trim().replace(/[<>]/g,"").slice(0,max); }

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = CONFIG.SPREADSHEET_ID || props.getProperty("SPREADSHEET_ID");
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create(CONFIG.SHOP_NAME + " - Attendance and Payroll");
    props.setProperty("SPREADSHEET_ID", ss.getId());
  }
  ensureSheet_(ss, CONFIG.SHEETS.STAFF, ["EmployeeID","Name","MonthlySalary","Active","FaceDescriptorJSON","CreatedAt"]);
  ensureSheet_(ss, CONFIG.SHEETS.ATTENDANCE, ["Timestamp","Date","EmployeeID","Name","Action","Source","Notes"]);
  ensureSheet_(ss, CONFIG.SHEETS.PAYROLL, ["Month","EmployeeID","Name","MonthlySalary","PresentDays","AbsentDays","PaidLeaveDays","OvertimeAmount","Advance","OtherDeductions","EstimatedPay","ApprovedPay","PaymentStatus","UpdatedAt"]);
  ensureSheet_(ss, CONFIG.SHEETS.SETTINGS, ["Key","Value"]);
  return ss;
}
function ensureSheet_(ss,name,headers) {
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
  return sh;
}
/**
 * API entry point for the server-side proxy.
 * Never put API_SECRET in the public GitHub frontend.
 */
function doPost(e) {
  try {
    const body = JSON.parse(
      e && e.postData ? e.postData.contents : "{}"
    );

    const expectedSecret = PropertiesService
      .getScriptProperties()
      .getProperty("1XYoVpoysnk2Da67UqrQQwiF6mvNXVsBympp1_LmcZ__EMA_a3TAuxK3q");

    if (!expectedSecret ||
        typeof body.apiKey !== "string" ||
        body.apiKey !== expectedSecret) {
      return jsonResponse_({
        ok: false,
        message: "Unauthorized"
      });
    }

    if (body.action === "ping") {
      return jsonResponse_({
        ok: true,
        message: "Attendance API is running"
      });
    }

    if (body.action === "attendance") {
      const result = verifyAndMarkAttendance(
        body.employeeId,
        body.attendanceAction,
        body.descriptor,
        0.48
      );

      return jsonResponse_(result);
    }

    return jsonResponse_({
      ok: false,
      message: "Unsupported action"
    });

  } catch (error) {
    console.error(error);
    return jsonResponse_({
      ok: false,
      message: "Request failed"
    });
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
