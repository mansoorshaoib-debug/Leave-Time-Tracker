/* =========================================================
   OTHTECH Leave Portal — Google Apps Script (paste in Code.gs)

   Sheet "Employees":    A = User Name, B = Password, C = Slack Name (optional)
                         C sirf tab likhein jab Slack ka naam User Name se khud match na ho.
   Sheet "Leaves":       A = ID, B = Date, C = Reason, D = Extra Leaves
   Sheet "Extra Leaves": A = ID, B = Date (off day worked), C = Work Done, D = Days
                         (this tab is created automatically)

   Rules
   - Every row in "Leaves" that has a Date = 1 leave taken.
   - Every row in "Extra Leaves" = extra leave earned (Days, default 1),
     shown with its date on the website.
   - A plain number in "Leaves" column D is also added as extra leave
     (shown as "Added by admin", without a date).
   - Total leaves = 10 + all extra leaves.

   Security
   - Website ab poora data nahi mangti. Login par sirf us employee ka data
     wapas jata hai, aur password yahin (server par) check hota hai.
   - Admin ka username/password bhi yahin check hota hai (neeche ADMIN_*).

   Attendance
   - Attendance.gs file Slack se data laati hai aur "Daily Summary" tab banati hai.
   - Login par us employee ki attendance bhi sath bheji jati hai.
   - Overtime: "Overtime Months" + "Earned Leaves" (permanent) aur
     "Overtime History" (aakhri 2 haftay) bhi sath bheje jate hain.
   ========================================================= */

const SHEET_ID = "105gYa5WSjHynwFGgmc5xnkG3CohtYh1ek4_ZDXVxWYQ";
const EXTRA_TAB = "Extra Leaves";

// Admin login — Project Settings > Script Properties mein ADMIN_USER / ADMIN_PASS
// rakh dein to woh use honge; warna yeh purane wale.
const ADMIN_USER_DEFAULT = "OTHTECH";
const ADMIN_PASS_DEFAULT = "admin";

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_({ ok: false, error: "Please refresh the website (new version)." });
}

/* Website yahan POST karti hai: { action: "login", id, pass } ya { action: "admin", user, pass } */
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (err) {}
  try {
    if (body.action === "login") return json_(employeeLogin_(body));
    if (body.action === "admin") return json_(adminLogin_(body));
    return json_({ ok: false, error: "Unknown request" });
  } catch (err) {
    return json_({ ok: false, error: "Server error: " + err.message });
  }
}

function employeeLogin_(body) {
  const id = String(body.id || "").trim();
  const pass = String(body.pass || "").trim();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const all = buildEmployees_(ss);
  const emp = all[id];
  if (!emp) return { ok: false, error: "Employee not found" };
  if (emp.password !== pass) return { ok: false, error: "Incorrect password" };
  const att = typeof getAttendanceData_ === "function" ? getAttendanceData_(ss, id) : {};
  const ot = typeof getOvertimeData_ === "function" ? getOvertimeData_(ss, id) : {};
  return { ok: true, employee: publicEmp_(emp, att[id] || [], ot[id]) };
}

function adminLogin_(body) {
  const props = PropertiesService.getScriptProperties();
  const U = props.getProperty("ADMIN_USER") || ADMIN_USER_DEFAULT;
  const P = props.getProperty("ADMIN_PASS") || ADMIN_PASS_DEFAULT;
  if (String(body.user || "").trim() !== U || String(body.pass || "").trim() !== P) {
    return { ok: false, error: "Invalid admin credentials" };
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const all = buildEmployees_(ss);
  const att = typeof getAttendanceData_ === "function" ? getAttendanceData_(ss) : {};
  const ot = typeof getOvertimeData_ === "function" ? getOvertimeData_(ss) : {};
  const employees = {};
  Object.keys(all).forEach(id => employees[id] = publicEmp_(all[id], att[id] || [], ot[id]));
  const issues = typeof getOpenIssues_ === "function" ? getOpenIssues_(ss, 50) : [];
  return { ok: true, employees: employees, issues: issues };
}

/* Password kabhi browser ko nahi bheja jata */
function publicEmp_(emp, attendance, overtime) {
  return { name: emp.name, leaves: emp.leaves, extraLeaves: emp.extraLeaves,
    extraManual: emp.extraManual, attendance: attendance,
    overtime: overtime || null };   // Overtime Months + Earned Leaves + Overtime History (Attendance.gs)
}

/* Run this once from the editor (select it → ▶ Run) to create the tab right away.
   It is also created automatically the first time the website loads. */
function setupExtraLeavesSheet() {
  getExtraSheet_(SpreadsheetApp.openById(SHEET_ID));
}

function getExtraSheet_(ss) {
  let sh = ss.getSheetByName(EXTRA_TAB);
  if (sh) return sh;

  sh = ss.insertSheet(EXTRA_TAB);
  sh.getRange("A1:D1")
    .setValues([["ID", "Date", "Work Done", "Days"]])
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#000000")
    .setHorizontalAlignment("center");
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 320);
  sh.setColumnWidth(4, 80);
  sh.getRange("B2:B").setNumberFormat("d-mmm-yyyy");
  sh.getRange("D2:D").setHorizontalAlignment("center");
  sh.getRange("A2").setNote("Write the employee's User Name exactly as in the Employees tab.\nDays: leave empty for 1, or write 0.5 / 2 etc.");
  return sh;
}

function fmtDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  return (v || "").toString().trim();
}

function buildEmployees_(ss) {
  const tz = ss.getSpreadsheetTimeZone();

  const empData = ss.getSheetByName("Employees").getDataRange().getValues();
  const leaveData = ss.getSheetByName("Leaves").getDataRange().getValues();
  const extraData = getExtraSheet_(ss).getDataRange().getValues();

  const employees = {};

  // Employees (A: User Name, B: Password)
  for (let i = 1; i < empData.length; i++) {
    const id = (empData[i][0] || "").toString().trim();
    if (!id) continue;
    employees[id] = {
      name: id,
      password: (empData[i][1] || "").toString().trim(),
      leaves: [],
      extraLeaves: [],
      extraManual: 0
    };
  }

  // Leaves (A: ID, B: Date, C: Reason, D: Extra Leaves number)
  for (let j = 1; j < leaveData.length; j++) {
    const id = (leaveData[j][0] || "").toString().trim();
    const emp = employees[id];
    if (!emp) continue; // skips section headers like "SHAN Leaves" and blank rows

    const date = leaveData[j][1];
    if (date !== "" && date !== null) {
      emp.leaves.push({
        date: fmtDate_(date, tz),
        reason: (leaveData[j][2] || "").toString().trim()
      });
    }

    const extra = parseFloat(leaveData[j][3]);
    if (!isNaN(extra)) emp.extraManual += extra;
  }

  // Extra Leaves (A: ID, B: Date, C: Work Done, D: Days)
  for (let k = 1; k < extraData.length; k++) {
    const id = (extraData[k][0] || "").toString().trim();
    const emp = employees[id];
    if (!emp) continue;

    const days = parseFloat(extraData[k][3]);
    emp.extraLeaves.push({
      date: fmtDate_(extraData[k][1], tz),
      reason: (extraData[k][2] || "").toString().trim(),
      days: isNaN(days) ? 1 : days
    });
  }

  return employees;
}
