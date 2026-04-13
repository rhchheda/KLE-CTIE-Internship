/*******************************
 * CTIE SYSTEM - MASTER SCRIPT
 *******************************/

/************ MENU ************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CTIE System")
    .addItem("Initialize System", "initializeSystem")
    .addToUi();
}

/************ INITIALIZER ************/
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createStudentMaster(ss);
  createSkillMaster(ss);
  createStartupRequestSheet(ss);
  createSelectionTrackerSheet(ss); // ✅ NEW

  protectStudentSheet();
  protectSkillSheet();
  protectStartupRequestSheet();

  SpreadsheetApp.getUi().alert("✅ System with Tracking Initialized");
}
function protectSelectionTracker() {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("SELECTION_TRACKER");

  const protection = sheet.protect();
  protection.removeEditors(protection.getEditors());

  protection.setUnprotectedRanges([sheet.getRange("A2:H1000")]);
}
function protectStartupRequestSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("STARTUP_REQUEST");

  const protection = sheet.protect();
  protection.removeEditors(protection.getEditors());

  protection.setUnprotectedRanges([sheet.getRange("A2:G1000")]);

  sheet.getRange("1:1").protect();
}

/************ STUDENT MASTER ************/
function createStudentMaster(ss) {
  const name = "STUDENT_MASTER";

  let sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet(name);

  const headers = [
    "StudentID","Name","Email","Phone","College",
    "Discipline","Year","CGPA","Skills","Tools",
    "Projects","Availability","ResumeLink","Timestamp"
  ];

  sheet.appendRow(headers);

  // Header styling
  const header = sheet.getRange(1,1,1,headers.length);
  header.setFontWeight("bold")
        .setBackground("#1a73e8")
        .setFontColor("white")
        .setHorizontalAlignment("center");

  // Freeze + filter
  sheet.setFrozenRows(1);
  header.createFilter();

  // Column widths
  const widths = [120,180,220,120,180,120,80,80,220,220,300,150,250,180];
  widths.forEach((w,i)=> sheet.setColumnWidth(i+1,w));

  // Alternating colors
  sheet.getRange("A2:N1000")
       .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // VALIDATIONS
  sheet.getRange("F2:F").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["CSE","ECE","ME","Civil"], true)
      .build()
  );

  sheet.getRange("G2:G").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["2","3","4"], true)
      .build()
  );

  sheet.getRange("L2:L").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Immediate","1 Month"], true)
      .build()
  );

  sheet.getRange("H2:H").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0,10)
      .build()
  );

  sheet.getRange("C2:C").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireTextIsEmail()
      .build()
  );

  sheet.getRange("N2:N").setNumberFormat("dd-mmm-yyyy hh:mm");

  // CONDITIONAL FORMATTING
  const rules = [];

  // High CGPA
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(8)
      .setBackground("#c6efce")
      .setRanges([sheet.getRange("H2:H")])
      .build()
  );

  // Low CGPA
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(5)
      .setBackground("#f4c7c3")
      .setRanges([sheet.getRange("H2:H")])
      .build()
  );

  // Missing Resume
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground("#fff3cd")
      .setRanges([sheet.getRange("M2:M")])
      .build()
  );

  // Missing Name/Email
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($B2="", $C2="")')
      .setBackground("#f8d7da")
      .setRanges([sheet.getRange("A2:N")])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  sheet.getRange("A:N").setVerticalAlignment("middle");
}

/************ SKILL MASTER ************/
function createSkillMaster(ss) {
  const name = "SKILL_MASTER";

  let sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet(name);

  sheet.appendRow(["Skill","Category"]);

  const header = sheet.getRange("A1:B1");
  header.setFontWeight("bold")
        .setBackground("#34a853")
        .setFontColor("white");

  const skills = [
    ["Python","Programming"],
    ["Java","Programming"],
    ["C++","Programming"],
    ["Embedded C","Embedded"],
    ["Arduino","Embedded"],
    ["MATLAB","Simulation"],
    ["AutoCAD","Mechanical"],
    ["SolidWorks","Mechanical"],
    ["Machine Learning","AI"],
    ["Data Analysis","Data"],
    ["HTML","Web"],
    ["CSS","Web"],
    ["JavaScript","Web"]
  ];

  sheet.getRange(2,1,skills.length,2).setValues(skills);

  sheet.setFrozenRows(1);
  sheet.getRange("A1:B1").createFilter();

  sheet.setColumnWidth(1,200);
  sheet.setColumnWidth(2,150);
}

/************ PROTECTION ************/
function protectStudentSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("STUDENT_MASTER");

  const protection = sheet.protect();
  protection.setDescription("Protected: Student Data");

  protection.removeEditors(protection.getEditors());

  const header = sheet.getRange("1:1");
  const headerProtection = header.protect();
  headerProtection.setDescription("Header Locked");

  protection.setUnprotectedRanges([sheet.getRange("A2:N1000")]);
}

function protectSkillSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("SKILL_MASTER");

  const protection = sheet.protect();
  protection.setDescription("Skill Master Locked");

  protection.removeEditors(protection.getEditors());
}
function createStartupRequestSheet(ss) {
  const name = "STARTUP_REQUEST";

  let sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet(name);

  const headers = [
    "RequestID","StartupName","Discipline","RequiredSkills",
    "MinCGPA","Positions","Status"
  ];

  sheet.appendRow(headers);

  const header = sheet.getRange("A1:G1");
  header.setFontWeight("bold")
        .setBackground("#673ab7")
        .setFontColor("white");

  sheet.setFrozenRows(1);
  header.createFilter();

  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,180);
  sheet.setColumnWidth(3,120);
  sheet.setColumnWidth(4,250);
  sheet.setColumnWidth(5,100);
  sheet.setColumnWidth(6,100);
  sheet.setColumnWidth(7,120);
}
function calculateSkillMatch(studentSkills, requiredSkills) {
  if (!studentSkills || !requiredSkills) return 0;

  const sSkills = studentSkills.toLowerCase().split(",").map(s => s.trim());
  const rSkills = requiredSkills.toLowerCase().split(",").map(s => s.trim());

  let matchCount = sSkills.filter(skill => rSkills.includes(skill)).length;

  return (matchCount / rSkills.length) * 100;
}
function calculateScore(student, request) {
  let score = 0;

  // 1. Skill Match (50%)
  const skillScore = calculateSkillMatch(student.skills, request.skills);
  score += skillScore * 0.5;

  // 2. Discipline Match (20%)
  if (student.discipline === request.discipline) {
    score += 20;
  }

  // 3. CGPA (15%)
  const cgpaScore = (student.cgpa / 10) * 15;
  score += cgpaScore;

  // 4. Availability (15%)
  if (student.availability === "Immediate") {
    score += 15;
  }

  return Math.round(score);
}
function getRankedStudents(requestID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const studentSheet = ss.getSheetByName("STUDENT_MASTER");
  const requestSheet = ss.getSheetByName("STARTUP_REQUEST");

  const students = studentSheet.getDataRange().getValues();
  const requests = requestSheet.getDataRange().getValues();

  const headers = students[0];
  const data = students.slice(1);

  // Find request
  const requestRow = requests.find(r => r[0] === requestID);
  if (!requestRow) throw new Error("Request not found");

  const request = {
    discipline: requestRow[2],
    skills: requestRow[3],
    minCGPA: requestRow[4]
  };

  let results = [];

  data.forEach(row => {
    const student = {
      id: row[0],
      name: row[1],
      email: row[2],
      discipline: row[5],
      cgpa: parseFloat(row[7]) || 0,
      skills: row[8],
      availability: row[11]
    };

    // Apply CGPA filter
    if (student.cgpa < request.minCGPA) return;

    const score = calculateScore(student, request);

    results.push({
      id: student.id,
      name: student.name,
      score: score,
      cgpa: student.cgpa,
      discipline: student.discipline
    });
  });

  // Sort descending
  results.sort((a, b) => b.score - a.score);

  return results;
}
function generateMatchResults(requestID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("MATCH_RESULTS");
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet("MATCH_RESULTS");

  const results = getRankedStudents(requestID);

  sheet.appendRow(["StudentID","Name","Score","CGPA","Discipline"]);

  const data = results.map(r => [
    r.id, r.name, r.score, r.cgpa, r.discipline
  ]);

  if (data.length > 0) {
    sheet.getRange(2,1,data.length,5).setValues(data);
  }

  // Formatting
  const header = sheet.getRange("A1:E1");
  header.setFontWeight("bold")
        .setBackground("#ff9800")
        .setFontColor("white");

  sheet.setFrozenRows(1);
  header.createFilter();

  sheet.autoResizeColumns(1,5);
}
function createSelectionTrackerSheet(ss) {
  const name = "SELECTION_TRACKER";

  let sheet = ss.getSheetByName(name);
  if (sheet) ss.deleteSheet(sheet);

  sheet = ss.insertSheet(name);

  const headers = [
    "SelectionID",
    "RequestID",
    "StudentID",
    "StudentName",
    "StartupName",
    "Status",
    "LastUpdated",
    "Remarks"
  ];

  sheet.appendRow(headers);

  // Header styling
  const header = sheet.getRange("A1:H1");
  header.setFontWeight("bold")
        .setBackground("#009688")
        .setFontColor("white");

  sheet.setFrozenRows(1);
  header.createFilter();

  sheet.setColumnWidth(1,130);
  sheet.setColumnWidth(2,130);
  sheet.setColumnWidth(3,130);
  sheet.setColumnWidth(4,180);
  sheet.setColumnWidth(5,180);
  sheet.setColumnWidth(6,150);
  sheet.setColumnWidth(7,180);
  sheet.setColumnWidth(8,250);

  // Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["APPLIED","SHORTLISTED","SELECTED","ONBOARDED","COMPLETED","REJECTED"],
      true
    )
    .build();

  sheet.getRange("F2:F").setDataValidation(statusRule);

  // Timestamp format
  sheet.getRange("G2:G").setNumberFormat("dd-mmm-yyyy hh:mm");

  // Conditional colors
  const rules = [];

  // Selected → Green
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("SELECTED")
      .setBackground("#c6efce")
      .setRanges([sheet.getRange("F2:F")])
      .build()
  );

  // Rejected → Red
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("REJECTED")
      .setBackground("#f4c7c3")
      .setRanges([sheet.getRange("F2:F")])
      .build()
  );

  // Shortlisted → Yellow
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("SHORTLISTED")
      .setBackground("#fff3cd")
      .setRanges([sheet.getRange("F2:F")])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
} 
function createSelectionEntries(requestID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const results = getRankedStudents(requestID);
  const tracker = ss.getSheetByName("SELECTION_TRACKER");

  const requestSheet = ss.getSheetByName("STARTUP_REQUEST");
  const requestRow = requestSheet.getDataRange().getValues()
    .find(r => r[0] === requestID);

  const startupName = requestRow[1];

  let rows = [];

  results.forEach((r, index) => {
    const selectionID = "SEL-" + (new Date().getTime() + index);

    rows.push([
      selectionID,
      requestID,
      r.id,
      r.name,
      startupName,
      "APPLIED",
      new Date(),
      ""
    ]);
  });

  if (rows.length > 0) {
    tracker.getRange(tracker.getLastRow()+1,1,rows.length,8).setValues(rows);
  }
}
function updateSelectionStatus(selectionID, newStatus, remarks="") {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("SELECTION_TRACKER");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === selectionID) {

      sheet.getRange(i+1,6).setValue(newStatus); // Status
      sheet.getRange(i+1,7).setValue(new Date()); // Timestamp
      sheet.getRange(i+1,8).setValue(remarks); // Remarks

      return "Updated Successfully";
    }
  }

  throw new Error("Selection ID not found");
}
function getFulfillmentStats(requestID) {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("SELECTION_TRACKER");

  const data = sheet.getDataRange().getValues();

  let stats = {
    applied: 0,
    shortlisted: 0,
    selected: 0,
    onboarded: 0
  };

  data.slice(1).forEach(row => {
    if (row[1] !== requestID) return;

    const status = row[5];

    if (status === "APPLIED") stats.applied++;
    if (status === "SHORTLISTED") stats.shortlisted++;
    if (status === "SELECTED") stats.selected++;
    if (status === "ONBOARDED") stats.onboarded++;
  });

  return stats;
}
function getStartupRequests() {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("STARTUP_REQUEST");

  const data = sheet.getDataRange().getValues();
  return data.slice(1);
}
function getCandidatesForRequest(requestID) {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("SELECTION_TRACKER");

  const data = sheet.getDataRange().getValues();

  return data.slice(1)
    .filter(row => row[1] === requestID)
    .map(row => ({
      selectionID: row[0],
      studentID: row[2],
      name: row[3],
      status: row[5],
      remarks: row[7]
    }));
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile("startup")
    .setTitle("Startup Dashboard");
}