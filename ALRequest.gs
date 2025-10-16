/**
 * Serves the HTML form for leave requests.
 */
const doGet = () => {
  return HtmlService.createHtmlOutputFromFile("ALRequestForm").setTitle(
    "Annual Leave Request Form"
  );
};

/**
 * Fetches employee data, colleagues, and team leads for the form.
 */
const getEmployeeData = (email) => {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic indexing ===
    const empCol = getColumnIndexes(empSheet);
    const employees = empSheet
      .getDataRange()
      .getValues()
      .slice(1)
      .filter((row) => (row[empCol["name"]] || "").toString().trim() !== "");

    // === 🔹 Find employee ===
    const emp = employees.find(
      (row) =>
        (row[empCol["email"]] || "").toString().trim().toLowerCase() ===
        email.toLowerCase()
    );

    if (!emp) throw new Error(`Employee with email "${email}" not found.`);

    // === 🔹 Colleagues (exclude self) ===
    const colleagues = employees
      .filter(
        (row) =>
          (row[empCol["email"]] || "").toString().trim().toLowerCase() !==
          email.toLowerCase()
      )
      .map((row) => (row[empCol["name"]] || "").toString().trim())
      .filter(Boolean);

    // === 🔹 Team Leads ===
    const teamLeads = getTeamLeadList(); // uses dynamic indexing already

    return {
      empName: (emp[empCol["name"]] || "").toString().trim(),
      empEmail: (emp[empCol["email"]] || "").toString().trim(),
      colleagues,
      teamLeads,
    };
  } catch (err) {
    Logger.log("Error in getEmployeeData: " + err);
    return null;
  }
};

/**
 * Handles leave request submission, updates sheets, and sends notifications.
 */
const sendAndUpdateALRequest = ({
  empName,
  empEmail,
  startDate,
  daysCount,
  leaveType,
  responsibleColleague,
  teamLead,
  vacationType,
}) => {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Dynamic accessors
    const { sheet: empSheet, col: empCol, data: employees } = getEmployees();
    const hrList = getHRList();
    const teamLeads = getTeamLeadList();
    const templates = getEmailTemplates();

    const hrEmails = hrList.map((h) => h.email);
    const teamLeadMng = teamLeads.find(
      (tl) =>
        tl.name?.toString().trim().toLowerCase() ===
        teamLead?.toString().trim().toLowerCase()
    );
    if (!teamLeadMng) throw new Error(`Team lead not found: ${teamLead}`);

    // Helper for finding template
    const findTemplate = (code) => {
      const tpl = templates.find(
        (t) => t.code?.toString().trim().toUpperCase() === code
      );
      if (!tpl) throw new Error(`Missing email template: ${code}`);
      return tpl;
    };

    // Email templates
    const alRequestHRNotify = findTemplate("AL_REQUEST_HR_NOTIFY");
    const alRequestTeamLeadNotify = findTemplate("AL_REQUEST_TEAMLEAD_NOTIFY");
    const compOffHRNotify = findTemplate("COMP_OFF_HR_NOTIFY");
    const compOffTeamLeadNotify = findTemplate("COMP_OFF_TEAMLEAD_NOTIFY");
    const sickLeaveHRnotify = findTemplate("SICK_LEAVE_HR_NOTIFY");
    const sickLeaveTeamleadNotify = findTemplate("SICK_LEAVE_TEAMLEAD_NOTIFY");

    // Dates
    const start = new Date(startDate);
    const end = new Date(start);
    end.setDate(end.getDate() + parseInt(daysCount) - 1);
    const dateStr = getFormattedDate(start);
    const endDateStr = getFormattedDate(end);

    const resColl =
      responsibleColleague?.toString().trim() !== ""
        ? responsibleColleague
        : "N/A";

    // === Loop through Employees dynamically ===
    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      if (!emp[empCol["email"]]) continue;

      if (
        emp[empCol["email"]].toString().trim().toLowerCase() ===
        empEmail.toString().trim().toLowerCase()
      ) {
        let totalLeaves = parseInt(emp[empCol["total_leaves"]] || 0);
        let leavesUsed = parseInt(emp[empCol["leaves_used"]] || 0);
        let remainingLeaves = parseInt(
          emp[empCol["remaining_leaves"]] || totalLeaves
        );

        // === Annual Leave ===
        if (leaveType === "Annual Leave") {
          empSheet
            .getRange(i + 2, empCol["total_leaves"] + 1, 1, 4)
            .setValues([
              [
                totalLeaves,
                leavesUsed + parseInt(daysCount),
                remainingLeaves - parseInt(daysCount),
                dateStr,
              ],
            ]);

          const hrTemplate = alRequestHRNotify;
          const tlTemplate = alRequestTeamLeadNotify;

          // Replace placeholders dynamically
          const replaceTokens = (str) =>
            str
              .replace(/\[EMP_NAME\]/gi, empName)
              .replace(/\[START_DATE\]/gi, dateStr)
              .replace(/\[END_DATE\]/gi, endDateStr)
              .replace(/\[DAYS_COUNT\]/gi, daysCount)
              .replace(/\[LEAVE_TYPE\]/gi, leaveType)
              .replace(/\[RES_COLL\]/gi, resColl)
              .replace(/\[TEAMLEAD_NAME\]/gi, teamLeadMng.name)
              .replace(/\[VACATION_TYPE\]/gi, vacationType);

          hrEmails.forEach((hr) =>
            sendEmail(
              hr,
              replaceTokens(hrTemplate.subject),
              replaceTokens(hrTemplate.body)
            )
          );
          sendEmail(
            teamLeadMng.email,
            replaceTokens(tlTemplate.subject),
            replaceTokens(tlTemplate.body)
          );

          spreadsheet
            .getSheetByName(AL_STATISTIC_SHEET)
            .appendRow([
              empName,
              empEmail,
              dateStr,
              endDateStr,
              leaveType,
              vacationType,
              daysCount,
            ]);
        }

        // === Sick Leave ===
        else if (leaveType === "Sick Leave") {
          const sickLeaveSheet = spreadsheet.getSheetByName(SICK_LEAVE_SHEET);
          sickLeaveSheet.appendRow([empName, empEmail, dateStr, endDateStr]);

          const hrTemplate = sickLeaveHRnotify;
          const tlTemplate = sickLeaveTeamleadNotify;

          const replaceTokens = (str) =>
            str
              .replace(/\[EMP_NAME\]/gi, empName)
              .replace(/\[START_DATE\]/gi, dateStr)
              .replace(/\[END_DATE\]/gi, endDateStr)
              .replace(/\[DAYS_COUNT\]/gi, daysCount)
              .replace(/\[LEAVE_TYPE\]/gi, leaveType)
              .replace(/\[RES_COLL\]/gi, resColl)
              .replace(/\[TEAMLEAD_NAME\]/gi, teamLeadMng.name)
              .replace(/\[VACATION_TYPE\]/gi, vacationType);

          hrEmails.forEach((hr) =>
            sendEmail(
              hr,
              replaceTokens(hrTemplate.subject),
              replaceTokens(hrTemplate.body)
            )
          );
          sendEmail(
            teamLeadMng.email,
            replaceTokens(tlTemplate.subject),
            replaceTokens(tlTemplate.body)
          );
        }

        // === Comp Off ===
        else {
          empSheet
            .getRange(i + 2, empCol["total_leaves"] + 1, 1, 3)
            .setValues([
              [
                totalLeaves + parseInt(daysCount),
                leavesUsed,
                remainingLeaves + parseInt(daysCount),
              ],
            ]);

          const hrTemplate = compOffHRNotify;
          const tlTemplate = compOffTeamLeadNotify;

          const replaceTokens = (str) =>
            str
              .replace(/\[EMP_NAME\]/gi, empName)
              .replace(/\[START_DATE\]/gi, dateStr)
              .replace(/\[DAYS_COUNT\]/gi, daysCount)
              .replace(/\[RES_COLL\]/gi, resColl)
              .replace(/\[TEAMLEAD_NAME\]/gi, teamLeadMng.name)
              .replace(/\[VACATION_TYPE\]/gi, vacationType);

          hrEmails.forEach((hr) =>
            sendEmail(
              hr,
              replaceTokens(hrTemplate.subject),
              replaceTokens(hrTemplate.body)
            )
          );
          sendEmail(
            teamLeadMng.email,
            replaceTokens(tlTemplate.subject),
            replaceTokens(tlTemplate.body)
          );

          spreadsheet
            .getSheetByName(AL_STATISTIC_SHEET)
            .appendRow([
              empName,
              empEmail,
              dateStr,
              endDateStr,
              leaveType,
              vacationType,
              daysCount,
            ]);
        }

        return {
          message: `${leaveType} request submitted successfully. You can now close this window.`,
        };
      }
    }

    // Employee not found
    return {
      error:
        "Something went wrong. Could not submit the request. Please contact HR Department.",
    };
  } catch (err) {
    Logger.log(err);
    return {
      error:
        "Something went wrong. Could not submit the request. Please contact HR Department.",
    };
  }
};
