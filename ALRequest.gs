/**
 * doGet()
 * Entry point for the web app. Serves the HTML form for annual leave requests.
 * This function is automatically triggered when the web app URL is accessed.
 * 
 * @returns {HtmlOutput} The rendered HTML form with title "Annual Leave Request Form"
 */
const doGet = () => {
  return HtmlService.createHtmlOutputFromFile("ALRequestForm").setTitle(
    "Annual Leave Request Form"
  );
};

/**
 * getEmployeeData(email)
 * Fetches employee data, colleagues, and team leads for the leave request form.
 * Validates that the employee exists and retrieves their information along with
 * a list of colleagues (for responsible colleague selection) and team leads (for approval).
 * 
 * @param {string} email - The employee's email address to look up
 * @returns {Object|null} Object containing empName, empEmail, colleagues array, and teamLeads array
 *                        Returns null if an error occurs
 */
const getEmployeeData = (email) => {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic indexing ===
    // Get column index mapping from sheet headers for flexible column references
    const empCol = getColumnIndexes(empSheet);
    // Fetch all employee data, skipping header row and filtering out empty names
    const employees = empSheet
      .getDataRange()
      .getValues()
      .slice(1)
      .filter((row) => (row[empCol["name"]] || "").toString().trim() !== "");

    // === 🔹 Find employee ===
    // Search for employee matching the provided email (case-insensitive)
    const emp = employees.find(
      (row) =>
        (row[empCol["email"]] || "").toString().trim().toLowerCase() ===
        email.toLowerCase()
    );

    if (!emp) throw new Error(`Employee with email "${email}" not found.`);

    // === 🔹 Colleagues (exclude self) ===
    // Get list of all colleagues except the current employee for responsible colleague selection
    const colleagues = employees
      .filter(
        (row) =>
          (row[empCol["email"]] || "").toString().trim().toLowerCase() !==
          email.toLowerCase()
      )
      .map((row) => (row[empCol["name"]] || "").toString().trim())
      .filter(Boolean);

    // === 🔹 Team Leads ===
    // Fetch all team leads from Settings sheet for approval selection
    const teamLeads = getTeamLeadList();

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
 * sendAndUpdateALRequest(data)
 * Handles leave request submission, updates employee leave records, and sends notifications.
 * Processes three types of leave: Annual Leave, Sick Leave, and Compensatory Days Off.
 * Updates employee data, appends records to tracking sheets, and sends templated emails
 * to HR and Team Leads for approval/notification.
 * 
 * @param {Object} data - Leave request details
 * @param {string} data.empName - Employee's full name
 * @param {string} data.empEmail - Employee's email address
 * @param {string} data.startDate - Leave start date
 * @param {number} data.daysCount - Total number of leave days
 * @param {string} data.leaveType - Type: "Annual Leave", "Sick Leave", or "Compensatory Days Off"
 * @param {string} data.responsibleColleague - Colleague(s) covering during leave
 * @param {string} data.teamLead - Team Lead email for approval
 * @param {string} data.vacationType - "Paid" or "Unpaid"
 * @returns {Object} Success/error message object
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

    // === 🔹 Retrieve data accessors ===
    // Get employee data and configuration lists from sheets
    const { sheet: empSheet, col: empCol, data: employees } = getEmployees();
    const hrList = getHRList();
    const teamLeads = getTeamLeadList();
    const templates = getEmailTemplates();

    // === 🔹 Prepare email recipients and validators ===
    // Extract HR email addresses for notification distribution
    const hrEmails = hrList.map((h) => h.email);
    // Find the specific team lead object for this leave request
    const teamLeadMng = teamLeads.find(
      (tl) =>
        tl.email?.toString().trim().toLowerCase() ===
        teamLead?.toString().trim().toLowerCase()
    );
    if (!teamLeadMng) throw new Error(`Team lead not found: ${teamLead}`);

    // === 🔹 Email template loader ===
    // Helper function to retrieve and validate email template by code
    const findTemplate = (code) => {
      const tpl = templates.find(
        (t) => t.code?.toString().trim().toUpperCase() === code
      );
      if (!tpl) throw new Error(`Missing email template: ${code}`);
      return tpl;
    };

    // === 🔹 Load email templates ===
    // Retrieve all required email templates for different leave types
    const alRequestHRNotify = findTemplate("AL_REQUEST_HR_NOTIFY");
    const alRequestTeamLeadNotify = findTemplate("AL_REQUEST_TEAMLEAD_NOTIFY");
    const compOffHRNotify = findTemplate("COMP_OFF_HR_NOTIFY");
    const compOffTeamLeadNotify = findTemplate("COMP_OFF_TEAMLEAD_NOTIFY");
    const sickLeaveHRnotify = findTemplate("SICK_LEAVE_HR_NOTIFY");
    const sickLeaveTeamleadNotify = findTemplate("SICK_LEAVE_TEAMLEAD_NOTIFY");

    // === 🔹 Calculate leave dates ===
    // Parse start date and calculate end date based on number of days
    const start = new Date(startDate);
    const end = new Date(start);
    end.setDate(end.getDate() + parseInt(daysCount) - 1);
    const dateStr = getFormattedDate(start);
    const endDateStr = getFormattedDate(end);

    // === 🔹 Prepare responsible colleague info ===
    // Set colleague name or N/A if not provided
    const resColl =
      responsibleColleague?.toString().trim() !== ""
        ? responsibleColleague
        : "N/A";

    // === 🔹 Loop through Employees dynamically ===
    // Find matching employee and process leave request based on type
    for (let i = 0; i < employees.length; i++) {
      const emp = employees[i];
      if (!emp[empCol["email"]]) continue;

      // === Check if this is the requesting employee ===
      if (
        emp[empCol["email"]].toString().trim().toLowerCase() ===
        empEmail.toString().trim().toLowerCase()
      ) {
        // === 🔹 Extract current leave balances ===
        let totalLeaves = parseInt(emp[empCol["total_leaves"]] || 0);
        let leavesUsed = parseInt(emp[empCol["leaves_used"]] || 0);
        let remainingLeaves = parseInt(
          emp[empCol["remaining_leaves"]] || totalLeaves
        );

        // === 🔹 Annual Leave Processing ===
        if (leaveType === "Annual Leave") {
          // Update employee sheet with new leave balances
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

          // === 🔹 Email template replacement ===
          // Replace all placeholder tokens with actual leave request data
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

          // === 🔹 Send HR notifications ===
          hrEmails.forEach((hr) =>
            sendEmail(
              hr,
              replaceTokens(hrTemplate.subject),
              replaceTokens(hrTemplate.body)
            )
          );
          // === 🔹 Send Team Lead notifications ===
          sendEmail(
            teamLeadMng.email,
            replaceTokens(tlTemplate.subject),
            replaceTokens(tlTemplate.body)
          );

          // === 🔹 Append to AL Statistic sheet ===
          // Log the leave request in the statistics tracking sheet
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

        // === 🔹 Sick Leave Processing ===
        else if (leaveType === "Sick Leave") {
          // Append to Sick Leaves sheet (no leave balance updates for sick leaves)
          const sickLeaveSheet = spreadsheet.getSheetByName(SICK_LEAVE_SHEET);
          sickLeaveSheet.appendRow([empName, empEmail, dateStr, endDateStr]);

          const hrTemplate = sickLeaveHRnotify;
          const tlTemplate = sickLeaveTeamleadNotify;

          // === 🔹 Email template replacement for sick leave ===
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

          // === 🔹 Send notifications ===
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

        // === 🔹 Compensatory Days Off (Comp Off) Processing ===
        else {
          // Update employee sheet: ADD days to total leaves (bonus days)
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

          // === 🔹 Email template replacement for comp off ===
          const replaceTokens = (str) =>
            str
              .replace(/\[EMP_NAME\]/gi, empName)
              .replace(/\[START_DATE\]/gi, dateStr)
              .replace(/\[DAYS_COUNT\]/gi, daysCount)
              .replace(/\[RES_COLL\]/gi, resColl)
              .replace(/\[TEAMLEAD_NAME\]/gi, teamLeadMng.name)
              .replace(/\[VACATION_TYPE\]/gi, vacationType);

          // === 🔹 Send notifications ===
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

          // === 🔹 Append to AL Statistic sheet ===
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

        // === 🔹 Return success message ===
        return {
          message: `${leaveType} request submitted successfully. You can now close this window.`,
        };
      }
    }

    // === 🔹 Employee not found handling ===
    // This should not occur if getEmployeeData validation is working properly
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

/**
 * submitFeedback(data)
 * Processes feedback submission from employees during probation period.
 * Stores feedback in the Feedback sheet with metadata including employee info,
 * feedback category, colleague name (if applicable), and submission timestamp.
 * 
 * @param {Object} data - Feedback submission details
 * @param {string} data.empName - Employee's full name
 * @param {string} data.empEmail - Employee's email address
 * @param {string} data.feedbackFor - Feedback category ("Self", "Colleague", "Company", "Process")
 * @param {string} data.colleagueName - Name of colleague if feedback is about a colleague (optional)
 * @param {string} data.feedbackText - Detailed feedback content
 * @returns {Object} Success or error message object
 */
const submitFeedback = (data) => {
  try {
    // === 🔹 Extract feedback data ===
    const { empName, empEmail, feedbackFor, colleagueName, feedbackText } =
      data;

    // === 🔹 Validate required fields ===
    if (!empEmail || !feedbackText || !feedbackFor) {
      throw new Error("Missing required feedback fields.");
    }

    // === 🔹 Access Feedback sheet ===
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = spreadsheet.getSheetByName(FEEDBACK_SHEET);

    if (!feedbackSheet) {
      throw new Error("Feedback sheet not found.");
    }

    // === 🔹 Get dynamic column indexes ===
    const colIndex = getColumnIndexes(feedbackSheet);

    // === 🔹 Prepare feedback data array ===
    // Initialize array with correct size based on number of columns
    const feedbackData = new Array(6);

    // === 🔹 Map feedback data to correct columns ===
    feedbackData[colIndex["employee_name"]] = empName;
    feedbackData[colIndex["employee_email"]] = empEmail;
    feedbackData[colIndex["feedback_about"]] = feedbackFor;
    feedbackData[colIndex["feedback_date"]] = getFormattedDate(new Date());
    feedbackData[colIndex["feedback"]] = feedbackText;
    feedbackData[colIndex["colleague_name"]] = colleagueName;

    // === 🔹 Append feedback to sheet ===
    feedbackSheet.appendRow(feedbackData);

    return { message: "Feedback submitted successfully." };
  } catch (err) {
    Logger.log(err);
    return {
      error: err.message || "Failed to submit feedback. Please contact HR.",
    };
  }
};
