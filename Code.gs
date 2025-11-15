/**
 * sendBirthdayNotifications()
 * Sends birthday notifications and reminders to colleagues and HR.
 * 
 * Workflow:
 * 1. On employee's birthday: sends greeting email to all colleagues
 * 2. Day before birthday (previous working day): notifies HR for preparation
 * 
 * Triggered daily by the autoTriggerMainFunction
 * 
 * @returns {void}
 * @throws {Error} If required sheets or email templates are missing
 */

const sendBirthdayNotifications = () => {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic indexing ===
    // Build column index map from headers for flexible column references
    const empCol = getColumnIndexes(empSheet);
    // Fetch all employee records (excluding header row)
    const employees = empSheet.getDataRange().getValues().slice(1);

    // === 🔹 Load HR email addresses ===
    // Extract email addresses from HR list for notification distribution
    const hrEmails = getHRList().map((h) => h.email);

    // === 🔹 Load and cache email templates ===
    // Build a lookup map for quick template access by code
    const templates = getEmailTemplates().reduce((acc, t) => {
      acc[t.code] = t;
      return acc;
    }, {});

    // === 🔹 Retrieve specific templates ===
    const birthdayEmail = templates.BIRTHDAY_ALERT;
    const birthdayHREmail = templates.BIRTHDAY_ALERT_HR;

    // === 🔹 Validate required templates exist ===
    if (!birthdayEmail || !birthdayHREmail)
      throw new Error("Missing BIRTHDAY_ALERT or BIRTHDAY_ALERT_HR templates.");

    // === 🔹 Get today's date for comparison ===
    const today = new Date();

    // === 🔹 Process each employee ===
    employees.forEach((emp) => {
      // === Extract employee data ===
      const name = (emp[empCol["name"]] || "").toString().trim();
      const email = (emp[empCol["email"]] || "").toString().trim();
      const birthdayStr = emp[empCol["birthday"]];

      // === Skip if missing required data ===
      if (!name || !email || !birthdayStr) return;

      // === Parse and format birthday ===
      const birthday = parseDate(birthdayStr);
      const thisYearBirthday = new Date(
        today.getFullYear(),
        birthday.getMonth(),
        birthday.getDate()
      );
      const birthdayFormatted = getFormattedDate(thisYearBirthday);

      // === 🎉 Notify colleagues on birthday ===
      // Send birthday greeting to all colleagues (excluding the birthday person)
      if (isSameDayMonthYear(today, thisYearBirthday)) {
        const colleagues = employees
          .filter((e) => (e[empCol["email"]] || "").toString().trim() !== email)
          .map((e) => e[empCol["email"]].toString().trim())
          .filter(Boolean);

        const subject = fillTemplate(birthdayEmail.subject, { EMP_NAME: name });
        const body = fillTemplate(birthdayEmail.body, {
          EMP_NAME: name,
          BIRTHDAY: birthdayFormatted,
        });

        colleagues.forEach((coll) => sendEmail(coll, subject, body));
        Logger.log(`Birthday email sent to colleagues of ${name}`);
      }

      // === 📅 Notify HR day before birthday (previous working day) ===
      // Gives HR time to prepare birthday arrangements (cake, gift, etc.)
      const notifyDate = getPreviousWorkingDay(thisYearBirthday);
      if (isSameDayMonthYear(today, notifyDate)) {
        const subject = fillTemplate(birthdayHREmail.subject, {
          EMP_NAME: name,
        });
        const body = fillTemplate(birthdayHREmail.body, {
          EMP_NAME: name,
          BIRTHDAY: birthdayFormatted,
        });

        hrEmails.forEach((hr) => sendEmail(hr, subject, body));
        Logger.log(`Birthday HR notification sent for ${name}`);
      }
    });
  } catch (err) {
    Logger.log("Error in sendBirthdayNotifications: " + err);
  }
};

/**
 * sendQuarterlyLeaveReminders()
 * Sends quarterly leave reminders to employees and HR.
 * 
 * Triggered on: 1st of March, June, September, and December
 * 
 * Employee notification: Reminds employees of available leaves and encourages planning
 * HR alert: Notifies HR if employee hasn't used at least 14 days (policy enforcement)
 * 
 * @returns {void}
 * @throws {Error} If required sheets or email templates are missing
 */
const sendQuarterlyLeaveReminders = () => {
  try {
    // === 🔹 Check if today is quarterly reminder date ===
    // Early exit if not the correct date (Mar 1, Jun 1, Sep 1, or Dec 1)
    if (!isTodayQuarterlyReminderDate()) return;

    // === 🔹 Access required sheets ===
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    // === 🔹 Validate sheets exist ===
    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic column indexing ===
    // Build column index map from headers for flexible column references
    const empCol = getColumnIndexes(empSheet);
    // Fetch all employee records (excluding header row)
    const employees = empSheet.getDataRange().getValues().slice(1);

    // === 🔹 Load HR email addresses ===
    // Extract email addresses from HR list for notification distribution
    const hrEmails = getHRList().map((h) => h.email);

    // === 🔹 Load and cache email templates ===
    // Build a lookup map for quick template access by code
    const templates = getEmailTemplates().reduce((acc, t) => {
      acc[t.code] = t;
      return acc;
    }, {});

    // === 🔹 Retrieve specific templates ===
    const quarterlyALTemplate = templates.QUARTERLY_AL_REMINDER;
    const obligatoryALTemplate = templates.HR_OBLIGATORY_AL_REMINDER;

    // === 🔹 Validate required templates exist ===
    if (!quarterlyALTemplate || !obligatoryALTemplate)
      throw new Error(
        "Missing QUARTERLY_AL_REMINDER or HR_OBLIGATORY_AL_REMINDER templates."
      );

    // === 🔹 Load web form link from script properties ===
    // Used in email templates to direct employees to leave request form
    const alRequestForm =
      PropertiesService.getScriptProperties().getProperty("WEB_FORM_LINK") ||
      "";

    // === 🔹 Validate required employee columns ===
    // Ensure all necessary columns exist in the Employees sheet
    const requiredEmpCols = [
      "name",
      "email",
      "total_leaves",
      "leaves_used",
      "remaining_leaves",
    ];
    requiredEmpCols.forEach((c) => {
      if (empCol[c] == null)
        throw new Error(`Missing '${c}' column in Employees sheet.`);
    });

    // === 🔹 Process each employee ===
    employees.forEach((emp) => {
      // === Extract employee leave data ===
      const name = (emp[empCol["name"]] || "").toString().trim();
      const email = (emp[empCol["email"]] || "").toString().trim();
      const totalLeaves = emp[empCol["total_leaves"]] || 0;
      const leavesUsed = emp[empCol["leaves_used"]] || 0;
      const remainingLeaves = emp[empCol["remaining_leaves"]] || 0;

      // === Skip if missing required data ===
      if (!name || !email) return;

      // === 📧 Send quarterly AL reminder to employee ===
      // Encourages employees to plan and use their leave
      const subject = fillTemplate(quarterlyALTemplate.subject, {
        EMP_NAME: name,
      });
      const body = fillTemplate(quarterlyALTemplate.body, {
        EMP_NAME: name,
        WEB_FORM_LINK: alRequestForm,
        TOTAL_AL: totalLeaves,
        AL_USED: leavesUsed,
        AL_REMAINING: remainingLeaves,
      });

      sendEmail(email, subject, body);
      Logger.log(`Quarterly AL reminder sent to ${name}`);

      // === 📧 Notify HR if leaves used < 14 ===
      // Company policy: employees should use at least 14 days per quarter
      if (parseInt(leavesUsed) < 14) {
        const hrSubject = fillTemplate(obligatoryALTemplate.subject, {
          EMP_NAME: name,
        });
        const hrBody = fillTemplate(obligatoryALTemplate.body, {
          EMP_NAME: name,
          TOTAL_AL: totalLeaves,
          AL_USED: leavesUsed,
          AL_REMAINING: remainingLeaves,
        });

        hrEmails.forEach((hr) => sendEmail(hr, hrSubject, hrBody));
        Logger.log(`HR notified for ${name} (leaves used < 14)`);
      }
    });
  } catch (err) {
    Logger.log("Error in sendQuarterlyLeaveReminders: " + err);
  }
};

/**
 * manageAnnualLeaves()
 * Handles annual leave milestones for employees.
 * Sends 6-month anniversary notification with leave request form link.
 * Resets annual leave balance on work anniversary (yearly).
 * 
 * @returns {void}
 * @throws {Error} If required sheets or email templates are missing
 */

const manageAnnualLeaves = () => {
  try {
    // === 🔹 Access required sheets ===
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    // === 🔹 Validate sheets exist ===
    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic column indexing ===
    // Build column index map from headers for flexible column references
    const empCol = getColumnIndexes(empSheet);
    // Fetch all employee records (excluding header row)
    const employees = empSheet.getDataRange().getValues().slice(1);

    // === 🔹 Load and cache email templates ===
    // Build a lookup map for quick template access by code
    const templates = getEmailTemplates().reduce((acc, t) => {
      acc[t.code] = t;
      return acc;
    }, {});

    // === 🔹 Retrieve specific templates ===
    const sixMonthsEmail = templates.SIX_MONTH_AL_REMINDER;
    const alResetAnniversaryEmail = templates.AL_RESET_ANNIVERSARY;

    // === 🔹 Validate required templates exist ===
    if (!sixMonthsEmail || !alResetAnniversaryEmail)
      throw new Error(
        "Missing SIX_MONTH_AL_REMINDER or AL_RESET_ANNIVERSARY templates."
      );

    // === 🔹 Load web form link from script properties ===
    // Used in email templates to direct employees to leave request form
    const alRequestForm =
      PropertiesService.getScriptProperties().getProperty("WEB_FORM_LINK") ||
      "";

    // === 🔹 Validate required employee columns ===
    // Ensure all necessary columns exist in the Employees sheet
    const requiredEmpCols = [
      "name",
      "email",
      "join_date",
      "total_leaves",
      "leaves_used",
      "remaining_leaves",
    ];
    requiredEmpCols.forEach((c) => {
      if (empCol[c] == null)
        throw new Error(`Missing '${c}' column in Employees sheet.`);
    });

    // === 🔹 Process each employee ===
    employees.forEach((row, idx) => {
      // === Extract employee data ===
      const name = (row[empCol["name"]] || "").toString().trim();
      const email = (row[empCol["email"]] || "").toString().trim();
      const joinDateStr = row[empCol["join_date"]];

      // === Skip if missing required data ===
      if (!name || !email || !joinDateStr) return;

      // === Parse join date ===
      const joinDate = parseDate(joinDateStr);

      // === 📅 6-month AL reminder ===
      // Sent on the exact 6-month anniversary of joining to inform about leave balance
      if (isSixMonthComplete(joinDate)) {
        // === Update leaves: Total 21, Used 0, Remaining 21 ===
        // Reset leave balance for 6-month milestone
        empSheet
          .getRange(idx + 2, empCol["total_leaves"] + 1, 1, 3)
          .setValues([[21, 0, 21]]);

        const subject = fillTemplate(sixMonthsEmail.subject, {
          EMP_NAME: name,
        });
        const body = fillTemplate(sixMonthsEmail.body, {
          EMP_NAME: name,
          WEB_FORM_LINK: alRequestForm,
        });

        sendEmail(email, subject, body);
        Logger.log(`Six-month AL reminder sent to ${name}`);
      }

      // === 📅 Anniversary AL reset ===
      // On work anniversary, reset the annual leave balance to starting amount
      if (isAnniversary(joinDate)) {
        // === Update leaves: Total 21, Used 0, Remaining 21 ===
        // Reset leave balance on anniversary (employee completes 1 year or more)
        empSheet
          .getRange(idx + 2, empCol["total_leaves"] + 1, 1, 3)
          .setValues([[21, 0, 21]]);

        // === Calculate years of service ===
        const years = anniversaryYears(joinDate);

        const subject = fillTemplate(alResetAnniversaryEmail.subject, {
          EMP_NAME: name,
        });
        const body = fillTemplate(alResetAnniversaryEmail.body, {
          EMP_NAME: name,
          ANNIVERSARY_YEARS: years,
          WEB_FORM_LINK: alRequestForm,
        });

        sendEmail(email, subject, body);
        Logger.log(
          `Anniversary AL reset email sent to ${name} (${years} years)`
        );
      }
    });
  } catch (err) {
    Logger.log("Error in manageAnnualLeaves: " + err);
  }
};

/**
 * manageProbationPeriod()
 * Manages the probation period workflow for new employees.
 * 
 * Workflow:
 * 1. 7 days before probation end: notifies HR and Team Lead for evaluation
 * 2. Creates probation record in Probation sheet with editable Result dropdown
 * 3. On probation end date: checks result and sends "Probation Passed" email
 * 
 * @returns {void}
 * @throws {Error} If required sheets or email templates are missing
 */

const manageProbationPeriod = () => {
  // === 🔹 Access required sheets ===
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
  const probationSheet = ss.getSheetByName(PROBATION_SHEET);

  // === 🔹 Validate sheets exist ===
  if (!empSheet || !probationSheet)
    throw new Error("Missing required sheet(s): Employees or Probation.");

  // === 🔹 Get data and dynamic column indexes ===
  // Build column index maps for flexible column references
  const empIndex = getColumnIndexes(empSheet);
  const probIndex = getColumnIndexes(probationSheet);
  // Fetch all employee and probation records (excluding headers)
  const employees = empSheet.getDataRange().getValues().slice(1);
  const probationData = probationSheet.getDataRange().getValues().slice(1);

  // === 🔹 Load HR and email templates ===
  // Get HR email addresses for notifications
  const hrEmails = getHRList();
  const emailTemplates = getEmailTemplates();

  // === 🔹 Template mapping for fast lookup ===
  // Build a lookup map for quick template access by code
  const templates = emailTemplates.reduce((acc, t) => {
    acc[t.code] = t;
    return acc;
  }, {});

  // === 🔹 Retrieve specific templates ===
  const probationHRnotify = templates.PROBATION_HR_NOTIFY;
  const probationTeamleadNotify = templates.PROBATION_TEAMLEAD_NOTIFY;
  const probationPassEmail = templates.PROBATION_PASSED;

  // === 🔹 Validate required templates exist ===
  if (!probationHRnotify || !probationTeamleadNotify || !probationPassEmail)
    throw new Error(
      "Missing one or more probation email templates in Settings sheet."
    );

  // === 🔹 Validate required columns ===
  // Ensure all necessary columns exist in the Employees sheet
  const requiredEmpCols = ["name", "email", "join_date", "department"];
  requiredEmpCols.forEach((c) => {
    if (empIndex[c] == null)
      throw new Error(`Missing '${c}' column in Employees sheet.`);
  });

  // === 🔹 Validate probation sheet columns ===
  // Ensure all necessary columns exist in the Probation sheet
  const requiredProbCols = ["employee_name", "result"];
  requiredProbCols.forEach((c) => {
    if (probIndex[c] == null)
      throw new Error(`Missing '${c}' column in Probation sheet.`);
  });

  // === 🔹 Get today's date for comparison ===
  const today = new Date();

  // === 🔹 Process each employee ===
  employees.forEach((row) => {
    // === Extract employee data ===
    const name = (row[empIndex["name"]] || "").toString().trim();
    const email = (row[empIndex["email"]] || "").toString().trim();
    const joinDateStr = row[empIndex["join_date"]];
    const department = (row[empIndex["department"]] || "")
      .toString()
      .trim()
      .toLowerCase();

    // === Skip if missing required data ===
    if (!name || !email || !joinDateStr) return;

    // === 🔹 Calculate probation dates ===
    // Standard probation period is 3 months (90 days)
    const joinDate = parseDate(joinDateStr);
    const probationEndDate = new Date(joinDate);
    probationEndDate.setMonth(probationEndDate.getMonth() + 3);

    // === 🔹 Calculate notification date (7 days before probation end) ===
    // Gives HR and Team Lead time to evaluate employee
    const probationNotifyDate = new Date(probationEndDate);
    probationNotifyDate.setDate(probationNotifyDate.getDate() - 7);

    // === 🔹 Build common email template replacements ===
    // These values are used in multiple email templates
    const commonMap = {
      EMP_NAME: name,
      JOIN_DATE: getFormattedDate(joinDate),
      PROBATION_END_DATE: getFormattedDate(probationEndDate),
    };

    // === 📅 Notify HR & TL before probation end ===
    // Triggered 7 days before probation end date
    if (isSameDayMonthYear(today, probationNotifyDate)) {
      // === HR Notification ===
      // Send email to all HR contacts with probation details
      const hrSubject = fillTemplate(probationHRnotify.subject, commonMap);
      const hrBody = fillTemplate(probationHRnotify.body, commonMap);
      hrEmails.forEach((hr) => sendEmail(hr.email, hrSubject, hrBody));

      // === TL Notification (by department) ===
      // Match team lead to employee's department for targeted notification
      const depLead =
        department === "service"
          ? SERVICE_DEP_LEAD
          : department === "client"
          ? CLIENT_DEP_LEAD
          : null;

      if (depLead) {
        const tlMap = { ...commonMap, TEAMLEAD_NAME: depLead.name };
        const tlSubject = fillTemplate(probationTeamleadNotify.subject, tlMap);
        const tlBody = fillTemplate(probationTeamleadNotify.body, tlMap);
        sendEmail(depLead.email, tlSubject, tlBody);
      }

      // === 🔹 Append row with dropdown in "Result" column ===
      // Creates a new entry in probation sheet for HR to fill in later
      const newRow = [
        name,
        email,
        getFormattedDate(joinDate),
        getFormattedDate(probationEndDate),
        getFormattedDate(probationNotifyDate),
      ];

      // === Append the row first ===
      const appendRowIndex = probationSheet.getLastRow() + 1;
      probationSheet.appendRow(newRow);

      // === Set dropdown for "Result" column ===
      // Allows HR to select from predefined options (Probation Passed/Not Passed)
      const resultColIndex = probIndex["result"] + 1; // +1 because sheet ranges are 1-based
      const resultCell = probationSheet.getRange(
        appendRowIndex,
        resultColIndex
      );

      // === Create data validation dropdown ===
      // Restricts Result column to specific values for data consistency
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Probation Passed", "Probation Not Passed"], true)
        .setAllowInvalid(false)
        .build();

      resultCell.setDataValidation(rule);
    }

    // === ✅ On actual end date → Send "Probation Passed" email ===
    // Triggered on the exact probation end date
    if (isSameDayMonthYear(today, probationEndDate)) {
      // === Find probation record for this employee ===
      const empProbRow = probationData.find(
        (r) => (r[probIndex["employee_name"]] || "").toString().trim() === name
      );

      // === Check if probation result is "Passed" ===
      if (empProbRow) {
        const result = (empProbRow[probIndex["result"]] || "")
          .toString()
          .toLowerCase();

        // === Send confirmation email if probation passed ===
        if (result === "probation passed") {
          const alRequestForm =
            PropertiesService.getScriptProperties().getProperty(
              "WEB_FORM_LINK"
            );

          const passMap = {
            EMP_NAME: name,
            WEB_FORM_LINK: alRequestForm || "",
          };

          const subject = fillTemplate(probationPassEmail.subject, passMap);
          const body = fillTemplate(probationPassEmail.body, passMap);

          sendEmail(email, subject, body);
        }
      }
    }
  });
};

/**
 * Main wrapper function for automatic trigger
 */

/**
 * autoTriggerMainFunction()
 * Main orchestrator function that triggers all HR automation workflows.
 * 
 * This is the entry point for time-based triggers. It executes sequentially:
 * 1. sendBirthdayNotifications() - Daily birthday greetings and HR prep
 * 2. sendQuarterlyLeaveReminders() - Quarterly leave balance reminders
 * 3. manageAnnualLeaves() - 6-month and anniversary milestones
 * 4. manageProbationPeriod() - Probation notifications and tracking
 * 
 * Error Handling:
 * - Each function is wrapped in try-catch to isolate failures
 * - Errors are logged to Apps Script Logger for audit trail
 * - Single function failure does not prevent other functions from executing
 * 
 * Trigger Setup:
 * - Configure in Apps Script Project Settings
 * - Time-based trigger: Daily between 12 AM - 1 AM GMT+8 recommended
 * - Executes daily at consistent time for reliable notifications
 * 
 * @returns {void}
 * @throws {Error} Logged for each failed sub-function but does not block execution
 */

const autoTriggerMainFunction = () => {
  // === 🎉 Execute Birthday Notifications ===
  try {
    sendBirthdayNotifications();
    Logger.log("✅ Birthday notifications executed successfully.");
  } catch (e) {
    Logger.log("❌ Birthday notifications error: " + e.message);
  }

  // === 📅 Execute Quarterly Leave Reminders ===
  try {
    sendQuarterlyLeaveReminders();
    Logger.log("✅ Quarterly leave reminders executed successfully.");
  } catch (e) {
    Logger.log("❌ Quarterly leave reminders error: " + e.message);
  }

  // === 📊 Execute Annual Leave Management ===
  try {
    manageAnnualLeaves();
    Logger.log("✅ Annual leave management executed successfully.");
  } catch (e) {
    Logger.log("❌ Annual leave management error: " + e.message);
  }

  // === 👤 Execute Probation Period Management ===
  try {
    manageProbationPeriod();
    Logger.log("✅ Probation period management executed successfully.");
  } catch (e) {
    Logger.log("❌ Probation period management error: " + e.message);
  }

  // === ✅ All workflows completed ===
  Logger.log("📌 Automation cycle complete.");
};
