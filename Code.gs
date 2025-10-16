/**
 * Send birthday notifications / reminders to colleagues and HR
 */

const sendBirthdayNotifications = () => {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic indexing ===
    const empCol = getColumnIndexes(empSheet);
    const employees = empSheet.getDataRange().getValues().slice(1);

    const hrEmails = getHRList().map((h) => h.email);

    const templates = getEmailTemplates().reduce((acc, t) => {
      acc[t.code] = t;
      return acc;
    }, {});

    const birthdayEmail = templates.BIRTHDAY_ALERT;
    const birthdayHREmail = templates.BIRTHDAY_ALERT_HR;

    if (!birthdayEmail || !birthdayHREmail)
      throw new Error("Missing BIRTHDAY_ALERT or BIRTHDAY_ALERT_HR templates.");

    const today = new Date();

    employees.forEach((emp) => {
      const name = (emp[empCol["name"]] || "").toString().trim();
      const email = (emp[empCol["email"]] || "").toString().trim();
      const birthdayStr = emp[empCol["birthday"]];

      if (!name || !email || !birthdayStr) return;

      const birthday = parseDate(birthdayStr);
      const thisYearBirthday = new Date(
        today.getFullYear(),
        birthday.getMonth(),
        birthday.getDate()
      );
      const birthdayFormatted = getFormattedDate(thisYearBirthday);

      // === 🎉 Notify colleagues on birthday ===
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
 * Sends quarterly leave reminders to employees and HR.
 */
const sendQuarterlyLeaveReminders = () => {
  try {
    if (!isTodayQuarterlyReminderDate()) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic column indexing ===
    const empCol = getColumnIndexes(empSheet);
    const employees = empSheet.getDataRange().getValues().slice(1);

    const hrEmails = getHRList().map((h) => h.email);

    const templates = getEmailTemplates().reduce((acc, t) => {
      acc[t.code] = t;
      return acc;
    }, {});

    const quarterlyALTemplate = templates.QUARTERLY_AL_REMINDER;
    const obligatoryALTemplate = templates.HR_OBLIGATORY_AL_REMINDER;

    if (!quarterlyALTemplate || !obligatoryALTemplate)
      throw new Error(
        "Missing QUARTERLY_AL_REMINDER or HR_OBLIGATORY_AL_REMINDER templates."
      );

    const alRequestForm =
      PropertiesService.getScriptProperties().getProperty("AL_REQUEST_FORM") ||
      "";

    // === 🔹 Validate required employee columns ===
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
      const name = (emp[empCol["name"]] || "").toString().trim();
      const email = (emp[empCol["email"]] || "").toString().trim();
      const totalLeaves = emp[empCol["total_leaves"]] || 0;
      const leavesUsed = emp[empCol["leaves_used"]] || 0;
      const remainingLeaves = emp[empCol["remaining_leaves"]] || 0;

      if (!name || !email) return;

      // === 📧 Send quarterly AL reminder to employee ===
      const subject = fillTemplate(quarterlyALTemplate.subject, {
        EMP_NAME: name,
      });
      const body = fillTemplate(quarterlyALTemplate.body, {
        EMP_NAME: name,
        AL_REQUEST_FORM: alRequestForm,
        TOTAL_AL: totalLeaves,
        AL_USED: leavesUsed,
        AL_REMAINING: remainingLeaves,
      });

      sendEmail(email, subject, body);
      Logger.log(`Quarterly AL reminder sent to ${name}`);

      // === 📧 Notify HR if leaves used < 14 ===
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
 * Send notification when employee completes 6 months
 * When employee has an anniversary, resert the annual leaves
 */

const manageAnnualLeaves = () => {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
    const settingSheet = ss.getSheetByName(SETTINGS_SHEET);

    if (!empSheet || !settingSheet)
      throw new Error("Missing Employees or Settings sheet.");

    // === 🔹 Dynamic column indexing ===
    const empCol = getColumnIndexes(empSheet);
    const employees = empSheet.getDataRange().getValues().slice(1);

    const templates = getEmailTemplates().reduce((acc, t) => {
      acc[t.code] = t;
      return acc;
    }, {});

    const sixMonthsEmail = templates.SIX_MONTH_AL_REMINDER;
    const alResetAnniversaryEmail = templates.AL_RESET_ANNIVERSARY;

    if (!sixMonthsEmail || !alResetAnniversaryEmail)
      throw new Error(
        "Missing SIX_MONTH_AL_REMINDER or AL_RESET_ANNIVERSARY templates."
      );

    const alRequestForm =
      PropertiesService.getScriptProperties().getProperty("AL_REQUEST_FORM") ||
      "";

    // === 🔹 Validate required employee columns ===
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
      const name = (row[empCol["name"]] || "").toString().trim();
      const email = (row[empCol["email"]] || "").toString().trim();
      const joinDateStr = row[empCol["join_date"]];

      if (!name || !email || !joinDateStr) return;

      const joinDate = parseDate(joinDateStr);

      // --- 6-month AL reminder ---
      if (isSixMonthComplete(joinDate)) {
        // Update leaves: Total 21, Used 0, Remaining 21
        empSheet
          .getRange(idx + 2, empCol["total_leaves"] + 1, 1, 3)
          .setValues([[21, 0, 21]]);

        const subject = fillTemplate(sixMonthsEmail.subject, {
          EMP_NAME: name,
        });
        const body = fillTemplate(sixMonthsEmail.body, {
          EMP_NAME: name,
          AL_REQUEST_FORM: alRequestForm,
        });

        sendEmail(email, subject, body);
        Logger.log(`Six-month AL reminder sent to ${name}`);
      }

      // --- Anniversary AL reset ---
      if (isAnniversary(joinDate)) {
        // Update leaves: Total 21, Used 0, Remaining 21
        empSheet
          .getRange(idx + 2, empCol["total_leaves"] + 1, 1, 3)
          .setValues([[21, 0, 21]]);

        const years = anniversaryYears(joinDate);

        const subject = fillTemplate(alResetAnniversaryEmail.subject, {
          EMP_NAME: name,
        });
        const body = fillTemplate(alResetAnniversaryEmail.body, {
          EMP_NAME: name,
          ANNIVERSARY_YEARS: years,
          AL_REQUEST_FORM: alRequestForm,
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

const manageProbationPeriod = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
  const probationSheet = ss.getSheetByName(PROBATION_SHEET);

  if (!empSheet || !probationSheet)
    throw new Error("Missing required sheet(s): Employees or Probation.");

  // === 🔹 Get data and dynamic column indexes ===
  const empIndex = getColumnIndexes(empSheet);
  const probIndex = getColumnIndexes(probationSheet);
  const employees = empSheet.getDataRange().getValues().slice(1);
  const probationData = probationSheet.getDataRange().getValues().slice(1);

  const hrEmails = getHRList();
  const emailTemplates = getEmailTemplates();

  // === 🔹 Template mapping for fast lookup ===
  const templates = emailTemplates.reduce((acc, t) => {
    acc[t.code] = t;
    return acc;
  }, {});

  const probationHRnotify = templates.PROBATION_HR_NOTIFY;
  const probationTeamleadNotify = templates.PROBATION_TEAMLEAD_NOTIFY;
  const probationPassEmail = templates.PROBATION_PASSED;

  if (!probationHRnotify || !probationTeamleadNotify || !probationPassEmail)
    throw new Error(
      "Missing one or more probation email templates in Settings sheet."
    );

  // === 🔹 Validate required columns ===
  const requiredEmpCols = ["name", "email", "join_date", "department"];
  requiredEmpCols.forEach((c) => {
    if (empIndex[c] == null)
      throw new Error(`Missing '${c}' column in Employees sheet.`);
  });

  const requiredProbCols = ["employee_name", "result"];
  requiredProbCols.forEach((c) => {
    if (probIndex[c] == null)
      throw new Error(`Missing '${c}' column in Probation sheet.`);
  });

  const today = new Date();

  employees.forEach((row) => {
    const name = (row[empIndex["name"]] || "").toString().trim();
    const email = (row[empIndex["email"]] || "").toString().trim();
    const joinDateStr = row[empIndex["join_date"]];
    const department = (row[empIndex["department"]] || "")
      .toString()
      .trim()
      .toLowerCase();

    if (!name || !email || !joinDateStr) return;

    const joinDate = parseDate(joinDateStr);
    const probationEndDate = new Date(joinDate);
    probationEndDate.setMonth(probationEndDate.getMonth() + 3);

    const probationNotifyDate = new Date(probationEndDate);
    probationNotifyDate.setDate(probationNotifyDate.getDate() - 7);

    const commonMap = {
      EMP_NAME: name,
      JOIN_DATE: getFormattedDate(joinDate),
      PROBATION_END_DATE: getFormattedDate(probationEndDate),
    };

    // === 📅 Notify HR & TL before probation end ===
    if (isSameDayMonthYear(today, probationNotifyDate)) {
      // HR Notification
      const hrSubject = fillTemplate(probationHRnotify.subject, commonMap);
      const hrBody = fillTemplate(probationHRnotify.body, commonMap);
      hrEmails.forEach((hr) => sendEmail(hr.email, hrSubject, hrBody));

      // TL Notification (by department)
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
      const newRow = [
        name,
        email,
        getFormattedDate(joinDate),
        getFormattedDate(probationEndDate),
        getFormattedDate(probationNotifyDate),
      ];

      // Append the row first
      const appendRowIndex = probationSheet.getLastRow() + 1;
      probationSheet.appendRow(newRow);

      // Set dropdown for "Result" column
      const resultColIndex = probIndex["result"] + 1; // +1 because sheet ranges are 1-based
      const resultCell = probationSheet.getRange(
        appendRowIndex,
        resultColIndex
      );

      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Probation Passed", "Probation Not Passed"], true)
        .setAllowInvalid(false)
        .build();

      resultCell.setDataValidation(rule);
    }

    // === ✅ On actual end date → Send "Probation Passed" email ===
    if (isSameDayMonthYear(today, probationEndDate)) {
      const empProbRow = probationData.find(
        (r) => (r[probIndex["employee_name"]] || "").toString().trim() === name
      );

      if (empProbRow) {
        const result = (empProbRow[probIndex["result"]] || "")
          .toString()
          .toLowerCase();

        if (result === "probation passed") {
          const alRequestForm =
            PropertiesService.getScriptProperties().getProperty(
              "AL_REQUEST_FORM"
            );

          const passMap = {
            EMP_NAME: name,
            AL_REQUEST_FORM: alRequestForm || "",
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

const autoTriggerMainFunction = () => {
  sendBirthdayNotifications();
  sendQuarterlyLeaveReminders();
  manageAnnualLeaves();
  manageProbationPeriod();
};
