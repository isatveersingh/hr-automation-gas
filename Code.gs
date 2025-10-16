/**
 * Send birthday notifications / reminders to colleagues and HR
 */

const sendBirthdayNotifications = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const empsheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET);
  const settingsheet = spreadsheet.getSheetByName(SETTINGS_SHEET);

  const emplastrow = empsheet.getLastRow();
  const employees = empsheet.getRange(`A2:H${emplastrow}`).getValues();
  const hrEmails = settingsheet
    .getRange("A2:B")
    .getValues()
    .filter((row) => row[0].toString().trim() !== "")
    .map((row) => row[1]);

  const emailTemplates = settingsheet.getRange("E2:G").getValues();

  const birthdayEmail = emailTemplates.find(
    (row) => row[0].toString().trim() === "BIRTHDAY_ALERT"
  );
  const birthdayHREmail = emailTemplates.find(
    (row) => row[0].toString().trim() === "BIRTHDAY_ALERT_HR"
  );

  const today = new Date();

  for (let emp of employees) {
    const [name, email, birthdayStr] = emp;

    if (name.toString().trim() === "") {
      continue;
    }

    const birthday = parseDate(birthdayStr);
    const thisYearBirthday = new Date(
      today.getFullYear(),
      birthday.getMonth(),
      birthday.getDate()
    );

    const thisYearBirthdayFormatted = getFormattedDate(thisYearBirthday);

    if (isSameDayMonthYear(today, thisYearBirthday)) {
      const colleagues = employees
        .filter((e) => e[1] !== email)
        .map((e) => e[1]);

      const subject = birthdayEmail[1]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name);
      const body = birthdayEmail[2]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name)
        .replace(/\[BIRTHDAY\]/gi, thisYearBirthdayFormatted);

      for (let coll of colleagues) {
        sendEmail(coll, subject, body);
      }
    }

    const notifydate = getPreviousWorkingDay(thisYearBirthday);

    if (isSameDayMonthYear(today, notifydate)) {
      const subject = birthdayHREmail[1]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name);
      const body = birthdayHREmail[2]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name)
        .replace(/\[BIRTHDAY\]/gi, thisYearBirthdayFormatted);

      for (let hr of hrEmails) {
        sendEmail(hr, subject, body);
      }
    }
  }
};

/**
 * Sends quarterly leave reminders to employees and HR.
 */
const sendQuarterlyLeaveReminders = () => {
  if (!isTodayQuarterlyReminderDate()) return;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const empsheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET);
  const settingsheet = spreadsheet.getSheetByName(SETTINGS_SHEET);

  const emplastrow = empsheet.getLastRow();
  const employees = empsheet.getRange(`A2:H${emplastrow}`).getValues();

  const hrEmails = settingsheet
    .getRange("A2:B")
    .getValues()
    .filter((row) => row[0].toString().trim() !== "")
    .map((row) => row[1]);
  const emailTemplates = settingsheet.getRange("E2:G").getValues();

  const quarterlyAlReminderEmail = emailTemplates.find(
    (row) => row[0].toString().trim() === "QUARTERLY_AL_REMINDER"
  );
  const obligatoryALReminder = emailTemplates.find(
    (row) => row[0].toString().trim() === "HR_OBLIGATORY_AL_REMINDER"
  );

  const alRequestForm =
    PropertiesService.getScriptProperties().getProperty("AL_REQUEST_FORM");

  for (let emp of employees) {
    const [
      name,
      email,
      _birthdayStr,
      _joinDateStr,
      totalLeaves,
      leavesUsed,
      remainingLeaves,
    ] = emp;

    if (name.toString().trim() === "") {
      continue;
    }

    const subject = quarterlyAlReminderEmail[1]
      .toString()
      .replace(/\[EMP_NAME\]/gi, name);

    const body = quarterlyAlReminderEmail[2]
      .toString()
      .replace(/\[EMP_NAME\]/gi, name)
      .replace(/\[AL_REQUEST_FORM\]/gi, alRequestForm)
      .replace(/\[TOTAL_AL\]/gi, totalLeaves || 0)
      .replace(/\[AL_USED\]/gi, leavesUsed || 0)
      .replace(/\[AL_REMAINING\]/gi, remainingLeaves || 0);

    sendEmail(email, subject, body);

    if (parseInt(leavesUsed || 0) < 14) {
      const hrSubject = obligatoryALReminder[1]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name);
      const hrBody = obligatoryALReminder[2]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name)
        .replace(/\[TOTAL_AL\]/gi, totalLeaves || 0)
        .replace(/\[AL_USED\]/gi, leavesUsed || 0)
        .replace(/\[AL_REMAINING\]/gi, remainingLeaves || 0);

      for (let hr of hrEmails) {
        sendEmail(hr, hrSubject, hrBody);
      }
    }
  }
};

/**
 * Send notification when employee completes 6 months
 * When employee has an anniversary, resert the annual leaves
 */

const manageAnnualLeaves = () => {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const empsheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET);
  const settingsheet = spreadsheet.getSheetByName(SETTINGS_SHEET);

  const emplastrow = empsheet.getLastRow();
  const employees = empsheet.getRange(`A2:H${emplastrow}`).getValues();

  const emailTemplates = settingsheet.getRange("E2:G").getValues();

  const alResetAnniversaryEmail = emailTemplates.find(
    (row) => row[0].toString().trim() === "AL_RESET_ANNIVERSARY"
  );

  const sixMonthsEmail = emailTemplates.find(
    (row) => row[0].toString().trim() === "SIX_MONTH_AL_REMINDER"
  );

  const alRequestForm =
    PropertiesService.getScriptProperties().getProperty("AL_REQUEST_FORM");

  for (let i = 0; i < employees.length; i++) {
    const [name, email, _birthdayStr, joinDateStr] = employees[i];

    if (name.toString().trim() === "") {
      continue;
    }

    const joinDate = parseDate(joinDateStr);

    if (isSixMonthComplete(joinDate)) {
      empsheet.getRange(`E${i + 2}:G${i + 2}`).setValues([[21, 0, 21]]);

      const subject = sixMonthsEmail[1]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name);

      const body = sixMonthsEmail[2]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name)
        .replace(/\[AL_REQUEST_FORM\]/gi, alRequestForm);

      sendEmail(email, subject, body);
    }

    if (isAnniversary(joinDate)) {
      empsheet.getRange(`E${i + 2}:G${i + 2}`).setValues([[21, 0, 21]]);
      const years = anniversaryYears(joinDate);
      const subject = alResetAnniversaryEmail[1]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name);

      const body = alResetAnniversaryEmail[2]
        .toString()
        .replace(/\[EMP_NAME\]/gi, name)
        .replace(/\[ANNIVERSARY_YEARS\]/gi, years)
        .replace(/\[AL_REQUEST_FORM\]/gi, alRequestForm);

      sendEmail(email, subject, body);
    }
  }
};

const manageProbationPeriod = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(EMPLOYEES_SHEET);
  const probationSheet = ss.getSheetByName(PROBATION_SHEET);

  if (!empSheet || !probationSheet)
    throw new Error("One or more required sheets are missing.");

  // === 🔹 Get data and dynamic column indexes ===
  const empIndex = getColumnIndexes(empSheet);
  const probIndex = getColumnIndexes(probationSheet);

  const employees = empSheet.getDataRange().getValues().slice(1);
  const probationData = probationSheet.getDataRange().getValues().slice(1);

  const hrEmails = getHRList();
  const emailTemplates = getEmailTemplates();

  // === 🔹 Validate templates ===
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
      throw new Error(`Missing '${c}' column in Employees sheet`);
  });

  const requiredProbCols = ["employee_name", "result"];
  requiredProbCols.forEach((c) => {
    if (probIndex[c] == null)
      throw new Error(`Missing '${c}' column in Probation sheet`);
  });

  // === 🔹 Process employees ===
  const today = new Date();
  const newProbationRows = [];

  employees.forEach((row) => {
    const name = (row[empIndex["name"]] || "").toString().trim();
    const email = (row[empIndex["email"]] || "").toString().trim();
    const joinDateStr = row[empIndex["join_date"]];
    const department = (row[empIndex["department"]] || "")
      .toString()
      .trim()
      .toLowerCase();

    if (!name || !email || !joinDateStr) return; // skip incomplete rows

    const joinDate = parseDate(joinDateStr);
    const probationEndDate = new Date(joinDate);
    probationEndDate.setMonth(probationEndDate.getMonth() + 3);

    const probationNotifyDate = new Date(probationEndDate);
    probationNotifyDate.setDate(probationNotifyDate.getDate() - 7);

    // === 📅 Notify HR & Team Leads 7 days before probation ends ===
    if (isSameDayMonthYear(today, probationNotifyDate)) {
      const commonMap = {
        EMP_NAME: name,
        JOIN_DATE: getFormattedDate(joinDate),
        PROBATION_END_DATE: getFormattedDate(probationEndDate),
      };

      const subject = fillTemplate(probationHRnotify.subject, commonMap);
      const body = fillTemplate(probationHRnotify.body, commonMap);

      hrEmails.forEach((hr) => sendEmail(hr.email, subject, body));
      Logger.log(`Probation HR notify sent for ${name}`);

      const depLead =
        department === "service"
          ? SERVICE_DEP_LEAD
          : department === "client"
          ? CLIENT_DEP_LEAD
          : null;

      if (depLead) {
        const tlMap = {
          ...commonMap,
          TEAMLEAD_NAME: depLead.name,
        };

        const tlSubject = fillTemplate(probationTeamleadNotify.subject, tlMap);
        const tlBody = fillTemplate(probationTeamleadNotify.body, tlMap);

        sendEmail(depLead.email, tlSubject, tlBody);
        Logger.log(`Team lead notify sent for ${name} (${depLead.name})`);
      }

      newProbationRows.push([
        name,
        getFormattedDate(joinDate),
        getFormattedDate(probationEndDate),
        getFormattedDate(probationNotifyDate),
      ]);
    }

    // === ✅ On actual probation end date, send "passed" email ===
    if (isSameDayMonthYear(today, probationEndDate)) {
      const empProbationRow = probationData.find(
        (r) => (r[probIndex["employee_name"]] || "").toString().trim() === name
      );

      if (empProbationRow) {
        const result = (empProbationRow[probIndex["result"]] || "")
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
          Logger.log(`Probation pass email sent to ${name} (${email})`);
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
