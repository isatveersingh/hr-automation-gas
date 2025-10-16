// Utility functions for date parsing, formatting, and HR logic

const getColumnIndexes = (sheet) => {
  // Get all headers (first row)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const colIndex = {};

  headers.forEach((header, i) => {
    // Normalize header:
    // - trim leading/trailing spaces
    // - replace multiple spaces/tabs with single underscore
    // - convert to lowercase
    const normalized = header.trim().replace(/\s+/g, "_").toLowerCase();

    colIndex[normalized] = i; // zero-based index for arrays
  });

  return colIndex;
};

const getHRList = (sheetName = "Settings") => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const col = getColumnIndexes(sheet);
  const data = sheet.getDataRange().getValues().slice(1); // skip header

  const hrList = data
    .filter((r) => r[col["hr_name"]] || r[col["hr_email"]])
    .map((r) => ({
      name: r[col["hr_name"]],
      email: r[col["hr_email"]],
    }))
    .filter((h) => h.email); // only keep those with email

  return hrList;
};

const getTeamLeadList = (sheetName = "Settings") => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const col = getColumnIndexes(sheet);
  const data = sheet.getDataRange().getValues().slice(1);

  const teamLeads = data
    .filter((r) => r[col["team_lead_name"]] || r[col["team_lead_email"]])
    .map((r) => ({
      name: r[col["team_lead_name"]],
      email: r[col["team_lead_email"]],
    }))
    .filter((tl) => tl.email);

  return teamLeads;
};

const getEmailTemplates = (sheetName = "Settings") => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const col = getColumnIndexes(sheet);
  const data = sheet.getDataRange().getValues().slice(1);

  const templates = data
    .filter(
      (r) =>
        r[col["email_code"]] || r[col["email_subject"]] || r[col["email_body"]]
    )
    .map((r) => ({
      code: r[col["email_code"]],
      subject: r[col["email_subject"]],
      body: r[col["email_body"]],
    }))
    .filter((t) => t.code);

  return templates;
};

/**
 * Parses a date string or Date object into a Date object.
 * Supports "dd/mm/yyyy", "dd-mm-yyyy", "dd.mm.yyyy" formats.
 */
const parseDate = (value) => {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.replace(/[.\-]/g, "/");
    const [day, month, year] = normalized.split("/");
    return new Date(year, month - 1, day); // month is 0-based
  }
};

/**
 * Checks if two dates are the same day (ignores year).
 */
const isSameDayMonth = (d1, d2) => {
  return d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
};

/**
 * Returns true if the given date is a weekend (Saturday or Sunday).
 */
const isWeekend = (date) => {
  const day = date.getDay();
  return day === 0 || day === 6;
};

/**
 * Returns the previous working day (skips weekends).
 */
const getPreviousWorkingDay = (date) => {
  let previousDay = new Date(date);
  previousDay.setDate(date.getDate() - 1);
  while (isWeekend(previousDay)) {
    previousDay.setDate(previousDay.getDate() - 1);
  }
  return previousDay;
};

/**
 * Formats a date as "dd.MM.yyyy" in GMT+8 timezone.
 */
const getFormattedDate = (date) => {
  return Utilities.formatDate(date, "GMT+8", "dd.MM.yyyy");
};

/**
 * Sends an email using MailApp with a custom sender name.
 */
const sendEmail = (email, subject, body) => {
  MailApp.sendEmail(email, subject, body, {
    name: EMAIL_SENDER,
  });
};

/**
 * Checks if today is the quarterly reminder date (1st of Mar, Jun, Sep, Dec).
 */
const isTodayQuarterlyReminderDate = () => {
  const today = new Date();
  const month = today.getMonth();
  const day = today.getDate();
  return day === 1 && [2, 5, 8, 11].includes(month); // Mar, Jun, Sep, Dec
};

/**
 * Checks if two dates are the same day, month, year.
 */
const isSameDayMonthYear = (d1, d2) => {
  return (
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate() &&
    d1.getFullYear() === d2.getFullYear()
  );
};

/**
 * Checks if today is the 6-month complete from the join date.
 */
const isSixMonthComplete = (joinDate) => {
  const sixMonthDate = new Date(joinDate);
  sixMonthDate.setMonth(sixMonthDate.getMonth() + 6);
  return isSameDayMonthYear(new Date(), sixMonthDate);
};

/**
 * Returns the number of full months between two dates.
 */
const monthsBetween = (d1, d2) => {
  const months =
    (d2.getFullYear() - d1.getFullYear()) * 12 +
    (d2.getMonth() - d1.getMonth());
  return months + (d2.getDate() >= d1.getDate() ? 0 : -1);
};

const isAnniversary = (joinDate) => {
  const today = new Date();
  const months = monthsBetween(joinDate, today);
  return months >= 12 && isSameDayMonth(joinDate, today);
};

const anniversaryYears = (joinDate) => {
  const today = new Date();

  return today.getFullYear() - joinDate.getFullYear();
};
