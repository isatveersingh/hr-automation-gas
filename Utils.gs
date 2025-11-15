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

const getHRList = () => {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
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

const getTeamLeadList = () => {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
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

const getEmailTemplates = () => {
  // === 🔹 Access Settings sheet ===
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  
  // === 🔹 Get column indexes and data ===
  const col = getColumnIndexes(sheet);
  const data = sheet.getDataRange().getValues().slice(1);

  // === 🔹 Extract email templates ===
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
 * getEmployees()
 * Convenience function to retrieve employee data with column indexes.
 * 
 * @returns {Object} Object with { sheet, col: columnIndexes, data: employeeArray }
 */
const getEmployees = () => {
  // === 🔹 Access Employees sheet ===
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEES_SHEET);
  
  // === 🔹 Get column indexes and data ===
  const col = getColumnIndexes(sheet);
  const data = sheet.getDataRange().getValues().slice(1); // skip header row
  
  return { sheet, col, data };
};

/**
 * fillTemplate(template, map)
 * Replaces placeholders in email templates with actual values.
 * Placeholders are in format [PLACEHOLDER_NAME].
 * Replacement is case-insensitive.
 * 
 * @param {string} template - Email template text with [PLACEHOLDER] tokens
 * @param {Object} map - Object mapping placeholder names to replacement values
 * @returns {string} Template with all placeholders replaced
 * @example
 * fillTemplate("Hello [EMP_NAME], your leave is on [START_DATE]", 
 *   { EMP_NAME: "John", START_DATE: "01.11.2025" })
 * // Returns: "Hello John, your leave is on 01.11.2025"
 */
const fillTemplate = (template, map) => {
  // === 🔹 Convert template to string and iterate through replacements ===
  let result = template.toString();
  for (const key in map) {
    // Replace all occurrences (case-insensitive) with actual values
    result = result.replace(new RegExp(`\\[${key}\\]`, "gi"), map[key]);
  }
  return result;
};

/**
 * parseDate(value)
 * Converts various date formats into a JavaScript Date object.
 * Supports multiple date formats for flexibility.
 * 
 * @param {Date|string} value - Date object or string in format "dd/mm/yyyy", "dd-mm-yyyy", or "dd.mm.yyyy"
 * @returns {Date} Parsed Date object
 * @example
 * parseDate("15/11/2025") // Returns Date for November 15, 2025
 * parseDate("15-11-2025") // Same result
 * parseDate(new Date()) // Returns same date object
 */
const parseDate = (value) => {
  // === 🔹 If already a Date object, return as is ===
  if (value instanceof Date) {
    return value;
  }

  // === 🔹 If string, parse to Date ===
  if (typeof value === "string") {
    // Normalize separators to forward slash
    const normalized = value.replace(/[.\-]/g, "/");
    const [day, month, year] = normalized.split("/");
    return new Date(year, month - 1, day); // month is 0-based in JS
  }
};

/**
 * isSameDayMonth(d1, d2)
 * Checks if two dates fall on the same day and month, ignoring the year.
 * Useful for checking birthdays or monthly recurring dates.
 * 
 * @param {Date} d1 - First date
 * @param {Date} d2 - Second date
 * @returns {boolean} True if same day and month
 */
const isSameDayMonth = (d1, d2) => {
  return d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
};

/**
 * isWeekend(date)
 * Determines if a given date is a weekend (Saturday or Sunday).
 * 
 * @param {Date} date - Date to check
 * @returns {boolean} True if Saturday (6) or Sunday (0)
 */
const isWeekend = (date) => {
  const day = date.getDay();
  return day === 0 || day === 6;
};

/**
 * getPreviousWorkingDay(date)
 * Calculates the previous working day, skipping weekends.
 * Used for scheduling notifications on business days.
 * 
 * @param {Date} date - Reference date
 * @returns {Date} Previous working day (Monday-Friday)
 */
const getPreviousWorkingDay = (date) => {
  // === 🔹 Start with previous day ===
  let previousDay = new Date(date);
  previousDay.setDate(date.getDate() - 1);
  
  // === 🔹 Skip backwards past weekends ===
  while (isWeekend(previousDay)) {
    previousDay.setDate(previousDay.getDate() - 1);
  }
  return previousDay;
};

/**
 * getFormattedDate(date)
 * Formats a date as "dd.MM.yyyy" in GMT+8 timezone.
 * Used consistently throughout the system for date display.
 * 
 * @param {Date} date - Date to format
 * @returns {string} Formatted date string (e.g., "15.11.2025")
 */
const getFormattedDate = (date) => {
  return Utilities.formatDate(date, "GMT+8", "dd.MM.yyyy");
};

/**
 * sendEmail(email, subject, body)
 * Sends an email using Google's MailApp with a custom sender name.
 * Uses the EMAIL_SENDER constant defined in Constants.gs.
 * 
 * @param {string} email - Recipient email address
 * @param {string} subject - Email subject line
 * @param {string} body - Email body content
 * @returns {void}
 */
const sendEmail = (email, subject, body) => {
  MailApp.sendEmail(email, subject, body, {
    name: EMAIL_SENDER,
  });
};

/**
 * isTodayQuarterlyReminderDate()
 * Checks if today is a quarterly reminder date.
 * Quarterly reminders are sent on the 1st of March, June, September, and December.
 * 
 * @returns {boolean} True if today is the 1st of Q1, Q2, Q3, or Q4
 */
const isTodayQuarterlyReminderDate = () => {
  const today = new Date();
  const month = today.getMonth();
  const day = today.getDate();
  // Mar=2, Jun=5, Sep=8, Dec=11 (0-indexed months)
  return day === 1 && [2, 5, 8, 11].includes(month);
};

/**
 * isSameDayMonthYear(d1, d2)
 * Checks if two dates are exactly the same (same day, month, and year).
 * 
 * @param {Date} d1 - First date
 * @param {Date} d2 - Second date
 * @returns {boolean} True if dates are identical
 */
const isSameDayMonthYear = (d1, d2) => {
  return (
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate() &&
    d1.getFullYear() === d2.getFullYear()
  );
};

/**
 * isSixMonthComplete(joinDate)
 * Checks if today is exactly 6 months after the join date.
 * Used to send 6-month milestone notifications to employees.
 * 
 * @param {Date} joinDate - Employee's join date
 * @returns {boolean} True if today is the 6-month anniversary
 */
const isSixMonthComplete = (joinDate) => {
  // === 🔹 Calculate 6-month anniversary date ===
  const sixMonthDate = new Date(joinDate);
  sixMonthDate.setMonth(sixMonthDate.getMonth() + 6);
  
  return isSameDayMonthYear(new Date(), sixMonthDate);
};

/**
 * monthsBetween(d1, d2)
 * Calculates the number of full months between two dates.
 * Accounts for day-of-month differences in the calculation.
 * 
 * @param {Date} d1 - Start date
 * @param {Date} d2 - End date
 * @returns {number} Number of full months between dates
 */
const monthsBetween = (d1, d2) => {
  // === 🔹 Calculate month difference ===
  const months =
    (d2.getFullYear() - d1.getFullYear()) * 12 +
    (d2.getMonth() - d1.getMonth());
  
  // === 🔹 Adjust if end date day is before start date day ===
  return months + (d2.getDate() >= d1.getDate() ? 0 : -1);
};

/**
 * isAnniversary(joinDate)
 * Checks if today is the employee's work anniversary (yearly).
 * Matches month and day, year difference must be >= 1 year.
 * 
 * @param {Date} joinDate - Employee's join date
 * @returns {boolean} True if today is work anniversary
 */
const isAnniversary = (joinDate) => {
  const today = new Date();
  const months = monthsBetween(joinDate, today);
  
  // === 🔹 Require at least 12 months and matching month/day ===
  return months >= 12 && isSameDayMonth(joinDate, today);
};

/**
 * anniversaryYears(joinDate)
 * Calculates how many complete years the employee has been with the company.
 * 
 * @param {Date} joinDate - Employee's join date
 * @returns {number} Number of years of service
 */
const anniversaryYears = (joinDate) => {
  const today = new Date();
  return today.getFullYear() - joinDate.getFullYear();
};
