/**
 * Global Constants for HR Automation System
 * These constants define sheet names and key contact information
 */

// === 🔹 Sheet Names === 
// All references to Google Sheets must use these exact names

/**
 * EMPLOYEES_SHEET
 * Main employee database containing:
 * - Employee name, email, join date
 * - Leave balances (total, used, remaining)
 * - Department information
 * - Birthday and other personal details
 */
const EMPLOYEES_SHEET = "HR List";

/**
 * SETTINGS_SHEET
 * Configuration sheet containing:
 * - HR department contacts (name, email)
 * - Team leads for approval workflows (name, email)
 * - Email templates with codes, subjects, and body content
 */
const SETTINGS_SHEET = "Settings";

/**
 * AL_STATISTIC_SHEET
 * Annual Leave statistics tracking sheet
 * Records all leave requests (Annual Leave and Compensatory Days Off)
 * Contains: Employee name, email, start date, end date, leave type, vacation type, days count
 */
const AL_STATISTIC_SHEET = "AL Statistic";

/**
 * SICK_LEAVE_SHEET
 * Sick leave records tracking sheet
 * Contains: Employee name, email, start date, end date
 */
const SICK_LEAVE_SHEET = "Sick Leaves";

/**
 * PROBATION_SHEET
 * Probation period tracking for new employees
 * Contains: Employee name, email, join date, probation end date, result (dropdown)
 * Used to monitor and manage employee probation workflows
 */
const PROBATION_SHEET = "Probation";

/**
 * FEEDBACK_SHEET
 * Employee feedback collection during probation
 * Contains: Employee name, email, feedback category, colleague name (optional), feedback text, date
 */
const FEEDBACK_SHEET = "Feedback";

// === 🔹 Email Configuration ===

/**
 * EMAIL_SENDER
 * Display name for all automated emails sent by the system
 * Appears as the sender in employee inboxes
 */
const EMAIL_SENDER = "Sino Services HR Department";

// === 🔹 Department Leadership Contacts ===
// These constants map departments to their respective leads for notifications

/**
 * SERVICE_DEP_LEAD
 * Service Department Team Lead contact
 * Used for:
 * - Probation notifications
 * - Department-specific HR communications
 * 
 * Update with actual team lead details when deploying
 */
const SERVICE_DEP_LEAD = {
  name: "Lead One",
  email: "demo.satveer@gmail.com",
};

/**
 * CLIENT_DEP_LEAD
 * Client Department Team Lead contact
 * Used for:
 * - Probation notifications
 * - Department-specific HR communications
 * 
 * Update with actual team lead details when deploying
 */
const CLIENT_DEP_LEAD = {
  name: "Iulia Avdeeva",
  email: "mail.satveer@gmail.com",
};
