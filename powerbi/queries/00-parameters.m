// ============================================================
// Power BI PARAMETERS
// Create each of these in Power Query: Home > Manage Parameters > New.
// Do NOT paste this whole file as one query - it is a description of
// the parameters you need to create, one at a time.
// ============================================================

// --- 1. TimesheetFolderPath  (Type: Text) -------------------
// Current value example:
//   C:\Users\<you>\OneDrive\Utilisation\timesheets
// Every Employment Hero timesheet CSV export gets dropped in here.
// One file per pay period is fine - the query unions them all and
// de-duplicates.

// --- 2. MappingFolderPath  (Type: Text) ---------------------
// Current value example:
//   C:\Users\<you>\OneDrive\Utilisation\mapping
// Holds worktype-utilisation-map.csv, employee-targets.csv and
// public-holidays.csv from the /powerbi/mapping folder of this repo.

// --- 3. FY_Start_Year  (Type: Whole Number) -----------------
// Example: 2024  -> calendar starts 1 July 2024

// --- 4. FY_End_Year  (Type: Whole Number) -------------------
// Example: 2027  -> calendar ends 30 June 2028

// --- 5. StandardDayHours  (Type: Decimal Number) ------------
// Example: 7.6
// Used to convert a public holiday into hours removed from capacity,
// and to spread weekly contracted hours across the working week.

// --- 6. PayrollApiBase  (Type: Text) ------------------------
// ONLY needed if you use 02-fact-timesheet-api.m instead of CSVs.
// Value: https://api.yourpayroll.com.au
// Keep this as a plain base URL with no query string - see the note
// in 02-fact-timesheet-api.m about Web.Contents and scheduled refresh.

// --- 7. BusinessId  (Type: Text) ----------------------------
// ONLY needed for the API path. Your Employment Hero Payroll
// business ID (the number in the URL when you are inside the business).
