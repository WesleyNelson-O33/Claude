// Utilisation model - measures, relationships and formatting.
// Run in Tabular Editor 2 (free, tabulareditor.com). Open your .pbix, then
// External Tools > Tabular Editor > "C# Script" tab. Paste this, press F5,
// then Ctrl+S to write it back. Leave Power BI Desktop open while you do.

Table m;
if (Model.Tables.Contains("_Measures")) m = Model.Tables["_Measures"];
else { m = Model.AddCalculatedTable("_Measures", "{1}"); m.Columns[0].IsHidden = true; }

var TotalHours = m.AddMeasure("Total Hours", "SUM ( Fact_Timesheet[Hours] )");
TotalHours.FormatString = "#,0.0";
TotalHours.Description = "Every timesheet hour, including leave and public holidays.";
var ChargeableHours = m.AddMeasure("Chargeable Hours", "SUM ( Fact_Timesheet[Chargeable Hours] )");
ChargeableHours.FormatString = "#,0.0";
ChargeableHours.Description = "Hours on jobs classified Chargeable.";
var ChargeableHoursOrdinary = m.AddMeasure("Chargeable Hours (Ordinary)", "CALCULATE ( [Chargeable Hours], Dim_PayType[Category] = \"Ordinary\" )");
ChargeableHoursOrdinary.FormatString = "#,0.0";
ChargeableHoursOrdinary.Description = "Chargeable hours at ordinary time. This is the utilisation numerator - overtime sits outside the ratio.";
var NonChargeableHours = m.AddMeasure("Non-Chargeable Hours", "CALCULATE ( [Total Hours] - [Chargeable Hours], Dim_PayType[Category] = \"Ordinary\" )");
NonChargeableHours.FormatString = "#,0.0";
NonChargeableHours.Description = "Ordinary worked hours that were not chargeable. Excludes leave, public holidays and overtime.";
var LeaveHours = m.AddMeasure("Leave Hours", "CALCULATE ( [Total Hours], Dim_PayType[Category] = \"Leave\" )");
LeaveHours.FormatString = "#,0.0";
var PublicHolidayHours = m.AddMeasure("Public Holiday Hours", "CALCULATE ( [Total Hours], Dim_PayType[Category] = \"Public Holiday\" )");
PublicHolidayHours.FormatString = "#,0.0";
var OvertimeHours = m.AddMeasure("Overtime Hours", "CALCULATE ( [Total Hours], Dim_PayType[Category] = \"Overtime\" )");
OvertimeHours.FormatString = "#,0.0";
OvertimeHours.Description = "Reported beside the ratio, never inside it. Overtime is excluded from both the numerator and the denominator.";
var WorkedHours = m.AddMeasure("Worked Hours", "[Chargeable Hours (Ordinary)] + [Non-Chargeable Hours]");
WorkedHours.FormatString = "#,0.0";
WorkedHours.Description = "Actual work. Excludes leave, public holidays and overtime.";
var PermanentCapacityHours = m.AddMeasure("Permanent Capacity Hours", "SUM ( Fact_Capacity[Capacity Hours] )");
PermanentCapacityHours.FormatString = "#,0.0";
PermanentCapacityHours.Description = "Contracted hours for permanents, on working days they were employed.";
var CasualCapacityHours = m.AddMeasure("Casual Capacity Hours", "CALCULATE ( [Worked Hours], Dim_Employee[Is Casual] = TRUE () )");
CasualCapacityHours.FormatString = "#,0.0";
CasualCapacityHours.Description = "For a casual, capacity is the hours you engaged them for. You carry no idle cost for a casual you did not call in.";
var CapacityHours = m.AddMeasure("Capacity Hours", "[Permanent Capacity Hours] + [Casual Capacity Hours]");
CapacityHours.FormatString = "#,0.0";
var AvailableHours = m.AddMeasure("Available Hours", "VAR Cap = [Capacity Hours]\nVAR Lv = [Leave Hours]\nRETURN MAX ( Cap - Lv, 0 )");
AvailableHours.FormatString = "#,0.0";
AvailableHours.Description = "Capacity less leave taken. Public holidays are NOT subtracted again - they are already excluded from Fact_Capacity.";
var Utilisation = m.AddMeasure("Utilisation %", "DIVIDE (\n    CALCULATE ( [Chargeable Hours (Ordinary)], Dim_Employee[Utilisation Scope] = \"Delivery\" ),\n    CALCULATE ( [Available Hours],   Dim_Employee[Utilisation Scope] = \"Delivery\" )\n)");
Utilisation.FormatString = "0.0%";
Utilisation.Description = "HEADLINE. Delivery staff only, casuals included. Overhead staff are excluded from the ratio but their hours are still reported.";
var UtilisationAllStaff = m.AddMeasure("Utilisation % (All Staff)", "DIVIDE ( [Chargeable Hours (Ordinary)], [Available Hours] )");
UtilisationAllStaff.FormatString = "0.0%";
UtilisationAllStaff.Description = "Every person including overhead roles. Materially lower - use for capacity planning, not performance.";
var BillableMix = m.AddMeasure("Billable Mix %", "DIVIDE ( [Chargeable Hours (Ordinary)], [Worked Hours] )");
BillableMix.FormatString = "0.0%";
BillableMix.Description = "Of the work actually done, how much was chargeable. Immune to timesheet-completion gaps.";
var Productive = m.AddMeasure("Productive %", "DIVIDE ( [Worked Hours], [Available Hours] )");
Productive.FormatString = "0.0%";
var TargetUtilisation = m.AddMeasure("Target Utilisation %", "DIVIDE (\n    SUMX ( Dim_Employee, [Available Hours] * Dim_Employee[Target Utilisation %] ),\n    SUMX ( Dim_Employee, [Available Hours] )\n)");
TargetUtilisation.FormatString = "0.0%";
TargetUtilisation.Description = "Availability-weighted, so a part-timer moves the group target in proportion to their capacity.";
var TargetChargeableHours = m.AddMeasure("Target Chargeable Hours", "SUMX ( Dim_Employee, [Available Hours] * Dim_Employee[Target Utilisation %] )");
TargetChargeableHours.FormatString = "#,0.0";
var UtilisationvsTargetpp = m.AddMeasure("Utilisation vs Target (pp)", "[Utilisation %] - [Target Utilisation %]");
UtilisationvsTargetpp.FormatString = "+0.0;-0.0;0.0";
UtilisationvsTargetpp.Description = "Percentage POINTS above or below target, not a percentage change.";
var CasualShareofDeliveredHours = m.AddMeasure("Casual Share of Delivered Hours", "DIVIDE ( CALCULATE ( [Worked Hours], Dim_Employee[Is Casual] = TRUE () ), [Worked Hours] )");
CasualShareofDeliveredHours.FormatString = "0.0%";
CasualShareofDeliveredHours.Description = "How much of the work is carried by casual labour.";
var CasualNonChargeable = m.AddMeasure("Casual Non-Chargeable %", "CALCULATE ( DIVIDE ( [Non-Chargeable Hours], [Worked Hours] ), Dim_Employee[Is Casual] = TRUE () )");
CasualNonChargeable.FormatString = "0.0%";
CasualNonChargeable.Description = "Paying casuals to do internal work is the expensive failure mode. Casual utilisation is ~100% by construction and tells you nothing.";
var UtilisationR13W = m.AddMeasure("Utilisation % R13W", "VAR W = DATESINPERIOD ( Dim_Date[Date], MAX ( Dim_Date[Date] ), -91, DAY )\nRETURN DIVIDE (\n    CALCULATE ( [Chargeable Hours (Ordinary)], W, Dim_Employee[Utilisation Scope] = \"Delivery\" ),\n    CALCULATE ( [Available Hours],  W, Dim_Employee[Utilisation Scope] = \"Delivery\" )\n)");
UtilisationR13W.FormatString = "0.0%";
UtilisationR13W.Description = "Rolling 13 weeks. Trend this, not a single fortnight.";
var UtilisationFYTD = m.AddMeasure("Utilisation % FYTD", "VAR FYc = MAX ( Dim_Date[FY] )\nVAR LastD = MAX ( Dim_Date[Date] )\nRETURN CALCULATE ( [Utilisation %], REMOVEFILTERS ( Dim_Date ), Dim_Date[FY] = FYc, Dim_Date[Date] <= LastD )");
UtilisationFYTD.FormatString = "0.0%";
UtilisationFYTD.Description = "Australian financial year to date, written without TOTALYTD's locale-sensitive year-end argument.";
var WIPHoursAdded = m.AddMeasure("WIP Hours Added", "CALCULATE ( [Chargeable Hours], Dim_Job[Billing Basis Final] = \"Project\" )");
WIPHoursAdded.FormatString = "#,0.0";
WIPHoursAdded.Description = "The monthly movement. Chargeable hours on date-coded project jobs. Recurring onsite contracts are billed monthly and never accrue.";
var WIPHourstoDate = m.AddMeasure("WIP Hours to Date", "CALCULATE (\n    [Chargeable Hours],\n    Dim_Job[Billing Basis Final] = \"Project\",\n    REMOVEFILTERS ( Dim_Date ),\n    Dim_Date[Date] <= MAX ( Dim_Date[Date] )\n)");
WIPHourstoDate.FormatString = "#,0.0";
WIPHourstoDate.Description = "Cumulative chargeable hours booked to open project jobs. NOT unbilled hours until Last Fully Billed Date is populated - label it 'hours booked to date'.";
var UnbilledWIPHours = m.AddMeasure("Unbilled WIP Hours", "SUMX (\n    VALUES ( Dim_Job[Job No] ),\n    VAR LastBilled = CALCULATE ( MAX ( Dim_Job[Last Fully Billed Date] ) )\n    VAR Basis      = CALCULATE ( SELECTEDVALUE ( Dim_Job[Billing Basis Final] ) )\n    VAR UpTo       = MAX ( Dim_Date[Date] )\n    RETURN\n        IF ( Basis <> \"Project\", 0,\n            CALCULATE ( [Chargeable Hours], REMOVEFILTERS ( Dim_Date ),\n                Dim_Date[Date] <= UpTo,\n                Dim_Date[Date] > IF ( ISBLANK ( LastBilled ), DATE(1900,1,1), LastBilled ) ) )\n)");
UnbilledWIPHours.FormatString = "#,0.0";
UnbilledWIPHours.Description = "The true balance: chargeable hours since the job was last fully billed. Populate Last Fully Billed Date in map_jobs.csv from the month the job's dollar balance on the WIP schedule returns to zero.";
var HoursWithNoJobNumber = m.AddMeasure("Hours With No Job Number", "CALCULATE ( [Total Hours], ISBLANK ( Fact_Timesheet[Job No] ) )");
HoursWithNoJobNumber.FormatString = "#,0.0";
HoursWithNoJobNumber.Description = "Should be zero. Timesheet lines that cannot be attributed to any job.";
var UnmatchedEmployeeHours = m.AddMeasure("Unmatched Employee Hours", "CALCULATE ( [Total Hours], ISBLANK ( Fact_Timesheet[Employee Key] ) )");
UnmatchedEmployeeHours.FormatString = "#,0.0";
UnmatchedEmployeeHours.Description = "Should be zero. Hours whose Employee Id / Name pair is not in map_employees.csv.";
var UnmappedJobHours = m.AddMeasure("Unmapped Job Hours", "SUMX ( FILTER ( Fact_Timesheet, NOT ISBLANK ( Fact_Timesheet[Job No] ) && ISBLANK ( RELATED ( Dim_Job[Job No] ) ) ), Fact_Timesheet[Hours] )");
UnmappedJobHours.FormatString = "#,0.0";
UnmappedJobHours.Description = "Should be zero. A job was used on a timesheet but is missing from map_jobs.csv.";
var UnmappedPayTypeHours = m.AddMeasure("Unmapped Pay Type Hours", "SUMX ( FILTER ( Fact_Timesheet, ISBLANK ( RELATED ( Dim_PayType[Pay Type Key] ) ) ), Fact_Timesheet[Hours] )");
UnmappedPayTypeHours.FormatString = "#,0.0";
UnmappedPayTypeHours.Description = "Should be zero.";
var TimesheetCoverage = m.AddMeasure("Timesheet Coverage %", "DIVIDE ( [Worked Hours] + [Leave Hours], [Permanent Capacity Hours] )");
TimesheetCoverage.FormatString = "0.0%";
TimesheetCoverage.Description = "Materially below 100% means missing timesheets, not low utilisation. Answer this before drawing any conclusion from Utilisation %.";
var LastTimesheetDate = m.AddMeasure("Last Timesheet Date", "MAX ( Fact_Timesheet[Work Date] )");
LastTimesheetDate.Description = "Read this before the utilisation number. A partial month against a full-month denominator is what made August 2026 read 50.2% when July read 72.6%.";
var DataAgedays = m.AddMeasure("Data Age (days)", "VAR L = [Last Timesheet Date]\nRETURN IF ( ISBLANK ( L ), BLANK (), INT ( TODAY () - L ) )");
DataAgedays.FormatString = "#,0";
var WorkingDaysWithData = m.AddMeasure("Working Days With Data", "CALCULATE ( DISTINCTCOUNT ( Fact_Timesheet[Date Key] ), Dim_Date[Is Weekday] = TRUE () )");
WorkingDaysWithData.FormatString = "#,0";
WorkingDaysWithData.Description = "Compare against Working Days In Period. If it is lower, you are looking at a partial month.";
var WorkingDaysInPeriod = m.AddMeasure("Working Days In Period", "CALCULATE ( COUNTROWS ( Dim_Date ), Dim_Date[Is Weekday] = TRUE () )");
WorkingDaysInPeriod.FormatString = "#,0";
var DataQualityFlag = m.AddMeasure("Data Quality Flag", "VAR I =\n    IF ( [Hours With No Job Number] > 0, 1, 0 )\n  + IF ( [Unmatched Employee Hours] > 0, 1, 0 )\n  + IF ( [Unmapped Job Hours] > 0, 1, 0 )\n  + IF ( [Working Days With Data] < [Working Days In Period], 1, 0 )\nRETURN IF ( I = 0, \"OK\", I & \" issue(s) - see Data Quality page\" )");

// relationships
var rels = new [] {
    new [] {"Fact_Timesheet","Date Key","Dim_Date","Date Key"},
    new [] {"Fact_Timesheet","Employee Key","Dim_Employee","Employee Key"},
    new [] {"Fact_Timesheet","Job No","Dim_Job","Job No"},
    new [] {"Fact_Timesheet","Pay Type Key","Dim_PayType","Pay Type Key"},
    new [] {"Fact_Capacity","Date Key","Dim_Date","Date Key"},
    new [] {"Fact_Capacity","Employee Key","Dim_Employee","Employee Key"},
};
foreach (var r in rels) {
    if (!Model.Tables.Contains(r[0]) || !Model.Tables.Contains(r[2])) continue;
    var ft = Model.Tables[r[0]]; var tt = Model.Tables[r[2]];
    if (!ft.Columns.Contains(r[1]) || !tt.Columns.Contains(r[3])) continue;
    var exists = Model.Relationships.Any(x => x.FromTable == ft && x.ToTable == tt
                 && x.FromColumn.Name == r[1] && x.ToColumn.Name == r[3]);
    if (exists) continue;
    var rel = Model.AddRelationship();
    rel.FromColumn = ft.Columns[r[1]];
    rel.ToColumn   = tt.Columns[r[3]];
    rel.FromCardinality = RelationshipEndCardinality.Many;
    rel.ToCardinality   = RelationshipEndCardinality.One;
    rel.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection;
    rel.IsActive = true;
}

// sort-by columns and the date table
if (Model.Tables.Contains("Dim_Date")) {
    var d = Model.Tables["Dim_Date"];
    d.DataCategory = "Time";
    if (d.Columns.Contains("Date")) ((Column)d.Columns["Date"]).IsKey = true;
    if (d.Columns.Contains("Month") && d.Columns.Contains("Month Sort"))
        d.Columns["Month"].SortByColumn = d.Columns["Month Sort"];
    if (d.Columns.Contains("Day Name") && d.Columns.Contains("Day of Week No"))
        d.Columns["Day Name"].SortByColumn = d.Columns["Day of Week No"];
}

// hide the plumbing so nobody drags a raw column onto a visual
var hide = new [] {
    new [] {"Dim_Date","Date Key"},
    new [] {"Dim_Employee","Employee Key"},
    new [] {"Dim_Employee","Hours"},
    new [] {"Dim_Employee","Rows"},
    new [] {"Dim_PayType","Counts Toward Capacity"},
    new [] {"Dim_PayType","Pay Type Key"},
    new [] {"Fact_Timesheet","Chargeable Category"},
    new [] {"Fact_Timesheet","Chargeable Hours"},
    new [] {"Fact_Timesheet","Date Key"},
    new [] {"Fact_Timesheet","Employee Id"},
    new [] {"Fact_Timesheet","Employee Key"},
    new [] {"Fact_Timesheet","Employee Name"},
    new [] {"Fact_Timesheet","Hours"},
    new [] {"Fact_Timesheet","Job Name Raw"},
    new [] {"Fact_Timesheet","Job No"},
    new [] {"Fact_Timesheet","Pay Type"},
    new [] {"Fact_Timesheet","Pay Type Key"},
    new [] {"Fact_Timesheet","Work Date"},
    new [] {"Fact_Capacity","Capacity Hours"},
    new [] {"Fact_Capacity","Date"},
    new [] {"Fact_Capacity","Date Key"},
    new [] {"Fact_Capacity","Employee Key"},
};
foreach (var h in hide) {
    if (Model.Tables.Contains(h[0]) && Model.Tables[h[0]].Columns.Contains(h[1]))
        Model.Tables[h[0]].Columns[h[1]].IsHidden = true;
}
