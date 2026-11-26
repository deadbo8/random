Pacific Gas and Electric Company

[INSERT PG&E LOGO]    [INSERT TCS LOGO]

Rate Calc Testing Tool – Tool 2 (CC&B)

Enterprise User Guide & Standard Operating Procedure

Document Information
Document Attribute	Details
Project Name	Rate Calc Testing Tool – Tool 2 (CC&B)
Value Stream	[Enter Value Stream]
Line of Business / Department	MTC – Customer Care
Prepared By	[Enter Author Name]
Reviewed / Approved By	[Enter Reviewer / Approver Names & Roles]
Document Date	[DD/MM/YYYY]
Tool Version	[Enter Version, e.g., v3.11.x]
Document Status	[Draft / Final for Approval]
Classification	Internal
Document Control & Revision History
Version	Date	Author	Summary of Changes
1.0	[DD/MM/YYYY]	[Author Name]	Initial corporate release covering Individual Rate Check, Bulk Rate Check, and IBC (CC&B 25.4 Validation).
Formal Approval Sign-Off
Role	Name	Signature / Approval Status	Date
QA Lead / Delivery Lead	[Enter Name]		
Business Owner / Product Lead	[Enter Name]		
Technical Architect / SME	[Enter Name]		
Table of Contents
Executive Overview & Purpose
Prerequisites, Package Setup & Launch
Oracle Login & Account Safety
Application Home Page & Navigation
Workflow 1: Individual Rate Check
Workflow 2: Bulk Rate Check
Workflow 3: IBC – Individual/Inter-Bill Comparison (CC&B 25.4)
Managing Configuration & Exception Rules (storage/)
Output Reports & How to Interpret Results
Troubleshooting & Error Resolution
Appendix: Glossary of CC&B & Tool Terms
1. Executive Overview & Purpose
1.1 Purpose

The Rate Calc Testing Tool – Tool 2 (CC&B) is a desktop web application used to verify PG&E's Oracle Customer Care & Billing (CC&B) bills, tariff rate releases, and the Oracle CC&B 25.4 Upgrade.

Manually verifying multi-tier utility bills, mid-cycle rate changes, Time-of-Use (TOU) intervals, Service Quantity Identifiers (SQIs), and General Ledger (GL) financial distributions is time-consuming and error-prone. Tool 2 automates rate verification and bill-to-bill comparison, producing audit-ready Excel workbooks and Executive Word reports for test sign-off.

1.2 Choosing the Right Workflow

Tool 2 provides three testing modes from one interface:

Operational Mode	What It Does	When to Use It
1. Individual Rate Check	Calculates expected charges for a single Bill Segment and compares them against CC&B billed calculation lines (CI_BSEG_CALC_LN).	Deep-dive defect investigation, single-account tariff checks, or what-if overrides.
2. Bulk Rate Check	Runs Individual Rate Check across dozens or hundreds of accounts from one uploaded Excel workbook.	Regression test cycles (Cycle 1, Cycle 2, Cycle 3) and major tariff releases.
3. IBC – Bill Comparison	Compares a Prod-generated (Cancelled) bill directly against a Non-Prod/LN-regenerated (New) bill across Headers, Reads, SQIs, Calc Lines, Financials, and Messages.	CC&B 25.4 Upgrade validation and cancel/rebill regression testing.

Supported Utility & Tariff Domains:

Electric Residential: Tiered, TOU, EV, and Baseline schedules (e.g., E1, EV2A, HETOUC, HE1, EM, EML, HEML, HES, ETL).
Commercial & Industrial: Small, Medium, and Large Commercial schedules (e.g., A1, A1F, HB1, HB1ST, HB6, HB19S, HB20S, HB20P).
Agricultural: Agricultural TOU and non-TOU schedules (e.g., HAGC, HAGA1, HAGB).
Gas & Net Energy Metering (NEM): Gas procurement/transportation/PPP surcharges (HG1) and NEM True-Up billing schedules.
2. Prerequisites, Package Setup & Launch
2.1 System & Network Prerequisites
Operating System: Windows 10 / Windows 11 (Corporate VDI or PG&E-managed laptop).
Browser: Microsoft Edge or Google Chrome.
Network Access: Active PG&E corporate network / VPN connection with read-only Oracle database credentials for the target CC&B test environments (OE, ON, or LN).
No Python Installation Required: Run the standalone executable (Rate-Calc-Tool.exe) directly.
2.2 Required Distribution Folder Structure

Before launching the tool, confirm that Rate-Calc-Tool.exe and the storage/ folder are in the same parent folder. The storage/ folder contains editable Excel configuration files so QA teams can update exception rules and rate lookups directly.

text
Rate_Calc_Tool_Distribution/
├── Rate-Calc-Tool.exe            # Application executable
├── storage/                      # REQUIRED configuration folder
│   ├── compare_exceptions.xlsx   # Approved CC&B 25.4 extra-line exception whitelist
│   └── db_storage.xlsx           # Rate schedules, UOM/TOU lookups, and season_config
├── uploads/                      # Auto-created folder for uploaded input files
└── output/                       # Auto-created folder for generated reports

Important: Do not move Rate-Calc-Tool.exe outside the folder containing storage/. If storage/db_storage.xlsx is missing, Rate Schedule autocomplete, UOM/TOU dropdowns, and IBC Rate Type/Meter Type filters will appear empty.

2.3 Launching the Application

Step 1: Start the Application

Double-click Rate-Calc-Tool.exe. A terminal window opens and starts the tool. Wait until the terminal displays:

text
Running on http://127.0.0.1:5000

[INSERT SCREENSHOT: Terminal window showing Rate-Calc-Tool.exe startup and Running on http://127.0.0.1:5000]

Step 2: Open the Web Interface

Hold Ctrl and click http://127.0.0.1:5000
 in the terminal window, or copy and paste the URL into Microsoft Edge or Google Chrome.

Note: Keep the terminal window open while using the tool. Closing the terminal stops the application.

3. Oracle Login & Account Safety

When connecting to corporate Oracle CC&B test environments (OE, ON, or LN), Tool 2 protects your Oracle account and test data with the following safeguards:

Upfront Credential Check: Before starting a Bulk Rate Check or IBC batch, the tool verifies your Oracle credentials once. If your username/password is wrong (ORA-01017) or your account is locked (ORA-28000), the run stops immediately with an alert so repeated batch attempts do not lock your Oracle account.
Immediate Stop on Login Failure: If an authentication error occurs mid-run, remaining items stop immediately without further login retries.
No Saved Passwords: Your Oracle username and password are kept only for your active session and are never written to disk, logs, or Excel/Word reports.
Read-Only Access: The tool only reads data from Oracle CC&B and never creates, modifies, or deletes billing records.
4. Application Home Page & Navigation

The top header bar spans the full width of the screen and contains:

PG&E Branding & CCB Indicator: Confirms you are working in the CC&B rate validation tool.
Software Version Badge: Displays the active release version (e.g., v3.11.x) for audit tracking in reports.
Mode Switcher Bar: Three buttons to switch between workflows at any time:
Individual Rate Check
Bulk Rate Check
IBC - Bill Comparison

[INSERT SCREENSHOT: Full-width Application Home Page showing PG&E Logo, Title, Version Badge, CCB Label, and the 3 Mode Buttons]

5. Workflow 1: Individual Rate Check
5.1 Purpose & Billing Rules Applied

Use Individual Rate Check to calculate expected tariff charges for a single Bill Segment (BSEG_ID) and compare them line-by-line against CC&B billed calculation lines (CI_BSEG_CALC_LN).

To align with Oracle CC&B billing behavior, the tool automatically applies these rules:

Month-End Rollover Rule: When a bill segment's Start Date falls on the last calendar day of a month (e.g., March 31 or April 30), tariff factors are evaluated as of the 1st of the following month.
Multi-Header Proration (H1 / H2 / H3): Bills spanning multiple months, seasonal crossovers, or mid-cycle rate changes are automatically split into up to three prorated calculation headers.
BILLDAYS Source-of-Truth: For Normal Bills, billed days are read directly from SQI_CD = 'BILLDAYS' in your SQI extract rather than simple date subtraction, preserving meter-route billing day counts.
5.2 Input Field Reference
Field Name	Format / Rule	Required	Description
Service Agreement	10-digit numeric ID (not all zeros)	Yes	CC&B Service Agreement (SA_ID). A green checkmark confirms valid format.
Bill Segment	12-digit numeric ID (not all zeros)	Yes	CC&B Bill Segment (BSEG_ID). A green checkmark confirms valid format.
Start Date / End Date	MM/DD/YYYY	Yes	Billing period start and end dates (Start Date must not be after End Date).
True Up Bill / True Up Date	Checkbox + Date	Optional	Check for NEM True-Up bills to display and enter the True Up Date.
Rate Schedule	Autocomplete Search	Yes	Type to search valid CC&B Rate Schedules loaded from storage/db_storage.xlsx.
Types of Bill	Radio Button	Yes	Choose Normal Bill or Rate Version Crossover / Seasonal Crossover (displays a second Rate Version extract field).
CCB Data Extract (BCS / BCS2)	.xlsx Upload or Fetch from DB	Yes	Primary (and secondary, for rate changes) Bill Factor extracts (CI_BF_VAL / CI_TOU_BF_VAL).
SQI Data	.xlsx Upload or Fetch from DB	Yes	Service Quantity extract (CI_BSEG_SQ). Fetching from DB also retrieves CI_BSEG_READ automatically.
Characteristic Override	Dynamic Rows (+ / -)	Optional	Override Service Agreement or Service Point characteristics for what-if testing.
Non-Interval Usage & Demand	Dynamic Rows (+ / -)	Optional	Manually enter UOM, TOU, and Measured Quantity values when testing usage overrides.
Developer Mode	Checkbox	Optional	Includes extra diagnostic details in the comparison output.

[INSERT SCREENSHOT: Individual Rate Check Form with Fetch from DB Modal Open]

5.3 Step-by-Step Procedure
Select the Individual Rate Check tab.
Enter the 10-digit Service Agreement, 12-digit Bill Segment, Start Date, End Date, Rate Schedule, and Types of Bill.
Provide the required extracts using either option:
Fetch from DB (Recommended): Click Fetch from DB next to CCB Data Extract (and Rate Version 2 if crossover) and SQI Data. Enter your Oracle Username, Password, and Environment (OE, ON, or LN), then click Fetch Extract. Your credentials remain active for the rest of the session.
Manual Upload: Click Choose File to select local .xlsx extract files.
Click Upload Files and wait for the green confirmation message.
Click Rate Check to calculate expected charges and display them grouped by segment (SEG1, SEG2).
(Optional) Click Download System Excel to export RCT_OUTPUT_{RateSchedule}_{BsegID}.xlsx.
Under Compare Bill, click Compare to compare the calculated lines against CC&B (CI_BSEG_CALC_LN), automatically applying approved exceptions from storage/compare_exceptions.xlsx.
Review the summary pop-up showing Matches and Mismatches, then click Download Comparison Excel for the full line-by-line diff report.
Click Reset at the bottom right to clear all inputs before testing another account.

[INSERT SCREENSHOT: Individual Rate Check Output Table & Compare Bill Action Buttons]

5.4 What to Infer from the Results
Calculated Charges Table (SEG1, SEG2): Verify that proration headers (SEG1, SEG2), billed days, usage quantities, and rate factors match what you expect for the billing dates and rate schedule.
Compare Pop-Up & Comparison Excel:
All Matches (0 Mismatches): Your rate calculation matches CC&B billed lines (or any extra $0.00 lines match an approved CC&B 25.4 exception). The bill passes rate validation.
Mismatches Found: Open the downloaded Comparison_{rate}_{bseg}.xlsx to see which specific rate component or amount differed between the tool's calculation and CC&B, and use Characteristic Override or Non-Interval Usage & Demand if testing root-cause hypotheses.
6. Workflow 2: Bulk Rate Check
6.1 Purpose & Overview

Use Bulk Rate Check to run rate calculations and CC&B comparisons across a batch of accounts from a single Excel file. For each row, the tool automatically retrieves tariff factors, TOU factors, SQIs, and meter reads from Oracle, reconciles the calculated charges against CI_BSEG_CALC_LN, and creates both per-account comparison workbooks and an overall summary report.

6.2 Batch Input Excel Specification

Create an .xlsx or .xls workbook with the following columns:

Column Header	Format / Accepted Values	Required	Example Value	Notes
S. No.	Integer	Optional	1	Sequence number for your own tracking.
service_agreement (or sa_id)	10-digit numeric ID	Yes	0008355018	Leading zeros are restored automatically if Excel strips them.
bill_segment (or bseg_id)	12-digit numeric ID	Yes	000472973307	Leading zeros are restored automatically if Excel strips them.
rate_schedule (or rate_type)	CC&B Rate Code	Yes	HB1ST, E1, EV2A, HAGC, HETOUC	Case-insensitive; extra spaces are trimmed automatically.
start_date	Date (M/D/YYYY or YYYY-MM-DD)	Yes	5/20/2026	Month-end rollover is applied automatically when applicable.
end_date	Date (M/D/YYYY or YYYY-MM-DD)	Yes	6/20/2026	Billing period end date.
bill_type	normal_bill, rate_value_change, seasonal_crossover	Yes	rate_value_change	Use rate_value_change or seasonal_crossover when two rate versions or seasons apply.
true_up_flag	0 or 1	Optional	0	Enter 1 for NEM True-Up bills; 0 otherwise.
true_up_date	Date or N/A	Optional	N/A	Required only when true_up_flag is 1.

[INSERT SCREENSHOT: Sample Bulk Input Excel (BILL_TEST_LN.xlsx) showing Column Headers and Sample Rows]

6.3 Step-by-Step Procedure
Click the Bulk Rate Check tab.
Select the target Environment (OE, ON, or LN) and enter your Oracle Username and Password.
Click Choose File inside the upload box and select your batch Excel file.
Click Run Bulk Rate Check.
The tool checks your Oracle login before starting the batch. If login fails, it stops immediately and displays an error alert.
Monitor progress using the progress bar, completion percentage, Processed / Matches / Mismatches counters, and the live Per Bill Line Breakdown table.
When processing reaches 100%:
Click individual links in the Report column to download and inspect specific accounts, or
Click Download All Comparison Reports (.zip) to download the summary workbook and all individual account reports together.

[INSERT SCREENSHOT: Bulk Rate Check Progress Bar, Stats Row, and Per Bill Line Breakdown Table]

6.4 What to Infer from the Results
Processed / Matches / Mismatches Counters: Show overall cycle pass/fail progress at a glance. Any non-zero Mismatches count indicates one or more bill segments need review.
Per Bill Line Breakdown Table: Identifies exactly which SA_ID / BSEG_ID rows passed or failed so you can immediately open the Report link for failing rows without sorting through passing accounts.
Bulk_Summary_{ver}_{id}.xlsx: Provides the consolidated pass/fail table across all uploaded accounts for test-cycle sign-off.
7. Workflow 3: IBC – Individual/Inter-Bill Comparison (CC&B 25.4)
7.1 Purpose & Upgrade Validation Flow

Use IBC (Individual/Inter-Bill Comparison) to validate CC&B bills during the Oracle CC&B 25.4 Upgrade. Instead of recalculating rates, IBC pulls two real bills from Oracle CC&B and compares them side-by-side across 7 billing areas.

Standard Upgrade Validation Process Flow:

Account moved from Prod to LN environment → Cancel and Rebill each Bill Segment (BSEG) → Run IBC with Cancelled BSEG/Bill ID (Prod-generated) and New BSEG/Bill ID (LN-generated).

How IBC Evaluates Rows (What to Expect):

Ignores Row Order / Sequence Shifts: Identical rows match even if CC&B 25.4 outputs them in a different order or assigns a different sequence number (SEQNO / GL_SEQ_NBR).
Pairs Modified Rows Side-by-Side: When a row differs between Old and New, IBC aligns both versions on the same row and highlights only the changed cell(s) in yellow.
Flags Missing / Extra Rows & 25.4 Exceptions: Dropped or unexpected extra rows are flagged as mismatch, while approved $0.00 extra lines in Calc Lines (defined in storage/compare_exceptions.xlsx) are classified as match (25.4 exception).
7.2 What IBC Validates (The 7 Criteria)

By default, criteria 1 through 5 are active when results load; criteria 6 and 7 can be toggled on or off at any time using the dashboard ribbons.

Seq	Functional Name (Shown in UI & Reports)	Technical CC&B Table (Shown on Hover)	Level	Default Active	What It Validates & What Mismatches Indicate
1	Bseg Headers	CI_BSEG_CALC	BSEG	Yes	Rate schedules (RS_CD), billing dates, and total segment calculated amounts.
2	Reads	CI_BSEG_READ	BSEG	Yes	Register start/end reads, TOU buckets, and measured usage quantities (MSR_QTY).
3	SQIs	CI_BSEG_SQ	BSEG	Yes	Service Quantity Identifiers (ALLKWH, BILLDAYS, Baseline, Indicators) checked to 6 decimal places.
4	Calc Lines	CI_BSEG_CALC_LN	BSEG	Yes	Bill line descriptions (DESCR_ON_BILL), CALC_AMT, AUDIT_CALC_AMT, and approved CC&B 25.4 extra-line exceptions.
5	Financials	CI_FT_GL (via CI_FT)	BSEG	Yes	General Ledger distributions (DST_ID, AMOUNT, GL_ACCT, CHAR_TYPE_CD, CHAR_VAL) for the billed segment (BS), excluding cancellation reversals (BX) so cancelled segments are not double-counted.
6	Bill Msgs	CI_BILL_MSGS	Bill	No	Bill-level customer messages and insert codes.
7	Bseg Msgs	CI_BSEG_MSG	BSEG	No	Segment-level billing messages.
7.3 Input Options, Validation Scopes & Input Safeguards

Validation Scope Dropdown

Bill Level (Bill + BSEG Validation): Compares both Bill-level and BSEG-level criteria, displaying both Bill-Level and BSEG-Level summaries and reports.
BSEG Level Only: Skips Bill-level checks (such as Bill Msgs and Bill-level tabs) and validates BSEG pairs directly.

Input Mode Dropdown

Batch Excel Upload (.xlsx) (Default): Click ⬇ User_Input Template in the top right of the IBC card to download the 2-column template:

New Bill id	cancelled Bill ID
239586762273	239616405323
202285011318	202452547935

Auto-Discovery: When you upload Bill IDs, the tool automatically discovers all child BSEGs, pairs each New_bseg ↔ old_bseg (by CAN_BSEG_ID or SA_ID), and fetches Rate_cd. (In BSEG Level Only mode, you may also paste New_bseg and old_bseg IDs directly).

Single Pair Entry (Manual IDs): Enter one 12-digit New Bill / BSEG ID and one 12-digit Cancelled Bill / BSEG ID directly on screen.

Built-In Input Safeguards

Identical IDs (New == Cancelled): Blocked automatically so you cannot accidentally compare a bill or segment against itself.
Invalid / Missing IDs: Non-numeric IDs, all-zero IDs, or IDs not found in Oracle CI_BSEG are never falsely marked as matched or mismatched:
If every row is invalid, validation stops and lists the row-by-row errors.
In mixed batches, invalid rows are skipped while valid rows run normally. An amber Warning Banner displays skipped rows in the UI, and a Skipped_Invalid_Input sheet is added to the Master Excel report.

[INSERT SCREENSHOT: IBC Input Configuration Card & Live 4-Step Progress Bar]

7.4 Step-by-Step Procedure
Click the IBC - Bill Comparison tab.
Select Environment (OE, ON, or LN), enter your Oracle Username and Password, choose Validation Scope, and select Input Mode.
Upload your User_Input Excel file (or enter a Single Pair) and click Run IBC Validation.
Monitor the 4-step progress tracker (1. Discover BSEGs → 2. Fetch SQL Tables → 3. Compare & Rules → 4. Build Excels) until the results dashboard loads.
Use the interactive dashboard controls (described in Section 7.5) to inspect results, filter criteria, and download your Excel/Word/ZIP deliverables.
7.5 How to Interpret the IBC Dashboard (What to Infer)
Table Filter Ribbons: Click any ribbon (Bseg Headers, Reads, SQIs, Calc Lines, Financials, Bill Msgs, Bseg Msgs) to include or exclude that criterion, and hover over a ribbon to view its CC&B table name.
What to infer: Turning criteria on or off updates all KPIs, charts, summary tables, the Table Inspector, and your downloaded reports immediately without re-running the comparison—use this to isolate whether a failure is in billing math (Calc Lines, SQIs), meter data (Reads), or accounting (Financials).
KPI Scorecard: Displays Bills / BSEGs, Total Rows, Matched Rows, Mismatched Rows, and 25.4 Exceptions.
What to infer: 25.4 Exceptions tracks approved $0.00 extra lines and does not count against pass rate. Focus investigation on Mismatched Rows.
Top Chart Row (3 Pie Charts — Strictly Match vs. Mismatch):
Bill/BSeg Pass % (Left): Toggle between BSEG and Bill to view whole-segment or whole-bill pass rates based on currently selected ribbons (hover shows the active criteria).
SQIs (Middle): Shows row-level Match vs. Mismatch for CI_BSEG_SQ.
Calc Lines (Right): Shows row-level Match vs. Mismatch for CI_BSEG_CALC_LN.
What to infer: Comparing Bill/BSeg Pass % against SQIs and Calc Lines tells you whether failures affect whole bills due to a single row mismatch or widespread line-level discrepancies.
Bottom Chart Row (Criteria Bar Chart): Displays Match vs. Mismatch row counts for each criterion, with All / Bill Only / BSEG Only toggles and hover tooltips.
What to infer: Quickly shows which specific CC&B area has the highest concentration of mismatches.
Collapsible Rate Schedule Breakdown: Click ▶ Expand to view a horizontal stacked bar chart of Matched vs. Mismatched BSEGs by Rate Schedule, with Type (Residential, Commercial, Agriculture), Meter Type (A, B, C, D), and Rate filters.
What to infer: Helps determine whether defects are isolated to a single tariff schedule, customer segment, or meter type rather than systemic across all rates.
Summary Tables & Live Table Inspector: Switch between Bill Level Summary, BSEG Level Summary, and Discovered IBC_Input. Click BSEGs (N) on a bill row to isolate its child segments, then click Inspect on a BSEG row to open the side-by-side Table Inspector (filterable by All, Mismatches Only, or Matches Only).
What to infer: Column 1 (Comparison) shows the row verdict, and yellow-highlighted cells pinpoint the exact column(s) whose value changed between the New and Cancelled bills.

[INSERT SCREENSHOT: IBC Full Results Dashboard showing Ribbons, 5 KPIs, 3 Pie Charts, and Bar Chart]

[INSERT SCREENSHOT: Expanded Rate Schedule Horizontal Stacked Bar Chart with Type & Meter Type Filters]

[INSERT SCREENSHOT: Live Side-by-Side Table Inspector showing Comparison in Column 1 and Yellow Mismatch Cell Highlights]

8. Managing Configuration & Exception Rules (storage/)

Both configuration workbooks live in the storage/ folder next to Rate-Calc-Tool.exe and can be updated whenever business rules or rate lists change.

8.1 Managing Approved CC&B 25.4 Exceptions (storage/compare_exceptions.xlsx)
Purpose: Whitelists business-approved $0.00 extra calculation lines introduced in CC&B 25.4 so they are reported as match (25.4 exception) rather than defects.
Updating Exceptions from the IBC Header:
Click ⬇ Exceptions (.xlsx) in the top right of the IBC card to download the active compare_exceptions.xlsx.
Add or edit approved line descriptions in the CCB_EXCEPTIONS sheet (supports exact text or %a wildcard matching, e.g., Footing: Solar Choice Program Charge%a).
Click ⬆ Update Exceptions and upload your edited workbook; the new rules take effect on your next run.
8.2 Maintaining Rate Categories & Lookups (storage/db_storage.xlsx)
season_config Sheet: Maps each Rate Schedule (Col B) to its Type (Col A, e.g., Residential, Commercial, Agriculture) and Meter Type (Col F, e.g., A, B, C, D), populating the IBC Rate Schedule Breakdown filters.
export, Export Worksheet_UOM, Export Worksheet_TOU, Rate_type Sheets: Populate the Individual Rate Check Rate Schedule autocomplete search, UOM/TOU dropdowns, and TOU rate detection.
9. Output Reports & How to Interpret Results
9.1 Generated Files Summary
Workflow	Output File	What It Contains & When to Use It
Individual	RCT_OUTPUT_{rate}_{bseg}.xlsx	Calculated rate line items generated by the tool.
Individual	Comparison_{rate}_{bseg}.xlsx	Side-by-side comparison of tool calculations vs. CC&B lines for defect analysis.
Bulk	Bulk_Summary_{ver}_{id}.xlsx	Aggregate pass/fail summary across all accounts in the batch.
Bulk	Bulk_Comparison_Results_{ver}_{id}.zip	Complete archive containing Bulk_Summary plus every individual account comparison workbook.
IBC	IBC_Master_Report_{id}.xlsx	Multi-tab summary workbook (User_Input, IBC_Input, IBC_Output_Bseg_level, IBC_Output_Bill_level with merged criteria headers, and Skipped_Invalid_Input when rows are skipped).
IBC	{New_bseg}vs{old_bseg}.xlsx & bill_{New}vs{Old}.xlsx	Detailed side-by-side workbooks with Column A = Comparison, tabs per functional criterion (Bseg Headers, Reads, SQIs, Calc Lines, Financials, etc.), and yellow mismatch cell highlights.
IBC	IBC_Executive_Report_{id}.docx	1-page stakeholder Word report in Aptos font with PG&E & TCS logos, tool version, active criteria, process flow, KPI scorecard, 3 aligned Pie Charts (Bill/BSeg Pass %, SQIs, Calc Lines), and the Criteria Bar Chart. Reflects your active Table Ribbons and BSEG/Bill toggle.
IBC	IBC_All_Reports_{id}.zip	Complete bundle containing the Master Excel, Executive Word Report, and all individual Bill/BSEG workbooks reflecting your currently selected Table Filter Ribbons.
9.2 Audit Color-Coding Legend & What Each Status Means
Color	Status Label	What to Infer
Green	match / MATCH	All compared values match identically between Baseline/Old and Target/New. No action required.
Light Blue	match (25.4 exception) / CCB EXCEPTION	An extra $0.00 line appears in CC&B 25.4 that matches an approved rule in compare_exceptions.xlsx. Treated as passing.
Red	mismatch / MISMATCH	A row value changed, a row is missing in Target, or an unapproved extra line appeared. Requires investigation.
Yellow Cell	(Cell Highlight)	Identifies the exact column(s) within a paired mismatch row whose value changed (e.g., INIT_SQ, AUDIT_CALC_AMT, AMOUNT).

[INSERT SCREENSHOT: Sample IBC Executive Word Report (.docx) and Side-by-Side Excel Sheet with Column A Comparison & Yellow Highlights]

10. Troubleshooting & Error Resolution
Symptom / Error Message	Cause	What to Do
Invalid Oracle Username or Password. Run Aborted. (ORA-01017)	Incorrect Oracle credentials entered for the selected environment (OE, ON, LN).	Re-enter the correct credentials and confirm the target environment. The tool stopped before running the batch so your account remains safe.
Oracle Account is LOCKED. Please contact your DBA. (ORA-28000)	Oracle account was already locked due to prior failed login attempts.	Contact your Oracle DBA / Helpdesk to unlock your account before retrying.
Rate Schedule autocomplete or Type / Meter Type filters are empty	storage/db_storage.xlsx is missing from the folder next to Rate-Calc-Tool.exe, or is currently open in Excel.	Close db_storage.xlsx in Excel and verify storage\db_storage.xlsx sits in the same folder as Rate-Calc-Tool.exe.
Validation aborted — Invalid, duplicate, or missing Bill/BSEG IDs	Same ID entered for both New and Cancelled, invalid/all-zero ID, or ID does not exist in the selected Oracle environment (OE/ON/LN).	Check the Bill/BSEG IDs in CC&B and confirm you selected the right environment (LN vs. OE/ON).
Buttons or UI changes not working after upgrading .exe	Browser is using cached UI files from an earlier version.	Press Ctrl + F5 (hard refresh) in Microsoft Edge or Google Chrome.
Browser cannot reach http://127.0.0.1:5000
	The Rate-Calc-Tool.exe terminal window was closed or has not finished starting.	Double-click Rate-Calc-Tool.exe and wait until Running on http://127.0.0.1:5000 appears in the terminal.
11. Appendix: Glossary of CC&B & Tool Terms
Term / Acronym	Definition
BCS / BCS2	Bill Factor Extract (CI_BF_VAL) for Rate Version 1 and Rate Version 2 (used in mid-cycle rate-change validations).
BSEG (BSEG_ID)	Bill Segment — a 12-digit identifier representing one service billing period on a CC&B Bill (e.g., Electric or Gas segment).
CC&B 25.4	Oracle Utilities Customer Care & Billing Version 25.4 target upgrade release.
CI_FT_GL / FT_TYPE_FLG	Financial Transaction General Ledger table; 'BS' indicates original Bill Segment accounting, while 'BX' indicates Cancellation reversal accounting.
DST_ID	Distribution Code governing General Ledger revenue and tax accounting allocations.
IBC	Individual / Inter-Bill Comparison — automated bill-to-bill and segment-to-segment comparison workflow.
NEM	Net Energy Metering (solar/renewable generation rate schedules with annual True-Up billing).
RC_SEQ	Rate Component Sequence — unique sequence number identifying a tariff calculation step within a Rate Schedule.
SA (SA_ID)	Service Agreement — a 10-digit identifier linking a customer account to a utility service and rate schedule.
SQI (SQI_CD)	Service Quantity Identifier — derived billing quantities such as ALLKWH, BILLDAYS, BLQ-QTY, and indicator flags.
TOU / UOM	Time-of-Use period (e.g., WPK, WPP, WOP, SPK) and Unit of Measure (e.g., KWH, KW, THM).
