# Handover: CTS payroll setup and award compliance

Last updated 8 November 2026. This covers the chat "Payroll data and award compliance". It is written so another agent can pick the work up without the chat history.

## Read this first

Danica Nelson is the Finance Operations Manager at Corporate Technology Services (CTS). She runs payroll in Employment Hero (EH) Payroll. Graham Cherry is the Executive General Manager and approves pay runs. Duncan Lugstein approves the bank file in CommBiz.

Follow the repo's CLAUDE.md for every reply. The main rules are short sentences, the answer first, no em dashes, no nested bullets, and no working shown unless she asks. She reads with text to speech.

Australian context. Every figure here is for payroll and statutory obligations. Say plainly when something comes from a web search rather than a source document, and never fill a gap with a guess.

## Where everything is

The register is `awards/CTS_Payroll_Setup_Register_21102026.xlsx` on branch `claude/gallant-planck-uztwlf`. It's the working file, and it includes Danica's own edits.

The award columns are filled by `awards/payroll_fill_award_columns.py`. Don't re-run it on the current file, because it would overwrite later edits. It's kept as a record only.

The Casual Rate Check tab is built by `awards/payroll_casual_rate_check.py`. It takes the register path and a Pay Run Audit xlsx. Re-running it rebuilds that tab only.

The Payroll Tracker template is on branch `claude/amazing-mayer-mwuadq` at `payroll/Payroll_Tracker_Template.xlsx`. It belongs to a different chat.

The source uploads from Danica live in the session uploads folder. They aren't in the repo. They are the EH Employee Details Report of 21/10/2026, the Pay Run Audit for the fortnight ending 16/10/2026, and Emmett Todd's payslip for 5/09/2026 to 18/09/2026.

## Register layout

These tabs are in order.

1. Summary. It holds the six numbered findings and how to use the Register.
2. Award References. It has Fair Work links. A Live Performance Award row was added at row 16.
3. Register. It has one row per employee, rows 2 to 80, with 79 staff.
4. Casual Rate Check. This one is new. It covers every casual, with their EH rate, the rate paid in pay run 8, the loading check, the award base rate and the result.
5. EH Export (Raw Data). Column N, "Pay condition rule set", was added by Danica.

On the Register tab, columns A to M are EH data, so don't edit them. Columns N to S are flag formulas. Column T is the priority and column U holds the notes.

The yellow columns V to Z are for Danica. V is the award, W is the classification, X is contracted hours, Y is contract extras and Z is action or status. Column AA is the "Under an award?" tick column that I added.

## What has been established

### The award

The EH agreements are built on the Live Performance Award 2020 [MA000081], Production and Support Staff. Danica's EH screenshot of Payroll Settings, Employment Agreements confirmed this on 21/10/2026.

The "Production & Support Staff - Level X" names are that award's own classifications, from clause 11.1. The award has Levels 1 to 8 and a Technical Manager set, each with a Full Time, a Part Time and a Casual agreement. "Your Agreements" is empty.

The payroll recordings back this up. Casuals follow the Live Performance Award, and clause 63 covers penalties. Full timers are paid overtime on the CTS internal Overtime Matrix, which is only compliant if it pays at least what the award would.

### Who is on what

There are 73 ticks in column AA and 6 people who are award free. Of the ticks, 67 are Live Performance Award at their current EH level, and 5 are the same award marked "check". The last 2 are under the Clerks Award.

The five "check" people are Daniel Sobkowski, Mitchell Healey, Tyler Wood, Simon Gruenefeld and Patrick Lee. Their roles might sit outside production work, or be award free management. Their contracts decide it.

The two Clerks Award people are Emelda Waiya and Isabella Nikodinovski, under the Clerks Private Sector Award [MA000002]. Their Clerks level still needs setting from their duties.

The six award free people are Danica Nelson, Duncan Lugstein, Gianmaria Yao, Graham Cherry, Jordan Sexty and Premila Lugstein. Danica confirmed that Premila does professional finance work, including BAS and tax.

I warned Danica that BAS work on its own doesn't take someone out of the Clerks Award. Premila's rate of 31.77 is below the Clerks Level 4 and Level 5 minimums. Danica kept her as award free. Don't reopen this unless she asks.

### Rule sets in EH (column N, EH Export tab)

Every Production & Support agreement holder is on the "Production & Support Staff" rule set.

Danica, Duncan and Gianmaria are on "Company made Rules".

Jordan Sexty is wrongly on "Production & Support Staff". He should be on "Company made Rules".

Graham Cherry's rule set is blank. Danica confirmed he should be "same as me", so "Company made Rules".

### Casual loading

EH adds the 25% loading on top of the rate in Register column G. Column G is the base rate before loading.

This was checked against the Pay Run Audit for the fortnight ending 16/10/2026. All 19 casuals with Casual Ordinary Hours were paid exactly column G times 1.25. Gabriela Panza Romanini Alchaar had Sunday hours only, paid at 2.25 times base, which is also correct.

The other 29 casuals weren't in that pay run, so their pay wasn't tested. They use the same pay category.

An earlier message in the chat said 35 casuals were underpaid. That was wrong, because it assumed column G already included the loading. It has been corrected in the file and in the chat.

### Award base rates used

These are from the Live Performance Award clause 11.1, for the first full pay period on or after 1 July 2026. Source: awards.fairwork.gov.au/MA000081.html.

- Level 3: 28.89 an hour.
- Level 4: 29.45 an hour.
- Level 5: 30.37 an hour.
- Level 6: 31.30 an hour.
- Level 7: 33.311 an hour.

The page didn't show a separate hourly figure for Level 8. Schedule A, the duties for each level, couldn't be loaded, so nobody's level has been checked against their duties.

The Clerks Award rates from clause 16.1, Table 3, are 31.11 for Level 3, 32.67 for Level 4 and 33.99 for Level 5.

## Open items, in priority order

### 1. Four casuals are paid below the award base rate

1. Emmett Todd, Level 4, Register row 29. His EH rate is 27.17 and should be at least 29.45. His payslip for 5/09 to 18/09/2026 confirms 33.96 an hour including 6.79 loading. He's owed 199.50 in wages plus 23.94 in super for the 70 hours up to 18/09/2026. The pay run ending 2/10/2026 hasn't been checked. He wasn't in the pay run ending 16/10/2026.
2. Jeffrey Kempin, Level 4, Register row 42. His EH rate is 28.12 and should be at least 29.45. He wasn't in pay run 8, and no payslip has been checked.
3. Morgan Palmer, Level 5, Register row 59. Her EH rate is 29.00 and should be at least 30.37. She wasn't in pay run 8, and no payslip has been checked.
4. Zed Norrie, Level 5, Register row 80. His EH rate is 29.00 and should be at least 30.37. He was paid 36.25 in pay run 8, which confirms it.

Back pay for all four runs from the first pay period on or after 1 July 2026. Before that date, the previous year's award rates apply, and those haven't been checked.

To work out each person's back pay, get their hours since 1 July from the payslips or the Pay Run Audits. Then multiply the hours by the shortfall, times 1.25 for loading, and add Super Guarantee at 12%. Give Danica the totals only.

Status: not yet done in EH, as far as this chat knows.

### 2. Alana Carling's missed leave, Register row 4

She's part time, but she was on a Casual NSW leave template. Graham fixed the template, but only for shifts on or after 16/10/2026, as his Teams message of 28/10/2026 says. Annual leave and personal leave must be added by hand for every part-time hour before that date. Graham only mentioned annual leave, but personal leave was missed too. Long service leave was accruing.

In pay run 8, she worked 46.75 Permanent Ordinary Hours between 7/10 and 15/10/2026. She's owed 3.60 hours of annual leave and 1.80 hours of personal leave for that run. The accrual rates match EH's for other permanents: annual leave 1/13 of ordinary hours, personal leave 1/26. Blake Crisford's accruals in the same audit prove this.

Still needed from Danica:
1. The date Alana changed from casual to part time.
2. The Pay Run Audits from that date up to the run ending 2/10/2026.

How to enter it, per the payroll recording in manuals/part5-narration-transcript.txt at about 11:40 to 13:30: in a pay run, open the employee, then Actions, then Accrue leave. Pick the leave type, enter the hours and add a note.

### 3. EH setup changes Danica needs to make

None of these have been confirmed as done.

1. Jordan Sexty: change his rule set to Company made Rules.
2. Graham Cherry: set his rule set to Company made Rules.
3. Premila Lugstein: take her off the Production & Support Level 5 agreement, so she's award free.
4. Emelda Waiya and Isabella Nikodinovski: move them to the Clerks Award at the right level.
5. The "Production & Support Staff - Level 6 - Part Time" agreement appears twice in EH. Nobody is currently on Level 6 Part Time. Compare the two copies, then ask EH support to remove one.

### 4. Items already on the Register from the first review

These are in columns T and U and haven't been worked through yet.

Damien Suchy is set up as Casual but is on a permanent agreement, rate template and pay category. Simon Gruenefeld is casual but on a Permanent NSW leave template. Some casuals have no leave template. Dale Cohen's leave template has no state. John Abbas Rizvi is in Melbourne but on an NSW template. Gianmaria Yao still shows as employed. Duncan Lugstein is on 26.08 an hour, which is low for an MD. Some job titles are missing.

### 5. Milo Rankin's per diem, not verified

On 29/10/2026, Danica asked for this to go in the notes for the pay run ending 30/10/2026: Milo Rankin, 1 night, 125.00, cost centre 26061607, APA Quarterly All Hands 4.11.26.

The reminder routine "Pay run FN ending 30/10/2026 - roster notes Wk3" could only be edited from the Payroll Tracker chat. So the request was sent to that chat by message instead.

The routine fired on 2/11/2026 without Milo's line in its text. Ask Danica whether Milo's 125 was paid in that run.

I also flagged that the CTS per diem in the recordings is 75 a night, and Danielle's December per diem uses 75 a night. Emelda has used 125 for single-day jobs before. Danica was told to confirm the amount with Emelda.

## Related chats and data

The Payroll Tracker chat holds pay run processing, the tracker workbook and the per diem reminders. It's a separate chat. Its reminder for Danielle Hurley's December per diem (7 nights, 525.00) is still scheduled.

The pay run 8 audit has every pay category and rate for the fortnight ending 16/10/2026. Casual overtime in EH is 1.75 times base and Sunday is 2.25 times base. Those multipliers haven't been checked against the award.

## Sources

- Live Performance Award 2020 [MA000081]: awards.fairwork.gov.au/MA000081.html and library.fairwork.gov.au/award/?krn=MA000081.
- Clerks Private Sector Award 2020 [MA000002]: awards.fairwork.gov.au/MA000002.html.
- Employment Hero help articles, for the menu paths. These are from web search summaries only, because the pages themselves returned 403.
- Payroll recordings in the repo: manuals/part3, part4 and part5 narration transcripts.
- Danica's EH screenshots, her Teams message from Graham, and the files she uploaded.
