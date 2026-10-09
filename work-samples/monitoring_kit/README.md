# Planting and Monitoring Data Kit

**Illustrative data.** Every site, plot, count and date is invented to show how the sheets work: 3 sites,
12 plots, 40 visits, with 8 kinds of entry error planted on purpose.

The workbook: `../../assets/files/planting-monitoring-kit-illustrative.xlsx` (Excel; opens in Google Sheets).

## Sheets

| Sheet | Purpose |
|---|---|
| Read_Me | How to use the kit, data rules, the 8 checks |
| Dashboard | Key figures and three charts |
| Settings | Report date, replanting trigger, plausible ranges, dropdown lists (named ranges) |
| Planting_Log | One row per plot planted |
| Monitoring | One row per plot visit, formula columns, QA flags |
| Plot_Status | One row per plot from its last clean visit: stocking, status, seedlings to replant, next action |
| Site_Summary | Totals per site, stocking by census month, losses by main cause |
| Field_Day_SOP | 18-step planting-day procedure with owner and record |
| Risk_Register | 12 risks scored likelihood x severity with controls |

## The 8 checks

Unknown plot ID, visit before planting, count mismatch, alive rose without replanting, duplicate visit,
height out of range, no observer, loss without cause. On the illustrative data they flag 9 rows.

## Data rules

1. Counts are cumulative since planting: alive + dead + missing = seedlings planted + replanted on earlier visits.
2. Replanting is recorded on the visit where it happens and enters the expected total from the next visit.
3. Plot status uses only rows that pass every check; flagged rows are corrected, never deleted.
4. Stocking (alive / design count) drives replanting; survival (alive / all seedlings ever planted) shows its cost.

## Scripts

- `make_kit.py` builds the workbook with openpyxl. Formulas use only COUNTIFS, SUMIFS, AVERAGEIFS, SUMPRODUCT
  and INDEX/MATCH, so they work in Excel and Google Sheets.
- `qa_check.py` reads only the input cells, recomputes every flag and plot status independently, and compares
  with the values Excel saved. Last run: 248 values checked, 0 mismatches.
