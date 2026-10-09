# Operations & Science portfolio: Les' Aullian Achmad Argito

Live site: https://lesaullianachmadargito.github.io/blue-carbon-ops-portfolio/

A job-application portfolio for an Operations & Science Intern role in blue carbon restoration. The site takes
the six responsibilities in the posting one at a time (field operations, project data, site assessment,
analysis and scientific writing, problem solving with community partners, SOPs and risk) and shows the evidence
for each, in five tabs.

## Work samples built for this application

| Folder | What it is |
|---|---|
| `work-samples/carbon_stock/` | Mangrove carbon stock from plot data: stems to biomass with the Komiyama et al. (2005) common equations, soil cores to organic carbon with two published LOI equations, per-hectare means with confidence intervals, and the number of plots a precision target needs. Python. |
| `work-samples/monitoring_kit/` | A planting and survival monitoring workbook (Excel, opens in Google Sheets) with 8 automatic data checks, replanting triggers, a dashboard, a planting-day SOP and a risk register, plus an independent Python check that recomputes every flag and status. |

**All work-sample data is illustrative.** It is generated or invented to show how the methods work and
describes no real site, partner or client. Parameters that must be replaced before real use are marked
`PLACEHOLDER` in the code.

## Run it

```
python work-samples/carbon_stock/carbon_stock.py
python work-samples/monitoring_kit/make_kit.py assets/files/planting-monitoring-kit-illustrative.xlsx
python work-samples/monitoring_kit/qa_check.py assets/files/planting-monitoring-kit-illustrative.xlsx
```

`carbon_stock.py` needs NumPy, SciPy and Matplotlib. `make_kit.py` needs openpyxl. `qa_check.py` reads the
values Excel saved, so open and save the workbook in Excel (or LibreOffice) once after building it.

## Notes

An independent portfolio, not affiliated with or endorsed by any employer named on the site. Student ID
numbers, signatures and villagers' faces are not published here.
