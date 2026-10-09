UOL LIVE EXCEL WORKBOOK CONNECTION

The dashboards load from the bundled files in data/ by default:
- Basic Reporting: 56_Submissions_Source.xlsx — All Submissions
- Executive Overview and Section 2: UOL_Separated_All_Activity_Data_Cleaned.xlsx — All Separated Activities and Internship - Placement
- Section 3: UOL_MoU_Country_Previous.xlsx — MoU Country (No Nulls)

To load newer Excel data, click Connect Excel at the top of any page. Import each corresponding workbook. The workbook is validated for expected sheet names and retained in browser IndexedDB (only in that browser). The dashboard refreshes and the PDF reports use the selected activity workbook. Use Restore bundled Excel to revert.

To update the published dashboard for ALL visitors, replace the relevant .xlsx file in the data/ folder on your hosting service, preserving its filename and sheet names. Browser overrides only affect your device/browser.

Important: Raw 56-submission forms are not interchangeable with the structured cleaned activity workbook. The overview requires the prepared activity workbook; the MoU section requires the country workbook. If the source submission data changes, regenerate the cleaned activity workbook before uploading it.

The site must be hosted on a regular web server or GitHub Pages for bundled workbook fetches. No RStudio or Power BI installation is required.
