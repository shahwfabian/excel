You are working on my Windows PC, which has desktop Microsoft Excel installed. Complete "Lab 3 – Creating More PivotCharts" by automating desktop Excel through COM from PowerShell. The deliverable is a finished workbook, not a report.

## Files
- Starting workbook: `lab03.xlsx`. Instructions: `lab03.pdf`. Background reading only: `ch04.pdf`.
- They should be in the current folder. If not, search my Downloads, Desktop and Documents folders. If you still can't find them, ask me.
- Output: a **new copy** named `lab03end_Fabian.xlsx`, saved in the same folder as `lab03.xlsx`.
- **Never modify or overwrite `lab03.xlsx`.** Copy it to `lab03end_Fabian.xlsx` first and work only on the copy.

## Hard rules
- Use only desktop Excel through COM (`New-Object -ComObject Excel.Application`).
- Do NOT use openpyxl, pandas, LibreOffice, ClosedXML, EPPlus, or manual XML/zip edits. They damage or delete the Data Model (Power Pivot), connections and PivotCharts.
- If Excel COM does not work, stop and tell me why. Do not build a substitute.
- Do not rebuild the workbook. Do not replace PivotCharts with ordinary charts.
- Keep everything that already exists: sheets, dashboard, charts, logo, Data Model, connections, PivotTables and pivot caches.
- **Do not refresh the Power Query queries or call `RefreshAll`.** The queries point to a file on the professor's PC (`C:\Users\drsjl\...\Sales Data.xlsx`), which doesn't exist here, so a refresh would fail and could empty the model. Only use data already in the Data Model.
- Set `$xl.DisplayAlerts = $false`. Always close the workbook and quit Excel in a `finally` block, even after an error. Release the COM objects so no hidden `EXCEL.EXE` stays running.
- Work step by step: inspect first, then build one chart at a time, then verify. Write each script to a `.ps1` file and run it with `powershell -ExecutionPolicy Bypass -File <script>.ps1`.

## What's already in lab03.xlsx
I inspected the file already. Confirm this yourself in Step 1 before changing anything.
- Sheets: `Sales Dashboard` (the dashboard, gridlines off) and `PivotTables` (empty).
- Existing objects on `Sales Dashboard`:
  - a university logo picture
  - PivotChart `cRevenueByYear` (column chart, about B4:D18)
  - PivotChart `cRevenueByOrderMethod` (doughnut chart, about A19:D34)
  - "Revenue", "Total Cost" and "Gross Margin" card cells along the top, in columns F, H and J
- Both existing PivotCharts come from the Data Model and have no visible PivotTable behind them.
- Connections: `Query - Dates`, `Query - Region`, `Query - Sales`, `Query - SalesManager`, and `ThisWorkbookDataModel`.
- Data Model tables: Dates (FullDate, Year, Start of Month, FY), Region (Country, Region), Sales (Revenue and other columns), SalesManager (Country, Sales Manager).
- Relationships:
  - Sales[Date] → Dates[FullDate]
  - Sales[Retailer country] → Region[Country]
  - Sales[Retailer country] → SalesManager[Country]
- Numbers to check against:
  - Total Sum of Revenue = **1,067,875,618.43**
  - Sales data runs from Jan 2022 to Dec 2024, so the month axis should show **36 months** (1/1/2022 … 12/1/2024)
  - **5 regions**: Asia, Europe, Latin America, North America, Oceania
  - **14 sales managers**
  - Correct ascending order by revenue, smallest to largest: Mia Jones, Lucas Maes, Carlos Silva, Navin Singh, Hugo García, Antonio González, Maria Romano, Jonas Müller, George Jones, Wang Shu, Camila Dubois, Sophia Evans, Pablo Ramirez, Emma Smith.
  - A stacked bar chart draws the first category at the bottom, so Emma Smith ends up at the top and Mia Jones at the bottom. That matches the figure in lab03.pdf.

## Step 1 – Inspect (read only)
Open `lab03.xlsx` read-only and print:
- the sheet names
- each `ChartObject` on `Sales Dashboard`: name, TopLeftCell, BottomRightCell, Left, Top, Width, Height, and whether `.Chart.PivotLayout` is set
- the shapes
- the workbook connections and their types
- `$wb.Model.ModelTables` with row counts
- `$wb.Model.ModelRelationships`
- the pivot caches, including which ones are OLAP / Data Model

Close without saving. Then read `lab03.pdf` to confirm the steps below match it.

## Step 2 – Create the copy
Copy `lab03.xlsx` to `lab03end_Fabian.xlsx`. Open the copy in Excel. It's fine to keep Excel visible so I can watch.

## Step 3 – PivotChart 1: Revenue by Month (lab pages 1–5)
Do this the same way as Insert → PivotChart → "Use this workbook's Data Model" → Existing Worksheet:
```
$conn = $wb.Connections.Item("ThisWorkbookDataModel")
$pc   = $wb.PivotCaches().Create(2, $conn, 6)   # 2 = xlExternal, 6 = version for Excel 2016+
$shp  = $pc.CreatePivotChart($ws, 4, <Left>, <Top>, <Width>, <Height>)   # $ws = 'Sales Dashboard', 4 = xlLine
$ch   = $shp.Chart
$pt   = $ch.PivotLayout.PivotTable
```
If `CreatePivotChart` with these arguments doesn't work in my Excel version, adapt it. The result must still be a PivotChart on the Data Model and placed on `Sales Dashboard`.

Fields:
- **Axis (Categories):** `CubeFields("[Dates].[Start of Month]").Orientation = 1` (xlRowField).
- **Legend (Series):** `CubeFields("[Region].[Region]").Orientation = 2` (xlColumnField).
- **Values:** Sum of `Sales[Revenue]`. Reuse the measure `[Measures].[Sum of Revenue]` if it already exists. If not, create it with `$pt.CubeFields.GetMeasure("[Sales].[Revenue]", -4157, "Sum of Revenue")` (-4157 = xlSum), then `$pt.AddDataField(<that measure>)`.
- **Remove extra categories:** if Excel automatically adds date-grouping fields (Year/Quarter/Month of Start of Month), remove them. The axis must show individual Start of Month dates (1/1/2022 …), like the lab figure.

Formatting:
- **Field buttons (step 7, Hide All):** `$ch.ShowAllFieldButtons = $false`.
- **Chart type (steps 8–9):** Line, `$ch.ChartType = 4` (xlLine).
- **Legend (step 10):** `$ch.HasLegend = $true` and `$ch.Legend.Position = -4160` (xlLegendPositionTop).
- **Vertical axis (steps 11–12):** `$ch.Axes(2).TickLabels.NumberFormat = '#,,"M"'` (2 = xlValue).
- **Legend font (step 13):** `$ch.Legend.Format.TextFrame2.TextRange.Font.Size = 12`.
- **No outline (step 14):** `$ch.ChartArea.Format.Line.Visible = 0` (msoFalse).
- **Title (step 15):** `$ch.HasTitle = $true` and `$ch.ChartTitle.Text = "Revenue by Month"`.
- **PivotChart Name (steps 16–17):** `$shp.Name = "cMonthlyRevenue"`. This is the name shown in PivotChart Options → PivotChart Name. After setting it, read the name back.

## Step 4 – PivotChart 2: Revenue by Sales Manager (lab pages 6–10)
Create a new PivotChart from the Data Model the same way as in Step 3.

Fields:
- **Axis:** `[SalesManager].[Sales Manager]` as xlRowField.
- **Legend (Series):** `[Region].[Region]` as xlColumnField.
- **Values:** Sum of Revenue.

Formatting:
- **Field buttons (step 5):** hide all.
- **Chart type (steps 6–7):** Stacked Bar, `$ch.ChartType = 58` (xlBarStacked).
- **Legend (step 8):** top.
- **Horizontal value axis (steps 9–10):** `$ch.Axes(2).TickLabels.NumberFormat = '#,,"M"'`. In a bar chart, xlValue is the horizontal axis.
- **Sort (steps 11–12):** ascending by Sum of Revenue. This is "More Sort Options → Ascending (A to Z) by Sum of Revenue".
  - Get the row field with `$pt.RowFields(1)`, or `$pt.PivotFields("[SalesManager].[Sales Manager].[Sales Manager]")`.
  - Call `.AutoSort(1, "Sum of Revenue")` (1 = xlAscending). If Excel wants the data field's full name, use `$pt.DataFields(1).Name`.
  - Read the order back and compare it with the expected order above.
- **No outline (step 13):** chart area line not visible.
- **Legend font (step 14):** 12.
- **Title (step 15):** "Revenue by Sales Manager". Make the chart tall enough that all 14 manager names and bars show. Check that no category label is skipped: `$ch.Axes(1).TickLabelSpacing` should be 1. Set it to 1 if it isn't.
- **PivotChart Name (steps 16–17):** `$shp.Name = "cRevenueBySalesManager"`.

## Step 5 – Layout (step 18 on pages 5 and 10)
Arrange the charts to match the example dashboard in lab03.pdf. Look at those figures before you do this.

**cMonthlyRevenue**
- Goes to the right of `cRevenueByYear` and under the Revenue / Total Cost cards.
- It covers columns F through H. Its top should line up with `cRevenueByYear` (about row 4) and its bottom with the bottom of `cRevenueByYear` (about row 18).
- Position it with the cells' `.Left`, `.Top` and `.Width` values, not guessed numbers. For example: Left = `$ws.Range("F4").Left`, Top = `cRevenueByYear.Top`, Width = right edge of column H minus Left, Height = `cRevenueByYear.Height`.

**cRevenueBySalesManager**
- Goes in the right-hand area to the right of the Gross Margin card column (column J). Start it at column K, about rows 19 to 36, to the right of and at the same level as `cRevenueByOrderMethod`.
- Make it wide enough for the legend and names (roughly columns K to P or Q), and tall enough for all 14 bars.

General:
- Don't move, resize or delete the logo, `cRevenueByYear`, `cRevenueByOrderMethod`, or the card cells.
- Nothing should overlap. Labels and legends must be readable.

## Step 6 – Verify in Excel before saving
Write a verification script that reads each value back from Excel and prints PASS or FAIL. Check:

1. `cMonthlyRevenue` and `cRevenueBySalesManager` exist on `Sales Dashboard`. `.Chart.PivotLayout` is not null for both. Each PivotTable's cache is OLAP (`PivotCache.OLAP = $true`) and uses the `ThisWorkbookDataModel` connection.
2. Fields:
   - chart 1: row = Start of Month, column = Region, data = Sum of Revenue
   - chart 2: row = Sales Manager, column = Region, data = Sum of Revenue
   - there are no extra row or column fields
3. Chart types are 4 (xlLine) and 58 (xlBarStacked).
4. Titles are exactly "Revenue by Month" and "Revenue by Sales Manager".
5. Legends are on top (-4160) with font size 12.
6. The value axis number format is exactly `#,,"M"` on both charts.
7. Field buttons are hidden: `ShowAllFieldButtons` is false, or every individual button property is false.
8. Chart area outlines are not visible.
9. Chart 2's category order equals the expected ascending order. Check it with `$pt.RowRange` values or `$ch.SeriesCollection(1).XValues`.
10. Both charts show data:
    - chart 1 has 5 series and 36 points
    - chart 2 has 5 series and 14 categories
    - the grand total is 1,067,875,618.43 (use `$pt.GetPivotData` or `$pt.DataBodyRange`)
11. Nothing is lost:
    - both sheets still exist
    - the logo, `cRevenueByYear` and `cRevenueByOrderMethod` are still there and still PivotCharts
    - all 5 connections are still there
    - `$wb.Model.ModelTables.Count` is 4, with the same row counts as in Step 1
    - the 3 relationships are still there
12. Export the dashboard to `lab03end_Fabian_preview.png`, or take a screenshot. Look at it to confirm it resembles the lab figure and is readable.

Fix any FAIL and run the verification again. Only save when everything passes.

## Step 7 – Save and report
- Save as `lab03end_Fabian.xlsx` (`$wb.SaveAs(<full path>, 51)`; 51 = xlOpenXMLWorkbook). Close Excel.
- Reopen the saved file in Excel and run the Step 6 checks again. This proves the save kept everything.
- Confirm `lab03.xlsx` is unchanged: compare its hash to one taken before you started.
- Give me a short checklist of each lab step and requirement, marked Verified in Excel, Failed, or Not verifiable, with the evidence (the value you read back). Don't mark anything as passed unless you read it back from Excel. Tell me the full path of the saved file and the preview image.
