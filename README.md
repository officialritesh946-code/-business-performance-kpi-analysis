# -business-performance-kpi-analysis
Retail supply chain KPI analysis — revenue, profit margin, growth rate, and region/product performance, built in Excel.
Business Performance Analysis — Retail Supply Chain KPIs

Author: Ritesh Kumar Task: Analyze business performance using KPIs such as revenue, profit, sales, customer count, growth rate, and product/region performance.

Dataset

Retail-Supply-Chain-Sales-Dataset.xlsx — a retail/superstore-style order-level transaction log covering 2014–2017.

9,994 line items across 5,009 distinct orders and 793 distinct customers
Order-level fields: Order ID, Order Date, Ship Date, Ship Mode, Customer, Segment, Region, State/City, Sales Rep
Product-level fields: Category, Sub-Category, Product Name
Transaction metrics: Sales, Quantity, Discount, Profit, Returned flag
A companion Calendar Table sheet for date/period lookups
Approach
Loaded and profiled the raw data in Python (pandas) to check date range, category/region values, missing data, and distributions before defining any KPI.
Defined each KPI explicitly before calculating it, in line with the task's hints:
Revenue = SUM(Sales)
Profit = SUM(Profit)
Profit Margin = Profit / Revenue
Growth Rate = (This Period − Last Period) / Last Period
Return Rate = Distinct Returned Orders / Distinct Orders
Rebuilt the KPIs as live Excel formulas (SUMIFS/AVERAGEIFS/COUNTIF) against a raw data sheet inside the workbook, rather than pasting static numbers, so the report recalculates if the underlying data changes.
Compared performance year-over-year, not just as one-time totals, to see whether revenue growth was actually translating into profit growth.
Broke performance down by region and by product category/sub-category, then cross-tabbed Region × Category to flag combinations underperforming relative to the overall profit margin.
Wrote up 5 business insights with concrete recommendations, each backed by a specific number from the workbook rather than a general observation.
Tools
Excel (openpyxl / LibreOffice-recalculated formulas) for the KPI workbook
Python (pandas) for data profiling, validation, and cross-checking every formula's output
Power BI was not used for this submission — analysis is fully reproducible from the Excel workbook alone; a Power BI version can be added if required
Files in this repository
File	Description
Retail-Supply-Chain-Sales-Dataset.xlsx	Source dataset (unchanged)
KPI_Summary_Report.xlsx	KPI summary, yearly trend, region breakdown, product breakdown, region×category matrix, and insights — all as live formulas over a raw-data sheet
Business_Insights_and_Recommendations.md	The 5 insights and recommendations, in narrative form
README.md	This file
Inside KPI_Summary_Report.xlsx
KPI Summary — total revenue, profit, margin, orders, customers, AOV, units sold, return rate, and YoY growth
Yearly Trend — revenue, profit, margin and growth by year, 2014–2017
Region Performance — revenue, profit, margin, discount and % of total revenue by region
Product Performance — revenue, profit, margin and units by category/sub-category, with category subtotals
Region-Category Matrix — every region × category combination, flagged "Below Average" where its margin sits under the overall company margin
Insights & Recommendations — the 5 write-ups, inside the workbook as well as in the separate markdown file
Final Outcome

Total revenue across the four years is $2.297M at a 12.5% overall profit margin, but that margin isn't even across the business:

Furniture is the only category that loses money overall (Tables and Bookcases are both net-negative), while Technology — especially Copiers — is the strongest performer.
Discounting above ~20% is where profit turns negative; discounts of 50%+ are sold well below cost.
Central is the weakest region, carrying roughly double the average discount rate of every other region, which drags its margin to the lowest of the four.
Revenue growth has outpaced profit growth since 2016 — margin peaked in 2016 at 13.4% and slipped to 12.7% in 2017 even as revenue kept climbing.

Full detail and recommendations for each point are in Business_Insights_and_Recommendations.md.

Content
Retail-Supply-Chain-Sales-Dataset.xlsx

XLSX
