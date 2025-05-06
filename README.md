# Excel Sales Dashboard in Excel
Excel coffee sales 

# Synopsis

**Problem**: The company lacked a clear, consolidated view of orders, products, and profit trends to guide business decisions.

**Solution**: An Excel dashboard was built using Power Query, PivotTables, and slicers to automate, aggregate, and visualize key metrics.

**Insights**: Robusta and Excelsa led sales, U.S. customers drove most profits, and loyalty members showed higher repeat behavior. 

**Recommendation**: Focus on loyalty programs, promote top-selling blends, and align inventory with seasonal profit trends.

# Introduction
Welcome to my Excel sales dashboard project. In this case study, I analyze coffee sales data to answer key business questions around customer behavior, product performance, and profitability.

To do this, I developed an interactive, dynamic dashboard in Excel using automated Power Query transformations, PivotTables, and slicers. The goal was to turn raw, disconnected sales data into actionable insights that drive better decision-making.

## Table of Contents

- [Synopsis](#synopsis)
- [Introduction](#introduction)
- [Background](#background)
- [Tools I Used](#tools-i-used)
- [How I Built It](#how-i-built-it)
- [The Analysis](#the-analysis)
- [Conclusion](#conclusion)
- [License](#license)


Download the Excel file here - select view raw to download the Excel document: [data folder](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/data/coffeeOrdersData-PQ.xlsm)

![Excel Dashboard](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/assets/excel-dashboard.png)

# Background
Many companies rely on Excel for day-to-day reporting, visualization and decision-making, so I built this dashboard to demonstrate how raw sales data can be transformed into actionable insights using the tools they already use.

This project was inspired by data and a tutorial from Mo Chen’s [Youtube channel](https://www.youtube.com/watch?v=m13o5aqeCbM&t=3188s). I expanded on the project it by adding automated Power Query transformations and custom VBA macros to enhance functionality and interactivity.

# Tools I used
- **Excel**: Dashboard design using PivotTables, slicers, conditional formatting, and advanced formulas including `XLOOKUP`, `INDEX`, and `IF` for dynamic calculations.  
- **Power Query (M)**: Automated data transformation and monthly profit calculations.  
- **VBA (Macros)**: Streamlined data refresh and automation tasks.  
- **PivotCharts**: Interactive visual summaries.  
- **Visual Studio Code (VS Code)**: Editing GitHub-facing files such as `README.md`, `.bas` (VBA modules), and license documentation.  
- **GitHub and Git**: Version control and sharing of analysis, visualizations, and code.  
- **ChatGPT**: Assisted with routine tasks and project efficiency.

# How I Built It
I began by cleaning and completing the source tables using Excel formulas such as `XLOOKUP`, `INDEX`, and `IF` to fill in missing values and derive necessary attributes.  
Next, I used **Power Query** to join the three primary tables; orders, customers and products, into a single unified query.  
I added calculated columns within Power Query to compute profits, monthly trends, and key metrics, automating data preparation across the file.  
Finally, I built interactive **PivotTables** and **PivotCharts** on top of this clean dataset to uncover trends by customer, product, and region, and to highlight top contributors and profit patterns.

I concluded the dashboard by creating a macro to print the repot to PDF. See the [macro code here in VBA](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/vba/SaveReportAsPDF.bas).

# The Analysis
This analysis explores sales and ordering trends within a coffee business using data extracted from a macro-enabled Excel workbook. Power Query was used to clean, transform, and combine the raw data tables, while Excel formulas, PivotTables, and slicers enabled interactive exploration of product performance, customer behavior, and profitability over time.

The analysis focused on the following key areas:

**Order Volume Over Time**
Monthly aggregation of coffee orders revealed clear seasonal patterns, with order volumes peaking in Q4. Time-series PivotCharts helped visualize these trends and highlight periods of high demand.

**Product-Level Insights**
Using frequency counts in PivotTables, I identified top-selling products. Robusta and Excelsa consistently led in volume.

**Customer Segmentation**
Customer data was grouped to analyze purchase frequency and total spend. Loyalty customers—those with 10+ orders—were responsible for a large share of revenue, highlighting the value of repeat buyers.

**Profitability and Margins**
Gross profit per product was calculated using item costs and sales prices. The dashboard surfaced several low-volume items with high margins, indicating potential opportunities for targeted promotions.

**Interactive Dashboard**
The final dashboard includes slicer-driven PivotTables and PivotCharts, allowing users to filter by date, customer type, and product category. Power Query was used to transform dates into month-level granularity for trend analysis and to automate data refresh and aggregation.

# Conclusion
**Problem**
The coffee business lacked a consolidated view of sales dynamics, customer behavior, and product profitability. Raw data was fragmented across multiple tables, making it difficult to generate timely, data-driven decisions.

**Solution**
I used Excel’s Power Query to automate data cleaning and merging, built calculated columns for profit and time-based metrics, and designed a dynamic dashboard using PivotTables, slicers, and charts. This transformed the dataset into an interactive analysis tool for uncovering patterns and performance drivers.

**Insights**

Robusta and Excelsa emerged as the top-selling blends.

U.S.based customers generated the highest profits.

Loyalty program members showed stronger repeat purchase behavior and higher overall spend.

Profitability spiked during colder months, revealing seasonal sales trends.

**Recommendation**
Double down on loyalty incentives to retain high-value customers, promote top-margin products like Robusta and Excelsa, and align inventory and staffing with seasonal demand patterns. Continue enhancing the dashboard for ongoing insight and explore automating future reporting workflows.

### License

Created by Harvest Mondello. You're welcome to use this project for **personal** or **educational** purposes! Feel free to explore, adapt, and learn from the code and visuals. Just note that **commercial use isn’t permitted** without permission. See [LICENSE.md](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/LICENSE.MD) for full details and contact info.
