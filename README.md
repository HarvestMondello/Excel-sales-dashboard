# Coffee Sales Dashboard in Excel

# Synopsis

**Problem**: The company lacked a clear, consolidated view of orders, products, and profit trends to guide business decisions.

**Solution**: An Excel dashboard was built using Power Query, PivotTables, and slicers to automate, aggregate, and visualize key metrics.

**Insights**: Robusta and Excelsa led sales, U.S. customers drove most profits, and loyalty members showed higher repeat behavior. 

**Recommendation**: Focus on loyalty programs, promote top-selling blends, and align inventory with seasonal profit trends.

# Introduction
Welcome to my Excel project. In this project I will look coffee sales to answer business questions. To answer theses questions I have developed an interactive dynamic Excel dashboard to analyze coffee sales, customer behavior, and profitability using automated data transformations and visual insights.

Download the Excel file here - select view raw to download the Excel document: [data folder](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/data/coffeeOrdersData-PQ.xlsm)

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


# Excel Dashboard

![Excel Dashboard](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/assets/excel-dashboard.png)


# How I Built It
I began by cleaning and completing the source tables using Excel formulas such as `XLOOKUP`, `INDEX`, and `IF` to fill in missing values and derive necessary attributes.  
Next, I used **Power Query** to join the three primary tables; orders, customers and products, into a single unified query.  
I added calculated columns within Power Query to compute profits, monthly trends, and key metrics, automating data preparation across the file.  
Finally, I built interactive **PivotTables** and **PivotCharts** on top of this clean dataset to uncover trends by customer, product, and region, and to highlight top contributors and profit patterns.

I concluded the dashboard by creating a macro to print the repot to PDF. See the [macro code here in VBA](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/vba/SaveReportAsPDF.bas).

# Conclusion
**Problem**: The company lacked a clear, consolidated view of orders, products, and profit trends to guide business decisions.

**Solution**: An Excel dashboard was built using Power Query, PivotTables, and slicers to automate, aggregate, and visualize key metrics.

**Insights**: Robusta and Excelsa led sales, U.S. customers drove most profits, and loyalty members showed higher repeat behavior. 

**Recommendation**: Focus on loyalty programs, promote top-selling blends, and align inventory with seasonal profit trends.

### License

Created by Harvest Mondello. You're welcome to use this project for **personal** or **educational** purposes! Feel free to explore, adapt, and learn from the code and visuals. Just note that **commercial use isn’t permitted** without permission. See [LICENSE.md](https://github.com/HarvestMondello/coffee-sales-dashboard/blob/main/LICENSE.MD) for full details and contact info.
