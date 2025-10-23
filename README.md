# **MainStreet-Renewal-Supply-Chain-Dashboard-Excel-BI-System-**
Interactive Excel dashboard developed for MainStreet Renewal’s Supply Chain and Field Services Department, integrating data from multiple SharePoints to monitor KPIs, team performance, and job status across U.S. markets. The tool provides weekly segmentation and KPI visualization using PivotTables, charts, VBA and custom VLOOKUP/XLOOKUP tools

## **_Data Privacy Notice_**
Due to confidentiality agreements with MainStreet Renewal, the datasets used in this project cannot be shared, as they contain sensitive client and property information.
The visuals shown here are authorized screenshots of the dashboards and analytical tools developed. Any potentially sensitive details (such as property addresses) have been hidden in accordance with company policy.

---

### Overview
Developed an interactive Excel-based Business Intelligence dashboard for the Supply Chain Construction and Field Services Department at MainStreet Renewal (U.S.).
The tool consolidates multi-source data from several company SharePoints into a single, structured environment — enabling transparent performance tracking, KPI monitoring, and data-driven decision-making across all U.S. markets.

### Key Features
- Multi-Source Data Integration:
  - Automated data consolidation from several SharePoint datasets into unified tables for real-time visibility across markets.

- Data Cleaning & Transformation
  - Applied best practices in data validation, normalization, and error handling, ensuring consistent and accurate analysis.
  - Added custom calculated fields for week-based reporting (Year-Week), allowing consistent week-over-week KPI review.

- Analytical Tools (VLOOKUP & XLOOKUP)
  - Designed helper sheets (“Paint Order Finder” & “Property Record Finder”) to instantly retrieve detailed order information (e.g., market, specialist, address, cycle time) by ID.
  - Implemented dynamic lookup formulas combining IF, VLOOKUP, and XLOOKUP for adaptive error control and automatic date-to-week conversion.

- Interactive Dashboard Design
  - Created fully interactive dashboards using PivotTables and PivotCharts, featuring:
    - KPI cards and dynamic metrics (e.g., Average Lead Time, Total Orders, Paint Volume, KPI Accomplishment Rate).
    - Filters and slicers by Market, Year-Week and Assigned Specialist for comparative analysis.
    - Donut charts, clustered columns, and summary tables showing team and market performance.

- Automation & UX Enhancements:
  - Added custom ActiveX scrollbars connected to ten-row summarized tables, programmed to automatically adjust their maximum value based on current pivot dataset size.
  - Implemented dynamic chart updates and conditional formatting, ensuring that KPI visuals highlight goal completion and team performance.


### Technologies & Skills
- **Software**: Microsoft Excel (Advanced), Power Query, VBA, SharePoint.
- **Techniques**: Data Consolidation, ETL (Extract-Transform-Load), Lookup Functions, VBA, PivotTables, ActiveX components, KPI Visualization, UX Design in Excel.
- **Key Concepts**: Data Validation, Dynamic Ranges, Week-based Time Analysis, Conditional Formatting, Dashboard Design Principles.

### Outcomes
- Reduced manual report preparation time by automating data consolidation and visualization processes.
- Enhanced operational transparency and decision-making efficiency for management.
- Provided a unified framework for weekly performance reviews.
- Empowered team leads and specialists to monitor their own KPIs and performance trends through intuitive dashboards.
