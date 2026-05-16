# Football-Transfer-Analysis-in-Excel
An in-depth data analysis of international football transfer economics across the 2021/2022 and 2022/2023 seasons - built entirely in Microsoft Excel.

**📌 Short Description / Purpose**
This project explores the economic patterns behind international football (soccer) player transfers using a multi-season dataset. It covers the full analytical workflow: raw data cleaning and preprocessing, cross-sheet lookups, multi-criteria aggregation using Excel functions, and interactive data visualization — all within Excel. The goal is to quantify how players and money flow between countries and associations, with a particular focus on European football markets.

**🗂️ Project Structure**
The workbook is organized across 6 sheets, each serving a distinct analytical purpose:

| Sheet | Purpose |
|---|---|
| `Database` | Master dataset of all transfer records for both seasons |
| `Countries` | Reference table mapping each country to its continent |
| `European Analysis` | Summary of total transfers in/out of Europe by season |
| `European Analysis Transfer` | Net transfer movements (count & value) for each European country |
| `Top 5 Country` | Ranking of top 5 European countries by incoming transfer spend in 2022/2023 |
| `Visualization Top 5 Countries` | Chart-ready table with transfer counts and average fees for the top 5 |

**🔧 Tech Stack / Formulas Used**
This project is built entirely in Microsoft Excel, leveraging the following features and functions:

1. Data Preparation
- Text to Columns — Split combined country-continent data in the `Countries` sheet into separate, usable columns
- Header Filters — Used to scan the `Season` column for data inconsistencies
- Find & Replace — Corrected erroneous season year entries across the database
- VLOOKUP / Cross-sheet referencing — Populated the `Continent` columns in the `Database` sheet from the `Countries` reference table

2. Excel Functions:
| Function | Usage |
|---|---|
| `SUMIFS` | Aggregate transfer counts and dollar values by season, country, and continent |
| `SUMIF` | Pull season-specific transfer data into the visualization table |
| `RANK` | Rank European countries by total incoming transfer spend in 2022/2023 |
| Arithmetic operators (`+`, `-`, `/`) | Calculate net transfers (incoming − outgoing) and average transfer fees |

3. Visualization:
Combo Chart (Clustered Column + Line with Secondary Axis) — Displays number of incoming transfers (primary axis) alongside average transfer fee per player (secondary axis) for the top 5 European countries

**📊 Features / Highlights**
1. 🔍 Data Cleaning & Preprocessing
* Identified and filled missing continent data for all countries using the Countries reference sheet.
* Corrected inconsistent season labels using Find & Replace.
* Ensured data integrity across both the 2021/2022 and 2022/2023 seasons before any analysis.

2. 🌍 European Transfer Summary (Aggregate View)
* Created a summary table showing the total number of incoming and outgoing transfers involving Europe for each season.
* Calculated the net transfer balance to determine whether Europe, overall, is a net importer or exporter of football talent.
* Powered entirely by SUMIFS with continent and season criteria.

3. 🗺️ Country-Level Transfer Breakdown
Built a detailed table listing every European country with:
* Number of incoming and outgoing transfers per season.
* Net player movement (positive = net importer, negative = net exporter).
* Total incoming and outgoing transfer spend (in USD).
* Net transfer value — showing which countries are net spenders vs. net earners.

4. 🏆 Top 5 Big Spenders — 2022/2023 Season
* Identified the top 5 European countries by total incoming transfer expenditure in 2022/2023 using the RANK function.
* For these countries, calculated the average transfer fee per player (total spend ÷ number of incoming transfers).
* Visualized the results in a dual-axis combo chart:
  
  Bar chart → Number of incoming transfers per country
  Line chart (secondary axis) → Average fee paid per incoming player

**💡 Key Insights Enabled by This Analysis**

* Which European countries are the biggest buyers vs. sellers in the transfer market.
* How transfer spending patterns shifted between the 2021/2022 and 2022/2023 seasons.
* The average price a top-spending country pays per player — a proxy for squad investment strategy.
* Whether Europe as a continent is a net importer or exporter of global football talent.

**📁 Dataset**
Source: Football transfer dataset (provided as part of the project)
Seasons covered: 2021/2022 and 2022/2023
Scope: International club-to-club transfers across global football associations
Key fields: Season, Transfers Incoming (Country), Transfers Outgoing (Country), Continent (Incoming & Outgoing), Number of Transfers, Total Club-to-Club Compensation (USD)


**🎓 Skills Demonstrated**
* Data preprocessing and error correction in Excel.
* Multi-criteria aggregation with SUMIFS and SUMIF.
* Cross-sheet data referencing and formula design.
* Ranking and filtering with RANK and header filters.
* Data visualization with combo charts and dual axes.
* Deriving business insights from raw sports economics data.

**📄 License**
This project was completed as part of an Introduction to Excel course assignment. Dataset and case study provided by the course instructor. For educational use only.
