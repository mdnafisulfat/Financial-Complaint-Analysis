<img width="1371" height="585" alt="image" src="https://github.com/user-attachments/assets/87ac651a-316e-4135-a549-f9af00f4c698" /># Financial Complaint Analysis

This project performs a complete analysis of financial complaints using **Microsoft Excel only**. The main goal is to identify which financial products are most problematic, where responses are slow, how frequent fraud complaints are, and how the company is performing overall.  

The dataset contains the following key columns:  
`Complaint ID, Submitted via, Date submitted, Date received, Year, Month, Resolution_Days, Delay_Category, State, State_Name, Product, Sub-product, Issue, Issue_Category, Sub-issue, Company public response, Company response to consumer, Response_Score, Timely response?, Timely response? (After TRIM)`

---

## 💡 Project Objective

The purpose of this project is to generate actionable insights from financial complaints data while showcasing advanced Excel skills. The analysis helps to understand:

- Which products are causing the most complaints  
- Which states or regions have slow responses or high complaint volume  
- Fraud-related complaint patterns  
- Company response performance and operational efficiency  

All of this is achieved using Excel alone, without SQL, Python, or Power BI.

---

## 🗂 Project Structure

### 1️⃣ Raw Data

The first sheet contains the original dataset to ensure data integrity. No formulas or transformations were applied.  

### 2️⃣ Cleaned & Reference Data

This sheet combines the cleaned data with reference mapping to enable feature engineering and dynamic classification.

**Reference Tables:**  
- **Response Score Map:** Converts company responses into numeric scores (e.g., "Closed with monetary relief" = 2, "Closed with explanation" = 1, "In progress" = 0).  
- **Issue Category Map:** Uses keywords to classify complaints automatically (e.g., "Fraud" → "Fraud Issue", "Incorrect" → "Credit Issue").  
- **State Map (Optional):** Maps state codes to human-readable state names.  

**Engineered Columns:**  
- **Year & Month:** Extracted from `Date submitted` to analyze trends.  
```excel
=YEAR([@[Date submitted]])
=TEXT([@[Date submitted]],"mmm")
```
- **Resolution_Days:** Measures operational efficiency.  
```excel
=[@[Date received]]-[@[Date submitted]]
```
- **Delay_Category:** Classifies response as Fast, Medium, or Slow.  
```excel
=IF([@Resolution_Days]<=1,"Fast",
 IF([@Resolution_Days]<=3,"Medium","Slow"))
```
- **Response_Score:** Uses XLOOKUP to assign numeric scores to responses.  
- **Issue_Category:** Automatically classifies complaints based on keywords using XLOOKUP and SEARCH.  
- **State_Name:** Optional, converts codes into readable state names.  

This sheet demonstrates feature engineering, logical functions, text manipulation, and lookup function expertise.

### 3️⃣ KPI Summary

The KPIs are calculated to provide business insights. Metrics include total complaints, timely response percentage, monetary relief percentage, average resolution days, fraud percentage, and average response score. These indicators reflect operational efficiency, complaint severity, and customer satisfaction proxies.  

**Example Formula:**  
```excel
=COUNTIF(TimelyColumn,"Yes")/Total
```

Skills showcased: COUNTIF, SUMIF, AVERAGE, percentage calculation, and business metric design.

### 4️⃣ Pivot Analysis

Pivot tables are used to analyze complaints in multiple dimensions. Examples include product-wise complaint counts, product vs. response distribution, product vs. average resolution days, state-wise complaints, and issue category distributions. These analyses help identify problematic products, complaint hotspots, and recurring issues.  

Skills demonstrated: Pivot tables, multi-level aggregation, grouping, and business insight extraction.

### 5️⃣ Trend Analysis

Time-series trends are analyzed to understand operational efficiency over time. Year-over-year growth and monthly complaint trends reveal seasonal spikes. Resolution trends highlight periods of faster or slower company response.  

**Year-over-Year Growth Formula:**  
```excel
=(CurrentYearTotal-PreviousYearTotal)/PreviousYearTotal
```

Skills demonstrated: Growth percentage calculation, time-series analysis, trend spotting.

### 6️⃣ Charts & Visuals

Visual storytelling is achieved using Excel charts. Bar charts display product-wise complaints, pie charts show response distribution, line charts track yearly trends, column charts compare states, and stacked charts illustrate product vs response. Charts are cleanly formatted with proper axes and labels for professional presentation.  

Skills demonstrated: Chart formatting, visual storytelling, axis labeling, clean and professional visualization.

### 7️⃣ Interactive Dashboard

An executive-ready dashboard summarizes key insights. The top section displays KPI cards, the middle section shows a yearly trend line, and the bottom section contains product charts, state charts, and issue category breakdowns. Slicers allow filtering by year, product, state, and issue category.  

The dashboard provides immediate insights into complaint trends, fraud risk, slow-performing products, and operational efficiency.  

Skills demonstrated: Dashboard design, slicer integration, interactive filtering, business storytelling.

https://github.com/mdnafisulfat/Financial-Complaint-Analysis/blob/main/Screenshot%202026-02-14%20191856.png?raw=true
https://github.com/mdnafisulfat/Financial-Complaint-Analysis/blob/main/Screenshot%202026-02-14%20192048.png?raw=true
https://github.com/mdnafisulfat/Financial-Complaint-Analysis/blob/main/Screenshot%202026-02-14%20192102.png?raw=true
https://github.com/mdnafisulfat/Financial-Complaint-Analysis/blob/main/Screenshot%202026-02-14%20192111.png?raw=true
https://github.com/mdnafisulfat/Financial-Complaint-Analysis/blob/main/Screenshot%202026-02-14%20192119.png?raw=true
https://github.com/mdnafisulfat/Financial-Complaint-Analysis/blob/main/Screenshot%202026-02-14%20192128.png?raw=true



---

## 🚀 Key Takeaways

This project proves that **Excel alone** can be used for end-to-end financial complaint analysis. Structured feature engineering, pivot tables, and charts enable clean, actionable insights, while KPI and dashboard design provide executive-ready reporting.

---

## 🛠 Tools Used

Microsoft Excel, Pivot Tables, XLOOKUP, Logical & Text Functions, Charts, Interactive Dashboard

---

> Portfolio-ready project demonstrating financial complaint analysis entirely in Excel, combining data cleaning, feature engineering, KPI calculation, pivot analysis, charts, and dashboard design.
