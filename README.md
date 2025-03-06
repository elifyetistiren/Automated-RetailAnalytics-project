
# Retail Sales Data Analysis and Automation Project

## <ins>Project Background</ins>

The company operates in the retail industry, selling a diverse range of products through its physical store network. While it has an established market presence, utilizing data-driven insights is essential for optimizing operations and enhancing customer satisfaction.  
To gain a deeper understanding of customer purchasing behavior, revenue patterns, and operational efficiency, this project focuses on analyzing daily sales data across multiple store locations over a one-week

period. By leveraging insights from sales transactions, store performance, customer demographics, and revenue trends, the goal is to identify key patterns that can drive strategic improvements.

Beyond data analysis, this project aims to build a fully automated process that streamlines data management and report generation. The system will be designed to automatically upload multiple input files to an Access database, process the data using SQL queries, and generate three types of weekly reports and distribute them using Outlook VBA:

- Excel report and visualizations  
-  PDF report  
- Automated batch of emails

![Animation2](https://github.com/user-attachments/assets/f3222c7a-0e0e-467e-aa99-5e1dd5080427)



All of these tasks will be done through VBA Macros, integrating MS Office tools. The project showcases how to efficiently handle data uploads, automate report generation, and improve decision-making by integrating various data sources into a seamless workflow. With sales data structured across multiple tables within an Access database, this automation will enhance operational efficiency and ensure timely reporting in a highly competitive retail environment.


<img src="https://github.com/user-attachments/assets/da40aaad-61f9-478f-ab67-cefd43d3b626" width="500">

Codes for automation and Macro files can be found in excel here [here](https://github.com/elifyetistiren/Automated-RetailAnalytics-project/blob/main/Master_control_file.xlsm)  

## <ins>Overview</ins>  

This project analyzes key performance indicators (KPIs) to evaluate store performance, revenue trends, and customer purchasing behavior. By assessing daily sales volume, sales value, and customer footfall, we can identify high-performing stores and measure target achievement. The insights gained will support pricing optimization, inventory management, and customer engagement strategies.  

SQL queries for analyses can be found [here](https://github.com/elifyetistiren/Automated-RetailAnalytics-project/blob/main/queries.sql) 

The analysis integrates multiple datasets to explore the following KPI categories:

### Sales Performance KPIs
1. **Total Sales Revenue** – Sum of all sales transactions over a period.  
2. **Sales per Shop** – Revenue generated per shop to assess individual store performance.  
3. **Sales per Region** – Geographical breakdown of revenue performance.  
4. **Discount Impact on Revenue** – Percentage of revenue lost due to applied discounts.  

### Customer Analytics KPIs
1. **Average Customer Spend** – Revenue per unique customer.  

### Target vs. Actual Performance KPIs
1. **Target Achievement Rate** – Actual sales as a percentage of the set performance targets.  
2. **Best/Worst Performing Shops** – Ranking based on target achievement.  

### Profitability KPIs
1. **Return Rate Impact** – Revenue lost due to returned products.  

Based on these findings, the project will provide data-driven recommendations to enhance business performance and decision-making.

## <ins>Technology Stack</ins>  

- **Microsoft Excel (VBA):** Automates data processing, calculations, and reporting.  
- **SQL Server / MySQL:** Stores and retrieves structured data for analysis.  
- **VBA (Visual Basic for Applications):** Orchestrates automation logic for data movement, analytics, and reporting.  
- **Outlook VBA:** Sends automated reports via email.  

## <ins>Installation & Setup </ins> 

### 1. Prerequisites  
Ensure the following are installed and configured:

- Microsoft Excel (Enable Macros & VBA)  
- SQL Server / MySQL (Database for Data Processing)  
- Microsoft Outlook (For Report Distribution)  

### 2. Enable VBA Macros in Excel  
1. Open **Excel → Options → Trust Center → Trust Center Settings**  
2. Enable **"Trust access to the VBA project object model"**  

## <ins>Data Structure</ins>  

The sales database is structured to facilitate comprehensive analysis of transaction data, store performance, and customer interactions. It consists of the following key tables, containing 7,900 rows in total:

![ERD2](https://github.com/user-attachments/assets/13a50c94-a7ea-4d68-bb25-b2166fd3fb78)

- **tbSales** – Stores detailed transaction records, including sales quantity, price, discounts, and total revenue.  
- **tbShops** – Contains information about various store locations, including shop names, locations, and managers.  
- **tbCustomer** – Holds customer-related data, including customer names, relationships, and start dates.  
- **tbPerformance** – Tracks sales targets and performance metrics for different store locations.  

<img src="https://github.com/user-attachments/assets/b67cd57d-6df5-4b04-bdc4-2fe14ee3fc8c" width="500">

## <ins>Data Characteristics and Cleaning Notes</ins>  

During data exploration using Excel and SQL, the following key observations were noted for accurate analysis:

- Returned quantities are recorded as **-1**, allowing direct summation of sales values without additional filtering.  
- Net sales value should be calculated using the formula:  
  `Net Sales = Price × (1 - Discount %) × Quantity`  
- Sales values with **"Reserved"** status should be included along with **"Paid"** sales, as revenue is recognized for these transactions.  
- No duplicate sales records were found, ensuring data integrity.  

This structured data model enables efficient analysis of sales trends, customer behavior, and store performance.

## <ins>Insights and Market Trend Analysis Based on Provided Data</ins>  

<img src="https://github.com/user-attachments/assets/475bb322-27b9-4898-abfb-e942e7ef2c4a" width="1000">



# Sales Performance Analysis

## 1. Sales Performance KPIs: Target vs Actual

### **Top Performers**
- **Extreme Shop** exceeded its target with **113% of target achieved**, generating **$311,655.33** in total sales, the highest of all stores.
- **Strong Performance in Extreme Shops:** Extreme Shop is leading the pack with sales exceeding their target by a large margin. This store has likely **optimized promotions and customer engagement strategies** to outperform expectations.

### **Underperformers**
- **Snow Giant** only achieved **75% of its target**, bringing in **$261,722.46** in sales, falling short significantly.
- **Mountain Heaven** underperformed with **66% of its target**, making only **$279,740.52** in sales, which requires a review of its sales strategy or product offerings.
- **Concerns for Snow Giant and Mountain Heaven:** The underperformance in these stores could be attributed to a mismatch in **optimized pricing, customer preferences, poor product-market fit, or ineffective marketing strategies**. These stores could benefit from **better pricing, targeted promotional campaigns, or product adjustments** to better align with local demand.

![weekly sales by shop](https://github.com/user-attachments/assets/f10d10a1-84ef-4a12-9c1e-d304d6439e38)

---


## 2. Customer Analytics KPIs

### **Average Transaction Value (ATV)**
- **Extreme Shop** reported the highest **ATV of $307.35**, which indicates that customers are **spending more per transaction**, possibly due to premium offerings or upselling strategies. Extreme Shop is likely capitalizing on **upselling, high-value items, and customer loyalty programs** to increase the average spend per customer.
- **Mountain Heaven** recorded the lowest **ATV at $301.12**, which could reflect a focus on **lower-priced items or smaller purchase volumes**. The lower ATV at Mountain Heaven suggests that they may be catering to a more **price-sensitive customer base**. They could explore strategies like **bundling or offering loyalty programs** to increase average customer spend.


---
## 3. Profitability KPIs

### **Return Rates Impact**
- **Snow Giant** suffered the highest impact from returns, with **18% of revenue lost** due to returns. This could signal **product quality issues, wrong customer expectations, or misaligned marketing**.
- **Extreme Shop II** saw a **17% return rate**, indicating potential **satisfaction or quality concerns** that need addressing to maintain profitability.
- Both stores could benefit from **post-purchase engagement, clearer return policies, and better quality control** to reduce this revenue loss.

### **Discount Impact on Revenue**
- **Extreme Shop** lost **47% of its revenue** due to **discounts**, reflecting its aggressive **discounting strategy**. However, this was accompanied by high sales performance, so the strategy likely worked.
- **Mountain Heaven** had the highest **discount impact at 49%**, which, paired with its underperformance, indicates that **discounts may not be yielding desired results**. Over-discounting without a **strong sales foundation** can lead to **diminished profitability**.


<img src="https://github.com/user-attachments/assets/bce0fc79-a078-4fdd-a566-f9c63312662e" width="600">

---

## 4. Weekly Customer Revenue Overview

- **Delta** contributed the highest revenue of **€390,103.94**, showing strong **customer engagement and loyalty**.
- **Alpha** also performed well, generating **€369,956.33** in revenue. Both **Delta and Alpha** bring in significant revenue, showing that **personalized marketing and high-value customer targeting** are effective.  
- **Gamma** contributed the lowest revenue at **€282,550.77**, suggesting that either their **target market or promotional strategies** need refinement. Gamma's lower performance may reflect **ineffective engagement or lack of personalized offers**. The store could benefit from **data-driven customer segmentation and tailored promotions** to enhance its contribution.

---

## 5. Sales Performance by Region

- **Paris:** Leading in sales with **$466,599.19**, showing strong **market demand and successful sales strategies**.
- **Tokyo:** Lowest sales at **$261,722.46**, indicating possible **market challenges**.
- **Paris' High Performance:** Paris is clearly outperforming the other regions. If this is a trend, it suggests that the **sales strategies in Paris** are well-aligned with customer preferences. The store might benefit from **scaling successful tactics in Paris to other regions**.
- **Tokyo's Struggles:** Tokyo’s relatively low performance may indicate that **additional research into local consumer behavior** is needed. The store could potentially **adapt its offerings** to suit local tastes better, **adjust its marketing strategies**, or **increase brand awareness and seasonal promotions** that can boost engagement in these regions.

<img src="https://github.com/user-attachments/assets/73310f53-fb39-4232-851c-a20cde4c1046" width="600">

---

## 6. Sales Count by Hour

- **Peak Sales Hours:** The highest sales counts occur at **8 AM, 11 AM, and 4 PM (16:00)**, suggesting these are the busiest times, likely due to **morning rush, lunch breaks, and late afternoon shopping**.
- **Low Sales Hours:** **12 PM and 3 PM (15:00)** show significantly lower sales, possibly indicating a **midday lull** where fewer customers are making purchases.

---

## Strategic Recommendations

1. **Optimize for High-Performing Stores:** Replicate **Extreme Shop** and **Hidden Rock** strategies across underperforming stores like **Mountain Heaven and Snow Giant**.
2. **Address Return Issues:** Focus on improving **product quality and post-purchase satisfaction** to reduce return rates, especially for **Snow Giant and Mountain Heaven**.
3. **Refine Discount Strategies:** Review the **discount effectiveness** in stores like **Mountain Heaven** and ensure that discounting is paired with **strong customer demand**.
4. **Enhance Customer Loyalty Programs:** **Alpha and Delta** are the top revenue generators—**targeting these customers** with **loyalty programs or exclusive offers** can help maintain high levels of engagement.
5. **Investigate Underperforming Segments and Align with Regional Culture Fit:** **Gamma’s low revenue** suggests a need for **targeted marketing, personalized promotions, and better customer insights**. There might be a **cultural, pricing, or product mismatch** that needs addressing to **boost sales**.
6. **Maximize Sales Efficiency:** Increase **staffing levels during peak hours (8 AM, 11 AM, and 4 PM)** to enhance customer service, while implementing **targeted promotions during low-traffic hours (12 PM & 3 PM)** to boost engagement. Additionally, ensure **high-demand products are well-stocked before peak hours** to meet demand effectively and **reduce missed sales opportunities**.



## Conclusion

The data reveals **strong market performance in some stores**, but **underperformance in others** suggests the need for **adjustments in sales strategies, discount management, and customer engagement practices**. Moving forward, **optimizing product offerings, improving return processes, and focusing on targeted customer strategies during low traffic periods** while understanding **local customer behavior** will likely drive **higher sales and profitability** across all stores.
