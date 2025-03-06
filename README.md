
# Retail Sales Data Analysis and Automation Project

## Project Background

The company operates in the retail industry, selling a diverse range of products through its physical store network. While it has an established market presence, utilizing data-driven insights is essential for optimizing operations and enhancing customer satisfaction.  
To gain a deeper understanding of customer purchasing behavior, revenue patterns, and operational efficiency, this project focuses on analyzing daily sales data across multiple store locations over a one-week period. By leveraging insights from sales transactions, store performance, customer demographics, and revenue trends, the goal is to identify key patterns that can drive strategic improvements.

Beyond data analysis, this project aims to build a fully automated process that streamlines data management and report generation. The system will be designed to automatically upload multiple input files to an Access database, process the data using SQL queries, and generate three types of weekly reports and distribute them using Outlook VBA:

- Excel report and visualizations  
- PDF report  
- Automated batch of emails  

All of these tasks will be done through VBA Macros, integrating MS Office tools. The project showcases how to efficiently handle data uploads, automate report generation, and improve decision-making by integrating various data sources into a seamless workflow. With sales data structured across multiple tables within an Access database, this automation will enhance operational efficiency and ensure timely reporting in a highly competitive retail environment.

## Overview  

This project analyzes key performance indicators (KPIs) to evaluate store performance, revenue trends, and customer purchasing behavior. By assessing daily sales volume, sales value, and customer footfall, we can identify high-performing stores and measure target achievement. The insights gained will support pricing optimization, inventory management, and customer engagement strategies.  

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

## Technology Stack  

- **Microsoft Excel (VBA):** Automates data processing, calculations, and reporting.  
- **SQL Server / MySQL:** Stores and retrieves structured data for analysis.  
- **VBA (Visual Basic for Applications):** Orchestrates automation logic for data movement, analytics, and reporting.  
- **Outlook VBA:** Sends automated reports via email.  

## Installation & Setup  

### 1. Prerequisites  
Ensure the following are installed and configured:

- Microsoft Excel (Enable Macros & VBA)  
- SQL Server / MySQL (Database for Data Processing)  
- Microsoft Outlook (For Report Distribution)  
- Windows Task Scheduler (For Scheduling Automation Runs)  

### 2. Enable VBA Macros in Excel  
1. Open **Excel → Options → Trust Center → Trust Center Settings**  
2. Enable **"Trust access to the VBA project object model"**  

## Data Structure  

The sales database is structured to facilitate comprehensive analysis of transaction data, store performance, and customer interactions. It consists of the following key tables, containing 7,900 rows in total:

- **tbSales** – Stores detailed transaction records, including sales quantity, price, discounts, and total revenue.  
- **tbShops** – Contains information about various store locations, including shop names, locations, and managers.  
- **tbCustomer** – Holds customer-related data, including customer names, relationships, and start dates.  
- **tbPerformance** – Tracks sales targets and performance metrics for different store locations.  

## Data Characteristics and Cleaning Notes  

During data exploration using Excel and SQL, the following key observations were noted for accurate analysis:

- Returned quantities are recorded as **-1**, allowing direct summation of sales values without additional filtering.  
- Net sales value should be calculated using the formula:  
  `Net Sales = Price × (1 - Discount %) × Quantity`  
- Sales values with **"Reserved"** status should be included along with **"Paid"** sales, as revenue is recognized for these transactions.  
- No duplicate sales records were found, ensuring data integrity.  

This structured data model enables efficient analysis of sales trends, customer behavior, and store performance.

## Insights and Market Trend Analysis Based on Provided Data  

### 1. Sales Performance KPIs: Target vs Actual  

#### **Top Performers:**  
- **Extreme Shop** exceeded its target with **113%** of target achieved, generating **$311,655.33** in total sales, the highest of all stores.  
- **Strong Performance in Extreme Shops:** Extreme Shop is leading the pack with sales exceeding their target by a large margin. This store has likely optimized promotions and customer engagement strategies.  

#### **Underperformers:**  
- **Snow Giant** only achieved **75%** of its target, bringing in **$261,722.46** in sales, falling short significantly.  
- **Mountain Heaven** underperformed with **66%** of its target, making only **$279,740.52** in sales, which requires a review of its sales strategy or product offerings.  

### 2. Customer Analytics KPIs  

#### **Average Transaction Value (ATV):**  
- **Extreme Shop** reported the highest **ATV of $307.35**, which indicates that customers are spending more per transaction, possibly due to premium offerings or upselling strategies.  
- **Mountain Heaven** recorded the lowest **ATV at $301.12**, which could reflect a focus on lower-priced items or smaller purchase volumes.  

### 3. Profitability KPIs  

#### **Return Rates Impact:**  
- **Snow Giant** suffered the highest impact from returns, with **18%** of revenue lost due to returns.  
- **Extreme Shop II** saw a **17%** return rate, indicating potential satisfaction or quality concerns.  

#### **Discount Impact on Revenue:**  
- **Extreme Shop** lost **47%** of its revenue due to discounts, reflecting its aggressive discounting strategy.  
- **Mountain Heaven** had the highest discount impact at **49%**, which, paired with its underperformance, indicates that discounts may not be yielding desired results.  

### 4. Weekly Customer Revenue Overview  

- **Delta** contributed the highest revenue of **€390,103.94**, showing strong customer engagement and loyalty.  
- **Gamma** contributed the lowest revenue at **€282,550.77**, suggesting that either their target market or promotional strategies need refinement.  

### 5. Sales Performance by Region  

- **Paris:** Leading in sales with **$466,599.19**, showing strong market demand and successful sales strategies.  
- **Tokyo:** Lowest sales at **$261,722.46**, indicating possible market challenges.  

### 6. Sales Count by Hour  

- **Peak Sales Hours:** The highest sales counts occur at **8 AM, 11 AM, and 4 PM**, suggesting these are the busiest times.  
- **Low Sales Hours:** **12 PM and 3 PM** show significantly lower sales.  

## Strategic Recommendations  

1. **Optimize for High-Performing Stores:** Replicate Extreme Shop strategies across underperforming stores.  
2. **Address Return Issues:** Focus on improving product quality and post-purchase satisfaction to reduce return rates.  
3. **Refine Discount Strategies:** Ensure discounting is paired with strong customer demand.  
4. **Enhance Customer Loyalty Programs:** Target high-value customers with loyalty programs.  
5. **Investigate Underperforming Segments:** Conduct targeted marketing and personalized promotions.  
6. **Adjust Staffing and Promotions:** Increase staffing during peak hours and implement targeted promotions.  

## Conclusion  

The data reveals strong market performance in some stores, but underperformance in others suggests the need for adjustments in sales strategies, discount management, and customer engagement practices. Moving forward, optimizing product offerings, improving return processes, and focusing on targeted customer strategies will drive higher sales and profitability.  