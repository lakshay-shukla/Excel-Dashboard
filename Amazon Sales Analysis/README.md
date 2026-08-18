# 📊 Amazon Product Scraped Data Analysis & Cleaning (Excel Project)

A comprehensive data cleaning and exploratory data analysis (EDA) project based on raw, web-scraped Amazon product data. This project focuses on transforming messy unstructured text into clean, structured metrics ready for business intelligence dashboards, using strictly **Microsoft Excel**.

---

## 🛠️ Project Overview & Objectives
Raw web-scraped data often contains inconsistent formats, trailing invisible spaces, and mixed data types. The goal of this project was to:
1. **Clean and standardize** raw columns (pricing, badges, delivery, coupons).
2. **Engineer features** into standard binary/categorical formats (`Yes`/`No`) for better analysis.
3. **Build interactive Excel Pivot Tables and visual dashboards** to extract actionable business insights without using any external coding languages.

*Note: The raw dataset used for this project was sourced from Kaggle.*

---

## 🧹 Data Cleaning & Feature Engineering Steps

Here is a breakdown of the major data transformations performed using Excel formulas and tools:

| Original Column | Issue / Mess in Raw Data | Excel Cleaning Technique & Formulas Used | Final Format |
| :--- | :--- | :--- | :--- |
| **Price on Variant** | Included wildcard text, `$` symbols, commas, and hidden spaces. | Used `Find & Replace` (`*$`), comma removal, and `=IFERROR(VALUE(H2),"")` | Clean Numeric |
| **Is Best Seller** | Mixed with random promotional badges and text ("Save 10%", "Ends in"). | Conditional check using `=IF(J2="Best Seller", "Yes", "No")` | Binary (`Yes`/`No`) |
| **Is Couponed** | Mixed percentages (`%`), dollar values (`$`), and "No Coupon". | String matching using `=IF(LEFT(K2, 4)="Save", "Yes", "No")` | Binary (`Yes`/`No`) |
| **Buy Box Availability** | Contained "add to cart" and "no" states. | Standardized using `=IF(L2="add to cart", "Yes", "No")` | Binary (`Yes`/`No`) |
| **Sustainability Badge** | Contained zero, certifications, and irrelevant tags (e.g., Alexa). | Logical check using `=IF(OR(M2=0, M2=""), "No", "Yes")` | Binary (`Yes`/`No`) |
| **Delivery Details** | Mixed delivery pricing, dates, and availability strings. | Multi-column splitting using `SEARCH`, `ISNUMBER`, and conditional filters. | Structured Metrics |

---

## 📈 Dashboard & Pivot Table Insights
The cleaned dataset was analyzed using **Excel Pivot Tables** and visualized via custom Clustered and Combo charts. Key insights extracted include:
* **Logistics Matrix:** Evaluating the relationship between free delivery availability and product performance.
* **Coupon Impact Analysis:** Tracking how promotional incentives affect product visibility and sales.
* **Interactive Slicers:** Dynamic filtering implemented for seamless dashboard navigation (filtering by Best Seller, Sponsored, etc.).

### 🖼️ Dashboard Preview
![Amazon Sales Dashboard](Amazon%20Sales%20Dashboard.png)

---

## 🚀 Tools & Technologies Used
This project was executed entirely using **Microsoft Excel**, showcasing advanced data manipulation capabilities without relying on programming languages.
* **Data Cleaning & Manipulation:** Advanced Formulas (`IF`, `VALUE`, `ISNUMBER`, `SEARCH`, `OR`, `LEFT`), Data formatting.
* **Data Modeling:** Excel Tables and dynamically connected Data Sources.
* **Data Visualization:** Pivot Tables, Clustered Column Charts, Combo Charts (with Secondary Axis), Custom Slicers, and Dark Theme UI/UX.

---

## 📂 Repository Structure
```text
├── data/
│   ├── raw_amazon_data.csv          # Original messy scraped dataset
│   └── cleaned_amazon_data.xlsx     # Final processed dataset with pivot tables
├── dashboards/
│   └── Amazon_Sales_Dashboard.png        # Visual snapshot of the Excel dashboard
└── README.md                        # Project documentation
