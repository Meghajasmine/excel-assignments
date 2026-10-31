Data Exploration Using Excel

## 📌 Project Overview
As an aspiring Data Analyst, I conducted this exploratory data analysis (EDA) and data transformation project to demonstrate core spreadsheet analytical competencies. The project analyzes a retail inventory catalog of **34 items**, focusing on key price distributions, category-level conditional totals, business logic classification, and text manipulation on structured product codes.

---

## 🎯 Analytical Objectives
- **Data Summarization:** Calculate catalog aggregates including total value, unit count, and average item price.
- **Extreme Value Identification:** Pinpoint upper and lower cost boundaries using boundary functions.
- **Categorical Logic:** Implement conditional logic (`IF`) to segment catalog offerings into high-tier vs. standard-tier goods.
- **Conditional Analytics:** Measure specific department value and volume with criteria-based aggregation (`SUMIF`, `COUNTIF`).
- **Feature Extraction:** Dissect composite string identifiers into dedicated temporal and geopolitical attributes using string extraction formulas (`LEFT`, `RIGHT`, `MID`).
- **Table Structuring:** Organize raw tabular data into an active Excel Table structure (`Sales_DB`) for dynamic structured references.

---

## 🛠️ Toolset & Excel Capabilities
- **Software:** Microsoft Excel
- **Table Name:** `Sales_DB`
- **Core Functions Utilized:** `SUM`, `COUNTA`, `AVERAGE`, `MIN`, `MAX`, `IF`, `SUMIF`, `COUNTIF`, `LEFT`, `RIGHT`, `MID`

---

## 📂 Repository Organization

```text
├── README.md                              # Comprehensive project documentation
├── Excel_Data_Exploration_Raw.xlsx        # Unprocessed source dataset
├── Excel_Data_Exploration_Solution.xlsx   # Completed analytical workbook with formulas
└── Work_Description_and_Project_Title.pdf # 1-page executive summary & verification document
```

---

## 📋 Methodology & Applied Formulas

### 1. Catalog Summary Statistics
Applied on the **`Assignment`** summary sheet:

| Metric | Formula (`Sales_DB` Table) | Equivalent Range Formula | Output |
| :--- | :--- | :--- | :--- |
| **Total Price** | `=SUM(Sales_DB[Price ($)])` | `=SUM(D2:D35)` | **$10,100** |
| **Product Count** | `=COUNTA(Sales_DB[Product ID])` | `=COUNTA(A2:A35)` | **34 items** |
| **Average Price** | `=AVERAGE(Sales_DB[Price ($)])` | `=AVERAGE(D2:D35)` | **$297.06** |

---

### 2. Extreme Price Outliers
Determined minimum and maximum inventory prices:

- **Minimum Price:** `=MIN(Sales_DB[Price ($)])` $\rightarrow$ **$30** *(Adidas T-shirt)*
- **Maximum Price:** `=MAX(Sales_DB[Price ($)])` $\rightarrow$ **$1,000** *(Dell Laptop)*

---

### 3. Business Logic Implementation (`IF`)
On the **`Copy_of_Dataset`** sheet, added a calculated column **`Price Range`** (Column G):
- **Rule:** If unit price is $\$500$ or higher, label as `"High Price"`; otherwise `"Standard Price"`.

**Formula Applied:**
```excel
=IF([@[Price ($)]]>=500, "High Price", "Standard Price")
```

---

### 4. Segmented Aggregations (`SUMIF` & `COUNTIF`)
Extracted segment-specific answers for management reporting:

- **Electronics Inventory Subtotal:**
  ```excel
  =SUMIF(Sales_DB[Category], "Electronics", Sales_DB[Price ($)])
  ```
  *Result:* **$8,050** (represents ~79.7% of total price valuation)

- **Volume of Budget Products (< $100):**
  ```excel
  =COUNTIF(Sales_DB[Price ($)], "<100")
  ```
  *Result:* **11 items**

---

### 5. String Extraction & Column Engineering
The raw `Product ID` follows a compound format: `DD-MMM-CC` (e.g., `28-JAN-US`).

Three new dimension columns were created:
1. **Day** (Column H): First 2 characters
   ```excel
   =LEFT([@[Product ID]], 2)
   ```
2. **Country Code** (Column I): Last 2 characters
   ```excel
   =RIGHT([@[Product ID]], 2)
   ```
3. **Month** (Column J): 3 characters starting at index 4
   ```excel
   =MID([@[Product ID]], 4, 3)
   ```

---

## 💡 Key Analytical Findings
1. **Catalog Valuation:** The 34 catalog listings hold a cumulative unit value of **$10,100**, yielding a mean price of **$297.06**.
2. **Category Weight:** Electronics heavily skews inventory value, contributing **$8,050** across laptops, cameras, and premium devices.
3. **Price Tier Distribution:** Exactly **11 items** are priced below $100, providing an accessible low-ticket consumer tier, while top-bracket devices sit cleanly at or above $500.

---

## ✍️ Author Meghajasmine
