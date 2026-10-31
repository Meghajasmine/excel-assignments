# excel-assignments
Excel practice exercises and course assignments covering formulas, data analysis, functions, and reporting
📌 Project Overview

This repository contains my submission for Excel Assignment 1 - Data Exploration. In this project, I performed exploratory data analysis (EDA), data cleaning, attribute engineering, and structured statistical calculations on a retail product dataset containing 34 product records.

The primary objective is to demonstrate proficiency in core spreadsheet functionality, conditional logic, aggregated analysis, and text extraction using dynamic Excel functions and table referencing.

🎯 Key Objectives

Organize unstructured data into a structured Excel Table (Sales_DB).

Compute core descriptive statistics: Total Price, Product Count, and Average Price.

Determine catalog boundary values using MIN and MAX.

Implement business decision logic using the IF function to classify items by price tier.

Execute conditional aggregations using SUMIF and COUNTIF.

Extract granular attributes (Day, Month, Country Code) from composite Product IDs using text formulas (LEFT, RIGHT, MID).

🛠️ Tools & Technologies Used

Application: Microsoft Excel

Dataset Size: 34 Records, 10 Attributes (6 original + 4 engineered)

Table Name: Sales_DB

Core Functions: SUM, COUNT, COUNTA, AVERAGE, MIN, MAX, IF, SUMIF, COUNTIF, LEFT, RIGHT, MID

📂 Repository Structure

├── README.md                              # Complete project overview & documentation
├── Excel_Data_Exploration_Raw.xlsx        # Unprocessed source dataset
├── Excel_Data_Exploration_Solution.xlsx   # Solved workbook with Sales_DB table & Assignment sheet
└── Work_Description_and_Project_Title.pdf # 1-page executive summary & verification document



🗂️ Workbook Organization

The solution workbook contains three structured sheets:

Dataset: The untouched, raw data records preserving original values.

Copy_of_Dataset: Converted into an official Excel Table named Sales_DB with 4 newly engineered columns (G through J).

Assignment: The dedicated evaluation sheet containing all calculation tasks referencing the table dynamically.

⚙️ Methodology & Applied Formulas

1. Data Transformation & Attribute Engineering (Copy_of_Dataset Sheet)

The Product ID follows a standardized code format: DD-MMM-CC (e.g., 28-JAN-US). Four new calculated fields were added:

Column

Header

Applied Excel Formula

Purpose / Output Description

G

Price Range

=IF([@[Price ($)]]>=500, "High Price", "Standard Price")

Classifies items $\ge \$500$ as "High Price", else "Standard Price"

H

Day

=LEFT([@[Product ID]], 2)

Extracts the first 2 characters representing the day code

I

Country Code

=RIGHT([@[Product ID]], 2)

Extracts the last 2 characters representing the country

J

Month

=MID([@[Product ID]], 4, 3)

Extracts 3 characters starting at position 4 (month abbreviation)

2. Analytical Calculations & Business Questions (Assignment Sheet)

All answers are dynamically linked to the Sales_DB table in Column K:

S.No

Analytical Question

Table Formula

Standard Range Formula

Result

1

Total price of all products

=SUM(Sales_DB[Price ($)])

=SUM(D2:D35)

$10,100

2

Total number of products

=COUNTA(Sales_DB[Product ID])

=COUNTA(A2:A35)

34

3

Average price of the products

=AVERAGE(Sales_DB[Price ($)])

=AVERAGE(D2:D35)

$297.06

4

Minimum product price

=MIN(Sales_DB[Price ($)])

=MIN(D2:D35)

$30 (Adidas T-shirt)

5

Maximum product price

=MAX(Sales_DB[Price ($)])

=MAX(D2:D35)

$1,000 (Dell Laptop)

6

Total price of Electronics category

=SUMIF(Sales_DB[Category], "Electronics", Sales_DB[Price ($)])

=SUMIF(F2:F35, "Electronics", D2:D35)

$8,050

7

Count of products priced below $100

=COUNTIF(Sales_DB[Price ($)], "<100")

=COUNTIF(D2:D35, "<100")

11

📈 Key Insights & Findings

Revenue Concentration: Products in the Electronics category account for $8,050 out of the $10,100 total catalog valuation (~79.7%), driven by high-ticket items like laptops, smartphones, and cameras.

Price Stratification:

11 items (32.4%) sit below the $100 threshold, representing accessible consumer goods (apparel, kitchenware).

High-tier items exceeding the $500 threshold clearly separate premium hardware from everyday inventory.
