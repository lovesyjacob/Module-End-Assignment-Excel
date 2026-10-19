# Module-End-Assignment-Excel
# Healthcare Data Analysis and Insights
This assignment analyzes a comprehensive healthcare dataset to extract insights into patient health profiles, medical histories, and healthcare costs. The assignment involves data cleaning, transformation, exploration, and the creation of an interactive dashboard using Microsoft Excel.

## Phase 1: Data Cleaning

**Q: Check for the number of missing values marked with '?' in each column of the "Medical Examinations" Table and "Hospitalization Details" Table.**
**A:** Identified the number of missing values across both tables to quantify the data cleaning required. Missing values were found in `smoker` (Medical table) and `year`, `month`, `Hospital tier`, `City tier`, and `State ID` (Hospitalisation table).

* **Exact Function/Tool Used:** `Find All` (`Ctrl + F`) with the search criteria `~?` and Excel **Data Filters**.

**Q: Fill in the missing values of 'month' with Sep and 'year' with its average rounded to the nearest integer.**
**A:** Replaced the missing month values with "Sep". Calculated the average of the existing years to be 1983, and replaced the missing year values with this integer.

* **Exact Function/Tool Used:**
* To find the average year: `=ROUND(AVERAGE(year_range), 0)`
* To replace values: Excel **Data Filters** and standard data entry.

**Q: Determine the most frequently occurring values in the 'smoker', 'Hospital tier' and 'City tier' columns, and fill in the missing values accordingly.**
**A:** Determined the mode (most frequent text value) for the categorical columns. Replaced missing `smoker` values with `No`, and missing `Hospital tier` and `City tier` values with `tier - 2`.

* **Exact Function/Tool Used:** **PivotTables** (placing the column in both Rows and Values to get the count) and Excel **Data Filters** for replacement.

**Q: If any 'State ID' values are missing, consider filling them with 'Unknown' or using another appropriate strategy.**
**A:** Replaced all null instances in the State ID column with a standardized text string to maintain data integrity.

* **Exact Function/Tool Used:** Excel **Data Filters** and Find/Replace.

## Phase 2: Data Transformation

**Q: Split the 'names' column in the "Customer Names" Table into 3 meaningful columns: 'Title', 'First Name', and 'Last Name'.**
**A:** Dynamically parsed the names column by utilizing advanced text manipulation functions to isolate the Last Name by the comma, and split the remaining text into Title and First Name.

* **Exact Function/Tool Used:**
* **Title:** `=LET(x,TRIM(TEXTAFTER(B2,",")),TEXTBEFORE(x," "))`
* **First Name:** `=LET(x,TRIM(TEXTAFTER(B2,",")),TRIM(TEXTAFTER(x," ")))`
* **Last Name:** `=TRIM(TEXTBEFORE(B2,","))`

**Q: Convert the "NumberOfMajorSurgeries" column in the "Medical Examinations" Table to numerical data by replacing non-numeric characters with meaningful numerical values.**
**A:** Converted text strings such as "No major surgery" into numerical zeros to ensure the column could be used for mathematical aggregation.

* **Exact Function/Tool Used:** **Find and Replace** (`Ctrl + H`) to replace "No major surgery" with `0`, changing the column format to Number.

**Q: Check for inconsistencies in the 'Heart Issues' and 'smoker' columns and propose corrective actions if necessary.**
**A:** Standardized all text casing (e.g., converting "yes" or "YES" to "Yes") to ensure uniformity for grouping and PivotTable analysis.

* **Exact Function/Tool Used:** `=PROPER(cell_reference)` in a helper column, followed by Paste as Values.

**Q: Create a new column named "Weight Status" that categorizes BMI into different categories...**
**A:** Grouped continuous BMI numerical data into distinct categorical buckets based on standard health ranges.

* **Exact Function/Tool Used:** `=IF(BMI_cell<18.5, "Underweight", IF(BMI_cell<=24.9, "Normal Weight", IF(BMI_cell<=29.9, "Overweight", "Obesity")))`

**Q: Create a new column named "Diabetes Status" and fill it as per the information given...**
**A:** Grouped continuous HbA1C numerical data into distinct diabetic categories.

* **Exact Function/Tool Used:** `=IF(HbA1c_cell<5.7, "Normal", IF(HbA1c_cell<=6.4, "Prediabetes", "Diabetes"))`

**Q: Merge 'year', 'month' and 'date' columns in the "Hospitalization Details" Table into one column named 'Date of Birth' and format it in 'DD-MMM-YYYY' custom format.**
**A:** Constructed a valid Excel serial date from three separate columns containing year integers, month text strings, and day integers, then applied a custom visual format.

* **Exact Function/Tool Used:** `=DATE(year_cell, MONTH(DATEVALUE(month_cell&" 1")), date_cell)` formatted via **Format Cells > Custom** as `DD-MMM-YYYY`.

**Q: Calculate the 'Age' of each customer based on their 'Date of Birth' and the date of collection of the dataset, which is 8th June 2023.**
**A:** Calculated the exact patient age dynamically based on the fixed dataset collection date.

* **Exact Function/Tool Used:** `=DATEDIF(DOB_cell, DATE(2023,6,8), "Y")`

**Q: Format 'charges' column as currency ($).**
**A:** Applied financial formatting to healthcare billing data.

* **Exact Function/Tool Used:** Excel's **Currency Number Format** ($) via the Home Ribbon.

## 📈 Phase 3: Data Exploration, Analysis & Visualization

**Q: Create a new sheet named "Healthcare", combine all three tables into one, using Customer ID as the common column, utilizing VLOOKUP.**
**A:** Consolidated the relational tables into a unified master flat file by searching for the `Customer ID` primary key and returning the associated data from the target tables.

* **Exact Function/Tool Used:** Used `VLOOKUP` for column population. Example format used: `=VLOOKUP($A2, 'Table Name'!$A:$K, column_index, FALSE)`

**Q: What is the distribution of cancer history among smokers and non-smokers? (Analysis using Pie/Donut Chart)**
**A:** Evaluated the count of patients with a history of cancer, segmented by their smoking habits.

* **Exact Function/Tool Used:** **PivotTable** (Rows: `smoker`, Columns: `Cancer history`, Values: Count of `Customer ID`) and **Pie Chart**.

**Q: How does the total number of major surgeries and average HbA1C differ between patients with and without a history of transplants?**
**A:** Aggregated the total volume of surgeries and the average blood sugar levels, grouped by transplant history.

* **Exact Function/Tool Used:** **PivotTable** (Rows: `Any Transplants`, Values: Sum of `NumberOfMajorSurgeries` and Average of `HBA1C`) and **Donut Chart**.

**Q: How do healthcare charges vary based on different weight statuses and diabetes statuses? (Analysis using Column/Bar Chart)**
**A:** Analyzed the average healthcare charges intersecting `Weight Status` and `Diabetes Status`.

* **Exact Function/Tool Used:** **PivotTable** (Rows: `Weight Status`, Columns: `Diabetes Status`, Values: Average of `charges`) and **Clustered Column Chart**.

**Q: Can you compare the average charges for each hospital tier within different states?**
**A:** Compared the average financial charges across different `State IDs`, segmented by `Hospital tier`.

* **Exact Function/Tool Used:** **PivotTable** (Rows: `State ID`, Columns: `Hospital tier`, Values: Average of `charges`) and **Clustered Bar Chart**.

**Q: Is there any correlation between age and both BMI and HbA1C in the dataset? (Analysis using Line/Scatter Plot)**
**A:** Evaluated continuous variables against patient age to visualize linear trends.

* **Exact Function/Tool Used:** **Scatter Plot** mapping raw `Age` vs. `BMI` and `Age` vs. `HBA1C`, with a **Linear Trendline** applied.

**Q: Explore the relationship between age and healthcare charges.**
**A:** Visualized the trajectory of healthcare costs as patients age to determine financial correlation.

* **Exact Function/Tool Used:** **Scatter Plot** mapping raw `Age` vs. `charges`, with a **Linear Trendline** applied.

**Q: Dashboard Creation: Build an interactive dashboard that consolidates all key insights... Add slicers for "Weight Status" and "Diabetes Status"...**
**A:** Consolidated all visual charts into an interactive dashboard allowing dynamic, real-time filtering to compare health outcomes and financial charges simultaneously.

* **Exact Function/Tool Used:** **Slicer** for `Weight Status` and `Diabetes Status`, utilizing **Report Connections** to link all dashboard PivotTables.

**Q: Add slicers for the fields "Weight Status" and "Diabetes Status" to enable filtering across all visualizations, supporting comparison of health outcomes and charges based on body weight and diabetes condition.**
**A:** Inserted interactive slicers for both fields on the Dashboard. Configured the "Report Connections" settings for each slicer to link them to all PivotTables simultaneously. This enables dynamic, global filtering across the dashboard, allowing stakeholders to seamlessly compare health outcomes and financial charges across different body weight and diabetic profiles.Exact Function/Tool Used: Insert > Slicer (selecting Weight Status and Diabetes Status), followed by configuring Report Connections (checking all PivotTables).
