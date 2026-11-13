🏥 Healthcare Data Analysis Using Microsoft Excel

1. Project Overview

This project involves cleaning, transforming, analysing and visualising healthcare data using Microsoft Excel.

The objective was to work with multiple healthcare datasets, prepare them for analysis, combine relevant information and develop an interactive dashboard to identify insights into patient health conditions, demographics and hospitalisation expenses.

This project was completed as part of a practical Data Analytics assignment.

⸻

2. Tools & Technologies Used

* Microsoft Excel
* VLOOKUP
* Excel Formulas
* Data Cleaning
* Data Transformation
* PivotTables
* PivotCharts
* Excel Slicers
* Data Visualisation

⸻

3. Assignment Questions & Implementation

QUESTION 1 – DATA CLEANING

Objective

Identify missing values and inconsistencies in the healthcare datasets and prepare the data for further analysis.

Tasks Performed

1. Missing Value Identification

* Identified missing values represented by ? in the dataset.
* Calculated the number of missing values in each column.
* Created a Missing Value Audit worksheet to document the findings.

2. Handling Missing Values

* Replaced missing Month values with September (Sep).
* Replaced missing Year values with the rounded average Year.
* Replaced missing State ID values with Unknown.

3. Identifying Most Frequent Categories

* Determined the most frequently occurring Smoker category.
* Identified the most frequently occurring Hospital Tier.
* Identified the most frequently occurring City Tier.

Steps Followed in Excel

1. Inspected the original datasets.
2. Identified missing values using filtering and Excel functions.
3. Created a separate Missing Value Audit sheet.
4. Calculated missing-value counts.
5. Applied the required replacement values.
6. Verified the cleaned data.

Outcome: Prepared cleaner datasets with missing values handled according to the assignment instructions.

⸻

QUESTION 2 – DATA TRANSFORMATION

Objective

Transform the cleaned healthcare data into a structured format suitable for analysis.

Tasks Performed

1. Splitting Customer Names

* Separated customer names into three individual columns:
    * Title
    * First Name
    * Last Name

2. Data Type Conversion

* Converted NumberOfMajorSurgeries into a numerical data type.

3. Data Consistency Checks

* Checked inconsistencies involving:
    * Heart Issues
    * Smoker status

4. Creating Weight Status

* Created a new column named Weight Status using BMI values.
* Categorised patients according to their weight status.

5. Creating Diabetes Status

* Created a new column named Diabetes Status using HBA1C values.
* Categorised patients based on their diabetes status.

6. Creating Date of Birth

* Combined Year, Month and Date information into a single Date of Birth column.

7. Calculating Age

* Calculated patient age using 8 June 2023 as the reference date.

8. Formatting Healthcare Charges

* Formatted the Charges column as currency.

Steps Followed in Excel

1. Cleaned and separated customer name fields.
2. Converted numerical columns into the appropriate data types.
3. Checked data consistency.
4. Created Weight Status and Diabetes Status categories.
5. Combined date-related columns.
6. Calculated patient age.
7. Applied currency formatting.
8. Verified the transformed dataset.

Outcome: Created structured healthcare datasets ready for data integration and analysis.

⸻

4. DATA INTEGRATION USING VLOOKUP

After completing the cleaning and transformation tasks, relevant datasets were combined using Excel’s VLOOKUP function.

Implementation Steps

1. Identified the common Customer ID across the datasets.
2. Used VLOOKUP to retrieve the required patient information.
3. Integrated medical examination details with hospitalisation information.
4. Created a consolidated worksheet named VLOOKUP - Corrected.
5. Verified the retrieved values and checked the consolidated dataset.

Outcome: Created a combined healthcare dataset containing patient information, medical examination details and hospitalisation-related information.

⸻

5. DATA ANALYSIS USING PIVOTTABLES

After consolidating the data, PivotTables were created to analyse different healthcare metrics.

Analysis 1 – Weight Status Analysis

Objective

Understand patient distribution across different weight categories and compare their healthcare expenses.

Fields Analysed

* Weight Status
* Customer ID
* BMI
* HBA1C
* Charges

Excel Steps

1. Created a PivotTable using the consolidated healthcare dataset.
2. Added Weight Status to Rows.
3. Added Customer ID to Values and selected Count.
4. Added BMI to Values and selected Average.
5. Added HBA1C to Values and selected Average.
6. Added Charges to Values and selected Average.

⸻

Analysis 2 – Diabetes Status Analysis

Objective

Analyse patient distribution according to diabetes status and compare relevant health indicators.

Fields Analysed

* Diabetes Status
* Customer ID
* BMI
* HBA1C
* Charges

Excel Steps

1. Created a separate PivotTable.
2. Added Diabetes Status to Rows.
3. Added Customer ID to Values and selected Count.
4. Added BMI to Values and selected Average.
5. Added HBA1C to Values and selected Average.
6. Added Charges to Values and selected Average.

⸻

Analysis 3 – Hospital Tier and State Analysis

Objective

Compare average hospitalisation charges across different hospital tiers within different states.

Fields Analysed

* State ID
* Hospital Tier
* Charges

Excel Steps

1. Created a separate PivotTable.
2. Added State ID to Rows.
3. Added Hospital Tier to Columns.
4. Added Charges to Values.
5. Changed the calculation from Sum to Average.

Outcome: Created a state-wise comparison of average charges across Tier 1, Tier 2 and Tier 3 hospitals.

⸻

Analysis 4 – Cancer History Among Smokers

Objective

Analyse the distribution of cancer history among patients identified as smokers.

Fields Analysed

* Cancer History
* Customer ID
* Smoker

Excel Steps

1. Created a separate PivotTable.
2. Added Cancer History to Rows.
3. Added Customer ID to Values and selected Count.
4. Added Smoker to Filters.
5. Filtered the data to include smokers only.
6. Created a Doughnut Chart to visualise the results.
7. Added percentage data labels.

Outcome: Visualised the proportion of smokers with and without a recorded cancer history.

⸻

6. INTERACTIVE DASHBOARD CREATION

Objective

Develop an interactive Excel dashboard to present key healthcare metrics and analysis results in a single visual interface.

Step 1 – Dashboard Layout

Created a separate worksheet named Dashboard.

Added the dashboard title:

Healthcare Data Analysis Dashboard

Organised the dashboard into KPI cards and visualisation sections.

Step 2 – Key Performance Indicators (KPIs)

Created five KPI cards using Excel formulas.

KPI	Description
Total Patients	Total number of patient records
Average Age	Average patient age
Average BMI	Average Body Mass Index
Average HbA1c	Average HbA1c level
Average Charges	Average healthcare charges

Step 3 – Dashboard Charts

Created the following visualisations:

1. Average Healthcare Charges by Weight Status

* Compared average healthcare charges across weight categories.
* Used a Column Chart.

2. Average Healthcare Charges by Diabetes Status

* Compared average healthcare charges across diabetes categories.
* Used a Column Chart.

3. Cancer History Among Smokers

* Compared the proportion of smokers with and without cancer history.
* Used a Doughnut Chart with percentage labels.

Step 4 – Interactive Slicers

Created Excel slicers to make the dashboard interactive.

Slicer 1: Weight Status

* Normal Weight
* Overweight
* Obesity
* Underweight

Slicer 2: Diabetes Status

* Normal
* Prediabetes
* Diabetes

Connected the slicers to the relevant PivotTables and PivotCharts to enable interactive filtering.

Step 5 – Dashboard Validation

* Tested the Weight Status slicer.
* Tested the Diabetes Status slicer.
* Verified that the connected PivotCharts responded to filter selections.
* Reviewed the dashboard layout and visualisations.

⸻

7. PROJECT WORKBOOK STRUCTURE

The final Excel workbook contains worksheets for:

Worksheet	Description
Q1-Q2 Instructions	Assignment instructions
Missing Value Audit	Missing-value identification
Customer Names - Clean	Cleaned customer name information
Medical Examinations - Clean	Cleaned medical examination data
Hospitalisation - Clean	Cleaned hospitalisation data
VLOOKUP - Corrected	Consolidated healthcare dataset
Weight Analysis	Weight-related PivotTable analysis
Diabetes Analysis	Diabetes-related PivotTable analysis
Hospital Analysis	Hospital tier and state comparison
Cancer Analysis	Cancer history among smokers
Dashboard	Interactive healthcare dashboard

⸻

8. KEY LEARNING OUTCOMES

Through this project, I gained practical experience in:

* Identifying and handling missing data.
* Cleaning and transforming raw datasets.
* Working with multiple Excel worksheets.
* Applying VLOOKUP for data integration.
* Using Excel formulas for calculations.
* Creating calculated categories from healthcare indicators.
* Analysing data using PivotTables.
* Building PivotCharts.
* Creating interactive dashboards using Excel slicers.
* Presenting analytical findings through data visualisation.

⸻

9. PROJECT FILES

The repository includes:

* README.md – Project documentation.
* Healthcare_Analysis_Dashboard.xlsx – Completed Excel assignment workbook.
* Dashboard_Screenshot.png – Dashboard preview.

⸻

10. Conclusion

This project demonstrates the practical application of Microsoft Excel in the data analytics workflow, starting from raw data cleaning and transformation to data integration, analysis and interactive dashboard development.

It provided hands-on experience in converting raw healthcare information into structured datasets and presenting meaningful analytical insights through visualisations.

Project Category: Healthcare Data Analytics
Tool: Microsoft Excel
Project Type: Practical Data Analytics Assignment
