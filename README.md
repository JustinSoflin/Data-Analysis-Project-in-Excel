# Data Analysis Project in Excel
## Data Cleansing, Analysis, and Visualization <br>

<img width="1000" height="525" alt="image" src="https://github.com/user-attachments/assets/14763d4b-cff9-4473-80ab-c5b17b6f6480" /> <br>


In this project, I utilize Microsoft Excel to perform a complete end-to-end data analysis workflow — from cleaning raw data to conducting descriptive statistical analysis, building PivotTables, and developing interactive visualizations. The goal is to transform an unstructured dataset into meaningful insights that can guide data-driven decision-making.

---

## 📑 Table of Contents

[Step 1: Data Cleaning](#step-1-data-cleaning) | 
[Step 2: Descriptive Statistics](#step-2-descriptive-statistics) | 
[Step 3: PivotTable Analysis](#step-3-pivottable-analysis) | 
[Step 4: Interactive Visualization](#step-4-interactive-visualization) | 
[Final Outcome](#-final-outcome)

---

## 📊 Tools Utilized 

- `Microsoft Excel`: _used for data cleaning, transformation, and preparation for analysis_
- `PivotTables`: _enabled dynamic aggregation and summarization of key metrics_
- `Descriptive Statistics`: _provided summary metrics including mean, median, mode, standard deviation, and range_
- `Slicers`: _allowed interactive filtering of PivotTables and connected charts_
- `Charts & Dashboards`: _visualized data trends and created interactive dashboards linked to PivotTables_
- `Sample Dataset`: _used for demonstrating Excel’s analytical capabilities_


---

## Step 1: Data Cleaning

Raw data often contains inconsistencies, duplicates, and missing information that can interfere with analysis. To prepare the dataset, I first created a backup copy and then removed any duplicate rows using Data → Remove Duplicates.

<Br> <img width="1280" height="800" alt="image" src="https://github.com/user-attachments/assets/3a09278e-bd13-4b8c-ba84-095ebb445bd6" />


Next, I standardized categorical fields in the Gender column by converting abbreviated values into full text. This was accomplished using the following Excel shortcuts and functions:

`Ctrl + Shift + ↓` → to select the entire column

`Ctrl + H` → to open the Find and Replace tool

Find "M" → Replace with "Male"

Find "F" → Replace with "Female"

<Br> <img width="1280" height="800" alt="image" src="https://github.com/user-attachments/assets/3170842a-08c5-40b6-ab32-6eebb17bc8d9" />


The dataset included columns for birth year, month, and day, but no calculated age. To address this, I created an Age column by combining these values using the following Excel functions:

- `=DATE(year, month, day)` → to generate full birthdates

- `=TODAY()` → to return the current date

- `=YEARFRAC(birthdate, TODAY())` → to calculate age in years

<Br>

<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/90bdae5d-1006-4c8c-8710-bd1a5b691413" />

<Br> 

<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/3b24dbaf-698a-4ad4-837c-89527d7fc48c" /> 

<br>

<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/d4afb854-b539-4d05-a0ad-7e07de6c9982" />

<br>

## Step 2: Descriptive Statistics

Once the dataset was normalized, I used Excel’s Data Analysis ToolPak to generate descriptive statistics. This summary provided quick insights into central tendency, variability, and data distribution — including mean, median, mode, standard deviation, and range.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/4fa73a8f-1f27-47ab-afe7-986e168fe993" />

<br>

<img width="1279" height="348" alt="image" src="https://github.com/user-attachments/assets/aa10c0b1-40d5-4da7-a0fd-239fdfd44ab2" />

<br>

## Step 3: PivotTable Analysis

Next, I leveraged PivotTables to summarize key financial and demographic attributes in the dataset. This allowed for dynamic grouping, filtering, and aggregation of data.

The first PivotTable ranks the top 10 wealthiest individuals by total net worth.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/c8f32da1-e236-436f-acf2-459e602a17cd" />

The second PivotTable analyzes age distribution across wealth categories, revealing which age groups are most frequently represented among the wealthiest individuals.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/5683f4a0-c4af-4ca8-8fe3-0249856a4255" />

<br>

## Step 4: Interactive Visualization

To make the data exploration process more intuitive, I incorporated Slicers — allowing for real-time filtering by specific categories such as gender or category. These filters dynamically update all connected PivotTables simultaneously.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/409328e6-68bc-44c0-9e96-483633ef099d" />

Finally, I created interactive charts connected to the PivotTables. These visualizations automatically adjust in response to Slicer selections, enabling a more dynamic and exploratory analysis experience.

<br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/b309679e-df20-48d3-aa3c-d7a1a82a9a1c" /> <Br>

<Br>

## ✅ Final Outcome

By the end of the project, the dataset was fully transformed into a clean, analyzable format supported by descriptive metrics, PivotTables, and interactive dashboards — demonstrating how Excel can serve as both a data preparation and visualization tool.

