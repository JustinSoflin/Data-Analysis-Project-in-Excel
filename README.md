# Data Analysis Project in Excel
## Data Cleaning, Analysis, and Visualization <br>

<img width="1000" height="525" alt="image" src="https://github.com/user-attachments/assets/14763d4b-cff9-4473-80ab-c5b17b6f6480" /> <br>


In this project, I utilize Microsoft Excel to perform a complete end-to-end data analysis workflow — from cleaning raw data to conducting descriptive statistical analysis, building PivotTables, and developing interactive visualizations. The goal is to transform an unstructured dataset into meaningful insights that can guide data-driven decision-making.

---

## 📑 Table of Contents



---

## 📊 Tools Utilized

- `Microsoft Excel`: _data cleaning, transformation, and analysis_
- `Sample Dataset`: _used for demonstrating Excel’s analytical capabilities_

---

## Step 1: Data Cleaning

Raw data often contains inconsistencies, duplicates, and missing information that can interfere with analysis. To prepare the dataset, I first created a backup copy and then removed any duplicate rows using Data → Remove Duplicates.

<Br> <img width="1280" height="800" alt="image" src="https://github.com/user-attachments/assets/a13f8006-d4df-4843-8f43-8f0527a77ba8" />

Next, I standardized categorical fields — for example, converting abbreviated values in the Gender column from M and F to “Male” and “Female.” This was accomplished using Ctrl + Shift + ↓ to select the column and Ctrl + H (Find and Replace) to update the values.

<Br> <img width="1280" height="800" alt="image" src="https://github.com/user-attachments/assets/320d3ada-79d7-4f93-bc5a-d23fbe46cf82" />

The dataset included columns for birth year, month, and day, but no calculated age. To address this, I created an Age column by combining these values using the following Excel functions:

- `=DATE(year, month, day)` → to generate full birthdates

- `=TODAY()` → to return the current date

- `=YEARFRAC(birthdate, TODAY())` → to calculate age in years

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/d4afb854-b539-4d05-a0ad-7e07de6c9982" /> <Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/3b24dbaf-698a-4ad4-837c-89527d7fc48c" /> <Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/90bdae5d-1006-4c8c-8710-bd1a5b691413" />

<br>

## Step 2: Descriptive Statistics

Once the dataset was normalized, I used Excel’s Data Analysis ToolPak to generate descriptive statistics. This summary provided quick insights into central tendency, variability, and data distribution — including mean, median, mode, standard deviation, and range.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/4fa73a8f-1f27-47ab-afe7-986e168fe993" />

<br>

## Step 3: PivotTable Analysis

Next, I leveraged PivotTables to summarize key financial and demographic attributes in the dataset. This allowed for dynamic grouping, filtering, and aggregation of data.

The first PivotTable ranks the top 10 wealthiest individuals by total net worth.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/c8f32da1-e236-436f-acf2-459e602a17cd" />

The second PivotTable analyzes age distribution across wealth categories, revealing which age groups are most frequently represented among the wealthiest individuals.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/5683f4a0-c4af-4ca8-8fe3-0249856a4255" />

<br>

## Step 4: Interactive Visualization

To make the data exploration process more intuitive, I incorporated Slicers — allowing for real-time filtering by specific categories such as gender or region. These filters dynamically update all connected PivotTables simultaneously.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/409328e6-68bc-44c0-9e96-483633ef099d" />

Finally, I created interactive charts connected to the PivotTables. These visualizations automatically adjust in response to Slicer selections, enabling a more dynamic and exploratory analysis experience.

<Br> <img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/409328e6-68bc-44c0-9e96-483633ef099d" />

<Br>

## ✅ Final Outcome

<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/b309679e-df20-48d3-aa3c-d7a1a82a9a1c" />

By the end of the project, the dataset was fully transformed into a clean, analyzable format supported by descriptive metrics, PivotTables, and interactive dashboards — demonstrating how Excel can serve as both a data preparation and visualization tool.













# Data-Analysis-Project-in-Excel

tools used: 
excel _excel desc_
sample excel data file

1) data cleaning
data can be messy. start by 'cleansing' to make a workable format. after making a backup of the data, ensure there are no duplicate rows

<img width="1280" height="800" alt="image" src="https://github.com/user-attachments/assets/a13f8006-d4df-4843-8f43-8f0527a77ba8" />

changing the gender tab, M to Male and F to Female. CRTL+SHIFT+down arrow to highlight the entire column, CRT+H to open find and replace tool to rename the fields.

<img width="1280" height="800" alt="image" src="https://github.com/user-attachments/assets/320d3ada-79d7-4f93-bc5a-d23fbe46cf82" />

the data set includes birth year, month, and day, but nothing about age. i scrolled over to W to create an Age column and used the excel function =DATE() and select the birth year, month, and day columns to create a birthday. in the next column, i use the =TODAY() function, and in the final column, =YEARFRAC() along with the newly created today date and birthday to create and age.


<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/d4afb854-b539-4d05-a0ad-7e07de6c9982" />


<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/3b24dbaf-698a-4ad4-837c-89527d7fc48c" />



<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/90bdae5d-1006-4c8c-8710-bd1a5b691413" />


2) Data Analysis
data analysis and descriptive statistics tools

<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/4fa73a8f-1f27-47ab-afe7-986e168fe993" />

pivot table to filter data to show the names and final worth of the top 10 wealthiest people


<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/c8f32da1-e236-436f-acf2-459e602a17cd" />

using pivot tables again, i include the ranges and counts of age to show which age groups are most likely to be the wealthiest

<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/5683f4a0-c4af-4ca8-8fe3-0249856a4255" />


3) Data Visualization
inserting slicers and choosing some categories allows us to filter the pivot tables in real time. ensure the tables are properly connected.


<img width="1280" height="764" alt="image" src="https://github.com/user-attachments/assets/409328e6-68bc-44c0-9e96-483633ef099d" />

finally, i add graphs for the final step of visualization, which are still interactive and connected to the pivot tables and respond to category selections.

   
