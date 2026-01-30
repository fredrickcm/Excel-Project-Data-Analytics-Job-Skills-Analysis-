# 📊 Excel Salary Dashboard

![Salary Dashboard Preview](1_Salary_Dashboard.png)

## 📌 Introduction
This **Data Jobs Salary Dashboard** was created to help job seekers explore salary trends across data-related roles and ensure they are being fairly compensated.

The project is based on data from my **Excel course**, which focuses on building a strong foundation in data analysis using Excel. The dataset includes detailed information on job titles, salaries, locations, and required skills, all visualized in an interactive dashboard.

---

## 📁 Dashboard File
- **Final Dashboard:** `1_Salary_Dashboard.xlsx`

---

## 🧠 Excel Skills Used
The following Excel skills were utilized throughout the analysis:

- 📉 Charts & Data Visualization  
- 🧮 Formulas and Functions  
- ❎ Data Validation  

---

## 📊 Data Jobs Dataset
The dataset contains **real-world data science job information from 2023**, provided through my Excel course. It includes:

- 👨‍💼 Job Titles  
- 💰 Salaries  
- 📍 Locations  
- 🛠️ Skills  

---

## 🛠️ Dashboard Build

### 📉 Charts

#### 📊 Data Science Job Salaries – Bar Chart
![Salary by Job Title](Salary_Dashboard_Chart1.png)

- 🛠️ **Excel Features:** Bar chart with formatted salary values and optimized layout  
- 🎨 **Design Choice:** Horizontal bar chart for easy comparison  
- 📉 **Data Organization:** Job titles sorted by descending median salary  
- 💡 **Insights Gained:** Senior roles and Engineering positions typically earn more than Analyst roles  

---

#### 🗺️ Country Median Salaries – Map Chart
![Country Median Salaries](1_Salary_Dashboard_Chart2.png)

- 🛠️ **Excel Features:** Excel Map Chart to plot global median salaries  
- 🎨 **Design Choice:** Color-coded map to distinguish salary levels  
- 📊 **Data Representation:** Median salary by country  
- 👁️ **Visual Enhancement:** Immediate understanding of geographic salary trends  
- 💡 **Insights Gained:** Highlights global salary disparities and high/low-paying regions  

---

## 🧮 Formulas and Functions

### 💰 Median Salary by Job Title
```excel
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])) )*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
