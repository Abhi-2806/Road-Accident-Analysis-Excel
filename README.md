# 🚗 Road Accident Data Analysis & Interactive Excel Dashboard

## 📌Summary

Road traffic accidents present significant public safety, urban planning and emergency management challenges. This project provides a comprehensive, end-to-end analytical study of historical collision records using Microsoft Excel as the sole analytics engine.

All data cleaning, transformation, KPI formulation, cohort aggregations and executive dashboard design were performed natively within Excel, eliminating the need for external database tools or business intelligence platforms. The resulting workbook enables municipal traffic authorities and safety planners to isolate accident severity drivers, road surface hazards, lighting factors, and high-risk casualty profiles.

## 🎯 Business Problems & Analytical Objectives

1. Casualty Severity Profiling: Quantify the baseline proportion of fatal, serious, and slight casualties across overall incident records to gauge impact severity.

2. Temporal & Seasonal Patterns: Identify high-risk time windows by analyzing incident frequency across month and Year.

3. Environmental & Surface Impact: Correlate adverse driving conditions (wet/damp, snow/ice, dry surfaces) and lighting states (daylight vs. dark/unlit) with incident frequency and severity.

4. Vulnerable Road User Demographics: Segment casualty counts across vehicle classifications (two-wheelers, passenger cars, light/heavy commercial vehicles) and casualty groups.

5. Urban vs. Rural Disparity: Evaluate the relationship between speed limits, road type and fatality ratios between urban and rural zones.

## 📊 Sheet Structure

The workbook (Road Accident Data.xlsx) is structured into distinct, modular functional layers to ensure separation between raw records, processing calculations, and final presentation:

Road Accident Data.xlsx  
📋 Sheet 1: Dashboard   

        ├── Primary KPI Scorecards (CY Casualties, CY Accidents, YoY casualities)  
        ├── Monthly Trend Comparison (CY vs. PY Line Chart)  
        ├── Casualties by Vehicle Type (Horizontal Bar / Tree Map)  
        ├── Casualties by Road Type (Horizontal Bar Chart)  
        ├── Environmental Donut Charts (Road Surface & Light Conditions)  
        ├── Area Breakdown (Urban vs. Rural Donut Chart)  
        └── Synchronized Interactive Slicers (Year, Accident Severity, Road Type, Area)  

 ⚙️ Sheet 2: Data Sheet  

        ├── PT 1: CY & PY Total Casualties & YoY Variance Calculation  
        ├── PT 2: CY & PY Total Accidents & YoY Variance Calculation  
        ├── PT 3: Casualties by Accident Severity (Fatal, Serious, Slight)  
        ├── PT 4: Monthly Casualties Trend (Jan–Dec for 2021 vs. 2022)  
        ├── PT 5: Casualties by Vehicle Type  
        ├── PT 6: Casualties by Road Type (Single/Dual Carriageway, Roundabout, Slip Road)  
        ├── PT 7: Casualties by Road Surface (Dry, Wet/Damp, Snow, Frost/Ice, Flood)  
        ├── PT 8: Casualties by Light Conditions (Daylight, Darkness - Lights Lit/Unlit)  
        └── PT 9: Casualties by Area (Urban vs. Rural)  



### 🛠️ Tools, Functions & Excel Techniques Used

1. Data Cleaning & Feature Engineering

Date & Time Parsing: Extracted calendar hierarchies using =YEAR(), =MONTH(), =TEXT(Date, "mmmm").


2. Metric Calculations & Dynamic Modeling

Summary Counting & Aggregation: Calculated incident volumes and casualty totals using multi-criteria conditional functions:

=COUNTIFS(Severity_Range, "Fatal", Road_Surface, "Wet")

=SUMIFS(Casualties_Range, Year_Range, 2021, Area_Range, "Rural")

Period-over-Period Variance: Formulated Year-over-Year (YoY) casualty variance directly across summary cells:


3. Interactive Dashboard Design

Pivot Tables & Calculated Items: Generated modular summary views grouped by month & year, vehicle categories, and road type etc.

Interactive Slicers & Timeline Controls: Configured slicers linked via Report Connections across all workbook Pivot Charts to achieve synchronized multi-dimensional filtering.

Visual Components:

KPI Summary Cards: High-level metrics showing Current Year Casualties, Fatal count, and YoY Casualties.

Line & Area Trend Charts: Tracking month-over-month accident frequency and seasonal peaks.

Donut & Bar Charts: Profiling casualty splits by road surface, lighting conditions, and vehicle classification.



### 💡 Strategic Recommendations for Stakeholders

1. Infrastructure Safety Upgrades: Prioritize high-friction resurfacing and drainage improvements on single carriageways and unsegregated rural corridors to prevent skidding under wet/adverse road conditions.

2. Targeted Speed & Hazard Enforcement: Install automated speed enforcement and traffic calming on high-speed rural routes, where lower collision volumes paradoxically yield the highest fatality rates due to severe impact speeds.

3. Low-Light & Night Visibility Interventions: Deploy high-output LED street lighting and reflective hazard signage along high-density urban corridors to mitigate severe nighttime and low-light collisions.

4. Vulnerable Road User Protection: Establish dedicated, physically separated transit lanes and protected intersections in dense commercial and residential zones to shield pedestrians, cyclists, and two-wheelers from passenger and commercial traffic.


5. Data-Driven Policy Enforcement: Focus traffic patrols and seasonal safety campaigns around high-risk collision spikes identified in late autumn and winter months (October–December).

### 📂 Repository File Structure
```
│
├── Road Accident Data.xlsx      # Primary Excel workbook (Dashboard, Data Analysis, Data)
│
├── images/                        
│   ├── Road-Accident-Analysis-Dashboard.png
│   ├── Road-Accident-Analysis-Data Sheet.png
│
├── dashboard/                  # Microsoft Excel dashboard/data file
│   └── Road Accident Data.xlsx
```

### 🚀 How to View & Explore the Dashboard

Clone the Repository:

git clone https://github.com/your-username/road-accident-data-analysis.git


Open the Workbook:

Open Road Accident Data.xlsx in Microsoft Excel 2016, 2019, 2021, or Microsoft 365 (for full Slicer and modern dynamic chart support).

Interact with the Dashboard:

Navigate to the Dashboard worksheet.

Click any Slicer button (e.g., Year, Accident Severity, Road Type) to dynamically filter all visualizations and KPI cards simultaneously.

Note:- Link on Dashboard for official "GOVERNMENT OF INDIA
MINISTRY OF ROAD TRANSPORT AND HIGHWAYS" has been provided, showing the road accidents from 2020-2024.

### 👤 Author & Contact

**Abhishek Pathak**  
Data Analyst  
📧 Email: wayoflife507@gmail.com  
🔗 Contact No. 6307017308  
🔗 [LinkedIn](https://www.linkedin.com/in/abhishekpathak2806/)  

