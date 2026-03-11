# CRS-Score-Analysis-Tool-Excel

An interactive Excel-based CRS (Comprehensive Ranking System) score analysis tool designed to help users analyze how different factors affect their CRS score for Canadian immigration. 

![DashBoard Overview](/Images/DashBoard%20Overview.png)
## ⚠️ Disclaimer: 
- This tool does not provide immigration advice and should not be considered a substitute for a licensed immigration consultant or lawyer. The purpose of this project is purely analytical and educational, allowing users to explore CRS scoring trends and understand how different inputs influence their score.
- The CRS scoring logic and datasets were derived from publicly available information on the official IRCC website.
This tool is intended only for educational and analytical purposes.
For official immigration advice, please consult:
A licensed Immigration Consultant (RCIC)
A qualified Immigration Lawyer
The official IRCC website

## 👁️How to Use
**Reset Data | Enter Data | Analyze Trends**
![Tool Demo](/Images/GIF/CRS%20Tool%20Demo%20(1).gif)

[*Dashboard*](https://github.com/Gopi467420/CRS-Score-Analysis-Tool-Excel-/blob/main/CRS%20Macro.xlsm)



## 💡 Key Features
#### ✔ Interactive CRS input system 
#### ✔ Dynamic score calculation 
#### ✔ CRS scoring trend visualization 
#### ✔ Data-driven insights 
#### ✔ Fully Excel-based solution 
#### ✔ No external dependencies

## 📊 Project Purpose

### The goal of this project is to help individuals: 
 -  Understand how CRS points are distributed across different factors.
 - See how changes in KPI's like age, education, language scores, and experience impact CRS.
 - Visualize trends in CRS scoring components.
 - Identify potential ways to improve their CRS score.
  - Users can use these insights to determine which factors may help them improve their CRS score.

 **<Instead of only showing a final score, this tool focuses on data visualization and trend analysis, allowing users to understand where they stand in the CRS system>**



## 📂 Data Source
 - The CRS scoring data used in this project was obtained from the official Immigration, Refugees and Citizenship Canada (IRCC) website.
**Official source:** [Comprehensive Ranking System (CRS) criteria](https://www.canada.ca/en/immigration-refugees-citizenship/services/immigrate-canada/express-entry/check-score/crs-criteria.html#spouse)

 - The data was imported using the **From Web** tool in Excel, So that if in future if the data is updated in the website it updates when teh data is refreshed in Excel.
 - It was extracted as structured into Excel tables
Transformed into usable datasets for analysis
⚙️ How the Tool Works
The tool follows a structured process:

## 1️⃣ Data Extraction
CRS scoring information was collected from official IRCC resources.

## 2️⃣ Data Transformation & Table Structuring
 - Using Excel's Data → Transform tools, the raw data was cleaned
Structured converted into organized tables named for easier referencing
- The CRS scoring datasets were converted into Excel Tables to enable:
Structured references
Easier calculations
Dynamic data usage in formulas and charts 

![Transformed Data](/Images/Transforming%20Data.png)

<p align="center"><b>   
Transforming Data into Meaningful Tables
</b><p>

## 3️⃣ Data Validation Inputs
 - User inputs are handled using Data Validation dropdown menus placed in the Data Validation Sheet, replicating the type of questions found on the CRS calculator on the IRCC website.

 ![Data Validation](/Images/Data%20Validation.png)

<p align="center"><b>   
Data Validation Drop Down Data
</b><p>

 ## 4️⃣Questions
  The input questions were modeled after the official IRCC CRS Score Calculator and implemented using Excel data validation dropdowns, allowing users to select options similar to those found on the official calculator while simplifying the interface for analysis.
  ### These inputs allow users to select values such as:
  -  Age
  -  Education level
  -  Language proficiency
  -  Canadian work experience
  -  Foreign work experience
  -  Spouse Language ability and Canadian work experience
  -  Marital status
  - Other CRS factors

 
## 5️⃣  CRS Trend Visualization
### For each CRS scoring category, trendline charts were created to show:
- How points change across different values
- The relationship between input variables and CRS score
- The current user's position relative to the scoring curve
- Charts allow users to visually analyze:
- How CRS points increase or decrease
- Which factors provide higher scoring opportunities and ignoring which don't
- Where they currently stand in the scoring system

## 6️⃣Named Formula System
The Excel Name Manager was used extensively to create reusable formulas.
This approach allows:
Modular calculations
Cleaner formulas
Easier maintenance and updates
Named formulas help calculate CRS scores dynamically based on user input.

![Name Manager](/Images/Name%20Manager.png)
<p align="center"><b>   
Name Manager for formulas
</b><p>


## 7️⃣ Excel Functions, Formulas and VBA Used


Excel Formulas Used

**LAMBDA:** 
Used to create reusable custom formulas for repeated CRS calculations across different scoring sections.

**TEXTJOIN:**
Used to combine text values dynamically, especially where multiple values needed to be merged into a single lookup-friendly format.

**XLOOKUP:**
Used to retrieve CRS point values from structured tables based on user-selected inputs.

**FILTER:**
Used to return dynamic subsets of data from tables for charting, analysis, and dropdown-dependent outputs.

**INDEX:**
Used in cases where specific values needed to be returned from structured datasets based on row or column position.

### Custom Formulas
The Excel logic was designed so that user inputs from dropdowns could flow into reusable formulas and return the correct CRS-related values from transformed IRCC data tables. This helped reduce repeated formula writing and made the workbook easier to update and scale.

**FFGREATERTHANALLFORANLANGUAGE:** Checks whether all four language abilities — listening, speaking, reading, and writing — are greater than or equal to a given comparison value for the selected language type.

#### This function was mainly used for:

evaluating language-based conditions

returning boolean results (TRUE / FALSE)

supporting chart logic and score analysis

``` 
=LAMBDA(LanguageType,ComparisonValue,
    LET(
        FrenchPos,
        MATCH(
            LanguageType,
            XLOOKUP(
                'Language CLB Calculations'!$C$8:$C$9,
                'Language CLB Calculations'!$M$6:$M$10,
                'Language CLB Calculations'!$L$6:$L$10
            ),
            0
        ),
        AND(
            XLOOKUP(INDEX('Language CLB Calculations'!$D$8:$D$9,FrenchPos),XLOOKUP(INDEX('Language CLB Calculations'!$A$8:$A$9,FrenchPos),'Language CLB Calculations'!$G$14:$H$14,'Language CLB Calculations'!$G$15:$H$22),'Language CLB Calculations'!$J$15:$J$22)>=ComparisonValue,
            XLOOKUP(INDEX('Language CLB Calculations'!$E$8:$E$9,FrenchPos),XLOOKUP(INDEX('Language CLB Calculations'!$A$8:$A$9,FrenchPos),'Language CLB Calculations'!$G$14:$H$14,'Language CLB Calculations'!$G$15:$H$22),'Language CLB Calculations'!$J$15:$J$22)>=ComparisonValue,
            XLOOKUP(INDEX('Language CLB Calculations'!$F$8:$F$9,FrenchPos),XLOOKUP(INDEX('Language CLB Calculations'!$A$8:$A$9,FrenchPos),'Language CLB Calculations'!$G$14:$H$14,'Language CLB Calculations'!$G$15:$H$22),'Language CLB Calculations'!$J$15:$J$22)>=ComparisonValue,
            XLOOKUP(INDEX('Language CLB Calculations'!$G$8:$G$9,FrenchPos),XLOOKUP(INDEX('Language CLB Calculations'!$A$8:$A$9,FrenchPos),'Language CLB Calculations'!$G$14:$H$14,'Language CLB Calculations'!$G$15:$H$22),'Language CLB Calculations'!$J$15:$J$22)>=ComparisonValue
        )
    )
)

```



**FFGetPointsFor:** 
Returns the CRS points for a single value from the official score table.

#### This function supports both:

exact-match lookups

range-based lookups

It was used to simplify score retrieval from structured CRS tables.
```
=LAMBDA(Value,Table,LookupCol,ReturnCol,IsRange,
    IFERROR(
        IF(
            IsRange,
            INDEX(
                INDEX(Table,,ReturnCol),
                XMATCH(
                    1,
                    MAP(
                        INDEX(Table,,LookupCol),
                        LAMBDA(r,--AND(
                            Value>=VALUE(TRIM(TEXTBEFORE(r,"-"))),
                            Value<=VALUE(TRIM(TEXTAFTER(r,"-")))
                        ))
                    )
                )
            ),
            XLOOKUP(Value, INDEX(Table,,LookupCol), INDEX(Table,,ReturnCol))
        ),
        0
    )
)
```

**FFGetValue:** Retrieves a matching value from a table using XLOOKUP and returns one of two adjacent columns depending on a condition.

This function was useful when selecting different outputs from the same lookup table based on user input or scoring conditions.
```
=LAMBDA(CellValue,Table,Condition,
    XLOOKUP(
        CellValue,
        INDEX(Table,,1),
        IF(Condition, INDEX(Table,,2), INDEX(Table,,3))
    )
)
```


### Excel VBA Used in the Project
The goal of the VBA in this project was not to replace formulas, but to enhance the user experience and simplify interaction with the workbook.

**VBA was used for tasks such as:**

- resetting selected input cells

- restoring default values

- supporting dropdown/input workflows

- improving sheet interaction and dashboard behavior

- helping control chart or interface updates where needed





## 8️⃣Chart Design


### To visualize CRS scoring trends, the charts combine:
- XY Scatter Charts (to plot score current positions)
- Line Charts (to show CRS point trends)
### This hybrid chart approach allows the tool to show:
- The CRS scoring curve
- The user's current score position along the curve
## 📈 Example Visualizations
### The dashboard includes  visualizations for multiple CRS components, such as:
- Age vs CRS Points
- Education vs CRS Points
- Language Scores vs CRS Points
- Work Experience vs CRS Points
- Each chart helps users understand how their choices influence CRS scoring 
outcomes.

## 🖼️ Screenshots Dashboard Overview

<table align="center">
<tr>
<td align="center">
<img src="/Images/Age%20Trend.png" width="600"><br>
<b>Age Points Trend</b>
</td>

<td align="center">
<img src="/Images/First%20Language%20Trend.png" width="600"><br>
<b>First Language Points Trend</b>
</td>
</tr>

<tr>
<td align="center">
<img src="/Images/Higher%20Education%20Trend.png" width="600"><br>
<b>Higher Education Points Trend</b>
</td>

<td align="center">
<img src="/Images/Spouce%20points%20Trend.png" width="600"><br>
<b>Spouse Points Trend</b>
</td>
</tr>
</table>

## 🧰 Tools & Technologies

### This project was built entirely using Microsoft Excel. Main Excel features used:
- Excel Tables
- Excel VBA
- Data Validation
- Named Ranges
- Name Manager
- Excel Charts
- XY Scatter Charts
- Line Charts
- Formula-based calculations
- Structured references

## 👨‍💻 Author : **Gopi Velmurugan**
- [**Linkedin**](https://www.linkedin.com/in/gopi-velmurugan-65249716a/)   
- [**Gmail**](mailto:gopiofficialca20@gmail.com)








