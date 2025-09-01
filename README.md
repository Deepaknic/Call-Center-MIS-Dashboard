# Call-Center-MIS-Dashboard
Excel-based MIS Dashboard for Call Center performance tracking using Pivot Tables, Charts, and VBA automation

**Project Overview**
This project is a Call Center Management Information System (MIS) Dashboard built in Excel.
It provides insights into employee performance using formulas, Pivot Tables, Pivot Charts, and VBA macros.
The dashboard enables:

    •	Tracking employee working hours & efficiency
    •	Monitoring breaks and active time with customers
    •	Analyzing attended calls & performance trends
    •	Automating reports with VBA macros
________________________________________
**Features**

    •	Employee detail extraction (Name, ID, Region)
    •	Total working hours & efficiency calculation
    •	Break analysis & attended calls tracking
    •	Pivot Tables for performance insights
    •	Dynamic charts with filters (Top N employees)
    •	VBA Macros for automation (Show/Hide charts, update Top N)
________________________________________
**Files Included**

    •	Call Center Projectt.xlsm → Main Excel dashboard (macro-enabled)
    •	Documentation → Formulas, steps, and VBA scripts explained
________________________________________
**Technology Used**

    •	Microsoft Excel (with VBA enabled)
    •	Pivot Tables & Charts
    •	Formulas (VLOOKUP, TIMEVALUE, MID, LEFT, TRIM)
    •	VBA Macros
________________________________________
**How It Works**
**1. Employee Details Extraction**
**Column	Formula	Purpose**

    Employee Name	=TRIM(MID(A2,FIND("_",A2)+1,FIND("(",A2)-FIND("_",A2)-1))	Extracts employee name
    Region	=LEFT(A2,FIND("_",A2)-1)	Extracts region
    Employee ID	=TRIM(MID(A2,FIND("(",A2)+1,FIND(")",A2)-FIND("(",A2)-1))	Extracts employee ID
**2. Working Hours & Efficiency**

    •	Working Hours: =TIMEVALUE(VLOOKUP(A2,EData,2,0))
    •	Efficiency: =TIMEVALUE(VLOOKUP(A2,EData,3,0))
    •	Active with Customers: =VLOOKUP(A2,EData,6,0)
**3. Breaks & Calls**

•	Total Break Taken: =TIMEVALUE(VLOOKUP(A2,BData,9,0))
•	Attended Calls: =VLOOKUP(A2,OData,2,0)
________________________________________
**VBA Automation**

Macro: Update Top N Employees
Sub Report_1()
    'Updates PivotTable to show top N employees by working hours
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Emp. Name").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Emp. Name").PivotFilters. _
        Add2 Type:=xlTopCount, DataField:=ActiveSheet.PivotTables("PivotTable1"). _
        PivotFields("Sum of Total No of Working Hours"), Value1:=Sheet5.Range("D1").Value
End Sub

**Macro: Show/Hide Charts**

Sub OpenEye()
    'Show chart and spinner
    ActiveSheet.Shapes.Range(Array("Open Eye")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Close Eye")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Chart 1")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Spinner 2")).Visible = msoTrue
End Sub

Sub CloseEye()
    'Hide chart and spinner
    ActiveSheet.Shapes.Range(Array("Open Eye")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Close Eye")).Visible = msoTrue
    ActiveSheet.Shapes.Range(Array("Chart 1")).Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("Spinner 2")).Visible = msoFalse
End Sub
________________________________________
**KPIs in Dashboard**

•	Total Working Hours
•	Average Efficiency (%)
•	Total Break Time Taken
•	Total Attended Calls
•	Active Time with Customers
________________________________________
**Dashboard Preview**

<img width="985" height="602" alt="image" src="https://github.com/user-attachments/assets/919d566a-f6a0-4a73-9891-5900c653fa3e" />

 ________________________________________
**How to Use**

1.	Download the file Call Center Projectt.xlsm
2.	Open in Excel (enable macros when prompted)
3.	Navigate through dashboard sheets
4.	Use filters, and buttons for interactivity
________________________________________
Author
Developed by Deepaknic

