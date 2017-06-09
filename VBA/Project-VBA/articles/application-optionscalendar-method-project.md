---
title: Application.OptionsCalendar Method (Project)
keywords: vbapj.chm649
f1_keywords:
- vbapj.chm649
ms.prod: project-server
api_name:
- Project.Application.OptionsCalendar
ms.assetid: bde3b645-3417-ee45-57b5-0109bc7b17ad
ms.date: 06/08/2017
---


# Application.OptionsCalendar Method (Project)

Sets options for the calendar of the active project.


## Syntax

 _expression_. **OptionsCalendar**( ** _StartWeekOnMonday_**, ** _StartYearIn_**, ** _StartTime_**, ** _FinishTime_**, ** _HoursPerDay_**, ** _HoursPerWeek_**, ** _SetDefaults_**, ** _StartWeekOn_**, ** _UseFYStartYear_**, ** _DaysPerMonth_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartWeekOnMonday_|Optional|**Boolean**|**True** if the calendar week starts on Monday. **False** if the calendar week starts on Sunday. If _StartWeekOn_ is specified, _StartWeekOnMonday_ is ignored. (The _StartWeekOn_ argument is a better way to specify the start of the week.)|
| _StartYearIn_|Optional|**Long**|The first month of the fiscal year. Can be one of the  **[PjMonth](pjmonth-enumeration-project.md)** constants.|
| _StartTime_|Optional|**Variant**|The default start time for working days.|
| _FinishTime_|Optional|**Variant**|The default finish time for working days.|
| _HoursPerDay_|Optional|**Double**|The default number of work hours per day.|
| _HoursPerWeek_|Optional|**Double**|The default number of work hours per week.|
| _SetDefaults_|Optional|**Boolean**|**True** if the values of _StartYearIn_,  _StartTime_,  _FinishTime_,  _HoursPerDay_,  _HoursPerWeek_,  _StartWeekOn_, and  _UseFYStartYear_ are used as the default values for new projects. The default value is **False**.|
| _StartWeekOn_|Optional|**Long**|The first day of the week. Can be one of the  **[PjWeekday](pjweekday-enumeration-project.md)** constants.|
| _UseFYStartYear_|Optional|**Boolean**|**True** if a fiscal year is determined by the year of the first month of that fiscal year. **False** if determined by the last month of the fiscal year.For example, if  _StartYearIn_ is pjJuly (to denote July 2012) and _UseFYStartYear_ is **True**, the fiscal year ending in June 2012 would be FY2012.|
| _DaysPerMonth_|Optional|**Double**|The default number of work days per month.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the setting on the  **Schedule** tab of the **Project Options** dialog box.

Using the  **OptionsCalendar** method without specifying any arguments displays the **Project Options** dialog box with the **General** tab selected.


## Example

The following example sets the first month of the fiscal year to April, default number of work hours per day to 4 hours and default number of work hours per week to 20 hours.


```vb
Sub Options_Calendar() 
    Dim HoursDay As Double 
    Dim HoursWeek As Double 
 
    HoursDay = 4 
    HoursWeek = 20 
 
    OptionsCalendar StartYearIn:=pjApril, HoursPerDay:=HoursDay, HoursPerWeek:=HoursWeek 
End Sub
```


