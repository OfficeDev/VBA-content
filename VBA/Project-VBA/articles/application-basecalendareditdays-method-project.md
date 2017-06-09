---
title: Application.BaseCalendarEditDays Method (Project)
keywords: vbapj.chm615
f1_keywords:
- vbapj.chm615
ms.prod: project-server
api_name:
- Project.Application.BaseCalendarEditDays
ms.assetid: 3a65015e-c174-985a-5235-099db363c003
ms.date: 06/08/2017
---


# Application.BaseCalendarEditDays Method (Project)

Changes one or more days in a base calendar.


## Syntax

 _expression_. **BaseCalendarEditDays**( ** _Name_**, ** _StartDate_**, ** _EndDate_**, ** _WeekDay_**, ** _Working_**, ** _From1_**, ** _To1_**, ** _From2_**, ** _To2_**, ** _From3_**, ** _To3_**, ** _Default_**, ** _From4_**, ** _To4_**, ** _From5_**, ** _To5_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|**String**. The name of the base calendar to change.|
| _StartDate_|Optional|**Variant**|The first date to change. If  **StartDate** is specified without **EndDate**, that date is the only day affected. If **WeekDay** is specified, **StartDate** is ignored.|
| _EndDate_|Optional|**Variant**|The last date to change. If  **EndDate** is specified without **StartDate**, that date is the only day affected. If **WeekDay** is specified, **EndDate** is ignored.|
| _WeekDay_|Optional|**Long**|The weekday to change. If  **StartDate** or **EndDate** is specified, **WeekDay** is ignored. Can be one of the **[PjWeekday](pjweekday-enumeration-project.md)** constants.|
| _Working_|Optional|**Boolean**|**True** if the days are working days.|
| _From1_|Optional|**Variant**|The start time of the first shift.|
| _To1_|Optional|**Variant**|The end time of the first shift.|
| _From2_|Optional|**Variant**|The start time of the second shift.|
| _To2_|Optional|**Variant**|The end time of the second shift.|
| _From3_|Optional|**Variant**|The start time of the third shift.|
| _To3_|Optional|**Variant**|The end time of the third shift.|
| _Default_|Optional|**Boolean**|Resets the dates specified by  **StartDate** and **EndDate**, or by **WeekDay**, to the default values. If **Working** is specified, **Default** is ignored.|
| _From4_|Optional|**Variant**|The start time of the fourth shift.|
| _To4_|Optional|**Variant**|The end time of the fourth shift.|
| _From5_|Optional|**Variant**|The start time of the fifth shift.|
| _To5_|Optional|**Variant**|The end time of the fifth shift.|

### Return Value

 **Boolean**


## Example

The following example makes Wednesday a nonworking day in the Standard calendar.


```vb
Sub MakeWednesdaysNonWorking() 
 BaseCalendarEditDays Name:="Standard", Weekday:=pjWednesday, Working:=False 
End Sub
```

The following example makes the days from 2/10/97 through 2/12/97 nonworking days in the Standard calendar.




```vb
Sub MakeSelectedDaysNonWorking() 
 BaseCalendarEditDays Name:="Standard", StartDate:="2/10/97", EndDate:="2/12/97", Working:=False 
End Sub
```


