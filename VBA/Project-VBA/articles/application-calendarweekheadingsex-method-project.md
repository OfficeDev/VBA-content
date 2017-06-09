---
title: Application.CalendarWeekHeadingsEx Method (Project)
keywords: vbapj.chm2341
f1_keywords:
- vbapj.chm2341
ms.prod: project-server
api_name:
- Project.Application.CalendarWeekHeadingsEx
ms.assetid: af964116-1d0e-7ab8-4674-4418c1c80f9c
ms.date: 06/08/2017
---


# Application.CalendarWeekHeadingsEx Method (Project)

Customizes headings in the Calendar.


## Syntax

 _expression_. **CalendarWeekHeadingsEx**( ** _MonthTitle_**, ** _WeekTitle_**, ** _DayTitle_**, ** _ShowPreview_**, ** _DaysPerWeek_**, ** _ShowTitleBeginningEndDates_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MonthTitle_|Optional|**Long**|The format of the month title. Can be one of the [PjMonthLabel](pjmonthlabel-enumeration-project.md) constants.|
| _WeekTitle_|Optional|**Long**|The format of week titles. Can be one of the [PjDateLabel](pjdatelabel-enumeration-project.md) constants.|
| _DayTitle_|Optional|**Long**|The format of day titles. Can be one of the [PjDayLabel](pjdaylabel-enumeration-project.md) constants.|
| _ShowPreview_|Optional|**Boolean**|**True** if the next and previous months are previewed.|
| _DaysPerWeek_|Optional|**Integer**|The number of days per week to display. Can be set to 5 or 7.|
| _ShowTitleBeginningEndDates_|Optional|**Boolean**|**True** if the beginning and end date titles are shown.|

### Return Value

 **Boolean**


## Remarks

Using the  **CalendarWeekHeadingsEx** method without specifying any arguments displays the **Timescale** dialog box with the **Week Headings** tab selected.


