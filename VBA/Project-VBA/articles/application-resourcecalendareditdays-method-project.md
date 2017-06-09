---
title: Application.ResourceCalendarEditDays Method (Project)
keywords: vbapj.chm620
f1_keywords:
- vbapj.chm620
ms.prod: project-server
api_name:
- Project.Application.ResourceCalendarEditDays
ms.assetid: 0dc0172f-bc49-347a-7c46-f6a6dc608d8f
ms.date: 06/08/2017
---


# Application.ResourceCalendarEditDays Method (Project)

Edits days in a resource calendar.


## Syntax

 _expression_. **ResourceCalendarEditDays**( ** _ProjectName_**, ** _ResourceName_**, ** _StartDate_**, ** _EndDate_**, ** _WeekDay_**, ** _Working_**, ** _Default_**, ** _From1_**, ** _To1_**, ** _From2_**, ** _To2_**, ** _From3_**, ** _To3_**, ** _From4_**, ** _To4_**, ** _From5_**, ** _To5_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProjectName_|Required|**String**|The name of the project containing the resource calendar to edit.|
| _ResourceName_|Required|**String**|The name of the resource to edit.|
| _StartDate_|Optional|**Variant**|The first date to edit.|
| _EndDate_|Optional|**Variant**|The last date to edit.|
| _WeekDay_|Optional|**Long**|The weekday to edit. If StartDate and EndDate are specified, WeekDay is ignored. Can be one of the  **[PjWeekday](pjweekday-enumeration-project.md)** constants.|
| _Working_|Optional|**Boolean**|**True** if the days are working days. If Default is **True**, Working is ignored.|
| _Default_|Optional|**Boolean**|**True** if the resource calendar uses the values in the corresponding base calendar as defaults. The default value is **False**.|
| _From1_|Optional|**Variant**|The start time of the first shift.|
| _To1_|Optional|**Variant**|The end time of the first shift.|
| _From2_|Optional|**Variant**|The start time of the second shift.|
| _To2_|Optional|**Variant**|The end time of the second shift.|
| _From3_|Optional|**Variant**|The start time of the third shift.|
| _To3_|Optional|**Variant**|The end time of the third shift.|
| _From4_|Optional|**Variant**|The start time of the fourth shift.|
| _To4_|Optional|**Variant**| The end time of the fourth shift.|
| _From5_|Optional|**Variant**|The start time of the fifth shift.|
| _To5_|Optional|**Variant**|The end time of the fifth shift.|

### Return Value

 **Boolean**


## Remarks

The  **ResourceCalendarEditDays** method returns a trappable error (error code 1101) when applied to material resources.


