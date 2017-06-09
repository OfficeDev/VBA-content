---
title: Application.FilePageSetupCalendar Method (Project)
keywords: vbapj.chm2361
f1_keywords:
- vbapj.chm2361
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupCalendar
ms.assetid: 50f4ab0a-ffb4-2bff-44af-82b674de7c4c
ms.date: 06/08/2017
---


# Application.FilePageSetupCalendar Method (Project)

Sets up the Calendar for printing.


## Syntax

 _expression_. **FilePageSetupCalendar**( ** _Name_**, ** _MonthsPerPage_**, ** _WeeksPerPage_**, ** _ScreenWeekHeight_**, ** _OnlyDaysInMonth_**, ** _OnlyWeeksInMonth_**, ** _MonthPreviews_**, ** _MonthTitle_**, ** _AdditionalTasks_**, ** _GroupAdditionalTasks_**, ** _PrintNotes_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**Variant**|Sets up the Calendar for printing.|
| _MonthsPerPage_|Optional|**Integer**|The number of months to print on each page. Can be 1 or 2. The  **MonthsPerPage** argument is required if **OnlyDaysInMonth** or **OnlyWeeksInMonth** is specified.|
| _WeeksPerPage_|Optional|**Integer**|The number of weeks to print on each page.|
| _ScreenWeekHeight_|Optional|**Boolean**|**True** if the week height displayed on screen is used for the printout.|
| _OnlyDaysInMonth_|Optional|**Boolean**|**True** if only the days in the month are printed. **False** if the days at the end of the previous month and at the start of the next month are printed in addition to the days in the current month. The **OnlyDaysInMonth** argument is ignored unless a value for **MonthsPerPage** is specified.|
| _OnlyWeeksInMonth_|Optional|**Boolean**|**True** if only the weeks that are fully contained in the month are printed. **False** if weeks that have one or more days in the month are printed. The **OnlyWeeksInMonth** argument is ignored unless a value for **MonthsPerPage** is specified.|
| _MonthPreviews_|Optional|**Boolean**|**True** if preview calendars for the previous and next months are printed.|
| _MonthTitle_|Optional|**Boolean**|**True** if the month's title is printed.|
| _AdditionalTasks_|Optional|**Boolean**|**True** if tasks that do not fit on the Calendar are printed. (Additional tasks appear at the end of the printout.)|
| _GroupAdditionalTasks_|Optional|**Boolean**|**True** if additional tasks are grouped by day.|
| _PrintNotes_|Optional|**Boolean**|**True** if the notes associated with each task are printed. Notes are printed at the end, after any additional tasks|

### Return Value

 **Boolean**


## Remarks

Using the  **FilePageSetupCalendar** method without specifying any arguments displays the **Page Setup** dialog box with the **View** tab selected. The **FilePageSetupCalendar** method is only available when the Calendar is the active view.


## Example

The following example sets up the calandar for printing with 2 months per page and with preview calendars for the previous and next months.


```vb
Sub File_PageSetupCalendar() 
 
 'Activate Calandar view 
 ViewApply Name:="&;Calendar" 
 FilePageSetupCalendar MonthsPerPage:=2, OnlyDaysInMonth:=False, MonthPreviews:=True 
End Sub
```


