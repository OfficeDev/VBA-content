
# Application.FilePageSetupCalendar Method (Project)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets up the Calendar for printing.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **FilePageSetupCalendar**( **_Name_**,  **_MonthsPerPage_**,  **_WeeksPerPage_**,  **_ScreenWeekHeight_**,  **_OnlyDaysInMonth_**,  **_OnlyWeeksInMonth_**,  **_MonthPreviews_**,  **_MonthTitle_**,  **_AdditionalTasks_**,  **_GroupAdditionalTasks_**,  **_PrintNotes_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Optional| **Variant**|Sets up the Calendar for printing.|
|MonthsPerPage|Optional| **Integer**|The number of months to print on each page. Can be 1 or 2. The  **MonthsPerPage** argument is required if **OnlyDaysInMonth** or **OnlyWeeksInMonth** is specified.|
|WeeksPerPage|Optional| **Integer**|The number of weeks to print on each page.|
|ScreenWeekHeight|Optional| **Boolean**| **True** if the week height displayed on screen is used for the printout.|
|OnlyDaysInMonth|Optional| **Boolean**| **True** if only the days in the month are printed. **False** if the days at the end of the previous month and at the start of the next month are printed in addition to the days in the current month. The **OnlyDaysInMonth** argument is ignored unless a value for **MonthsPerPage** is specified.|
|OnlyWeeksInMonth|Optional| **Boolean**| **True** if only the weeks that are fully contained in the month are printed. **False** if weeks that have one or more days in the month are printed. The **OnlyWeeksInMonth** argument is ignored unless a value for **MonthsPerPage** is specified.|
|MonthPreviews|Optional| **Boolean**| **True** if preview calendars for the previous and next months are printed.|
|MonthTitle|Optional| **Boolean**| **True** if the month's title is printed.|
|AdditionalTasks|Optional| **Boolean**| **True** if tasks that do not fit on the Calendar are printed. (Additional tasks appear at the end of the printout.)|
|GroupAdditionalTasks|Optional| **Boolean**| **True** if additional tasks are grouped by day.|
|PrintNotes|Optional| **Boolean**| **True** if the notes associated with each task are printed. Notes are printed at the end, after any additional tasks|

### Return Value

 **Boolean**


## Remarks
<a name="sectionSection1"> </a>

Using the  **FilePageSetupCalendar** method without specifying any arguments displays the **Page Setup** dialog box with the **View** tab selected. The **FilePageSetupCalendar** method is only available when the Calendar is the active view.


## Example
<a name="sectionSection2"> </a>

The following example sets up the calandar for printing with 2 months per page and with preview calendars for the previous and next months.


```
Sub File_PageSetupCalendar() 
 
 'Activate Calandar view 
 ViewApply Name:="&amp;Calendar" 
 FilePageSetupCalendar MonthsPerPage:=2, OnlyDaysInMonth:=False, MonthPreviews:=True 
End Sub
```

