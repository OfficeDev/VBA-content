---
title: CalendarView.DaysInMultiDayMode Property (Outlook)
keywords: vbaol11.chm2640
f1_keywords:
- vbaol11.chm2640
ms.prod: outlook
api_name:
- Outlook.CalendarView.DaysInMultiDayMode
ms.assetid: 1dcb2a69-93b9-432e-56ca-7e39b040dc6f
ms.date: 06/08/2017
---


# CalendarView.DaysInMultiDayMode Property (Outlook)

Returns or sets a  **Long** value that represents the number of consecutive days displayed in the **[CalendarView](calendarview-object-outlook.md)** object. Read/write


## Syntax

 _expression_ . **DaysInMultiDayMode**

 _expression_ A variable that represents a **CalendarView** object.


## Remarks

This property can be set to a value between 2 and 14. If this property is set to a value less than 2, the property is set to 2. If the this property is set to a value greater than 14, the property is set to 14. The default value for this property is 5.


## Example

The following Visual Basic for Applications (VBA) example creates a new  **CalendarView** object in the **Calendar** default folder, and then configures it to display 14 consecutive days in multi-day mode, with Outlook items displayed within an hourly time scale.


```vb
Sub CreateTwoWeekView() 
 
 Dim objNamespace As NameSpace 
 
 Dim objFolder As Folder 
 
 Dim objView As CalendarView 
 
 
 
 ' Obtain Folder object reference to the Calendar default folder. 
 
 Set objNamespace = Application.GetNamespace("MAPI") 
 
 Set objFolder = objNamespace.GetDefaultFolder(olFolderCalendar) 
 
 
 
 ' Create a new CalendarView object named "Two Weeks". 
 
 Set objView = objFolder.Views.Add("Two Weeks", _ 
 
 olCalendarView, _ 
 
 olViewSaveOptionAllFoldersOfType) 
 
 
 
 ' Configure the new CalendarView object. 
 
 With objView 
 
 ' Display the view in multi-day mode. 
 
 .CalendarViewMode = olCalendarViewMultiDay 
 
 
 
 ' Display 14 consecutive days in multi-day 
 
 ' mode. 
 
 .DaysInMultiDayMode = 14 
 
 ' Set the time scale for the view in one-hour 
 
 ' intervals. 
 
 .DayWeekTimeScale = olTimeScale60Minutes 
 
 
 
 ' Save and apply the new CalendarView object. 
 
 .Save 
 
 .Apply 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[CalendarView Object](calendarview-object-outlook.md)

