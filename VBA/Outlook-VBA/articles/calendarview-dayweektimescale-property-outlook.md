---
title: CalendarView.DayWeekTimeScale Property (Outlook)
keywords: vbaol11.chm3025
f1_keywords:
- vbaol11.chm3025
ms.prod: outlook
api_name:
- Outlook.CalendarView.DayWeekTimeScale
ms.assetid: 94f2aad5-6699-82e9-40a4-3c3c13d80684
ms.date: 06/08/2017
---


# CalendarView.DayWeekTimeScale Property (Outlook)

Returns or sets an  **[OlDayWeekTimeScale](oldayweektimescale-enumeration-outlook.md)** constant that represents the scale used to represent time periods in a **[CalendarView](calendarview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **DayWeekTimeScale**

 _expression_ A variable that represents a **CalendarView** object.


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

