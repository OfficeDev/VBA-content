---
title: CalendarView.DisplayedDates Property (Outlook)
keywords: vbaol11.chm3028
f1_keywords:
- vbaol11.chm3028
ms.prod: outlook
api_name:
- Outlook.CalendarView.DisplayedDates
ms.assetid: 45d77ff9-b93e-4439-3594-ff9dcf1f180b
ms.date: 06/08/2017
---


# CalendarView.DisplayedDates Property (Outlook)

Returns a  **Variant** array containing strings that represent the days displayed in a **[CalendarView](calendarview-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **DisplayedDates**

 _expression_ A variable that represents a **CalendarView** object.


## Remarks

This property returns an array of date strings, in which each date string represents a day displayed in the  **CalendarView** object. The date strings are formatted using the short date format settings for the operating system.


## Example

The following Visual Basic for Applications (VBA) example obtains the value of the  **DisplayedDates** property from the current **CalendarView** object, then displays a dialog box with a summary of that property value.


```vb
Sub DisplayDayRange() 
 
 Dim objView As CalendarView 
 
 Dim varArray As Variant 
 
 
 
 ' Check if the current view is a calendar view. 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olCalendarView Then 
 
 
 
 ' Obtain a CalendarView object reference for the 
 
 ' current calendar view. 
 
 Set objView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 ' Obtain the DisplayedDates value, a string 
 
 ' array of dates representing the dates displayed 
 
 ' in the calendar view. 
 
 varArray = objView.DisplayedDates 
 
 
 
 ' If the example obtained a valid array, display 
 
 ' a dialog box with a summary of its contents. 
 
 If IsArray(varArray) Then 
 
 MsgBox "There are " &; _ 
 
 (UBound(varArray) - LBound(varArray)) + 1 &; _ 
 
 " days displayed, from " &; _ 
 
 varArray(LBound(varArray)) &; _ 
 
 " to " &; _ 
 
 varArray(UBound(varArray)) 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[CalendarView Object](calendarview-object-outlook.md)

