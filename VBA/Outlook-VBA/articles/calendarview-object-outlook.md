---
title: CalendarView Object (Outlook)
keywords: vbaol11.chm3208
f1_keywords:
- vbaol11.chm3208
ms.prod: outlook
api_name:
- Outlook.CalendarView
ms.assetid: 37e078b9-9fc6-5894-b043-06d7257666a8
ms.date: 06/08/2017
---


# CalendarView Object (Outlook)

Represents a view that displays Outlook items in a calendar format.


## Remarks

The  **CalendarView** object, derived from the **[View](view-object-outlook.md)** object, allows you to create customizable views that allow you to display Outlook items within a calendar, in one of several different modes.

Outlook provides several built-in  **CalendarView** objects, and you can also create custom **CalendarView** objects. Use the **[Add](http://msdn.microsoft.com/library/8005ca2e-8b28-1286-74d1-448f2a168c65%28Office.15%29.aspx)** method of the **[Views](http://msdn.microsoft.com/library/5dd7edc2-12a2-f4c2-d158-8053d80e8dc9%28Office.15%29.aspx)** collection to add a new **CalendarView** to a **[Folder](folder-object-outlook.md)** object. Use the **[Standard](http://msdn.microsoft.com/library/798b5dcd-9226-b0f9-032e-bcfa7b3e17ab%28Office.15%29.aspx)** property to determine if an existing **CalendarView** object is built-in or custom.

The  **CalendarView** object supports several different view modes, depending on the desired layout and time period in which to display Outlook items. Use the **[CalendarViewMode](http://msdn.microsoft.com/library/144e46ed-984f-fac0-fad3-0ff5ac9f2996%28Office.15%29.aspx)** property to set the view mode, the **[StartField](http://msdn.microsoft.com/library/085c6605-0bff-98a5-fb48-ce32b76037db%28Office.15%29.aspx)** property to specify the Outlook item property that contains the start date, and the **[EndField](http://msdn.microsoft.com/library/311994db-ef43-e49c-6f0e-9b346d0bb3ca%28Office.15%29.aspx)** property to specify the Outlook item property that contains the end date for Outlook items to be displayed.

If you set the  **CalendarViewMode** property to any value other than **olCalendarViewMonth**, you can use the **[DayWeekFont](http://msdn.microsoft.com/library/ddb6f65d-72e2-d3f2-b10f-b3d8bc4d21b3%28Office.15%29.aspx)** and **[DayWeekTimeFont](http://msdn.microsoft.com/library/37ea6e1f-4148-3ab4-e0aa-48c49321ac91%28Office.15%29.aspx)** properties to configure the fonts used to display the day, date, and time labels in the view. Use the **[DayWeekTimeScale](http://msdn.microsoft.com/library/94f2aad5-6699-82e9-40a4-3c3c13d80684%28Office.15%29.aspx)** to configure the time scale used to display Outlook items within the view. If you set the **CalendarViewMode** to **olCalendarViewMultiDay**, you can use the **[DaysInMultiDayMode](http://msdn.microsoft.com/library/1dcb2a69-93b9-432e-56ca-7e39b040dc6f%28Office.15%29.aspx)** property to determine the number of days to display in the view.

If you set the  **CalendarViewMode** to **olCalendarViewMonth**, you can use the **[MonthFont](http://msdn.microsoft.com/library/b69d1690-d1a8-dbc0-3de4-86a8eb98a471%28Office.15%29.aspx)** property to configure the fonts used to display the month and day labels and the **[MonthShowEndTime](http://msdn.microsoft.com/library/19a92965-aa85-e1f6-9db6-ce85c7980d75%28Office.15%29.aspx)** to indicate whether the end time for is displayed in the view.

You can also configure how Outlook items appear within the  **CalendarView** object. Use the **[BoldSubjects](http://msdn.microsoft.com/library/b7bf5518-68d0-0a8a-98b2-94c267855f2b%28Office.15%29.aspx)** property to indicate whether subjects for Outlook items are displayed in bold and the **[BoldDatesWithItems](http://msdn.microsoft.com/library/4928abe0-c650-f09e-796c-5d931a1c6aae%28Office.15%29.aspx)** property to indicate whether dates in the Date Navigator that contain Outlook items are displayed in bold. Use the **[Filter](http://msdn.microsoft.com/library/c62e9521-e1aa-bfe8-5774-25c3227973b5%28Office.15%29.aspx)** property to determine which Outlook items to display in the view.

The definition for each  **CalendarView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](http://msdn.microsoft.com/library/f188b827-77c6-71da-0b36-972b16b843a8%28Office.15%29.aspx)** property to work with the XML definition for the **CalendarView** object.

Use the  **[Apply](http://msdn.microsoft.com/library/274edf67-7a3b-8132-3990-a07fa30b5024%28Office.15%29.aspx)** method to apply any changes made to the **CalendarView** object to the current view. Use the **[Save](http://msdn.microsoft.com/library/19cea2c8-39bd-875c-2cde-50d19f25f73b%28Office.15%29.aspx)** method to persist any changes made to the **CalendarView** object. Use the **[LockUserChanges](http://msdn.microsoft.com/library/b5102728-a0d4-6eb6-15ae-916644fe6f9c%28Office.15%29.aspx)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **CalendarView** objects, but you cannot delete them. Use the **[Delete](http://msdn.microsoft.com/library/90a07253-844e-d40b-6450-c97a9cf85c58%28Office.15%29.aspx)** method to delete a custom **CalendarView** object. Use the **[Reset](http://msdn.microsoft.com/library/222b2537-4d70-6a12-97f2-5034a262655b%28Office.15%29.aspx)** method to reset the properties of a built-in **CalendarView** object to their default values.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **CalendarView** object to show a single day, using an 8-point Verdana font to display items and a 16-point Verdana font to display time values and the Tasks header within the view.


```
Sub ConfigureDayViewFonts() 
 Dim objView As CalendarView 
 
 ' Check if the current view is a calendar view. 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 olCalendarView Then 
 
 ' Obtain a CalendarView object reference for the 
 ' current calendar view. 
 Set objView = _ 
 Application.ActiveExplorer.CurrentView 
 
 With objView 
 ' Set the calendar view to show a 
 ' single day. 
 .CalendarViewMode = olCalendarViewDay 
 
 ' Set the DayWeekFont to 8-point Verdana. 
 .DayWeekFont.Name = "Verdana" 
 .DayWeekFont.Size = 8 
 
 ' Set the DayWeekTimeFont to 16-point Verdana. 
 .DayWeekTimeFont.Name = "Verdana" 
 .DayWeekTimeFont.Size = 16 
 
 ' Save the calendar view. 
 .Save 
 End With 
 End If 
End Sub 

```


## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[CalendarView Object Members](http://msdn.microsoft.com/library/c8ee2de7-d65c-90b2-0d63-5fa584c7c500%28Office.15%29.aspx)
