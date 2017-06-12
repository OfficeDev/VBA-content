---
title: CalendarView.SelectedEndTime Property (Outlook)
keywords: vbaol11.chm3314
f1_keywords:
- vbaol11.chm3314
ms.prod: outlook
api_name:
- Outlook.CalendarView.SelectedEndTime
ms.assetid: cf617cf4-9c71-96ca-e8f5-52fa4596cb6b
ms.date: 06/08/2017
---


# CalendarView.SelectedEndTime Property (Outlook)

Returns a  **Date** that represents the end time of a selection in the **[CalendarView](calendarview-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **SelectedEndTime**

 _expression_ A variable that represents a **CalendarView** object.


## Remarks

The intent of the  **[SelectedStartTime](calendarview-selectedstarttime-property-outlook.md)** and the **SelectedEndTime** properties is to replicate, programmatically, the way that users create an appointment in the Microsoft Outlook user interface. Typically, a user selects a time range in the calendar view and then creates a new appointment by double clicking the selection or by clicking **New Appointment** in the **Home** tab of the ribbon. With these two properties of the **CalendarView** object, you can obtain the start time and the end time of any selection in that view programmatically. You can then programmatically create the **[AppointmentItem](appointmentitem-object-outlook.md)** object, setting the **[Start](appointmentitem-start-property-outlook.md)** and **[End](appointmentitem-end-property-outlook.md)** properties of the **AppointmentItem** object to the **SelectedStartTime** and **SelectedEndTime** properties respectively to reflect any user selection in the calendar view.

If the selection in the calendar view is a time range and is not an item,  **SelectedEndTime** returns a **Date** value equal to the end time of the selection.

If one or more items are selected in the calendar view,  **SelectedEndTime** returns a **Date** value equal to the end time of the first item in the selection of the explorer that displays the calendar view. That selection is specified by the **[Selection](explorer-selection-property-outlook.md)** property of the **[Explorer](explorer-object-outlook.md)** object.

To use this property on a  **CalendarView** object, obtain the **CalendarView** object from the **[CurrentView](explorer-currentview-property-outlook.md)** property of the active **[Explorer](explorer-object-outlook.md)** object (which can be returned by the **[Application.ActiveExplorer](application-activeexplorer-method-outlook.md)** method). There is a known issue with using this property on an **CalendarView** object obtained otherwise - using the **[CurrentView](folder-currentview-property-outlook.md)** property of the current **[Folder](folder-object-outlook.md)** object (returned by the **[Application.ActiveExplorer.CurrentFolder](explorer-currentfolder-property-outlook.md)** property).


## Example

The following code samples, in Visual Basic for Applications (VBA) and C#, show how to use the  **SelectedStartTime** and **SelectedEndTime** properties of the calendar view of the active explorer to initialize the start and end times of a new appointment. The following code sample is in VBA.


```vb
Sub CreateAppointmentUsingSelectedTime() 
 Dim datStart As Date 
 Dim datEnd As Date 
 Dim oView As Outlook.view 
 Dim oCalView As Outlook.CalendarView 
 Dim oExpl As Outlook.Explorer 
 Dim oFolder As Outlook.folder 
 Dim oAppt As Outlook.AppointmentItem 
 Const datNull As Date = #1/1/4501# 
 
 ' Obtain the calendar view using 
 ' Application.ActiveExplorer.CurrentFolder.CurrentView. 
 ' If you use oExpl.CurrentFolder.CurrentView, 
 ' this code will not operate as expected. 
 Set oExpl = Application.ActiveExplorer 
 Set oFolder = Application.ActiveExplorer.CurrentFolder 
 Set oView = oExpl.CurrentView 
 
 ' Check whether the active explorer is displaying a calendar view. 
 If oView.ViewType = olCalendarView Then 
 Set oCalView = oExpl.currentView 
 ' Create the appointment using the values in 
 ' the SelectedStartTime and SelectedEndTime properties as 
 ' appointment start and end times. 
 datStart = oCalView.SelectedStartTime 
 datEnd = oCalView.SelectedEndTime 
 Set oAppt = oFolder.items.Add("IPM.Appointment") 
 If datStart <> datNull And datEnd <> datNull Then 
 oAppt.Start = datStart 
 oAppt.End = datEnd 
 End If 
 oAppt.Display 
 End If 
End Sub
```

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.




```C#
private void CreateAppointmentUsingSelectedTime() 
{ 
 DateTime dateNull = 
 new DateTime(4501, 1, 1, 0, 0, 0); 
 Outlook.Explorer expl = Application.ActiveExplorer(); 
 Outlook.Folder folder = expl.CurrentFolder as Outlook.Folder; 
 Outlook.View view = expl.CurrentView as Outlook.View; 
 if (view.ViewType == Outlook.OlViewType.olCalendarView) 
 { 
 Outlook.CalendarView calView = view as Outlook.CalendarView; 
 DateTime dateStart = calView.SelectedStartTime; 
 DateTime dateEnd = calView.SelectedEndTime; 
 Outlook.AppointmentItem appt = 
 folder.Items.Add("IPM.Appointment") 
 as Outlook.AppointmentItem; 
 if (dateStart != dateNull &;&; dateEnd != dateNull) 
 { 
 appt.Start = dateStart; 
 appt.End = dateEnd; 
 } 
 appt.Display(false); 
 } 
} 

```


## See also


#### Concepts


[CalendarView Object](calendarview-object-outlook.md)

