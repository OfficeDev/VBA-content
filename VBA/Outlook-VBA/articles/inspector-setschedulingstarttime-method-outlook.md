---
title: Inspector.SetSchedulingStartTime Method (Outlook)
keywords: vbaol11.chm3555
f1_keywords:
- vbaol11.chm3555
ms.prod: outlook
api_name:
- Outlook.Inspector.SetSchedulingStartTime
ms.assetid: 22e6358a-9dba-7edb-fc5f-3a2a7326bece
ms.date: 06/08/2017
---


# Inspector.SetSchedulingStartTime Method (Outlook)

Sets the start time for a meeting item in the free/busy grid on the  **Scheduling Assistant** tab of the inspector.


## Syntax

 _expression_ . **SetSchedulingStartTime**( **_Start_** )

 _expression_ A variable that represents an **[Inspector](inspector-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Required| **Date**|The beginning of the time range that the  **Scheduling Assistant** tab of the inspector displays free/busy times for meeting attendees.|

## Remarks

The object specified by the  **[CurrentItem](inspector-currentitem-property-outlook.md)** property of the parent **[Inspector](inspector-object-outlook.md)** object must be an **[AppointmentItem](appointmentitem-object-outlook.md)** or **[MeetingItem](meetingitem-object-outlook.md)** . The **Scheduling Assistant** tab must be displayed in the inspector, otherwise Microsoft Outlook raises an error. If Outlook cannot display the **Scheduling Assistant** tab for that item type, Outlook displays the following error: **The scheduling start time can only be set when the Scheduling Assistant is displayed on a meeting item.**


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) shows how to use the  **SetSchedulingStartTime** method to set the scheduling start time on the **Scheduling Assistant** tab of an **AppointmentItem** . The appointment start time is set to one month from now, and the scheduling start time is also set to one month from now.


```vb
Sub DemoSetSchedulingStartTime() 
 
 Dim oAppt As Outlook.AppointmentItem 
 
 Dim oInsp As Outlook.inspector 
 
 
 
 ' Create and display appointment. 
 
 Set oAppt = Application.CreateItem(olAppointmentItem) 
 
 oAppt.MeetingStatus = olMeeting 
 
 oAppt.Subject = "Test Appointment" 
 
 oAppt.Start = DateAdd("m", 1, Now) 
 
 ' Display the appointment in the Appointment tab of the inspector. 
 
 oAppt.Display 
 
 
 
 Set oInsp = oAppt.GetInspector 
 
 ' Switch to the Scheduling Assistant tab in that inspector. 
 
 oInsp.SetCurrentFormPage ("Scheduling Assistant") 
 
 ' Set the appointment start time in the Scheduling Assistant. 
 
 oInsp.SetSchedulingStartTime (DateAdd("m", 1, Now)) 
 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

