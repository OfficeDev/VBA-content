---
title: MeetingItem.GetAssociatedAppointment Method (Outlook)
keywords: vbaol11.chm1455
f1_keywords:
- vbaol11.chm1455
ms.prod: outlook
api_name:
- Outlook.MeetingItem.GetAssociatedAppointment
ms.assetid: 8344d40d-5c1d-ead3-87cb-fd795b831712
ms.date: 06/08/2017
---


# MeetingItem.GetAssociatedAppointment Method (Outlook)

Returns an  **[AppointmentItem](appointmentitem-object-outlook.md)** object that represents the appointment associated with the meeting request.


## Syntax

 _expression_ . **GetAssociatedAppointment**( **_AddToCalendar_** )

 _expression_ A variable that represents a **MeetingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AddToCalendar_|Required| **Boolean**| **True** to add the meeting to the default **Calendar** folder.|

### Return Value

An  **AppointmentItem** object that represents the associated appointment.


## Example

This Visual Basic for Applications (VBA) example finds a  **[MeetingItem](meetingitem-object-outlook.md)** in the default **Inbox** folder that has not been responded to yet and adds the associated appointment to the **Calendar** folder. It then responds to the sender by accepting the meeting.


```vb
Sub AcceptMeeting() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myMtgReq As Outlook.MeetingItem 
 
 Dim myAppt As Outlook.AppointmentItem 
 
 Dim myMtg As Outlook.MeetingItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myMtgReq = myFolder.Items.Find("[MessageClass] = 'IPM.Schedule.Meeting.Request'") 
 
 If TypeName(myMtgReq) <> "Nothing" Then 
 
 Set myAppt = myMtgReq.GetAssociatedAppointment(True) 
 
 Set myMtg = myAppt.Respond(olResponseAccepted, True) 
 
 myMtg.Send 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

