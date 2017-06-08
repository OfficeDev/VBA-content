---
title: AppointmentItem.MeetingStatus Property (Outlook)
keywords: vbaol11.chm883
f1_keywords:
- vbaol11.chm883
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.MeetingStatus
ms.assetid: cfd970cd-df6c-4537-0a17-b5adab3b667f
ms.date: 06/08/2017
---


# AppointmentItem.MeetingStatus Property (Outlook)

Returns or sets an  **[OlMeetingStatus](olmeetingstatus-enumeration-outlook.md)** constant specifying the meeting status of the appointment. Read/write.


## Syntax

 _expression_ . **MeetingStatus**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

Use this property to make a  **[MeetingItem](meetingitem-object-outlook.md)** object available for the appointment.


## Example

This Visual Basic for Applications example uses  **[CreateItem](application-createitem-method-outlook.md)** to create an appointment and uses **[MeetingStatus](appointmentitem-meetingstatus-property-outlook.md)** to set the meeting status to "Meeting" to turn it into a meeting request with both a required and an optional attendee.


```vb
Sub CreateAppt() 
 
 Dim myItem As Object 
 
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 
 myItem.MeetingStatus = olMeeting 
 
 myItem.Subject = "Strategy Meeting" 
 
 myItem.Location = "Conference Room B" 
 
 myItem.Start = #9/24/1997 1:30:00 PM# 
 
 myItem.Duration = 90 
 
 Set myRequiredAttendee = myItem.Recipients.Add("Nate Sun") 
 
 myRequiredAttendee.Type = olRequired 
 
 Set myOptionalAttendee = myItem.Recipients.Add("Kevin Kennedy") 
 
 myOptionalAttendee.Type = olOptional 
 
 Set myResourceAttendee = myItem.Recipients.Add("Conference Room B") 
 
 myResourceAttendee.Type = olResource 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

