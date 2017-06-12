---
title: AppointmentItem.Start Property (Outlook)
keywords: vbaol11.chm902
f1_keywords:
- vbaol11.chm902
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Start
ms.assetid: 1b869a9d-fe08-6efb-48b1-f33cf9ea0024
ms.date: 06/08/2017
---


# AppointmentItem.Start Property (Outlook)

Returns or sets a  **Date** indicating the starting date and time for the Outlook item. Read/write.


## Syntax

 _expression_ . **Start**

 _expression_ A variable that represents an **AppointmentItem** object.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](application-createitem-method-outlook.md)** to create an appointment and uses **[MeetingStatus](appointmentitem-meetingstatus-property-outlook.md)** to set the meeting status to "Meeting" and to make it a meeting request with both a required and an optional attendee.


```vb
Sub ScheduleMeeting() 
 
 Dim myItem as Outlook.AppointmentItem 
 
 Dim myRequiredAttendee As Outlook.Recipient 
 
 Dim myOptionalAttendee As Outlook.Recipient 
 
 Dim myResourceAttendee As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 
 myItem.MeetingStatus = olMeeting 
 
 myItem.Subject = "Strategy Meeting" 
 
 myItem.Location = "Conference Room B" 
 
 myItem.Start = #9/24/2003 1:30:00 PM# 
 
 myItem.Duration = 90 
 
 Set myRequiredAttendee = myItem.Recipients.Add ("Nate Sun") 
 
 myRequiredAttendee.Type = olRequired 
 
 Set myOptionalAttendee = myItem.Recipients.Add ("Kevin Kennedy") 
 
 myOptionalAttendee.Type = olOptional 
 
 Set myResourceAttendee = myItem.Recipients.Add("Conference Room B") 
 
 myResourceAttendee.Type = olResource 
 
 myItem.Display 
 
End Sub
```


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)
#### Other resources



[How to: Import Appointment XML Data into Outlook Appointment Objects](http://msdn.microsoft.com/library/ecfd3849-877b-01ad-2b76-1a54e980f6e2%28Office.15%29.aspx)

