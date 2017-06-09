---
title: Create an Appointment as a Meeting on the Calendar
ms.prod: outlook
ms.assetid: 130b6ae1-d1a4-3805-7e9c-75543b93fff5
ms.date: 06/08/2017
---


# Create an Appointment as a Meeting on the Calendar

This topic shows a Visual Basic for Applications (VBA) procedure,  `CreateAppt`, that programmatically creates an appointment, sets various properties, and sends the appointment to request a meeting.  `CreateAppt` uses the **[CreateItem](application-createitem-method-outlook.md)** method to create an **[AppointmentItem](appointmentitem-object-outlook.md)** object. It sets the **[MeetingStatus](appointmentitem-meetingstatus-property-outlook.md)** property of the **AppointmentItem** to **olMeeting** to indicate the appointment as a meeting request, and sets a required attendee, an optional attendee, and a meeting location as a resource. The example then displays and sends the appointment item.


```vb
Sub CreateAppt() 
 Dim myItem As Object 
 Dim myRequiredAttendee, myOptionalAttendee, myResourceAttendee As Outlook.Recipient 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 myItem.MeetingStatus = olMeeting 
 myItem.Subject = "Strategy Meeting" 
 myItem.Location = "Conf Rm All Stars" 
 myItem.Start = #9/24/2009 1:30:00 PM# 
 myItem.Duration = 90 
 Set myRequiredAttendee = myItem.Recipients.Add("Nate Sun") 
 myRequiredAttendee.Type = olRequired 
 Set myOptionalAttendee = myItem.Recipients.Add("Kevin Kennedy") 
 myOptionalAttendee.Type = olOptional 
 Set myResourceAttendee = myItem.Recipients.Add("Conf Rm All Stars") 
 myResourceAttendee.Type = olResource 
 myItem.Display 
 myItem.Send 
End Sub
```


