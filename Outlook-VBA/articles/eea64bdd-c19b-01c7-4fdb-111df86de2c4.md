
# AppointmentItem.Duration Property (Outlook)

 **Last modified:** July 28, 2015

Returns or sets a  **Long** indicating the duration (in minutes) of the ** [AppointmentItem](204a409d-654e-27aa-643a-8344c631b82d.md)**. Read/write.

## Syntax

 _expression_. **Duration**

 _expression_A variable that represents an  **AppointmentItem** object.


## Example

This Visual Basic for Applications example uses  ** [Application.CreateItem](e5fbf367-db16-5042-823e-68e6b805e612.md)**to create an appointment and uses  ** [AppointmentItem.MeetingStatus](cfd970cd-df6c-4537-0a17-b5adab3b667f.md)**to set the meeting status to "Meeting" to turn it into a meeting request with both a required and an optional attendee.


```
Sub ScheduleMeeting() 
 
 Dim myItem as AppointmentItem 
 
 Dim myRequiredAttendee As Recipient 
 
 Dim myOptionalAttendee As Recipient 
 
 Dim myResourceAttendee As Recipient 
 
 
 
 Set myItem = Application.CreateItem(olAppointmentItem) 
 
 myItem.MeetingStatus = olMeeting 
 
 myItem.Subject = "Strategy Meeting" 
 
 myItem.Location = "Conference Room B" 
 
 myItem.Start = #9/24/2002 1:30:00 PM# 
 
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


 [AppointmentItem Object](204a409d-654e-27aa-643a-8344c631b82d.md)
#### Other resources


 [AppointmentItem Object Members](c72c459d-6d3c-7a05-aa4a-b1b767ddc0b2.md)
