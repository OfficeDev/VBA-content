---
title: RecurrencePattern.EndTime Property (Outlook)
keywords: vbaol11.chm276
f1_keywords:
- vbaol11.chm276
ms.prod: outlook
api_name:
- Outlook.RecurrencePattern.EndTime
ms.assetid: 7babda13-9e57-4c80-1ab3-56025753ed9d
ms.date: 06/08/2017
---


# RecurrencePattern.EndTime Property (Outlook)

Returns or sets a  **Date** indicating the end time for a recurrence pattern. Read/write.


## Syntax

 _expression_ . **EndTime**

 _expression_ A variable that represents a **RecurrencePattern** object.


## Remarks

This property is only valid for appointments. 

When you create a  **[RecurrencePattern](recurrencepattern-object-outlook.md)** object and no time zones have been specified for the appointment, **[StartTime](recurrencepattern-starttime-property-outlook.md)** and **EndTime** of the **RecurrencePattern** object are based on the time zone specified by **[Application.TimeZones.CurrentTimeZone](timezones-currenttimezone-property-outlook.md)** .

If you want to create a recurring appointment for a particular time zone, you should first create an  **[AppointmentItem](appointmentitem-object-outlook.md)** , set **[AppointmentItem.StartTimeZone](appointmentitem-starttimezone-property-outlook.md)** , and then call **[AppointmentItem.GetRecurrencePattern](appointmentitem-getrecurrencepattern-method-outlook.md)** . The **RecurrencePattern** object returned will have both **StartTime** and **EndTime** based on the time zone specified by **AppointmentItem.StartTimeZone** . Note that in the **Appointment Recurrence** dialog box, the time indicated as **Start** is **RecurrencePattern.StartTime** which is based on **AppointmentItem.StartTimeZone** , but the time indicated as **End** is not always the same as **RecurrencePattern.EndTime** which is based on **AppointmentItem.StartTimeZone** ; the displayed time value is based on **[AppointmentItem.EndTimeZone](appointmentitem-endtimezone-property-outlook.md)** .


## See also


#### Concepts


[RecurrencePattern Object](recurrencepattern-object-outlook.md)

