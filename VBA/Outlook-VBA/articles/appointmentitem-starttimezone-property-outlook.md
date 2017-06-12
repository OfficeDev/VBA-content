---
title: AppointmentItem.StartTimeZone Property (Outlook)
keywords: vbaol11.chm3275
f1_keywords:
- vbaol11.chm3275
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.StartTimeZone
ms.assetid: 3259fa91-5f6c-b899-9bfc-2ac669911271
ms.date: 06/08/2017
---


# AppointmentItem.StartTimeZone Property (Outlook)

Returns or sets a  **[TimeZone](timezone-object-outlook.md)** value that corresponds to the time zone for the start time of the appointment. Read/write.


## Syntax

 _expression_ . **StartTimeZone**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

The time zone information is used to map the appointment to the correct UTC time when the appointment is saved, and into the correct local time when the item is displayed in the calendar.

Changing  **StartTimeZone** affects the value of **[AppointmentItem.Start](appointmentitem-start-property-outlook.md)** which is always represented in the local time zone, **[Application.TimeZones.CurrentTimeZone](timezones-currenttimezone-property-outlook.md)** .

Depending on the circumstances, changing the  **StartTimeZone** may or may not cause Outlook to recalculate and update the **[AppointmentItem.StartInStartTimeZone](appointmentitem-startinstarttimezone-property-outlook.md)** .

As an example, in the appointment inspector, if you are the organizer of an appointment with a start time at 1 P.M. PST and end time at 3 P.M. PST, changing the appointment to have an  **StartTimeZone** of EST will result in an appointment lasting from 1 P.M. EST to 3 P.M. PST, with the **StartInStartTimeZone** remaining as 1 P.M. However, if you are not the organizer, then changing the **StartTimeZone** from PST to EST will cause Outlook to recalculate and update the **StartInStartTimeZone** , and the appointment will last from 4 P.M. EST to 3 P.M. PST.

Another example is changing the  **StartTimeZone** resulting in an appointment end time that occurs before a previously set appointment start time, in which case Outlook will recalculate and update the **StartInStartTimeZone** . For example, an appointment with a start time at 1 P.M. EST and end time at 3 P.M. EST has its **StartTimeZone** changed to PST. If Outlook did not recalculate the **StartInStartTimeZone** , the appointment would have a start time at 1 P.M. PST, which is equivalent to 4 P.M. EST, and which would occur before the end time of 3 P.M. EST. In practice, however, changing the **StartTimeZone** would result in Outlook recalculating and updating the **StartInStartTimeZone** to 10 A.M. (in the **StartTimeZone** PST).


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

