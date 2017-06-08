---
title: AppointmentItem.EndTimeZone Property (Outlook)
keywords: vbaol11.chm3276
f1_keywords:
- vbaol11.chm3276
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.EndTimeZone
ms.assetid: 8f33d93f-c0fe-fda1-608d-dec7fb86c732
ms.date: 06/08/2017
---


# AppointmentItem.EndTimeZone Property (Outlook)

Returns or sets a  **[TimeZone](timezone-object-outlook.md)** value that corresponds to the end time of the appointment. Read/write.


## Syntax

 _expression_ . **EndTimeZone**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

The time zone information is used to map the appointment to the correct UTC time when the appointment is saved, and into the correct local time when the item is displayed in the calendar.

 Changing **EndTimeZone** affects the value of **[AppointmentItem.End](appointmentitem-end-property-outlook.md)** which is always represented in the local time zone, **[Application.TimeZones.CurrentTimeZone](timezones-currenttimezone-property-outlook.md)** .

Depending on the circumstances, changing the  **EndTimeZone** may or may not cause Outlook to recalculate and update the **[AppointmentItem.EndInEndTimeZone](appointmentitem-endinendtimezone-property-outlook.md)** .

As an example, in the appointment inspector, if you are the organizer of an appointment with a start time at 1 P.M. EST and end time at 3 P.M. EST, changing the appointment to have an  **EndTimeZone** of PST will result in an appointment lasting from 1 P.M. EST to 3 P.M. PST, with the **EndInEndTimeZone** remaining as 3 P.M. However, if you are not the organizer, then changing the **EndTimeZone** from EST to PST will cause Outlook to recalculate and update the **EndInEndTimeZone** , and the appointment will last from 1 P.M. EST to 12 P.M. PST.

Another example is changing the  **EndTimeZone** resulting in an appointment end time that occurs before a previously set appointment start time, in which case Outlook will recalculate and update the **EndInEndTimeZone** . For example, an appointment with a start time at 1 P.M. PST and end time at 3 P.M. PST has its **EndTimeZone** changed to EST. If Outlook did not recalculate the **EndInEndTimeZone** , the appointment would have an end time at 3 P.M. EST, which is equivalent to 12 P.M. PST, and which would occur before the start time of 1 P.M. PST. In practice, however, changing the **EndTimeZone** would result in Outlook recalculating and updating the **EndInEndTimeZone** to 6 P.M. (in the **EndTimeZone** EST).


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

