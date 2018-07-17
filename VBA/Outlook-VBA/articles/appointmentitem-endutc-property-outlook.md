---
title: AppointmentItem.EndUTC Property (Outlook)
keywords: vbaol11.chm3272
f1_keywords:
- vbaol11.chm3272
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.EndUTC
ms.assetid: c741e893-3a29-10cc-0730-a0796d8c2e4c
ms.date: 06/08/2017
---


# AppointmentItem.EndUTC Property (Outlook)

Returns or sets a  **Date** value that represents the end date and time of the appointment expressed in the Coordinated Univeral Time (UTC) standard. Read/write.


## Syntax

 _expression_ . **EndUTC**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

Changing the value for the  **[AppointmentItem.End](appointmentitem-end-property-outlook.md)** property or the **[AppointmentItem.EndTimeZone](appointmentitem-endtimezone-property-outlook.md)** property will cause Outlook to recalculate the **EndUTC** .


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

