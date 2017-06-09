---
title: AppointmentItem.StartUTC Property (Outlook)
keywords: vbaol11.chm3271
f1_keywords:
- vbaol11.chm3271
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.StartUTC
ms.assetid: 8bfbf95f-bd88-acdc-f592-c41b454afe4b
ms.date: 06/08/2017
---


# AppointmentItem.StartUTC Property (Outlook)

Returns or sets a  **Date** value that represents the start date and time of the appointment expressed in the Coordinated Univeral Time (UTC) standard. Read/write.


## Syntax

 _expression_ . **StartUTC**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

Changing the value for the  **[AppointmentItem.Start](appointmentitem-start-property-outlook.md)** property or the **[AppointmentItem.StartTimeZone](appointmentitem-starttimezone-property-outlook.md)** property will cause Outlook to recalculate the value of **StartUTC** .


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

