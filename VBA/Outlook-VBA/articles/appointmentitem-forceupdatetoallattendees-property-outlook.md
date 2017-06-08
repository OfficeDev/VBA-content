---
title: AppointmentItem.ForceUpdateToAllAttendees Property (Outlook)
keywords: vbaol11.chm3226
f1_keywords:
- vbaol11.chm3226
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ForceUpdateToAllAttendees
ms.assetid: fe926820-2694-9aa3-8359-cc2ed3ac2f32
ms.date: 06/08/2017
---


# AppointmentItem.ForceUpdateToAllAttendees Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether updates to the[AppointmentItem](appointmentitem-object-outlook.md) object should be sent to all attendees. Read/write.


## Syntax

 _expression_ . **ForceUpdateToAllAttendees**

 _expression_ An expression that returns an **AppointmentItem** object.


## Remarks

Normally, updates are sent to attendees only if the time or location of an appointment item is changed. Setting this property to  **True** forces an update to be sent to all attendees, even if no changes to the time or location have occurred.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

