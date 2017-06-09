---
title: AppointmentItem.Resources Property (Outlook)
keywords: vbaol11.chm899
f1_keywords:
- vbaol11.chm899
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Resources
ms.assetid: 9b989d76-6897-cd2d-9156-fd7391dad8c1
ms.date: 06/08/2017
---


# AppointmentItem.Resources Property (Outlook)

Returns a semicolon-delimited  **String** of resource names for the meeting. Read/write.


## Syntax

 _expression_ . **Resources**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

This property contains the display names only. The  **[Recipients](recipients-object-outlook.md)** collection should be used to modify the resource recipients. Resources are added as **[BCC](mailitem-bcc-property-outlook.md)** recipients to the collection.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

