---
title: AppointmentItem.IsConflict Property (Outlook)
keywords: vbaol11.chm918
f1_keywords:
- vbaol11.chm918
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.IsConflict
ms.assetid: d0c14fa2-6bfe-29e8-e68b-3eff01a8bd70
ms.date: 06/08/2017
---


# AppointmentItem.IsConflict Property (Outlook)

Returns a  **Boolean** that determines if the item on the local computer is different from the copy on the server. Read-only.


## Syntax

 _expression_ . **IsConflict**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True** .

This property does not indicate whether the appointment item has a time conflict with another appointment in the calendar.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

