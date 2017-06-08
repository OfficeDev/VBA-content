---
title: AppointmentItem.GlobalAppointmentID Property (Outlook)
keywords: vbaol11.chm924
f1_keywords:
- vbaol11.chm924
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.GlobalAppointmentID
ms.assetid: 3a5e210a-5298-8977-d6e4-dc49a59bdd78
ms.date: 06/08/2017
---


# AppointmentItem.GlobalAppointmentID Property (Outlook)

Returns a  **String** value that represents a unique global identifier for the **[AppointmentItem](appointmentitem-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **GlobalAppointmentID**

 _expression_ An expression that returns an **AppointmentItem** object.


## Remarks

There are situations where the entry ID of  **AppointmentItem** objects may change, such as when an item is moved to a different folder or to a different store. Entry IDs can also change when a user performs certain functions in Outlook, such as exporting and then reimporting data.

Therefore, each Outlook appointment item is assigned a Global Object ID, a unique global identifier which does not change during those situations. The Global Object ID is a MAPI property that Outlook uses to correlate meeting updates and responses with a particular meeting on the calendar. The Global Object ID is the same across all copies of the item.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

