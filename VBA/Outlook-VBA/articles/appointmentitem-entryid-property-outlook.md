---
title: AppointmentItem.EntryID Property (Outlook)
keywords: vbaol11.chm851
f1_keywords:
- vbaol11.chm851
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.EntryID
ms.assetid: 8f4160de-0840-902a-589e-bce80797b6f5
ms.date: 06/08/2017
---


# AppointmentItem.EntryID Property (Outlook)

Returns a  **String** representing the unique Entry ID of the object. Read-only.


## Syntax

 _expression_ . **EntryID**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagEntryId** .

A MAPI store provider assigns a unique ID string when an item is created in its store. Therefore, the  **EntryID** property is not set for an Outlook item until it is saved or sent. The Entry ID changes when an item is moved into another store, for example, from your **Inbox** to a Microsoft Exchange Server public folder, or from one Personal Folders (.pst) file to another .pst file. Solutions should not depend on the **EntryID** property to be unique unless items will not be moved. The **EntryID** property returns a MAPI long-term Entry ID. For more information about long- and short-term EntryIDs, search http://msdn.microsoft.com for **PidTagEntryId** .

Furthermore, when you call the  **AppointmentItem.Respond** method with the **olMeetingAccepted** or **olMeetingTentative** parameter, Outlook will create a new appointment item that duplicates the original appointment item. The new item will have a different Entry ID. Outlook will then remove the original item. You should no longer use the Entry ID of the original item, but instead use **AppointmentItem.EntryID** to obtain the Entry ID for the new item for any subsequent needs. This is to ensure that this appointment item will be properly synchronized on your calendar if more than one client computer accesses your calendar but may be offline using the cache mode occasionally.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

