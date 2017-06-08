---
title: SharingItem.EntryID Property (Outlook)
keywords: vbaol11.chm606
f1_keywords:
- vbaol11.chm606
ms.prod: outlook
api_name:
- Outlook.SharingItem.EntryID
ms.assetid: fca59e3a-8a62-244b-65f2-19b5b69b209b
ms.date: 06/08/2017
---


# SharingItem.EntryID Property (Outlook)

Returns a  **String** representing the unique Entry ID of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **EntryID**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagEntryId** .

A MAPI store provider assigns a unique ID string when an item is created in its store. Therefore, the  **EntryID** property is not set for an Outlook item until it is saved or sent. The Entry ID changes when an item is moved into another store, for example, from your **Inbox** to a Microsoft Exchange Server public folder, or from one Personal Folders (.pst) file to another .pst file. Solutions should not depend on the **EntryID** property to be unique unless items will not be moved. The **EntryID** property returns a MAPI long-term Entry ID. For more information about long- and short-term EntryIDs, search http://msdn.microsoft.com for **PidTagEntryId** .


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

