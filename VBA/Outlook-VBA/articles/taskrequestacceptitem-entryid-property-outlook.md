---
title: TaskRequestAcceptItem.EntryID Property (Outlook)
keywords: vbaol11.chm1783
f1_keywords:
- vbaol11.chm1783
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.EntryID
ms.assetid: e8304c17-cd0d-8379-4513-ed9d4e40b319
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.EntryID Property (Outlook)

Returns a  **String** representing the unique Entry ID of the object. Read-only.


## Syntax

 _expression_ . **EntryID**

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagEntryId** .

A MAPI store provider assigns a unique ID string when an item is created in its store. Therefore, the  **EntryID** property is not set for an Outlook item until it is saved or sent. The Entry ID changes when an item is moved into another store, for example, from your **Inbox** to a Microsoft Exchange Server public folder, or from one Personal Folders (.pst) file to another .pst file. Solutions should not depend on the **EntryID** property to be unique unless items will not be moved. The **EntryID** property returns a MAPI long-term Entry ID. For more information about long- and short-term EntryIDs, search http://msdn.microsoft.com for **PidTagEntryId** .


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

