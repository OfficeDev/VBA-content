---
title: SharingItem.Parent Property (Outlook)
keywords: vbaol11.chm596
f1_keywords:
- vbaol11.chm596
ms.prod: outlook
api_name:
- Outlook.SharingItem.Parent
ms.assetid: 78d6d287-9623-0ed0-eab6-75a0a57d0c6c
ms.date: 06/08/2017
---


# SharingItem.Parent Property (Outlook)

Returns the parent  **Object** of the specified **[SharingItem](sharingitem-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

If the  **SharingItem** was just created, this property returns a **[Folder](folder-object-outlook.md)** object representing the **Inbox** folder. Otherwise, this property returns a **Folder** object representing the folder in which the **SharingItem** was saved.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

