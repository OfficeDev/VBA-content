---
title: SharingItem.RemoteID Property (Outlook)
keywords: vbaol11.chm695
f1_keywords:
- vbaol11.chm695
ms.prod: outlook
api_name:
- Outlook.SharingItem.RemoteID
ms.assetid: 07b0ba28-f560-7cee-bfc9-38fa073d8669
ms.date: 06/08/2017
---


# SharingItem.RemoteID Property (Outlook)

Returns a  **String** that represents the unique identifier of the sharing context for a **[SharingItem](sharingitem-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **RemoteID**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

This property contains either a GUID or EntryID for the sharing context contained within the  **SharingItem** object.

This property is set to an empty string if the  **[Type](sharingitem-type-property-outlook.md)** property of the **SharingItem** object is set to **olSharingMsgTypeRequest** .


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

