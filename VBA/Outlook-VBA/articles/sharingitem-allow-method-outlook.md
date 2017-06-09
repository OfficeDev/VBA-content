---
title: SharingItem.Allow Method (Outlook)
keywords: vbaol11.chm691
f1_keywords:
- vbaol11.chm691
ms.prod: outlook
api_name:
- Outlook.SharingItem.Allow
ms.assetid: 8f47e300-86d0-b90c-a41d-05bddec743f4
ms.date: 06/08/2017
---


# SharingItem.Allow Method (Outlook)

Allows a sharing request and sends a sharing response to the sender of the  **[SharingItem](sharingitem-object-outlook.md)** .


## Syntax

 _expression_ . **Allow**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

The  **Allow** method can only be called on **SharingItem** objects with a **[Type](sharingitem-type-property-outlook.md)** property value of **olSharingMsgTypeRequest** or **olSharingMsgTypeInviteAndRequest** .

The  **Type** property of the sharing response sent when this method is called is set to **olSharingMsgTypeResponseAllow** .


 **Note**  Sharing is allowed immediately after this method is called, regardless of whether the sharing response was received.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

