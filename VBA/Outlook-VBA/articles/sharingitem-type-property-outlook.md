---
title: SharingItem.Type Property (Outlook)
keywords: vbaol11.chm701
f1_keywords:
- vbaol11.chm701
ms.prod: outlook
api_name:
- Outlook.SharingItem.Type
ms.assetid: 1077b74f-38ee-8932-792d-64033bc66525
ms.date: 06/08/2017
---


# SharingItem.Type Property (Outlook)

Returns or sets an  **[OlSharingMsgType](olsharingmsgtype-enumeration-outlook.md)** constant that indicates the type of sharing message represented by the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **Type**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

An error occurs if you attempt to set this property after the sharing message has been sent or received, or if you attempt to set this property to  **olSharingMsgTypeResponseAllow** or **olSharingMsgTypeResponseDeny** .


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

