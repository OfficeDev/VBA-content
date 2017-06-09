---
title: SharingItem.ReadReceiptRequested Property (Outlook)
keywords: vbaol11.chm643
f1_keywords:
- vbaol11.chm643
ms.prod: outlook
api_name:
- Outlook.SharingItem.ReadReceiptRequested
ms.assetid: fa8f3b1c-77a6-1620-f0dd-7cf0bd6f64a3
ms.date: 06/08/2017
---


# SharingItem.ReadReceiptRequested Property (Outlook)

Returns a  **Boolean** value that indicates **true** if a read receipt has been requested by the sender.


## Syntax

 _expression_ . **ReadReceiptRequested**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagReadReceiptRequested** . This property is read/write for **[SharingItem](sharingitem-object-outlook.md)** objects that have been created but have not been sent or posted; it is read-only for sent **SharingItem** objects.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

