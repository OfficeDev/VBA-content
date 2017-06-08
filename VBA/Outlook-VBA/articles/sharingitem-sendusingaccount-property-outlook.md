---
title: SharingItem.SendUsingAccount Property (Outlook)
keywords: vbaol11.chm703
f1_keywords:
- vbaol11.chm703
ms.prod: outlook
api_name:
- Outlook.SharingItem.SendUsingAccount
ms.assetid: 32eb7889-e01a-6b03-ddeb-0447da2dc655
ms.date: 06/08/2017
---


# SharingItem.SendUsingAccount Property (Outlook)

Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[SharingItem](sharingitem-object-outlook.md)** is to be sent. Read/write.


## Syntax

 _expression_ . **SendUsingAccount**

 _expression_ An expression that returns a **SharingItem** object.


## Remarks

The  **SendUsingAccount** property can be used to specify the account that should be used to send the **SharingItem** when the **[Send](sharingitem-send-method-outlook.md)** method is called. This property returns **Null** ( **Nothing** in Visual Basic) if the **SharingItem** is a received item, or if the account specified for the **SharingItem** no longer exists.

This property is read-only if the  **SharingItem** is a received item, or if the **SharingItem** has already been sent (the **[Sent](sharingitem-sent-property-outlook.md)** property of the object is set to **True** .)


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

