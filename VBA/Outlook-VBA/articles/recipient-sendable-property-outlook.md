---
title: Recipient.Sendable Property (Outlook)
keywords: vbaol11.chm3476
f1_keywords:
- vbaol11.chm3476
ms.prod: outlook
api_name:
- Outlook.Recipient.Sendable
ms.assetid: ba6c3f35-5e51-f502-fb74-5403de3411e9
ms.date: 06/08/2017
---


# Recipient.Sendable Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether a meeting request can be sent to the **[Recipient](recipient-object-outlook.md)** . Read/write


## Syntax

 _expression_ . **Sendable**

 _expression_ A variable that represents a **[Recipient](recipient-object-outlook.md)** object.


## Remarks

This property corresponds to the MAPI property  **PidTagRecipientFlags** . It returns **True** if **PidTagRecipientFlags** is equal to 0x00000001. Setting the property changes **PidTagRecipientFlags** accordingly.

This property applies only to a recipient of a meeting request. If the recipient is not on a meeting request, getting and setting this property does not do anything.


## See also


#### Concepts


[Recipient Object](recipient-object-outlook.md)

