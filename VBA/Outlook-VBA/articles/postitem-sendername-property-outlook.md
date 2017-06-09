---
title: PostItem.SenderName Property (Outlook)
keywords: vbaol11.chm1550
f1_keywords:
- vbaol11.chm1550
ms.prod: outlook
api_name:
- Outlook.PostItem.SenderName
ms.assetid: cee9b0ac-1528-1387-48db-b31d58d691ca
ms.date: 06/08/2017
---


# PostItem.SenderName Property (Outlook)

Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.


## Syntax

 _expression_ . **SenderName**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagSenderName** .

If you wish to retrieve the fully qualified e-mail address of the sender, use the  **[SenderEmailAddress](postitem-senderemailaddress-property-outlook.md)** property.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

