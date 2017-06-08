---
title: MailItem.SentOn Property (Outlook)
keywords: vbaol11.chm1359
f1_keywords:
- vbaol11.chm1359
ms.prod: outlook
api_name:
- Outlook.MailItem.SentOn
ms.assetid: 477d7f13-af24-dca7-9845-1a3669093972
ms.date: 06/08/2017
---


# MailItem.SentOn Property (Outlook)

Returns a  **Date** indicating the date and time on which the Outlook item was sent. Read-only.


## Syntax

 _expression_ . **SentOn**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagClientSubmitTime** . When you send an item using the object's **[Send](mailitem-send-method-outlook.md)** method, the transport provider sets the **[ReceivedTime](mailitem-receivedtime-property-outlook.md)** and **SentOn** properties for you.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

