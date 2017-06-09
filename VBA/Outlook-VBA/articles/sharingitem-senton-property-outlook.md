---
title: SharingItem.SentOn Property (Outlook)
keywords: vbaol11.chm662
f1_keywords:
- vbaol11.chm662
ms.prod: outlook
api_name:
- Outlook.SharingItem.SentOn
ms.assetid: 9b744303-42e2-9c38-b203-db6f8260d474
ms.date: 06/08/2017
---


# SharingItem.SentOn Property (Outlook)

Returns a  **Date** indicating the date and time on which the **[SharingItem](sharingitem-object-outlook.md)** was sent. Read-only.


## Syntax

 _expression_ . **SentOn**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagClientSubmitTime** . When you send an item using the object's **[Send](sharingitem-send-method-outlook.md)** method, the transport provider sets the **[ReceivedTime](sharingitem-receivedtime-property-outlook.md)** and **SentOn** properties for you.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

