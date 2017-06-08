---
title: MeetingItem.SentOn Property (Outlook)
keywords: vbaol11.chm1452
f1_keywords:
- vbaol11.chm1452
ms.prod: outlook
api_name:
- Outlook.MeetingItem.SentOn
ms.assetid: 361dfa26-6514-cc3a-aa1b-240728ac0dd9
ms.date: 06/08/2017
---


# MeetingItem.SentOn Property (Outlook)

Returns a  **Date** indicating the date and time on which the Outlook item was sent. Read-only.


## Syntax

 _expression_ . **SentOn**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagClientSubmitTime** . When you send a meeting request item using the object's **[Send](meetingitem-receivedtime-property-outlook.md)** method, the transport provider sets the **[ReceivedTime](mailitem-receivedtime-property-outlook.md)** and **SentOn** properties for you.


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

