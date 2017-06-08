---
title: MeetingItem.SenderName Property (Outlook)
keywords: vbaol11.chm1450
f1_keywords:
- vbaol11.chm1450
ms.prod: outlook
api_name:
- Outlook.MeetingItem.SenderName
ms.assetid: 07dd4ff2-36cd-cfbd-3b48-08e60f0aed78
ms.date: 06/08/2017
---


# MeetingItem.SenderName Property (Outlook)

Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.


## Syntax

 _expression_ . **SenderName**

 _expression_ A variable that represents a **MeetingItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagSenderName** .

If you wish to retrieve the fully qualified e-mail address of the sender, use the  **[SenderEmailAddress](meetingitem-senderemailaddress-property-outlook.md)** property.


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

