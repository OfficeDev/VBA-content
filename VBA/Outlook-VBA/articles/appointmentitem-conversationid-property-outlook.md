---
title: AppointmentItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3469
f1_keywords:
- vbaol11.chm3469
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.ConversationID
ms.assetid: 6897e23d-1d1d-f8fb-fbab-aa19242f4e7f
ms.date: 06/08/2017
---


# AppointmentItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[AppointmentItem](appointmentitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **AppointmentItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

