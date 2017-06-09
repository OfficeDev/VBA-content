---
title: SharingItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3497
f1_keywords:
- vbaol11.chm3497
ms.prod: outlook
api_name:
- Outlook.SharingItem.ConversationID
ms.assetid: f60a9a2e-76f7-0ff3-8b9d-70d3a4bc3f94
ms.date: 06/08/2017
---


# SharingItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[SharingItem](sharingitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **SharingItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

