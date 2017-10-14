---
title: DistListItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3510
f1_keywords:
- vbaol11.chm3510
ms.prod: outlook
api_name:
- Outlook.DistListItem.ConversationID
ms.assetid: 8acbf4e8-d3ec-713c-7987-ba254e2373fb
ms.date: 06/08/2017
---


# DistListItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[DistListItem](distlistitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **DistListItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **DistListItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

