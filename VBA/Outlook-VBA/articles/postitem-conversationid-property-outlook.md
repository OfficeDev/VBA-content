---
title: PostItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3473
f1_keywords:
- vbaol11.chm3473
ms.prod: outlook
api_name:
- Outlook.PostItem.ConversationID
ms.assetid: 102f64a0-2188-3731-eb13-95bc41da4e37
ms.date: 06/08/2017
---


# PostItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[PostItem](postitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **PostItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

