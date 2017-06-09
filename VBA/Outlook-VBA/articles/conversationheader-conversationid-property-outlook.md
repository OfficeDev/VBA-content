---
title: ConversationHeader.ConversationID Property (Outlook)
keywords: vbaol11.chm3542
f1_keywords:
- vbaol11.chm3542
ms.prod: outlook
api_name:
- Outlook.ConversationHeader.ConversationID
ms.assetid: 2c359158-58e1-d40f-e8c5-b765e944e8c8
ms.date: 06/08/2017
---


# ConversationHeader.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies the **[Conversation](conversation-object-outlook.md)** object to which this conversation header belongs. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **[ConversationHeader](conversationheader-object-outlook.md)** object.


## Remarks

This property associates the conversation header with other items in the same conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .


## See also


#### Concepts


[ConversationHeader Object](conversationheader-object-outlook.md)

