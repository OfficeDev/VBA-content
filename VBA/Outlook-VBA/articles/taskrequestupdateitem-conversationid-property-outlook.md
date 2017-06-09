---
title: TaskRequestUpdateItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3506
f1_keywords:
- vbaol11.chm3506
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.ConversationID
ms.assetid: e70b6b6d-c6ba-4097-ab83-b1d826b1a6d5
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[TaskRequestUpdateItem](taskrequestupdateitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **TaskRequestUpdateItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

