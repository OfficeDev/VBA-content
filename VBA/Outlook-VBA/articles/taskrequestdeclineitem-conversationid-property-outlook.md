---
title: TaskRequestDeclineItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3503
f1_keywords:
- vbaol11.chm3503
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.ConversationID
ms.assetid: 14638aa8-8e39-bde9-19d1-3f082f57c9e2
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[TaskRequestDeclineItem](taskrequestdeclineitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **TaskRequestDeclineItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

