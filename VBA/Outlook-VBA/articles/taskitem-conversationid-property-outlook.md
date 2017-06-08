---
title: TaskItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3474
f1_keywords:
- vbaol11.chm3474
ms.prod: outlook
api_name:
- Outlook.TaskItem.ConversationID
ms.assetid: 69b28ef6-5521-944c-f908-df715e837c36
ms.date: 06/08/2017
---


# TaskItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[TaskItem](taskitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **TaskItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

