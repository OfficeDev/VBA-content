---
title: TaskRequestAcceptItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3501
f1_keywords:
- vbaol11.chm3501
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.ConversationID
ms.assetid: 0cd2c84f-0332-73aa-097e-5944bf6258c8
ms.date: 06/08/2017
---


# TaskRequestAcceptItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[TaskRequestAcceptItem](taskrequestacceptitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **TaskRequestAcceptItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **TaskRequestAcceptItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[TaskRequestAcceptItem Object](taskrequestacceptitem-object-outlook.md)

