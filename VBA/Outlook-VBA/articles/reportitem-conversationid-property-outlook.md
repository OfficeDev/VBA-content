---
title: ReportItem.ConversationID Property (Outlook)
keywords: vbaol11.chm3493
f1_keywords:
- vbaol11.chm3493
ms.prod: outlook
api_name:
- Outlook.ReportItem.ConversationID
ms.assetid: b642a06e-94f0-b615-1806-fdd5ae881d48
ms.date: 06/08/2017
---


# ReportItem.ConversationID Property (Outlook)

Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[ReportItem](reportitem-object-outlook.md)** object belongs to. Read-only.


## Syntax

 _expression_ . **ConversationID**

 _expression_ A variable that represents a **ReportItem** object.


## Remarks

This property associates items with a conversation. These items and the conversation all have the same value in their  **ConversationID** property.

This property corresponds with the MAPI property  **PidTagConversationId** .

If the  **ReportItem** object is created in a version of Microsoft Outlook earlier than Outlook 2013, or if Outlook is running in online mode against a version of Microsoft Exchange Server earlier than Microsoft Exchange Server 2010, this property returns the same value as the **[ConversationTopic](appointmentitem-conversationtopic-property-outlook.md)** property.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

