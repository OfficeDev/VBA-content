---
title: Store.IsConversationEnabled Property (Outlook)
keywords: vbaol11.chm3518
f1_keywords:
- vbaol11.chm3518
ms.prod: outlook
api_name:
- Outlook.Store.IsConversationEnabled
ms.assetid: ce333881-a5f3-2115-0ae4-296d15c4bead
ms.date: 06/08/2017
---


# Store.IsConversationEnabled Property (Outlook)

Returns a  **Boolean** value that is **True** if the store supports Conversation view. Read-only.


## Syntax

 _expression_ . **IsConversationEnabled**

 _expression_ A variable that represents a **[Store](store-object-outlook.md)** object.


## Remarks

 A store supports Conversation view if the store is a POP, IMAP, or PST store, or if it runs a version of Microsoft Exchange Server that is at least Microsoft Exchange Server 2010. A store also supports Conversation view if the store is running Microsoft Exchange Server 2007, the version of Outlook is at least Outlook, and Outlook is running in cached mode.

If a store supports conversations, calling the  **GetConversation** method of an item in the store returns a **[Conversation](conversation-object-outlook.md)** object for the item. If the store does not support conversations, **GetConversation** returns **Null** ( **Nothing** in Visual Basic) for items in the store.


## See also


#### Concepts


[Store Object](store-object-outlook.md)

