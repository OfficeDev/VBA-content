---
title: PostItem.GetConversation Method (Outlook)
keywords: vbaol11.chm3407
f1_keywords:
- vbaol11.chm3407
ms.prod: outlook
api_name:
- Outlook.PostItem.GetConversation
ms.assetid: f5912703-d064-5e4a-b3ae-5091d791cb87
ms.date: 06/08/2017
---


# PostItem.GetConversation Method (Outlook)

Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.


## Syntax

 _expression_ . **GetConversation**

 _expression_ A variable that represents a **[PostItem](postitem-object-outlook.md)** object.


### Return Value

A  **Conversation** object that represents the conversation to which this item belongs.


## Remarks

 **GetConversation** returns **Null** ( **Nothing** in Visual Basic) if no conversation exists for the item. No conversation exists for an item in the following scenarios:


- The item has not been saved. An item can be saved programmatically, by user action, or by auto-save.
    
- For an item that can be sent (for example, a mail item, appointment item, or contact item), the item has not been sent.
    
- Conversations have been disabled through the Windows registry.
    
- The store does not support Conversation view (for example, Outlook is running in classic online mode against a version of Microsoft Exchange earlier than Microsoft Exchange Server 2010). Use the  **[IsConversationEnabled](store-isconversationenabled-property-outlook.md)** property of the **[Store](store-object-outlook.md)** object to determine whether the store supports Conversation view.
    



## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

