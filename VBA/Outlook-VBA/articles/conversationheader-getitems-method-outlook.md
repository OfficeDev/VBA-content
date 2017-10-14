---
title: ConversationHeader.GetItems Method (Outlook)
keywords: vbaol11.chm3544
f1_keywords:
- vbaol11.chm3544
ms.prod: outlook
api_name:
- Outlook.ConversationHeader.GetItems
ms.assetid: 018fab26-3cdc-cd39-4a16-fb2a26ae237f
ms.date: 06/08/2017
---


# ConversationHeader.GetItems Method (Outlook)

Obtains a  **[SimpleItems](simpleitems-object-outlook.md)** collection that contains all of the items in the conversation that reside in the same folder as the selected conversation header.


## Syntax

 _expression_ . **GetItems**

 _expression_ A variable that represents a **[ConversationHeader](conversationheader-object-outlook.md)** object.


### Return Value

A  **SimpleItems** collection of items that belong to the same conversation and reside in the same folder as the conversation header.


## Remarks

The  **SimpleItems** collection only contains conversation items in the folder that contains the conversation header. The **SimpleItems** collection does not return cross-folder conversation items. If you must access cross-folder content, use the **[Conversation](conversation-object-outlook.md)** object.

If no conversation items exist in the same folder as the conversation header,  **GetItems** returns a **SimpleItems** collection with the **[SimpleItems.Count](simpleitems-count-property-outlook.md)** property equal to 0.


## See also


#### Concepts


[ConversationHeader Object](conversationheader-object-outlook.md)
#### Other resources


[How to: Obtain and Enumerate Selected Conversations](http://msdn.microsoft.com/library/3bba1e98-b2eb-c53d-354a-bdd899b65a59%28Office.15%29.aspx)


