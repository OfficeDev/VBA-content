---
title: ConversationHeader.Parent Property (Outlook)
keywords: vbaol11.chm3545
f1_keywords:
- vbaol11.chm3545
ms.prod: outlook
api_name:
- Outlook.ConversationHeader.Parent
ms.assetid: 2f465ae5-18a9-ad77-4419-eb8ec81acb2f
ms.date: 06/08/2017
---


# ConversationHeader.Parent Property (Outlook)

Returns the parent  **Object** of the specified object. Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **[ConversationHeader](conversationheader-object-outlook.md)** object.


## Remarks

The parent of the  **ConversationHeader** object returns a **[Selection](selection-object-outlook.md)** object.

 The returned **Selection** object contains only **[ConversationHeader](conversationheader-object-outlook.md)** objects. Getting the **Parent** property is equivalent to calling the **[Selection.GetSelection](selection-getselection-method-outlook.md)** method with the **olConversationHeaders** argument.


## See also


#### Concepts


[ConversationHeader Object](conversationheader-object-outlook.md)

