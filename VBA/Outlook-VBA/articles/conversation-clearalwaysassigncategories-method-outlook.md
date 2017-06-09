---
title: Conversation.ClearAlwaysAssignCategories Method (Outlook)
keywords: vbaol11.chm3489
f1_keywords:
- vbaol11.chm3489
ms.prod: outlook
api_name:
- Outlook.Conversation.ClearAlwaysAssignCategories
ms.assetid: 0494d8af-6569-c03d-99b1-be332c000985
ms.date: 06/08/2017
---


# Conversation.ClearAlwaysAssignCategories Method (Outlook)

Removes all categories from all items in the conversation and stops the action of always assigning categories to items in the conversation.


## Syntax

 _expression_ . **ClearAlwaysAssignCategories**( **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](store-object-outlook.md)**|Specifies the store from which categories of items that belong to the conversation should be removed.|

## Remarks

If the store specified by the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the category removal action will apply to items of the conversation in the default delivery store.

After you apply the  **ClearAlwaysAssignCategories** method on a conversation, the **[GetAlwaysAssignCategories](conversation-getalwaysassigncategories-method-outlook.md)** method will return **Null** ( **Nothing** in Visual Basic) for that conversation. Categories on existing items are cleared, and no categories are assigned to new items in the conversation.

If the  **[SetAlwaysAssignCategories](conversation-setalwaysassigncategories-method-outlook.md)** method has not been applied to a conversation, **ClearAlwaysAssignCategories** does not remove any categories.


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

