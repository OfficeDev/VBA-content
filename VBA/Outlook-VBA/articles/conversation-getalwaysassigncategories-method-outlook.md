---
title: Conversation.GetAlwaysAssignCategories Method (Outlook)
keywords: vbaol11.chm3439
f1_keywords:
- vbaol11.chm3439
ms.prod: outlook
api_name:
- Outlook.Conversation.GetAlwaysAssignCategories
ms.assetid: d09ae8ff-b725-cc09-9408-7b9039ee129f
ms.date: 06/08/2017
---


# Conversation.GetAlwaysAssignCategories Method (Outlook)

Returns a  **String** that indicates the category or categories that are assigned to all new items that arrive in the conversation.


## Syntax

 _expression_ . **GetAlwaysAssignCategories**( **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](store-object-outlook.md)**|Specifies the store to which categories of items that belong to the conversation should be returned.|

### Return Value

A  **String** that contains one or more categories that are assigned to items in the conversation.


## Remarks

Multiple categories are delimited by commas in the string of category names that this property returns. To convert the string of category names to an array of category names, use the Microsoft Visual Basic  **Split** function.

If the store specified by the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the method returns a string of categories that are applied to conversation items in the default delivery store.

If the  **[SetAlwaysAssignCategories](conversation-setalwaysassigncategories-method-outlook.md)** method has not been applied to a conversation, **GetAlwaysAssignCategories** returns **Null** ( **Nothing** in Visual Basic).

To stop the action of always assigning categories, use the  **[ClearAlwaysAssignCategories](conversation-clearalwaysassigncategories-method-outlook.md)** method. After the **ClearAlwaysAssignCategories** method has been called, **GetAlwaysAssignCategories** returns an empty string.


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

