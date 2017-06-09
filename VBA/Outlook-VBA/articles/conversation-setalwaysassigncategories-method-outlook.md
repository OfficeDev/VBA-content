---
title: Conversation.SetAlwaysAssignCategories Method (Outlook)
keywords: vbaol11.chm3444
f1_keywords:
- vbaol11.chm3444
ms.prod: outlook
api_name:
- Outlook.Conversation.SetAlwaysAssignCategories
ms.assetid: 9b19f083-3aa9-8a0b-ea91-ff52fe46ad35
ms.date: 06/08/2017
---


# Conversation.SetAlwaysAssignCategories Method (Outlook)

Applies one or more categories to all existing items and future items of the conversation.


## Syntax

 _expression_ . **SetAlwaysAssignCategories**( **_Categories_** , **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Categories_|Required| **String**|A comma-delimited string of one or more category names that are always assigned to all items in the conversation.|
| _Store_|Required| **[Store](store-object-outlook.md)**|The store in which items of the conversation should always be assigned the categories specified by the  _Categories_ parameter.|

## Remarks

If the store specified by the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the method returns a string of categories that are applied to conversation items in the default delivery store.

The  **[ItemChange](items-itemchange-event-outlook.md)** event of the **[Items](items-object-outlook.md)** object occurs when you call the **SetAlwaysAssignCategories** method on a conversation.

To determine existing master categories for the current user, examine the  **[Categories](store-categories-property-outlook.md)** property of the **[Store](store-object-outlook.md)** object that is specified by the _Store_ parameter. If one or more categories specified by the _Categories_ parameter do not exist in the master categories collection, the categories will be assigned to the conversation but will not be added to the master categories collection.

To determine the existing categories that are always assigned to items of the conversation in the specified store, use the  **[GetAlwaysAssignCategories](conversation-getalwaysassigncategories-method-outlook.md)** method.

If  **SetAlwaysAssignCategories** is called more than once, the result is cumulative. For example, if you call **SetAlwaysAssignCategories** specifying the category ?Important? and then call **SetAlwaysAssignCategories** again specifying the categories "Business" and "Social", the categories that are always assigned are "Important", "Business", and "Social".

To stop the action of always assigning categories, use the  **[ClearAlwaysAssignCategories](conversation-clearalwaysassigncategories-method-outlook.md)** method. After the **ClearAlwaysAssignCategories** method has been called, **GetAlwaysAssignCategories** returns an empty string.

 The **SetAlwaysAssignToCategories** method ignores any category names that are empty strings. For example, if the _Categories_ parameter is set to the string "Work,,Play", "Work" and "Play" are assigned to the conversation and the empty string category is ignored.


## Example

The following Visual Basic for Applications (VBA) example shows how to assign categories to all existing and new items that arrive in the conversation of a specific mail item. The code example,  `DemoSetAlwaysAssignCategories`, chooses the first mail item displayed in the Reading Pane as the specific mail item.  `DemoSetAlwaysAssignCategories` verifies that conversations are enabled in the store for the selected mail item, obtains the conversation object for that mail item if a conversation exists, and uses **SetAlwaysAssignToCategories** to set the two categories "Best Practices" and "OOM" to all existing items and future items of that conversation.


```vb
Sub DemoSetAlwaysAssignCategories() 
 Dim oMail As Outlook.MailItem 
 Dim oConv As Outlook.Conversation 
 Dim oStore As Outlook.Store 
 ' Get the item displayed in the Reading Pane. 
 Set oMail = ActiveExplorer.Selection(1) 
 Set oStore = oMail.Parent.Store 
 If oStore.IsConversationEnabled Then 
 Set oConv = oMail.GetConversation 
 If Not (oConv Is Nothing) Then 
 Dim oFolder As Outlook.folder 
 oConv.SetAlwaysAssignCategories "Best Practices; OOM", oStore 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

