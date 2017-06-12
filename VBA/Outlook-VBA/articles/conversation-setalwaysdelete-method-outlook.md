---
title: Conversation.SetAlwaysDelete Method (Outlook)
keywords: vbaol11.chm3445
f1_keywords:
- vbaol11.chm3445
ms.prod: outlook
api_name:
- Outlook.Conversation.SetAlwaysDelete
ms.assetid: f13fce28-864e-a607-304d-a3722845cdd8
ms.date: 06/08/2017
---


# Conversation.SetAlwaysDelete Method (Outlook)

Specifies a setting for the specified delivery store that indicates whether all existing items and all new items that arrive in the conversation are always moved to the Deleted Items folder in the specified delivery store.


## Syntax

 _expression_ . **SetAlwaysDelete**( **_AlwaysDelete_** , **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlwaysDelete_|Required| **[OlAlwaysDeleteConversation](olalwaysdeleteconversation-enumeration-outlook.md)**|A constant that indicates whether all existing and new items that arrive in the conversation are always moved to the Deleted Folder of the store specified by the  _Store_ parameter.|
| _Store_|Required| **[Store](store-object-outlook.md)**|Specifies the store that contains the Deleted Items folder to which existing and new items of the conversation are to be moved.|

## Remarks

The  **SetAlwaysDelete** method operates on conversation items in the delivery store specified by the _Store_ parameter. If the store specified by the _Store_ parameter represents a non-delivery store such as an archive .pst store, the action is applied to conversation items in the default delivery store.

If the  _AlwaysDelete_ parameter is **olAlwaysDelete** , conversation items are moved to the Deleted Items folder for the specfied store. In this case, the items are not permanently deleted, unless the user has specified a separate option to permanently delete items when Microsoft Outlook shuts down.

If  **SetAlwaysDelete** returns **olDoNotDelete** , existing conversation items and new items that arrive in the conversation are not moved to the Deleted Items folder in the specified delivery store, and existing conversation items in the Deleted Items folder are moved to the Inbox.


## Example

The following Visual Basic for Applications (VBA) example shows how to set the always-delete setting for the conversation of a specific mail item. The code example,  `DemoSetAlwaysDelete`, chooses the first mail item displayed in the Reading Pane as the specific mail item.  `DemoSetAlwaysDelete` verifies that conversations are enabled in the store for the mail item, obtains the conversation that involves that mail item if a conversation exists, and uses **SetAlwaysDelete** to always move existing and new items for that conversation to the Deleted Items folder in the same store.


```vb
Sub DemoSetAlwaysDelete() 
 Dim oMail As Outlook.MailItem 
 Dim oConv As Outlook.Conversation 
 Dim oStore As Outlook.Store 
 
 ' Get the item displayed in the Reading Pane. 
 Set oMail = ActiveExplorer.Selection(1) 
 Set oStore = oMail.Parent.Store 
 If oStore.IsConversationEnabled Then 
 Set oConv = oMail.GetConversation 
 If Not (oConv Is Nothing) Then 
 oConv.SetAlwaysDelete _ 
 olAlwaysDelete, oStore 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

