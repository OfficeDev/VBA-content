---
title: Conversation.GetAlwaysDelete Method (Outlook)
keywords: vbaol11.chm3440
f1_keywords:
- vbaol11.chm3440
ms.prod: outlook
api_name:
- Outlook.Conversation.GetAlwaysDelete
ms.assetid: 95843bf3-7fff-fab0-ca7b-014ba290d718
ms.date: 06/08/2017
---


# Conversation.GetAlwaysDelete Method (Outlook)

Returns a constant in the  **[OlAlwaysDeleteConversation](olalwaysdeleteconversation-enumeration-outlook.md)** enumeration that indicates whether all new items that join the conversation are always moved to the **Deleted Items** folder in the specified delivery store.


## Syntax

 _expression_ . **GetAlwaysDelete**( **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](store-object-outlook.md)**|Specifies the store that holds the  **Deleted Items** folder to which items of the conversation are moved.|

### Return Value

A constant from the  **OlAlwaysDeleteConversation** enumeration that indicates whether all new items of the conversation are always moved to the Deleted Items folder of the specified delivey store.


## Remarks

 If the _Store_ parameter specifies a non-delivery store such as an archive .pst store, the **GetAlwaysDelete** method returns a constant from **OlAlwaysDeleteConversation** that applies to conversation items in the default delivery store. Items on a non-delivery store are not moved to the **Deleted Items** folder for the default delivery store.

If  **GetAlwaysDelete** returns **olAlwaysDelete** , items of the conversation are always moved to the **Deleted Items** folder for the store that contains the items. In a cross-store conversation, items are moved to the **Deleted Items** folder for the store that contains the items. When **GetAlwaysDelete** returns **olAlwaysDelete** , the **[GetAlwaysMoveToFolder](conversation-getalwaysmovetofolder-method-outlook.md)** method returns a folder object that represents the **Deleted Items** folder for the default store.

If  **GetAlwaysDelete** returns **olAlwaysDeleteUnsupported** , the specified store does not support the action of always moving items to the **Deleted Items** folder of that store.

If  **GetAlwaysDelete** returns **olDoNotDelete** , new items that arrive in the conversation are not moved to the **Deleted Items** folder on the specified delivery store, and existing conversation items in the **Deleted Items** folder are moved to the **Inbox**.


## Example

The following Microsoft Visual Basic for Application (VBA) example shows how to verify the always-delete setting of the conversation of a selected mail item. The code example,  `DemoGetAlwaysDelete`, verifies that conversations are enabled in the default store, obtains the conversation that involves the first mail item displayed in the Reading Pane if a conversation exists, uses  **GetAlwaysDelete** to obtain the always-delete setting, and displays the setting.


```vb
Sub DemoGetAlwaysDelete() 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oConv As Outlook.Conversation 
 
 Dim oStore As Outlook.Store 
 
 Dim intValue As Integer 
 
 
 
 ' Get the item displayed in Reading Pane. 
 
 Set oMail = ActiveExplorer.Selection(1) 
 
 
 
 If Application.Session.DefaultStore.IsConversationEnabled Then 
 
 Set oConv = oMail.GetConversation 
 
 If Not (oConv Is Nothing) Then 
 
 intValue = _ 
 
 oConv.GetAlwaysDelete(Application.session.DefaultStore) 
 
 If intValue = _ 
 
 Outlook.OlAlwaysDeleteConversation.olAlwaysDelete Then 
 
 Debug.Print "olAlwaysDelete" 
 
 ElseIf intValue = _ 
 
 Outlook.OlAlwaysDeleteConversation.olAlwaysDeleteUnsupported Then 
 
 Debug.Print "olAlwaysDeleteUnsupported" 
 
 ElseIf intValue = _ 
 
 Outlook.OlAlwaysDeleteConversation.olDoNotDelete Then 
 
 Debug.Print "olDoNotDelete" 
 
 End If 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

