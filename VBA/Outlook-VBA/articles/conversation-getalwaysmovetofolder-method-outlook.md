---
title: Conversation.GetAlwaysMoveToFolder Method (Outlook)
keywords: vbaol11.chm3441
f1_keywords:
- vbaol11.chm3441
ms.prod: outlook
api_name:
- Outlook.Conversation.GetAlwaysMoveToFolder
ms.assetid: ecad049d-338b-d5e0-f241-a9dddaeae316
ms.date: 06/08/2017
---


# Conversation.GetAlwaysMoveToFolder Method (Outlook)

Returns a  **[Folder](folder-object-outlook.md)** object that indicates the folder in the specified delivery store to which new items that arrive in the conversation are always moved.


## Syntax

 _expression_ . **GetAlwaysMoveToFolder**( **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](store-object-outlook.md)**|The store where the folder to which conversation items are moved resides.|

### Return Value

A  **Folder** object in the specified store to which all new items that arrive in the conversation are always moved.


## Remarks

If the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the **GetAlwaysMoveToFolder** method returns a **Folder** object that applies to conversation items on the default delivery store.

If no folder, other than the  **Deleted Items** folder, has been specified to always move conversation items into, the **GetAlwaysMoveToFolder** method returns **Null** ( **Nothing** in Visual Basic).


## Example

The following Microsoft Visual Basic for Application (VBA) example shows how to find the folder into which new items that arrive in the conversation of the first mail item displayed in the Reading Pane are always moved. The code example,  `DemoGetAlwaysMoveToFolder`, verifies that conversations are enabled in the store for the selected mail item, obtains the conversation object for that mail item if a conversation exists, uses  **GetAlwaysMoveToFolder** to obtain the folder, and displays the folder name.


```vb
Sub DemoGetAlwaysMoveToFolder() 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oConv As Outlook.Conversation 
 
 Dim oStore As Outlook.Store 
 
 
 
 ' Get Item displayed in Reading Pane. 
 
 Set oMail = ActiveExplorer.Selection(1) 
 
 Set oStore = oMail.Parent.Store 
 
 If oStore.IsConversationEnabled Then 
 
 Set oConv = oMail.GetConversation 
 
 If Not (oConv Is Nothing) Then 
 
 Dim oFolder As Outlook.folder 
 
 Set oFolder = _ 
 
 oConv.GetAlwaysMoveToFolder(oStore) 
 
 If Not (oFolder Is Nothing) Then 
 
 Debug.Print "MoveToFolder: " &; oFolder.name 
 
 Else 
 
 Debug.Print "MoveToFolder action not set" 
 
 End If 
 
 End If 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

