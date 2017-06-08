---
title: Conversation.SetAlwaysMoveToFolder Method (Outlook)
keywords: vbaol11.chm3430
f1_keywords:
- vbaol11.chm3430
ms.prod: outlook
api_name:
- Outlook.Conversation.SetAlwaysMoveToFolder
ms.assetid: 52658b6d-c22c-a0e4-3743-4fe742bfbf9e
ms.date: 06/08/2017
---


# Conversation.SetAlwaysMoveToFolder Method (Outlook)

Sets a  **[Folder](folder-object-outlook.md)** object that indicates the folder to which all existing conversation items and new items that arrive in the conversation are always moved.


## Syntax

 _expression_ . **SetAlwaysMoveToFolder**( **_MoveToFolder_** , **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MoveToFolder_|Required| **Folder**|Specifies the folder to which all existing items and new items that arrive in the conversation are always moved.|
| _Store_|Required| **Store**|Specifies the store that contains the folder to which items of the conversation are moved.|

## Remarks

The  **SetAlwaysMoveToFolder** method operates on conversation items in the delivery store specified by the _Store_ parameter. If the _Store_ parameter represents a non-delivery store such as an archive .pst store, the move action will apply to conversation items in the default delivery store.

If the  _MoveToFolder_ parameter specifies an invalid folder that does not exist, has been moved, or is read-only, Outlook will raise an error.

To stop the always-move-to-folder action for conversations items in a store, call the  **[StopAlwaysMoveToFolder](conversation-stopalwaysmovetofolder-method-outlook.md)** method.




 **Note**  Setting the Deleted Items folder as the  _MoveToFolder_ parameter in **SetAlwaysMoveToFolder** is not equivalent to calling **[SetAlwaysDelete](conversation-setalwaysdelete-method-outlook.md)** on the same store and conversation. Setting the _MoveToFolder_ parameter to the Deleted Items folder results in the **[GetAlwaysDelete](conversation-getalwaysdelete-method-outlook.md)** method returning the value **olDoNotDelete** .

The  **[BeforeItemMove](folder-beforeitemmove-event-outlook.md)** event of the **Folder** object occurs when you call **SetAlwaysMoveToFolder** .


## Example

The following Visual Basic for Applications (VBA) example shows how to set the folder to which existing conversation items and new items that arrive in the conversation of a specific mail item are always moved. The code example,  `DemoSetAlwaysMoveToFolder`, chooses the first mail item displayed in the Reading Pane as the specific mail item, and the folder named "1-Reference" under the Inbox folder as the folder to move the conversation items to.  `DemoSetAlwaysMoveToFolder` verifies that conversations are enabled in the store for the selected mail item, obtains the conversation object for that mail item if a conversation exists, and uses **SetAlwaysMoveToFolder** to always move all existing conversation items and new items that arrive in the conversation to the specified folder.


```vb
Sub DemoSetAlwaysMoveToFolder() 
 Dim oMail As Outlook.MailItem 
 Dim oConv As Outlook.Conversation 
 Dim oStore As Outlook.Store 
 Dim oFolder As Outlook.Folder 
 
 ' Obtain a reference to the folder where conversation items will be moved. 
 Set oFolder = _ 
 Application.Session.GetDefaultFolder(olFolderInbox).Folders("1-Reference") 
 ' Get the Item displayed in the Reading Pane. 
 Set oMail = ActiveExplorer.Selection(1) 
 Set oStore = oFolder.Store 
 If oStore.IsConversationEnabled Then 
 Set oConv = oMail.GetConversation 
 If Not (oConv Is Nothing) Then 
 oConv.SetAlwaysMoveToFolder oFolder, oStore 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

