---
title: Conversation.StopAlwaysDelete Method (Outlook)
keywords: vbaol11.chm3432
f1_keywords:
- vbaol11.chm3432
ms.prod: outlook
api_name:
- Outlook.Conversation.StopAlwaysDelete
ms.assetid: c759c9c8-bc43-ad5e-954c-88494c3dc4a6
ms.date: 06/08/2017
---


# Conversation.StopAlwaysDelete Method (Outlook)

Stops the action of always moving conversation items in the specified store to the Deleted Items folder in that store.


## Syntax

 _expression_ . **StopAlwaysDelete**( **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](store-object-outlook.md)**|Specifies the store to which the stop-always-delete action applies.|

## Remarks

If the always-delete action has not been turned on,  **StopAlwaysDelete** does not carry out any action.

If the always-delete action has been turned on (by calling the [SetAlwaysDelete](conversation-setalwaysdelete-method-outlook.md) method, **StopAlwaysDelete** moves existing conversation items in the Deleted Items folder to the Inbox.

After calling the  **StopAlwaysDelete** method for a conversation in a store, calling the **[GetAlwaysDelete](conversation-getalwaysdelete-method-outlook.md)** method on that conversation and store returns the constant **olDoNotDelete** .

If the store specified by the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the stop-always-delete action is applied to conversation items in the default delivery store.

Calling this method on a conversation that is already in the Deleted Items folder in the specified store returns an error.


## Example

The following Visual Basic for Applications (VBA) example shows how to stop the always-delete action for the conversation of a specific mail item. The code example,  `DemoStopAlwaysDelete`, chooses the first mail item displayed in the Reading Pane as the specific mail item.  `DemoStopAlwaysDelete` verifies that conversations are enabled on the store for the mail item, obtains the conversation that involves that mail item if a conversation exists, and uses **SetAlwaysDelete** to stop the always-delete action for that conversation on that store.


```vb
Sub DemoStopAlwaysDelete() 
 Dim oMail As Outlook.MailItem 
 Dim oConv As Outlook.Conversation 
 Dim oStore As Outlook.Store 
 
 ' Get the item displayed in the Reading Pane. 
 Set oMail = ActiveExplorer.Selection(1) 
 Set oStore = oMail.Parent.Store 
 If oStore.IsConversationEnabled Then 
 Set oConv = oMail.GetConversation 
 If Not (oConv Is Nothing) Then 
 oConv.StopAlwaysDelete oStore 
 End If 
 End If 
End Sub
```


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

