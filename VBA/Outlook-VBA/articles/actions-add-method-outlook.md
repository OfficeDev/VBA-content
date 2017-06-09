---
title: Actions.Add Method (Outlook)
keywords: vbaol11.chm151
f1_keywords:
- vbaol11.chm151
ms.prod: outlook
api_name:
- Outlook.Actions.Add
ms.assetid: aaf539c4-d60a-867f-086b-3cef7632a6f2
ms.date: 06/08/2017
---


# Actions.Add Method (Outlook)

Creates a new action in the  **[Actions](actions-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**

 _expression_ A variable that represents an **Actions** object.


### Return Value

An  **[Action](action-object-outlook.md)** object that represents the new action.


## Example

This VBA example creates a new mail message and uses the  **Add** method to add an **[Action](action-object-outlook.md)** to it. To run this example without any errors, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub AddAction() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myAction As Outlook.Action 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myAction = myItem.Actions.Add 
 
 myAction.Name = "Link Original" 
 
 myAction.ShowOn = olMenuAndToolbar 
 
 myAction.ReplyStyle = olLinkOriginalItem 
 
 myItem.To = "Dan Wilson" 
 
 myItem.Send 
 
End Sub
```


## See also


#### Concepts


[Actions Object](actions-object-outlook.md)

