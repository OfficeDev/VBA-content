---
title: Action.Execute Method (Outlook)
keywords: vbaol11.chm23
f1_keywords:
- vbaol11.chm23
ms.prod: outlook
api_name:
- Outlook.Action.Execute
ms.assetid: 29dd0c5c-ed5f-b2cc-45b0-1c8c348239bb
ms.date: 06/08/2017
---


# Action.Execute Method (Outlook)

Executes the action for the specified item.


## Syntax

 _expression_ . **Execute**

 _expression_ An expression that returns a **Action** object.


### Return Value

An  **Object** value that represents the Outlook item created by the action upon execution.


## Example

This Visual Basic for Applications (VBA) example uses the  **Execute** method to look through all the actions for the given e-mail message and executes the action called "Reply."


```vb
Sub SendReply() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim MyItem As Outlook.MailItem 
 
 Dim myItem2 As Outlook.MailItem 
 
 Dim myAction As Outlook.Action 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 On Error GoTo ErrorHandler 
 
 Set MyItem = Application.ActiveInspector.CurrentItem 
 
 For Each myAction In MyItem.Actions 
 
 If myAction.Name = "Reply" Then 
 
 Set myItem2 = myAction.Execute 
 
 myItem2.Send 
 
 Exit For 
 
 End If 
 
 Next myAction 
 
 Exit Sub 
 
ErrorHandler: 
 
 MsgBox "There is no current item." 
 
End Sub
```


## See also


#### Concepts


[Action Object](action-object-outlook.md)

