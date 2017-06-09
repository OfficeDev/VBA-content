---
title: MailItem.CustomAction Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.CustomAction
ms.assetid: 2068586f-bdab-a786-d933-4e32117bb4f8
ms.date: 06/08/2017
---


# MailItem.CustomAction Event (Outlook)

Occurs when a custom action of an item (which is an instance of the parent object) executes.


## Syntax

 _expression_ . **CustomAction**( **_Action_** , **_Response_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Action_|Required| **Object**|The  **[Action](action-object-outlook.md)** object.|
| _Response_|Required| **Object**|The newly created item resulting from the custom action.|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the custom action is not completed.|

## Remarks

The  **Action** object and the newly created item resulting from the custom action are passed to the event.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the custom action operation is not completed.


## Example

This Visual Basic for Applications (VBA) example uses the  **CustomAction** event to set the **Subject** property on the response item. Execute the `AddAction` procedure before executing the `Initialize_Handler` to create an item with a custom event called 'Link Original'.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Sub AddAction() 
 
 Dim myAction As Outlook.Action 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myAction = myItem.Actions.Add 
 
 myAction.Name = "Link Original" 
 
 myAction.ShowOn = olMenuAndToolbar 
 
 myAction.ReplyStyle = olLinkOriginalItem 
 
 myItem.To = "Dan Wilson" 
 
 myItem.Subject = "Before" 
 
 myItem.Send 
 
End Sub 
 
 
 
Sub Initialize_Handler() 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
End Sub 
 
 
 
Private Sub myItem_CustomAction(ByVal Action As Object, ByVal Response As Object, Cancel As Boolean) 
 
 Select Case Action.Name 
 
 Case "Link Original" 
 
 Response.Subject = "Changed by VB Script" 
 
 Case Else 
 
 End Select 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

