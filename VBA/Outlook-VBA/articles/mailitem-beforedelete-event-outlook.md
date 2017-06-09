---
title: MailItem.BeforeDelete Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeDelete
ms.assetid: 10fb2ac0-0382-2d7b-13ab-3edf06e50c81
ms.date: 06/08/2017
---


# MailItem.BeforeDelete Event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

 _expression_ . **BeforeDelete**( **_Item_** , **_Cancel_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when an e-mail message, distribution list, journal entry, task, contact, or post are deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## Example

The following Visual Basic for Applications (VBA) example prompts the user regarding whether to delete the item currently open. For this example to run, you need to have an open e-mail item that can be deleted. If you click  **No**, the item will not be deleted. If this event is canceled, Microsoft Outlook displays an error message. Therefore, you need to capture this event in your code. One way to do this is shown below. The sample code must be placed in a class module such as  `ThisOutlookSession`, and the  `DeleteMail()` procedure should be called before the event procedure can be called by Outlook.


```vb
Public WithEvents myItem As Outlook.MailItem 
 
 
 
Public Sub DeleteMail() 
 
 Const strCancelEvent = "Application-defined or object-defined error" 
 
 On Error GoTo ErrHandler 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 myItem.Delete 
 
 Exit Sub 
 
 
 
ErrHandler: 
 
 MsgBox Err.Description 
 
 If Err.Description = strCancelEvent Then 
 
 MsgBox "The event was cancelled." 
 
 End If 
 
 'If you want to execute the next instruction 
 
 Resume Next 
 
 'Otherwise it will finish here 
 
End Sub 
 
 
 
Private Sub myItem_BeforeDelete(ByVal Item As Object, Cancel As Boolean) 
 
 'Prompts the user before deleting an item 
 
 Dim strPrompt As String 
 
 
 
 'Prompt the user for a response 
 
 strPrompt = "Are you sure you want to delete the item?" 
 
 If MsgBox(strPrompt, vbYesNo + vbQuestion) = vbNo Then 
 
 'Don't delete the item 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

