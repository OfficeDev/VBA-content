---
title: Results.ItemRemove Event (Outlook)
keywords: vbaol11.chm516
f1_keywords:
- vbaol11.chm516
ms.prod: outlook
api_name:
- Outlook.Results.ItemRemove
ms.assetid: 95f59319-3182-5b2e-977f-d61512106090
ms.date: 06/08/2017
---


# Results.ItemRemove Event (Outlook)

Occurs when an item is deleted from the specified collection.


## Syntax

 _expression_ . **ItemRemove**

 _expression_ A variable that represents a **Results** object.


## Remarks

This event does not run when the last item in a Personal Folders file (.pst) is deleted, or if 16 or more items are deleted at once from a .pst file, Microsoft Exchange mailbox, or an Exchange public folder.

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example optionally sends a notification message to a workgroup when the user removes a contact from the default  **Contacts** folder. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlItems As Outlook.Items 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items 
 
End Sub 
 
 
 
Private Sub myOlItems_ItemRemove() 
 
 Dim myOlMItem As Outlook.MailItem 
 
 If MsgBox("Do you want to notify the Sales Team?", vbYesNo + vbQuestion) = vbYes Then 
 
 Set myOlMItem = Application.CreateItem(olMailItem) 
 
 myOlMItem.To = "Sales Team" 
 
 myOlMItem.Subject = "Remove Contact" 
 
 myOlMItem.Body = "Remove the following contact from your list:" 
 
 myOlMItem.Display 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Results Object](results-object-outlook.md)

