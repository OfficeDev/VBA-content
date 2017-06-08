---
title: Results.ItemAdd Event (Outlook)
keywords: vbaol11.chm514
f1_keywords:
- vbaol11.chm514
ms.prod: outlook
api_name:
- Outlook.Results.ItemAdd
ms.assetid: b867fb25-9a66-1a80-4bf6-b1f4814a6d2e
ms.date: 06/08/2017
---


# Results.ItemAdd Event (Outlook)

Occurs when one or more items are added to the specified collection.


## Syntax

 _expression_ . **ItemAdd**( **_Item_** )

 _expression_ A variable that represents a **Results** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item that was added.|

## Remarks

This event does not run when a large number of items are added to the folder at once. It is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

In this Visual Basic for Applications (VBA) example, when a new contact is added to the  **Contacts** folder, the contact item is attached to an e-mail message and sent to a distribution list named "Sales Team". The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlItems As Outlook.Items 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items 
 
End Sub 
 
 
 
Private Sub myOlItems_ItemAdd(ByVal Item As Object) 
 
 Dim myOlMItem As Outlook.MailItem 
 
 Dim myOlAtts As Outlook.Attachments 
 
 Set myOlMItem = myOlApp.CreateItem(olMailItem) 
 
 myOlMItem.Save 
 
 
 
 Set myOlAtts = myOlMItem.Attachments 
 
 ' Add new contact to attachments in mail message 
 
 myOlAtts.Add Item, olByValue 
 
 myOlMItem.To = "Sales Team" 
 
 myOlMItem.Subject = "New contact" 
 
 myOlMItem.Send 
 
End Sub
```


## See also


#### Concepts


[Results Object](results-object-outlook.md)

