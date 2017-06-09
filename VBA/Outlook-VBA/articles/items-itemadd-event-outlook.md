---
title: Items.ItemAdd Event (Outlook)
keywords: vbaol11.chm314
f1_keywords:
- vbaol11.chm314
ms.prod: outlook
api_name:
- Outlook.Items.ItemAdd
ms.assetid: e46f5958-aff8-3a6b-b3df-5c4352b6c3d9
ms.date: 06/08/2017
---


# Items.ItemAdd Event (Outlook)

Occurs when one or more items are added to the specified collection. This event does not run when a large number of items are added to the folder at once. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Syntax

 _expression_ . **ItemAdd**( **_Item_** )

 _expression_ A variable that represents an **Items** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item that was added.|

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


[Items Object](items-object-outlook.md)

