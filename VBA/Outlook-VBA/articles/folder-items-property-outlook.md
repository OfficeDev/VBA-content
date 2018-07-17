---
title: Folder.Items Property (Outlook)
keywords: vbaol11.chm1990
f1_keywords:
- vbaol11.chm1990
ms.prod: outlook
api_name:
- Outlook.Folder.Items
ms.assetid: 441820e7-5fe8-e5ef-83c0-9c87fd3dc9e3
ms.date: 06/08/2017
---


# Folder.Items Property (Outlook)

Returns an  **[Items](items-object-outlook.md)** collection object as a collection of Outlook items in the specified folder. Read-only.


## Syntax

 _expression_ . **Items**

 _expression_ A variable that represents a **Folder** object.


## Remarks

The index for the  **Items** collection starts at 1, and the items in the **Items** collection object are not guaranteed to be in any particular order.


## Example

This Visual Basic for Applications (VBA) example uses the  **Items** property to obtain the collection of **[ContactItem](contactitem-object-outlook.md)** objects from the default Contacts folder.


```vb
Sub ContactDateCheck() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myContacts As Outlook.Items 
 
 Dim myItems As Outlook.Items 
 
 Dim myItem As Object 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myContacts = myNamespace.GetDefaultFolder(olFolderContacts).Items 
 
 Set myItems = myContacts.Restrict("[LastModificationTime] > '01/1/2003'") 
 
 For Each myItem In myItems 
 
 If (myItem.Class = olContact) Then 
 
 MsgBox myItem.FullName &; ": " &; myItem.LastModificationTime 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

