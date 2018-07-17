---
title: Folder.FolderPath Property (Outlook)
keywords: vbaol11.chm2006
f1_keywords:
- vbaol11.chm2006
ms.prod: outlook
api_name:
- Outlook.Folder.FolderPath
ms.assetid: 40a588fa-0962-bc01-f8ac-39f0bab2092c
ms.date: 06/08/2017
---


# Folder.FolderPath Property (Outlook)

Returns a  **String** that indicates the path of the current folder. Read-only.


## Syntax

 _expression_ . **FolderPath**

 _expression_ A variable that represents a **Folder** object.


## Example

The following example displays information about the default Contacts folder. The subroutine accepts a  **[Folder](folder-object-outlook.md)** object and displays the folder's name, path, and address book information.


```vb
Sub Folderpaths() 
 
 Dim nmsName As NameSpace 
 
 Dim fldFolder As Folder 
 
 
 
 'Create namespace reference 
 
 Set nmsName = Application.GetNamespace("MAPI") 
 
 'create folder instance 
 
 Set fldFolder = nmsName.GetDefaultFolder(olFolderContacts) 
 
 'call sub program 
 
 Call FolderInfo(fldFolder) 
 
End Sub 
 
 
 
Sub FolderInfo(ByVal fldFolder As Folder) 
 
 'Displays information about a given folder 
 
 MsgBox fldFolder.Name &; "'s current path is " &; _ 
 
 fldFolder.FolderPath &; _ 
 
 ". The current address book name is " &; fldFolder.AddressBookName &; "." 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

