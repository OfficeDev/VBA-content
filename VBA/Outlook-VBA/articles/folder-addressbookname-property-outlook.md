---
title: Folder.AddressBookName Property (Outlook)
keywords: vbaol11.chm2004
f1_keywords:
- vbaol11.chm2004
ms.prod: outlook
api_name:
- Outlook.Folder.AddressBookName
ms.assetid: e80535e9-216f-03a6-36a1-3776b5862e96
ms.date: 06/08/2017
---


# Folder.AddressBookName Property (Outlook)

Returns or sets a  **String** that indicates the Address Book name for the **[Folder](folder-object-outlook.md)** object representing a Contacts folder. Read/write.


## Syntax

 _expression_ . **AddressBookName**

 _expression_ A variable that represents a **Folder** object.


## Remarks

If you try to set the  **AddressBookName** property in a non-Contacts folder, an error will be returned.


## Example

The following example changes the Address Book name for the default Contacts folder and displays the new name to the user. The subroutine accepts the folder object and a string representing the new address book name.


```vb
Sub BookName() 
 
 Dim nmsName As Outlook.NameSpace 
 
 Dim fldFolder As Outlook.Folder 
 
 Dim strAns As String 
 
 
 
 'Create a reference to namepsace 
 
 Set nmsName = Application.GetNamespace("MAPI") 
 
 'Create an instance of the Contacts folder 
 
 Set fldFolder = nmsName.GetDefaultFolder(olFolderContacts) 
 
 'Prompt user for input 
 
 strAns = InputBox("Type the name of the new address book") 
 
 'Call Sub procedure 
 
 Call Changebook(fldFolder, strAns) 
 
End Sub 
 
 
 
Sub Changebook(ByRef fldFolder As Folder, ByVal strName As String) 
 
 'Changes the name of the address book for a given folder 
 
 'Set address book name to user input 
 
 fldFolder.AddressBookName = strName 
 
 'Display message to user 
 
 MsgBox ("The new address book name for the " &; fldFolder.Name &; " folder is " _ 
 
 &; strName &; ".") 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

