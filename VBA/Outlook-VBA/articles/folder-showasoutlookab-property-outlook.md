---
title: Folder.ShowAsOutlookAB Property (Outlook)
keywords: vbaol11.chm2005
f1_keywords:
- vbaol11.chm2005
ms.prod: outlook
api_name:
- Outlook.Folder.ShowAsOutlookAB
ms.assetid: bb74591b-a3ea-efbd-e7b2-f374f1974be8
ms.date: 06/08/2017
---


# Folder.ShowAsOutlookAB Property (Outlook)

Returns or sets a  **Boolean** variable that specifies whether the contact items folder will be displayed as an address list in the Outlook Address Book. Read/write.


## Syntax

 _expression_ . **ShowAsOutlookAB**

 _expression_ A variable that represents a **Folder** object.


## Remarks

If you set the  **ShowAsOutlookAB** property of a contact items folder to **False** , it will not be available in the drop-down list under **Address Book** in the **Address Book** dialog box.

 **ShowAsOutlookAB** does not support folders on another Exchange user's mailbox, for example, a Contacts folder that is shared by another user. Getting or setting this property on a such a folder will not produce any desired results.


## Example

The following Visual Basic for Applications (VBA) example creates a reference to the default Contacts folder and modifies its  **ShowAsOutlookAB** property to be displayed as an Address Book.


```vb
Sub ShowAsAddressBookChange() 
 
 Dim nmsName As Outlook.Namespace 
 
 Dim fldFolder As Outlook.Folder 
 
 
 
 'Create instance of namespace 
 
 Set nmsName = Application.GetNamespace("Mapi") 
 
 Set fldFolder = nmsName.GetDefaultFolder(olFolderContacts) 
 
 'Display the folder as Outlook Address Book 
 
 fldFolder.ShowAsOutlookAB = True 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

