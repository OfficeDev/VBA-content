---
title: Folders Object (Outlook)
keywords: vbaol11.chm2997
f1_keywords:
- vbaol11.chm2997
ms.prod: outlook
api_name:
- Outlook.Folders
ms.assetid: 0c814c3c-74fc-414c-982d-a0097fcb35c2
ms.date: 06/08/2017
---


# Folders Object (Outlook)

Contains a set of  **[Folder](folder-object-outlook.md)** objects that represent all the available Outlook folders in a specific subset at one level of the folder tree.


## Remarks

Use the  **[Folders](namespace-folders-property-outlook.md)** property to return the **Folders** object from a **[NameSpace](namespace-object-outlook.md)** object or another **Folder** object.

Use  **Folders** ( _index_ ), where _index_ is the name or index number, to return a single **Folder** object. Folder names are case-sensitive.


## Example

The following Visual Basic for Applications (VBA) example returns the folder named Old Contacts.


```vb
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderContacts) 
 
Set myNewFolder = myFolder.Folders("Old Contacts")
```

The following Visual Basic for Applications example returns the first folder.






```vb
Set myNewFolder = myFolder.Folders(1)
```


## See also


#### Other resources



[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)

