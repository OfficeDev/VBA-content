---
title: Folder.AddToPFFavorites Method (Outlook)
keywords: vbaol11.chm2003
f1_keywords:
- vbaol11.chm2003
ms.prod: outlook
api_name:
- Outlook.Folder.AddToPFFavorites
ms.assetid: d3926957-bf6d-ad4d-9c24-bfc5037ba9fd
ms.date: 06/08/2017
---


# Folder.AddToPFFavorites Method (Outlook)

Adds a Microsoft Exchange public folder to the public folder's Favorites folder.


## Syntax

 _expression_ . **AddToPFFavorites**

 _expression_ A variable that represents a **Folder** object.


## Example

The following Visual Basic for Applications (VBA) example adds the public folder GroupDiscussion to the user's Favorites folder by using the  **AddToPFFavorites** method. To run this example, you need to replace 'GroupDiscussion' with a valid public folder name.


```vb
Sub AddToFavorites() 
 
 'Adds a Public Folder to the list of favorites 
 
 Dim objFolder As Outlook.Folder 
 
 Set objFolder = Application.Session.GetDefaultFolder _ 
 
 (olPublicFoldersAllPublicFolders).Folders.Item("GroupDiscussion") 
 
 objFolder.AddToPFFavorites 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

