---
title: Folders.Item Method (Outlook)
keywords: vbaol11.chm44
f1_keywords:
- vbaol11.chm44
ms.prod: outlook
api_name:
- Outlook.Folders.Item
ms.assetid: 96a462c2-fa55-62dc-48a4-6464966b84ce
ms.date: 06/08/2017
---


# Folders.Item Method (Outlook)

Returns a  **[Folder](folder-object-outlook.md)** object from the collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **Folders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either the index number of the object, or a value used to match the default property of an object in the collection.|

### Return Value

A  **Folder** object that represents the specified object.


## Example

The following example adds the public folder  **Internal** to the user's **Favorites** folder by using the **AddToPFFavorites** method.


```vb
Sub AddToFavorites() 
 
 'Adds a Public Folder to the List of favorites 
 
 Dim objFolder As Outlook.Folder 
 
 Set objFolder = Application.Session.GetDefaultFolder(olPublicFoldersAllPublicFolders).Folders.Item("GroupDiscussion").Folders.Item("Standards").Folders.Item("Internal") 
 
 objFolder.AddToPFFavorites 
 
End Sub
```


## See also


#### Concepts


[Folders Object](folders-object-outlook.md)

