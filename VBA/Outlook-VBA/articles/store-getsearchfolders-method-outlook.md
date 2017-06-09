---
title: Store.GetSearchFolders Method (Outlook)
keywords: vbaol11.chm807
f1_keywords:
- vbaol11.chm807
ms.prod: outlook
api_name:
- Outlook.Store.GetSearchFolders
ms.assetid: aed6ba0b-5e20-adb9-6f62-d030a0de2e0b
ms.date: 06/08/2017
---


# Store.GetSearchFolders Method (Outlook)

Returns a  **[Folders](folders-object-outlook.md)** collection object that represents the search folders defined for the **[Store](store-object-outlook.md)** object.


## Syntax

 _expression_ . **GetSearchFolders**

 _expression_ A variable that represents a **Store** object.


### Return Value

A  **Folders** collection object that represents all the search folders for the **Store** object.


## Remarks

 **GetSearchFolders** returns all the visible active search folders for the **Store** . It does not return uninitialized or aged out search folders.

 **GetSearchFolders** returns a **Folders** collection object with **[Folders.Count](folders-count-property-outlook.md)** equal zero (0) if no search folders have been defined for the **Store** .

For a  **Folders** collection object that represents a collection of search folders, **[Folders.Parent](folders-parent-property-outlook.md)** returns the same object as **[Store.GetRootFolder](store-getrootfolder-method-outlook.md)** . **[Folder.Folders](folder-folders-property-outlook.md)** returns **Null** ( **Nothing** in Visual Basic).


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) enumerates the search folders on all stores for the current session.


```vb
Sub EnumerateSearchFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oSearchFolders As Outlook.folders 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oSearchFolders = oStore.GetSearchFolders 
 
 For Each oFolder In oSearchFolders 
 
 Debug.Print (oFolder.FolderPath) 
 
 Next 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[Store Object](store-object-outlook.md)

