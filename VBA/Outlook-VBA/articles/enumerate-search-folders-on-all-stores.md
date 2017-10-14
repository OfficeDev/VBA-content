---
title: Enumerate Search Folders on All Stores
ms.prod: outlook
ms.assetid: 513b0a63-1c0f-480c-214d-7a30be137875
ms.date: 06/08/2017
---


# Enumerate Search Folders on All Stores

This topic describes a code sample that enumerates the search folders on all stores for the current session.


1. The code sample begins by getting all the stores for the current session using the  **[NameSpace.Stores](namespace-stores-property-outlook.md)** property of the current session, `Application.Session`.
    
2. For each store of this session, it uses  **[Store.GetSearchFolders](store-getsearchfolders-method-outlook.md)** to obtain the collection of search folders for that store.
    
3. For each collection of search folders, it displays the name of each folder.
    

## Remarks

To run this code sample, place the code in the built-in  **ThisOutlookSession** module. Run the `EnumerateSearchFoldersInStores` procedure:


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


