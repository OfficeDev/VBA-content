---
title: Enumerate Folders on All Stores
ms.prod: outlook
ms.assetid: 9c78ecee-7b9b-bec0-5510-3224cd9aa1fd
ms.date: 06/08/2017
---


# Enumerate Folders on All Stores

This topic shows a code sample that enumerates all folders on all stores for a session.


1. The code sample begins by getting all the stores for the current session using the  **[NameSpace.Stores](namespace-stores-property-outlook.md)** property of the current session, `Application.Session`.
    
2. For each store of this session, it uses  **[Store.GetRootFolder](store-getrootfolder-method-outlook.md)** to obtain the folder at the root of the store.
    
3. For the root folder of each store, it iteratively calls the  `EnumerateFolders` procedure until it has visited and displayed the name of each folder in that tree.
    

## Remarks

To run this code sample, place the code in the built-in  **ThisOutlookSession** module. Run the `EnumerateFoldersInStores` procedure:


```vb
Sub EnumerateFoldersInStores() 
 Dim colStores As Outlook.Stores 
 Dim oStore As Outlook.Store 
 Dim oRoot As Outlook.Folder 
 
 On Error Resume Next 
 Set colStores = Application.Session.Stores 
 For Each oStore In colStores 
 Set oRoot = oStore.GetRootFolder 
 Debug.Print (oRoot.FolderPath) 
 EnumerateFolders oRoot 
 Next 
End Sub 
 
Private Sub EnumerateFolders(ByVal oFolder As Outlook.Folder) 
 Dim folders As Outlook.folders 
 Dim Folder As Outlook.Folder 
 Dim foldercount As Integer 
 
 On Error Resume Next 
 Set folders = oFolder.folders 
 foldercount = folders.Count 
 'Check if there are any folders below oFolder 
 If foldercount Then 
 For Each Folder In folders 
 Debug.Print (Folder.FolderPath) 
 EnumerateFolders Folder 
 Next 
 End If 
End Sub
```


