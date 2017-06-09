---
title: Store.GetRootFolder Method (Outlook)
keywords: vbaol11.chm806
f1_keywords:
- vbaol11.chm806
ms.prod: outlook
api_name:
- Outlook.Store.GetRootFolder
ms.assetid: 09da4d57-c33d-6946-cc21-7233e89efb10
ms.date: 06/08/2017
---


# Store.GetRootFolder Method (Outlook)

Returns a  **[Folder](folder-object-outlook.md)** object representing the root-level folder of the **[Store](store-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **GetRootFolder**

 _expression_ A variable that represents a **Store** object.


### Return Value

A  **Folder** object that represents the folder at the root of that **Store** .


## Remarks

You can use the  **GetRootFolder** method to enumerate the subfolders of the root folder of the **Store** . Unlike **[NameSpace.Folders](namespace-folders-property-outlook.md)** which contains all folders for all stores in the current profile, **Store.GetRootFolder.Folders** allows you to enumerate all folders for a given **Store** object in the current profile.

The  **[Parent](folder-parent-property-outlook.md)** property of the root folder of a store returns the string "Mapi".

The root folder for the Exchange Public Folder store is the folder  **Public Folders**. This folder is returned by the call to  `Application.Session.GetDefaultFolder(olPublicFoldersAllPublicFolders)`.

 **GetRootFolder** returns an error if the service provider does not support root folders.


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) starts at the root-level folder of each  **Store** in a **[Stores](stores-object-outlook.md)** collection for a session, and enumerates all folders on all stores for that session.


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


## See also


#### Concepts


[Store Object](store-object-outlook.md)

