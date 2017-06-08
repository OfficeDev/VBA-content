---
title: SharedWorkspaceFolders.Add Method (Office)
keywords: vbaof11.chm269003
f1_keywords:
- vbaof11.chm269003
ms.prod: office
api_name:
- Office.SharedWorkspaceFolders.Add
ms.assetid: 5b941034-502b-b2a5-c6b3-aed57bc2a578
ms.date: 06/08/2017
---


# SharedWorkspaceFolders.Add Method (Office)

Adds a folder to the document library in a shared workspace. Returns a  **[SharedWorkspaceFolder](sharedworkspacefolder-object-office.md)** object.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Add**( **_FolderName_**, **_ParentFolder_** )

 _expression_ Required. A variable that represents a **[SharedWorkspaceFolders](sharedworkspacefolders-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FolderName_|Required|**String**|The name of the folder to be added to the current shared workspace.|
| _ParentFolder_|Optional|**SharedWorkspaceFolder**|The subfolder in which to place the new folder, if not the main document library folder within the shared workspace. Add the folder to the main document library folder by leaving this optional argument empty.|

## Example

The following example adds a new folder to the folders collection of the shared workspace.


```
    Dim swsFolder As Office.SharedWorkspaceFolder 
    Set swsFolder = ActiveWorkbook.SharedWorkspace.Folders.Add("MyNewFolder") 
    MsgBox "New folder: " &amp; swsFolder.FolderName, _ 
        vbInformation + vbOKOnly, _ 
        "New Folder in Shared Workspace" 
    Set swsFolder = Nothing 

```


## See also


#### Concepts


[SharedWorkspaceFolders Object](sharedworkspacefolders-object-office.md)
#### Other resources


[SharedWorkspaceFolders Object Members](sharedworkspacefolders-members-office.md)

