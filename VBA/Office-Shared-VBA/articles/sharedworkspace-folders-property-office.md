---
title: SharedWorkspace.Folders Property (Office)
keywords: vbaof11.chm276005
f1_keywords:
- vbaof11.chm276005
ms.prod: office
api_name:
- Office.SharedWorkspace.Folders
ms.assetid: aaba6357-fff5-f3d2-e7d7-6453183864e3
ms.date: 06/08/2017
---


# SharedWorkspace.Folders Property (Office)

Gets a  **[SharedWorkspaceFolders](sharedworkspacefolders-object-office.md)** collection that represents the list of subfolders in the document library associated with the current shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Folders**

 _expression_ A variable that represents a **SharedWorkspace** object.


## Remarks

The  **SharedWorkspaceFolders** collection does not include the root document library folder itself, which by default is named "Shared Documents".


## Example

The following example lists the subfolders in the current shared workspace.


```
    Dim swsFolders As Office.SharedWorkspaceFolders 
    Set swsFolders = ActiveWorkbook.SharedWorkspace.Folders 
    MsgBox "There are " &amp; swsFolders.Count &amp; _ 
        " folder(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFolders = Nothing 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

