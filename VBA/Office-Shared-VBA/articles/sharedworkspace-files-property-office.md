---
title: SharedWorkspace.Files Property (Office)
keywords: vbaof11.chm276004
f1_keywords:
- vbaof11.chm276004
ms.prod: office
api_name:
- Office.SharedWorkspace.Files
ms.assetid: e4a2f80e-5cb7-8ff2-3ab7-2b8c2d9d3cfb
ms.date: 06/08/2017
---


# SharedWorkspace.Files Property (Office)

Provides access to the  **SharedWorkspaceFile** objects in the **SharedWorkspace**. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Files**

 _expression_ A variable that represents a **[SharedWorkspace](sharedworkspace-object-office.md)** object.


## Example

The following example lists the files saved in the current shared workspace.


```
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " &amp; swsFiles.Count &amp; _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

