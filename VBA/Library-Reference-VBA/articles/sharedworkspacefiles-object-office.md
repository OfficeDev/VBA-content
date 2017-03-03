---
title: SharedWorkspaceFiles Object (Office)
keywords: vbaof11.chm267000
f1_keywords:
- vbaof11.chm267000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceFiles
ms.assetid: 5e2937f7-f794-dffb-a1ec-69ea9a9e3546
---


# SharedWorkspaceFiles Object (Office)

A collection of the  **[SharedWorkspaceFile](sharedworkspacefile-object-office.md)** objects in the current shared workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Files](sharedworkspace-files-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceFiles** collection.


```vb
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " &; swsFiles.Count &; _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

