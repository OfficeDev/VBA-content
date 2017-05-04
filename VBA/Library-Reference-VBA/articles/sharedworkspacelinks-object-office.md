---
title: SharedWorkspaceLinks Object (Office)
keywords: vbaof11.chm271000
f1_keywords:
- vbaof11.chm271000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceLinks
ms.assetid: b226b376-9d8c-659a-9551-6341bbebed6f
---


# SharedWorkspaceLinks Object (Office)

A collection of the  **[SharedWorkspaceLink](sharedworkspacelink-object-office.md)** objects in the current shared workspace.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Links](sharedworkspace-links-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceLinks** collection.


```vb
    Dim swsLinks As Office.SharedWorkspaceLinks 
    Set swsLinks = ActiveWorkbook.SharedWorkspace.Links 
    MsgBox "There are " &; swsLinks.Count &; _ 
        " link(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsLinks = Nothing 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

