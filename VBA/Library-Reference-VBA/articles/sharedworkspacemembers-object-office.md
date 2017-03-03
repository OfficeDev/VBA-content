---
title: SharedWorkspaceMembers Object (Office)
keywords: vbaof11.chm273000
f1_keywords:
- vbaof11.chm273000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceMembers
ms.assetid: 2d0e6ce0-79ef-3030-b1af-465428314b15
---


# SharedWorkspaceMembers Object (Office)

A collection of the  **[SharedWorkspaceMember](sharedworkspacemember-object-office.md)** objects in the current shared workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Members](sharedworkspace-members-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceMembers** collection.


```vb
    Dim swsMembers As Office.SharedWorkspaceMembers 
    Set swsMembers = ActiveWorkbook.SharedWorkspace.Members 
    MsgBox "There are " &; swsMembers.Count &; _ 
        " member(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsMembers = Nothing 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

