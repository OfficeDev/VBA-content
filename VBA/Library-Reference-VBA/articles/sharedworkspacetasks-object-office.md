---
title: SharedWorkspaceTasks Object (Office)
keywords: vbaof11.chm265000
f1_keywords:
- vbaof11.chm265000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SharedWorkspaceTasks
ms.assetid: de26341f-44d1-131e-1dbe-e31f3f68e312
---


# SharedWorkspaceTasks Object (Office)

A collection of the  **[SharedWorkspaceTask](sharedworkspacetask-object-office.md)** objects in the current shared workspace site.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Example

Use the  **[Tasks](sharedworkspace-tasks-property-office.md)** property of the **[SharedWorkspace](sharedworkspace-object-office.md)** object to return a **SharedWorkspaceTasks** collection.


```vb
    Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " &; swsTasks.Count &; _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

