---
title: SharedWorkspace.Tasks Property (Office)
keywords: vbaof11.chm276003
f1_keywords:
- vbaof11.chm276003
ms.prod: office
api_name:
- Office.SharedWorkspace.Tasks
ms.assetid: 9f7fa28d-f442-cbec-de7c-9109cc3e6f2e
ms.date: 06/08/2017
---


# SharedWorkspace.Tasks Property (Office)

Gets a  **[SharedWorkspaceTasks](sharedworkspacetasks-object-office.md)** collection that represents the list of tasks in the current shared workspace. Read-only.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Tasks**

 _expression_ A variable that represents a **SharedWorkspace** object.


## Example

The following example lists the tasks in the current shared workspace.


```
   Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " &amp; swsTasks.Count &amp; _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## See also


#### Concepts


[SharedWorkspace Object](sharedworkspace-object-office.md)
#### Other resources


[SharedWorkspace Object Members](sharedworkspace-members-office.md)

