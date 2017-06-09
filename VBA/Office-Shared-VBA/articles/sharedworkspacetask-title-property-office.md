---
title: SharedWorkspaceTask.Title Property (Office)
keywords: vbaof11.chm264001
f1_keywords:
- vbaof11.chm264001
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.Title
ms.assetid: 038d24fe-5afa-c61d-16e7-7a8c8fca2ccf
ms.date: 06/08/2017
---


# SharedWorkspaceTask.Title Property (Office)

Sets or gets the title of a  **SharedWorkspaceTask** object. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **Title**

 _expression_ A variable that represents a **SharedWorkspaceTask** object.


### Return Value

String


## Remarks

The  **Title** property is the single required property of a shared workspace task. Use the optional **Description** property to provide or return additional information about the task.


## Example

The following example displays a list of the titles of all tasks in the current shared workspace.


```
 Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTasks As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        strTasks = strTasks &amp; swsTask.Title &amp; vbCrLf 
    Next 
    MsgBox strTasks, vbInformation + vbOKOnly, _ 
        "Tasks in Shared Workspace" 
    Set swsTask = Nothing 
 

```


## See also


#### Concepts


[SharedWorkspaceTask Object](sharedworkspacetask-object-office.md)
#### Other resources


[SharedWorkspaceTask Object Members](sharedworkspacetask-members-office.md)

