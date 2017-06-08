---
title: SharedWorkspaceTask.DueDate Property (Office)
keywords: vbaof11.chm264006
f1_keywords:
- vbaof11.chm264006
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.DueDate
ms.assetid: 86ef146e-7528-9dfb-646f-8412abade012
ms.date: 06/08/2017
---


# SharedWorkspaceTask.DueDate Property (Office)

Gets or sets the optional due date and time of a  **SharedWorkspaceTask** object. Read/write.


 **Note**  Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. **DueDate**()

 _expression_ An expression that returns a **SharedWorkspaceTask** object.


## Example

The following example sets the DueDate of all tasks in a shared workspace to 12:00 noon on December 31, 2005 and uploads these changes to the server using the  **Save** method.


```
Dim swsTask As Office.SharedWorkspaceTask 
    Const dtmNewDueDate As Date = #12/31/2005 12:00:00 PM# 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        swsTask.DueDate = dtmNewDueDate 
        swsTask.Save 
    Next 
    Set swsTask = Nothing
```


## See also


#### Concepts


[SharedWorkspaceTask Object](sharedworkspacetask-object-office.md)
#### Other resources


[SharedWorkspaceTask Object Members](sharedworkspacetask-members-office.md)

