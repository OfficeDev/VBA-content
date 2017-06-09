---
title: Application.OpenUndoTransaction Method (Project)
ms.prod: project-server
api_name:
- Project.Application.OpenUndoTransaction
ms.assetid: b94b2c87-786c-46d6-50d3-d20614493f8f
ms.date: 06/08/2017
---


# Application.OpenUndoTransaction Method (Project)

Create an undo transaction set for a series of operations.


## Syntax

 _expression_. **OpenUndoTransaction**( ** _Label_**, ** _guid_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Label_|Required|**String**|Name that appears in the drop-down list for the  **Undo Transaction** command.|
| _guid_|Optional|**Variant**|GUID that uniquely identifies Label.|

## Remarks

The  **OpenUndoTransaction** method is used in conjunction with **[CloseUndoTransaction](application-closeundotransaction-method-project.md)** method. You can use **OpenUndoTransaction** and **CloseUndoTransaction** on a single command or on a group of commands.

You cannot nest one undo transaction within another.


## Example

The following example demonstrates using the  **OpenUndoTransaction** method to create an undo transaction set. After you run the macro, the task named **Task outside transaction** shows as the item **Insert Task** in the **Undo** drop-down list on the **Quick Access Toolbar**. The six tasks named  **UndoMe 1** to **UndoMe 6** show as **Create 6 tasks** in the **Undo** list.


```vb
Sub CreateTasksWithUndoTransaction() 
    ActiveProject.Tasks.Add "Task outside transaction" 
    Application.OpenUndoTransaction "Create 6 tasks" 
    Dim i As Integer 
    For i = 1 To 6 
        ActiveProject.Tasks.Add "UndoMe " &; i 
    Next 
    Application.CloseUndoTransaction  
End Sub
```


