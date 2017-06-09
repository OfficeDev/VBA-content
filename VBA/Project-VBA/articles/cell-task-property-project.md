---
title: Cell.Task Property (Project)
ms.prod: project-server
api_name:
- Project.Cell.Task
ms.assetid: ba23b56f-e817-1ea3-bed6-b83342c2bded
ms.date: 06/08/2017
---


# Cell.Task Property (Project)

Gets a  **[Task](task-object-project.md)** object representing the task in the active cell. Read-only **Task**.


## Syntax

 _expression_. **Task**

 _expression_ A variable that represents a **Cell** object.


## Example

The following example displays the names of all tasks at the same outline level as the selected task.


```vb
Sub Siblings() 
 
 Dim MyParent As Task 
 Dim Sibling As Task 
 Dim Temp As String 
 
 Set MyParent = ActiveCell.Task.OutlineParent 
 
 For Each Sibling In MyParent.OutlineChildren 
 Temp = Sibling.Name &; ListSeparator &; " " &; Temp 
 Next Sibling 
 
 Temp = Left$(Temp, Len(Temp) - Len(ListSeparator &; " ")) 
 MsgBox Temp 
 
End Sub
```


