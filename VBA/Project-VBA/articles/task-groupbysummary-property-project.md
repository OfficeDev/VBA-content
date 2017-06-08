---
title: Task.GroupBySummary Property (Project)
ms.prod: project-server
api_name:
- Project.Task.GroupBySummary
ms.assetid: c86393b7-e123-b627-0762-475cef921fdf
ms.date: 06/08/2017
---


# Task.GroupBySummary Property (Project)

 **True** if the selected item in a task view is in a group summary row; otherwise, **false**. Read-only **Boolean**.


## Syntax

 _expression_. **GroupBySummary**

 _expression_ A variable that represents a **Task** object.


## Remarks

When you apply a  **Group by** command to a task view, the group summary rows show the group definition in the **Task Name** column. If a selected cell is in a group summary row, the **GroupBySummary** property is **True**.

The  **GroupBySummary** property is accessible through the `ActiveCell.Task` property, not through `ActiveProject.Tasks(x)`.


## Example

The following example applies the Duration grouping to the Gantt Chart view, and then selects the first cell in each row of the view and tests whether the row is a group summary. The process continues until the row is empty, and then shows a message box with the test results for each row.


```vb
Sub ShowGroupByItems() 
 Dim isValid As Boolean 
 Dim tsk As Task 
 Dim rowType As String 
 Dim msg As String 
 
 isValid = True 
 msg = "" 
 
 ActiveProject.Views("Gantt Chart").Apply 
 GroupApply Name:="Duration" 
 Application.SelectBeginning 
 
 ' When a cell in an empty row is selected, accessing the ActiveCell.Task 
 ' property results in error 1004. 
 On Error Resume Next 
 
 ' Loop until a cell in an empty row is selected. 
 While isValid 
 Set tsk = ActiveCell.Task 
 
 If Err.Number > 0 Then 
 isValid = False 
 Debug.Print Err.Number 
 Err.Number = 0 
 Else 
 If tsk.GroupBySummary Then 
 rowType = "' is a group-by summary row." 
 Else 
 rowType = "' is a task row." 
 End If 
 
 msg = msg &; "Task name: '" &; tsk.Name &; rowType &; vbCrLf 
 SelectCellDown 
 End If 
 Wend 
 
 MsgBox msg, vbInformation, "GroupBy Summary for Tasks" 
 
End Sub
```


