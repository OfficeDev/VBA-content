---
title: Cell.Resource Property (Project)
ms.prod: project-server
api_name:
- Project.Cell.Resource
ms.assetid: 17514412-363a-dd2d-f0b5-97b8fb5d41cc
ms.date: 06/08/2017
---


# Cell.Resource Property (Project)

Gets a  **[Resource](resource-object-project.md)** object representing the resource in the active cell. Read-only **Resource**.


## Syntax

 _expression_. **Resource**

 _expression_ A variable that represents a **Cell** object.


## Example

The following example applies the Complete and Incomplete Resources grouping to the Resource Sheet view, and then selects the first cell in each row of the view and tests whether the row is a group summary. The process continues until the row is empty, and then shows a message box with the test results for each row.


```vb
Sub ShowGroupByItems() 
 Dim isValid As Boolean 
 Dim res As Resource 
 Dim rowType As String 
 Dim msg As String 
 
 isValid = True 
 msg = "" 
 
 ActiveProject.Views("Resource Sheet").Apply 
 GroupApply Name:="Complete and Incomplete Resources" 
 Application.SelectBeginning 
 
 ' When a cell in an empty row is selected, accessing the ActiveCell.Resource 
 ' property results in error 1004. 
 On Error Resume Next 
 
 ' Loop until a cell in an empty row is selected. 
 While isValid 
 Set res = ActiveCell.Resource 
 
 If Err.Number > 0 Then 
 isValid = False 
 Debug.Print Err.Number 
 Err.Number = 0 
 Else 
 If res.GroupBySummary Then 
 rowType = "' is a group-by summary row." 
 Else 
 rowType = "' is a resource row." 
 End If 
 
 msg = msg &; "Resource name: '" &; res.Name &; rowType &; vbCrLf 
 SelectCellDown 
 End If 
 Wend 
 
 MsgBox msg, vbInformation, "GroupBy Summary for Resources" 
 
End Sub
```


