---
title: Resource.GroupBySummary Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.GroupBySummary
ms.assetid: 75bf2466-fa8f-8930-4b75-36198d9a6f4c
ms.date: 06/08/2017
---


# Resource.GroupBySummary Property (Project)

 **True** if the selected item in a resource view is in a group summary row; otherwise, **false**. Read-only **Boolean**.


## Syntax

 _expression_. **GroupBySummary**

 _expression_ A variable that represents a **Resource** object.


## Remarks

When you apply a  **Group by** command to a resource view, the group summary rows show the group definition in the **Resource Name** column. If a selected cell is in a group summary row, the **GroupBySummary** property is **True**.

The  **GroupBySummary** property is accessible through the `ActiveCell.Resource` property, not through `ActiveProject.Resources(x)`.


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


