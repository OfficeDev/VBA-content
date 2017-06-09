---
title: Application.SelectCellDown Method (Project)
keywords: vbapj.chm2050
f1_keywords:
- vbapj.chm2050
ms.prod: project-server
api_name:
- Project.Application.SelectCellDown
ms.assetid: 78754f19-651b-d614-fa69-5fccd6b3387c
ms.date: 06/08/2017
---


# Application.SelectCellDown Method (Project)

Selects cells directly below the current selection.


## Syntax

 _expression_. **SelectCellDown**( ** _NumCells_**, ** _Extend_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NumCells_|Optional|**Long**|The number of cells to select downward from the current selection. The default value is 1.|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the specified cell. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **SelectCellDown** method is not available when the Calendar, Network Diagram, or Resource Graph is the active view.


## Example

The following example applies the Complete and Incomplete Resources grouping to the Resource Sheet view, and then uses  **SelectCellDown** to select the first cell in each row and tests whether the row is a group summary. The process continues until the row is empty, and then shows a message box with the test results for each row.


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


