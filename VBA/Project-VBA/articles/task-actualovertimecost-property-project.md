---
title: Task.ActualOvertimeCost Property (Project)
ms.prod: project-server
api_name:
- Project.Task.ActualOvertimeCost
ms.assetid: 7e3b409e-3249-4fe1-b5a1-1b65646519b3
ms.date: 06/08/2017
---


# Task.ActualOvertimeCost Property (Project)

Gets the actual overtime cost for a task. Read-only  **Variant**.


## Syntax

 _expression_. **ActualOvertimeCost**

 _expression_ A variable that represents a **Task** object.


## Example

The following example shows the cost of overtime by calculating the total cost of tasks with overtime work, as well as breaking down the individual costs per task.


```vb
Sub PriceOfOvertime() 
 Dim T As Task 
 Dim Price As Variant, Breakdown As String 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.ActualOvertimeWork <> 0 Then 
 Price = Price + T.ActualOvertimeCost 
 Breakdown = Breakdown &; T.Name &; ": " &; _ 
 ActiveProject.CurrencySymbol &; _ 
 T.ActualOvertimeCost &; vbCrLf 
 End If 
 End If 
 Next T 
 
 If Breakdown <> "" Then 
 MsgBox Breakdown &; vbCrLf &; "Total: " &; _ 
 ActiveProject.CurrencySymbol &; Price 
 End If 
 
End Sub
```


