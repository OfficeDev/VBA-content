---
title: Task.ActualOvertimeWork Property (Project)
ms.prod: project-server
api_name:
- Project.Task.ActualOvertimeWork
ms.assetid: bbd2c42a-f6bb-1e0f-7e23-a76f78fe3a2e
ms.date: 06/08/2017
---


# Task.ActualOvertimeWork Property (Project)

Gets the actual overtime work (in minutes) for a task. Read-only  **Variant**.


## Syntax

 _expression_. **ActualOvertimeWork**

 _expression_ A variable that represents a **Task** object.


## Example

The following example shows the cost of overtime by calculating the total cost of tasks with overtime work, as well as breaking down the individual costs per task.


```vb
Sub PriceOfOvertime() 
 Dim T As Task 
 Dim Price As Variant 
 Dim Breakdown As String 
 
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


