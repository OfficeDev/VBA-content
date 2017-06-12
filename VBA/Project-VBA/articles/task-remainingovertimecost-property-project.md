---
title: Task.RemainingOvertimeCost Property (Project)
ms.prod: project-server
api_name:
- Project.Task.RemainingOvertimeCost
ms.assetid: 6e8d72fd-efac-ed22-9549-950bba1cfc84
ms.date: 06/08/2017
---


# Task.RemainingOvertimeCost Property (Project)

Gets the remaining overtime cost for the task. Read-only  **Variant**.


## Syntax

 _expression_. **RemainingOvertimeCost**

 _expression_ A variable that represents a **Task** object.


## Example

The following example returns the remaining overtime cost of each task in the active project.


```vb
Sub ReturnOvertimeCost() 
 Dim T As Task ' Task object used in For Each loop 
 Dim Results As String 
 
 For Each T In ActiveProject.Tasks 
 Results = Results &; T.Name &; ": " &; ActiveProject.CurrencySymbol &; _ 
 T.RemainingOvertimeCost &; ListSeparator &; " " 
 Next T 
 
 Results = Left$(Results, Len(Results) - Len(ListSeparator &; " ")) 
 
 MsgBox Results 
 
End Sub
```


