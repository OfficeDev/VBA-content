---
title: Task.FixedCost Property (Project)
keywords: vbapj.chm132231
f1_keywords:
- vbapj.chm132231
ms.prod: project-server
api_name:
- Project.Task.FixedCost
ms.assetid: 09fb9edb-00b6-d084-b0da-0b0fc5463960
ms.date: 06/08/2017
---


# Task.FixedCost Property (Project)

Gets or sets a fixed cost for a task. Read/write  **Variant**.


## Syntax

 _expression_. **FixedCost**

 _expression_ A variable that represents a **Task** object.


## Example

The following example increases the fixed costs of marked tasks by an amount specified by the user.


```vb
Sub IncreaseFixedCosts() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim Entry As String ' Amount to add to any existing fixed cost 
 
 Entry = InputBox$("Increase the fixed costs of marked tasks by what amount?") 
 
 ' If entry is invalid, display error message and exit Sub procedure. 
 If Not IsNumeric(Entry) Then 
 MsgBox ("You didn't enter a numeric value.") 
 Exit Sub 
 End If 
 
 ' Increase the fixed costs of marked tasks by the specified amount. 
 For Each T In ActiveProject.Tasks 
 If T.Marked Then 
 T.FixedCost = T.FixedCost + Val(Entry) 
 End If 
 Next T 
 
End Sub
```


