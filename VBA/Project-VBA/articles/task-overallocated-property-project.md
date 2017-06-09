---
title: Task.Overallocated Property (Project)
ms.prod: project-server
api_name:
- Project.Task.Overallocated
ms.assetid: bf030017-2774-939b-e0dd-70d66fb3dfa3
ms.date: 06/08/2017
---


# Task.Overallocated Property (Project)

 **True** if any of the assignments for a task is overallocated. Read-only **Boolean**.


## Syntax

 _expression_. **Overallocated**

 _expression_ A variable that represents a **Task** object.


## Example

The following example displays the percentage of resources in the active project that are overallocated.


```vb
Sub DisplayOverallocatedPercentage() 
 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim NOverallocated As Long ' Number of overallocated resources 
 
 For Each R In ActiveProject.Resources 
 If R.Overallocated Then NOverallocated = NOverallocated + 1 
 Next R 
 
 MsgBox (Str$((NOverallocated / ActiveProject.Resources.Count) * 100) _ 
 &; " percent (" &; Str$(NOverallocated) &; "/" &; Str$(ActiveProject.Resources.Count) _ 
 &; ")" &; " of the resources in this project are overallocated.") 
 
End Sub
```


