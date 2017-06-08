---
title: Assignment.Overallocated Property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Overallocated
ms.assetid: 739fcdcd-5ef0-754b-8868-ef3e0662a2e2
ms.date: 06/08/2017
---


# Assignment.Overallocated Property (Project)

 **True** if an assignment is overallocated. Read-only **Boolean**.


## Syntax

 _expression_. **Overallocated**

 _expression_ A variable that represents an **Assignment** object.


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


