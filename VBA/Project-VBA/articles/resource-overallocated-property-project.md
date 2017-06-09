---
title: Resource.Overallocated Property (Project)
ms.prod: project-server
api_name:
- Project.Resource.Overallocated
ms.assetid: 4cb06be7-0140-1bd0-3314-2a6b50d5a51b
ms.date: 06/08/2017
---


# Resource.Overallocated Property (Project)

 **True** if a resource is overallocated. Read-only **Boolean**.


## Syntax

 _expression_. **Overallocated**

 _expression_ A variable that represents a **Resource** object.


## Remarks

The  **Overallocated** property does not return any meaningful information for material resources.


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


