---
title: Group2.Index Property (Project)
ms.prod: project-server
api_name:
- Project.Group2.Index
ms.assetid: a7d4ec3e-825b-87c8-d7bb-a61984ba7ace
ms.date: 06/08/2017
---


# Group2.Index Property (Project)

Gets the index of a  **Group2** object in a **ResourceGroups2** collection or **TaskGroups2** collection. Read-only **Long**.


## Syntax

 _expression_. **Index**

 _expression_ An expression that returns a **Group2** object.


## Example

The following example displays the name of each  **Group2** object in the **TaskGroups2** collection in the **Immediate** window.


```vb
Sub ListTaskGroups() 

 Dim groupIndex As Integer 

 Dim numTaskGroups As Integer 

 

 numTaskGroups = ActiveProject.TaskGroups2.Count 

 

 For groupIndex = 1 To numTaskGroups 

 Debug.Print ActiveProject.TaskGroups2(groupIndex).Name 

 Next groupIndex 

End Sub
```


## See also


#### Concepts


[Group2 Object](group2-object-project.md)

