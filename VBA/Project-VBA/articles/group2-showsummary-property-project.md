---
title: Group2.ShowSummary Property (Project)
ms.prod: project-server
api_name:
- Project.Group2.ShowSummary
ms.assetid: 8cc3401e-ade3-c561-d561-e98a79e7bb22
ms.date: 06/08/2017
---


# Group2.ShowSummary Property (Project)

 **True** if summary tasks are displayed in a task view that is organized by group. Read/write **Boolean**.


## Syntax

 _expression_. **ShowSummary**

 _expression_ An expression that returns a **Group2** object.


## Example

The following example displays the name of the second  **Group2** object in the **TaskGroups2** collection, and then displays the setting for the **ShowSummary** property in the **Immediate** window.


```vb
Debug.Print ActiveProject.TaskGroups2(2).Name 

Debug.Print activeproject.TaskGroups2(2).ShowSummary
```


## See also


#### Concepts


[Group2 Object](group2-object-project.md)

